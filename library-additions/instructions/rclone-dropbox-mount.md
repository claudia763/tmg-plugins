# TMG Dropbox access on the agent server — the rclone FUSE mount (CURRENT)

How the agent server actually reaches TMG Dropbox as of 2026-08-07: an
**rclone FUSE mount**, not the native Dropbox client. Read this first for
anything Dropbox-infrastructure related. It **supersedes**
`dropbox-headless-linux-setup.md`, which describes the abandoned native-client
approach and will actively mislead you if you follow it (see §5).

For *which* files go in *which* category folder, see
`dropbox-deal-folder-filing.md` — that is the job-level convention and is
unaffected by this change.

## 1. Check state — the four facts that matter

```bash
export XDG_RUNTIME_DIR=/run/user/1000
systemctl --user is-active rclone-dropbox.service        # active
grep '^DROPBOX_DIR=' /home/claudia/email-cowork-server/.env
ls -d "/home/claudia/mnt/TMG Dropbox"                    # mount present
/home/claudia/bin/rclone about TMG: --config /home/claudia/.config/rclone/rclone.conf
```

`rclone about` answering with a quota (5 TiB total / ~1.2 TiB used on the TMG
team account) is the real proof of authenticated access. A mounted-looking
directory is not proof on its own — an empty or stale mount still `ls`es.

## 2. What is set up

- **Remote:** `TMG:` in `/home/claudia/.config/rclone/rclone.conf`, already
  authorized against the TMG team Dropbox. No human re-auth needed.
- **Mount:** `/home/claudia/mnt/TMG Dropbox`, from user unit
  `rclone-dropbox.service` — **enabled**, and `loginctl` linger is on for
  `claudia`, so it comes back after reboot.
- **Mount flags in use:** `--vfs-cache-mode full --vfs-cache-max-size 2G
  --vfs-cache-max-age 24h --dir-cache-time 5m --poll-interval 1m`.
  On-demand/online-only: the 1.2 TiB team folder is **not** stored locally, so
  it fits on this 75 GB host.
- **Mirror wiring:** `DROPBOX_DIR` in the mail server `.env` is set to
  `/home/claudia/mnt/TMG Dropbox/- Underwritings/- Claudia Outputs`, the Linux
  equivalent of Dmytro's Windows path
  `C:\Users\dmytr\TMG Dropbox\- Underwritings\- Claudia Outputs`.

## 3. Working with the mount — the traps

- **`find`, `du -sh` and `ls -R` over the mount will hang your job.** It is a
  network filesystem holding 1.2 TiB; a recursive walk fetches metadata for
  every object. Always scope to one directory, and wrap mount reads in
  `timeout`. To inspect deeply, query the remote instead of the mount:
  `rclone lsl "TMG:/- Underwritings/..."`.
- **Directory timestamps read as `Jan 1 2000` and sizes as `0`.** That is an
  rclone artifact, not corruption, and it means **directory mtimes tell you
  nothing** about when something was filed. Use `rclone lsl` for real file
  times.
- **Verify writes against the remote, not the mount.** With
  `--vfs-cache-mode full`, a fresh write lands in the local cache first and
  `ls` on the mount shows it immediately whether or not it uploaded. The only
  honest confirmation is `rclone lsl` on the `TMG:` remote showing the file
  with the expected byte size.

## 4. Filing a job's files by hand (when the mirror is off)

The mail server reads `.env` **once at start**, so a newly-set `DROPBOX_DIR`
does nothing until it restarts — and that restart is root-owned
(`/etc/systemd/system/claudia.service`) and must not happen mid-job. Until a
human restarts it, jobs must file into Dropbox themselves.

Reproduce exactly what `server.js` would have done:

```
<DROPBOX_DIR>/<Mon-YY>/<YYYY-MM-DD> <sanitized subject> (uid<N>)/
```

- `<Mon-YY>` e.g. `Aug-26`; `<YYYY-MM-DD>` is the date; subject is sanitized by
  `name.replace(/[^\w.\- ()]/g, "_")` truncated to 60 chars — so `Re: foo`
  becomes `Re_ foo`.
- Use the **thread's original subject** (the first message's), not the current
  `Re:` one, so successive hand-filed rounds land in one folder.
- **Expect one duplicate folder the first time the mirror goes live.**
  `server.js` reuses a remembered folder (`thread.dropboxDir` in
  `jobs/threads.json`, written at line ~673) — but it only records that after a
  mirror actually runs. A thread filed by hand while the mirror was off has no
  such record, so the first mirrored round builds a fresh name from the *then*
  current subject, e.g. `... Re_ checking dropbox intergration (uid63)`
  alongside your `... checking dropbox intergration (uid63)`. Don't try to
  pre-seed `threads.json`: the running server keeps that state in memory and
  rewrites the whole file at job end, clobbering the edit. Just merge the two
  folders on the next round and delete the empty one.
- Inside it: the job's `dropbox/` staging tree copied wholesale, raw
  attachments under `- Info for Buyers/Raw (from seller)/`, and the Notes
  `.docx` at the folder root.

## 5. Do NOT trust the old native-client path

The native client (`~/.dropbox-dist`, `dropbox.service`) is installed but
**deliberately disabled**, and its folder `/home/claudia/TMG Dropbox` no longer
exists. Consequences, all confirmed 2026-08-07:

- Starting `dropbox.service` puts it in a **crash loop**: it logs
  `[ALERT: Dropbox Folder Missing]`, exits 1, and systemd restarts it every
  15 s. Leave it disabled. If you start it to investigate, disable it again.
- Its journal says *"This computer was previously linked to
  dmytro.gladchenko@multifamilygrp.com's account"* — historical, and **not**
  evidence that syncing works now.
- **Never recreate `/home/claudia/TMG Dropbox` to "fix" the missing folder.**
  Re-pointing a linked client at an empty directory risks propagating
  deletions to the live company Dropbox. The rclone mount already provides
  access; there is nothing to fix.
- `scripts/finish_dropbox_setup.py` is written against the native client. Run
  today it reports **`BLOCKED: Dropbox is not linked yet`** and prints a
  `cli_link_nonce` URL. Both are wrong — Dropbox *is* reachable via rclone —
  and the nonce is freshly minted by `dropbox.py status` as a side effect, so
  it is not even the daemon's own. **Do not email that link to anyone.** The
  script now refuses to run when the rclone mount is present.

## 5b. Startup ordering — the mirror can silently switch itself off

`server.js` (lines ~87-95) checks at startup that the **parent** of
`DROPBOX_DIR` exists. If it does not, it logs
`WARNING: Dropbox parent folder not found ... mirror disabled`, sets
`CONFIG.dropboxDir = null`, and runs the rest of the day with mirroring off.

`claudia.service` is a root-owned system unit; the rclone mount is a **user**
unit. Nothing orders them, so after a reboot the mail server can start before
the mount is ready and disable the mirror for the whole session — with no
error in any reply email, because a disabled mirror simply appends nothing.

So after any reboot or restart, confirm from the server log that it printed
`Dropbox mirror enabled: ...` rather than the warning. If it warned, the fix is
just to restart `claudia.service` once the mount is up.

**Gotcha when you verify `.env` from inside a job.** A job shell inherits its
environment from the *running* mail server, so `DROPBOX_DIR` is already present
(and empty, if it was empty at server start). `dotenv` never overrides an
existing variable, so this reads back empty and looks like your edit failed:

```bash
node -e "require('dotenv').config(); console.log(process.env.DROPBOX_DIR)"   # ""
```

Clear the inherited value before testing — this is the honest check:

```bash
env -u DROPBOX_DIR node -e "require('dotenv').config(); console.log(process.env.DROPBOX_DIR)"
```

`claudia.service` sets only `HOME` and `PATH`, so a restarted server does read
the file. The stale empty value in your shell is also a handy confirmation of
what the *current* server started with.

## 6. Anti-footgun rule (unchanged, still the important one)

`server.js` appends *"Deal files were also filed in Dropbox: <path>"* to reply
emails whenever `DROPBOX_DIR` is set. Never point it at a directory that is not
genuinely synced. If `rclone-dropbox.service` is ever stopped for good, blank
`DROPBOX_DIR` in the same change — otherwise the server keeps telling clients
their files were filed while writing to a plain local folder.

An honest "Dropbox filing is not yet live" is recoverable; a false confirmation
is not.
