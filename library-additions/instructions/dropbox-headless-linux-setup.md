# Headless Dropbox client on a no-root Linux host

How to get the Dropbox **client itself** installed, running persistently and linked on the agent server — the infrastructure layer, with **no sudo**. Read this if `DROPBOX_DIR` is empty, if `~/Dropbox` does not exist, or if a job needs files to actually reach the company Dropbox. For *which* files go in *which* category folder once syncing works, see `dropbox-deal-folder-filing.md` — that is the job-level filing convention and assumes this setup is already done.

First established 2026-08-06 on `ubuntu-8gb-nbg1-2` (Ubuntu 26.04 LTS "Resolute Raccoon", Python 3.14.4, Node v22.23.2), as unprivileged user `claudia` (uid 1000). `sudo -n true` on this host answers `sudo: I'm sorry claudia. I'm afraid I can't do that` — assume no root, ever.

## Check before you fix

```bash
export XDG_RUNTIME_DIR=/run/user/1000
ls -d "$HOME/.dropbox-dist"                    # installed?
systemctl --user is-active dropbox.service     # running?
ls -d "$HOME"/*Dropbox* "$HOME/Dropbox" 2>/dev/null   # linked + synced?
grep '^DROPBOX_DIR=' /home/claudia/email-cowork-server/.env   # mirror wired up?
```

The last two are the ones that are usually still pending. A running daemon proves nothing on its own — an **unlinked** daemon runs happily forever and syncs nothing.

## 1. Install — no root needed

```bash
curl -sL "https://www.dropbox.com/download?plat=lnx.x86_64" -o /tmp/dropbox-lnx.tar.gz
tar -xzf /tmp/dropbox-lnx.tar.gz -C "$HOME"     # creates ~/.dropbox-dist
```

The URL redirects to the current build; at time of writing it resolved to `dropbox-lnx.x86_64-264.4.3385.tar.gz`, 90,371,562 bytes (87 MB). The daemon binary is `~/.dropbox-dist/dropboxd`.

**It needs no shared-library surgery.** It started on this host with zero missing `.so`s. Do *not* go looking for the `.deb`-extraction and `patchelf` dance from §3 of `no-root-toolchain-setup.md` — that trap is specific to Chromium. Dropbox ships its own bundled libraries; untar and run.

## 2. Run it persistently — `systemd --user` and the linger prerequisite

Out of the box `systemctl --user` fails:

```
Failed to connect to user scope bus via local transport: $DBUS_SESSION_BUS_ADDRESS and $XDG_RUNTIME_DIR not defined
```

`/run/user/1000` did not exist and `loginctl show-user claudia --property=Linger` reported `User ID 1000 is not logged in or lingering`. The useful discovery: **on this host the unprivileged user can enable linger for itself** — polkit permits self-linger, no root required.

```bash
loginctl enable-linger claudia
```

After that `/run/user/1000` exists, `Linger=yes`, and `systemctl --user is-system-running` reports `running`.

**In a non-interactive job shell `XDG_RUNTIME_DIR` is not set**, so every `systemctl --user` / `journalctl --user` call must be prefixed:

```bash
export XDG_RUNTIME_DIR=/run/user/1000
```

This is the same class of trap as §4 of `no-root-toolchain-setup.md`: env vars set for interactive shells do not survive into the non-interactive shells jobs actually run in. Export it inline in the command, do not assume a dotfile did it for you.

The unit installed at `~/.config/systemd/user/dropbox.service`, verbatim:

```
[Unit]
Description=Dropbox headless sync daemon (TMG Claudia agent server)
Documentation=https://www.dropbox.com/install-linux
After=network-online.target
Wants=network-online.target

[Service]
Type=simple
ExecStart=/home/claudia/.dropbox-dist/dropboxd
Restart=on-failure
RestartSec=15
Environment=HOME=/home/claudia
Nice=5

[Install]
WantedBy=default.target
```

Enable and start:

```bash
export XDG_RUNTIME_DIR=/run/user/1000
systemctl --user daemon-reload
systemctl --user enable --now dropbox.service
systemctl --user is-enabled dropbox.service   # enabled
systemctl --user is-active  dropbox.service   # active
```

## 3. Linking — the part an agent cannot finish alone

An unlinked daemon logs, every ~5 seconds:

```
This computer isn't linked to any Dropbox account...
Please visit https://www.dropbox.com/cli_link_nonce?nonce=<hex> to link this device.
```

Read the current one with:

```bash
export XDG_RUNTIME_DIR=/run/user/1000
journalctl --user -u dropbox.service --no-pager | grep -o 'https://www.dropbox.com/cli_link_nonce?nonce=[a-f0-9]*' | tail -1
```

**The nonce rotates on every daemon restart.** This was confirmed by experiment — three separate starts produced three different nonces (`31055bdb…`, `54c383fd…`, `d8ffa8a5…`). Consequences you must respect:

- The URL is valid only while *that* daemon instance keeps running. **Never restart the service after emailing a link** — it silently invalidates it, and the human gets a dead page with no error to explain it.
- Always re-read the nonce from the journal immediately before sending it. Never reuse one from an earlier job, from these notes, or from a previous message in the thread.
- Linking requires a human signed in to the Dropbox account. Do **not** attempt to log in on their behalf, and do not use credentials that arrive by email.

### Read the journal — both "status" commands are traps

Verified 2026-08-06. Do not check link state by shelling out to a status verb:

- **`dropboxd status` is not a status verb.** `dropboxd` only starts the daemon. With one already running it prints `Another instance of Dropbox (9362) is running!` and exits 0 — which reads as success and yields a **false PASS on "is it linked"**. This bug was caught in this job by a check that wrongly reported the account linked.
- **`dropbox.py status` mints a brand-new nonce on every invocation.** Three consecutive calls returned `94786a80…`, `f91713b9…`, `b0092807…` while the daemon's own journal line kept advertising the single stable `d8ffa8a5…`, `MainPID` unchanged and `NRestarts=0`. So "checking" with it hands you a URL that is *not* the one the daemon is advertising, and churns link requests as a side effect.

`journalctl` is read-only and reports the daemon's own stable nonce. Use it, and use the companion script (§4), which already implements this correctly.

Status as of this writing: **not yet linked**; link emailed to dmytro.gladchenko@multifamilygrp.com on 2026-08-06.

## 4. After linking — what still has to happen

- **Find the synced folder, do not assume it.** It appears as `~/Dropbox`, or for a team account as `~/<Team name> Dropbox`. Glob for it.
- **Check disk first.** This host has 75 GB total / 67 GB free. If the team folder is larger, use selective sync so only the needed folders come down. The official CLI helper is already installed at `~/bin/dropbox.py` (fetched in this job from `https://www.dropbox.com/download?dl=packages/dropbox.py`, 60 KB, Python 3):

  ```bash
  python3 ~/bin/dropbox.py help          # list subcommands
  python3 ~/bin/dropbox.py exclude list  # folders currently NOT synced
  ```

  Consult `dropbox.py help` for the exact selective-sync invocation rather than guessing a flag. Remember from §3 that `dropbox.py status` mints a nonce — it is fine once linked, but never use it as the unlinked-state check.
- **Wire up the mail server's mirror setting.** `DROPBOX_DIR` in `/home/claudia/email-cowork-server/.env` is currently **empty**, which disables mirroring. Its shipped default is a Windows path (`C:\Users\dmytr\TMG Dropbox\- Underwritings\- Claudia Outputs`) left over from Dmytro's PC. `server.js` enables the mirror only if the **parent** of `DROPBOX_DIR` already exists.
- Use the companion script rather than hand-editing `.env`:

  ```bash
  python3 /home/claudia/email-cowork-server/resources/library-additions/scripts/finish_dropbox_setup.py
  ```

- **A restart is required, and it is not yours to do mid-job.** `.env` is read once at process start (`import "dotenv/config"`), so the mail server must be restarted to pick up a new `DROPBOX_DIR`. It runs as the root-owned system unit `/etc/systemd/system/claudia.service` (`User=claudia`), so that restart needs root — and it must not be done during a job, since it would kill the running job, including your own. Set the value, then hand the restart to a human.

## 5. Anti-footgun rule

**Never point `DROPBOX_DIR` at an arbitrary local directory just to make the mirror switch on.** `server.js` appends `Deal files were also filed in Dropbox: <path>` to the reply email whenever the mirror is enabled. If that path is not genuinely a synced Dropbox folder, the agent has told a client their files were filed in Dropbox when nothing was synced anywhere.

Leave it disabled and visibly pending instead. An honest "Dropbox filing is not yet live" is recoverable; a false confirmation is not.

## Known-good versions (2026-08-06)

Dropbox client 264.4.3385 · Ubuntu 26.04 · Python 3.14.4 · Node v22.23.2 · systemd user service with linger enabled
