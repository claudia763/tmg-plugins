# No-root toolchain setup (email-cowork server)

How to install/repair the agent toolchain on a host where you have **no sudo** — covers the Python deps, npm globals, and the Chromium shared-library problem that breaks Playwright rendering. Read this if a job fails with `ModuleNotFoundError`, `externally-managed-environment`, or `error while loading shared libraries`.

This file covers the *document-production* toolchain only. For the Dropbox client — installing it, supervising it with `systemd --user`, and getting it linked so `DROPBOX_DIR` can be switched on — see `dropbox-headless-linux-setup.md`. (Useful cross-finding from that work: on this host `loginctl enable-linger claudia` succeeds **without root**, which makes `systemctl --user` services viable for anything that must survive a reboot.)

First established 2026-08-06 on `ubuntu-8gb-nbg1-2` (Ubuntu resolute, Python 3.14.4, Node v22.23.2).

## Check before you fix

Run the health check — it tests function, not just presence, so it catches the "imports fine but the binary won't start" case:

```bash
python3 /path/to/library-additions/scripts/toolchain_healthcheck.py
```

Exit 0 and `22/22 checks passed` means the box is job-ready.

## 1. Python packages — pip refuses to install

Ubuntu marks the system Python as externally managed (PEP 668), so a plain `pip install` fails. Install into the **user site directory**, which `python3` picks up automatically:

```bash
pip install --user --break-system-packages pymupdf python-docx pdfplumber rapidfuzz
```

`--user` keeps it in `~/.local/lib/python3.X/site-packages` and never touches system packages, so `--break-system-packages` here is only silencing the PEP 668 gate, not actually overwriting distro files. Prefer this over a venv: the library skills invoke bare `python3`, so a venv would need activating in every subprocess.

The full dependency set the brokerage skills need: `openpyxl pandas numpy pymupdf pdfplumber python-docx pillow rapidfuzz`.

## 2. npm globals — permission denied

`npm install -g` wants to write to `/usr/lib` and fails without root. Point npm at a user-owned prefix:

```bash
npm config set prefix "$HOME/.npm-global"
npm install -g docx playwright
```

Then make the modules resolvable **without** relying on `NODE_PATH` (see the warning in §4) by symlinking into `~/.node_modules`, one of Node's built-in global folders:

```bash
mkdir -p "$HOME/.node_modules"
ln -sfn "$HOME/.npm-global/lib/node_modules/playwright" "$HOME/.node_modules/playwright"
ln -sfn "$HOME/.npm-global/lib/node_modules/docx"       "$HOME/.node_modules/docx"
```

This matters because library scripts such as `shot_map.cjs` do a bare `require('playwright')`.

## 3. Chromium won't launch — the real trap

`npx playwright install chromium` downloads the browser fine, but the binary dies immediately:

```
error while loading shared libraries: libatk-1.0.so.0: cannot open shared object file
```

`playwright install-deps` fixes this normally — it runs `apt-get install` and needs root. Without root, fetch the `.deb`s (downloading needs no privileges) and unpack them into a user directory:

```bash
LIBDIR="$HOME/.local/chromium-deps"; mkdir -p "$LIBDIR/debs" "$LIBDIR/root"; cd "$LIBDIR/debs"
apt-get download libasound2t64 libatk1.0-0t64 libatk-bridge2.0-0t64 libatspi2.0-0t64 \
                 libgbm1 libpango-1.0-0 libxcomposite1 libxdamage1 libxfixes3 \
                 libthai0 libxi6 libxres1 libdatrie1 patchelf
for d in *.deb; do dpkg-deb -x "$d" "$LIBDIR/root"; done
```

That package list is the transitive closure found by iterating `ldd <chrome> | grep 'not found'`, re-running after each round. If a future Chromium build needs more, repeat that loop rather than guessing.

### Make it stick without environment variables

`LD_LIBRARY_PATH` works but is **fragile** — see §4. The durable fix is to copy the libraries next to each Chromium binary and give the binary an `$ORIGIN` RUNPATH, so the dynamic linker finds them unconditionally:

```bash
PE="$LIBDIR/root/usr/bin/patchelf"; SRC="$LIBDIR/root/usr/lib/x86_64-linux-gnu"
for d in "$HOME/.cache/ms-playwright/chromium-"*/chrome-linux64 \
         "$HOME/.cache/ms-playwright/chromium_headless_shell-"*/chrome-headless-shell-linux64; do
  mkdir -p "$d/tmgdeps"; cp -n "$SRC"/*.so* "$d/tmgdeps/"
  for so in "$d"/tmgdeps/*.so*; do "$PE" --set-rpath '$ORIGIN' "$so"; done
done
"$PE" --set-rpath '$ORIGIN:$ORIGIN/tmgdeps' "$HOME/.cache/ms-playwright/chromium-"*/chrome-linux64/chrome
"$PE" --set-rpath '$ORIGIN:$ORIGIN/tmgdeps' "$HOME/.cache/ms-playwright/chromium_headless_shell-"*/chrome-headless-shell-linux64/chrome-headless-shell
```

Note the single quotes on `'$ORIGIN'` — it is a literal token interpreted by the loader, not a shell variable.

Verify with the environment stripped, which is the condition that actually reproduces how jobs run:

```bash
env -i HOME="$HOME" PATH=/usr/bin:/bin ldd "$HOME/.cache/ms-playwright/chromium-"*/chrome-linux64/chrome | grep 'not found'
```

Silence means success.

**Re-run this section after every `playwright install`** — installing a new browser build drops in unpatched binaries and the `tmgdeps` folder/RUNPATH will be gone.

## 4. Why not just export LD_LIBRARY_PATH in .bashrc

Because it silently does not apply. Ubuntu's `~/.bashrc` starts with an early return for non-interactive shells:

```bash
case $- in *i*) ;; *) return;; esac
```

Anything appended below that line is skipped whenever a job shells out non-interactively — which is the normal case. Putting exports in `~/.profile` only helps *login* shells, so it fails the same way.

A `~/.tmg-toolchain.env` file sourced from `~/.profile` **and** from the very top of `~/.bashrc` (above the guard) is set up on this host as a convenience, but treat it as belt-and-suspenders only. Anything that must work unconditionally — Chromium's libraries, Node module resolution — should be fixed at the filesystem level per §2 and §3 instead of depending on env vars.

## Known-good versions (2026-08-06)

openpyxl 3.1.5 · pandas 3.0.5 · numpy 2.5.1 · PyMuPDF 1.28.0 · pdfplumber 0.11.10 · python-docx 1.2.0 · Pillow 12.3.0 · rapidfuzz 3.14.5 · Node v22.23.2 · Playwright chromium build 1234
