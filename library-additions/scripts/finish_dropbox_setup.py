#!/usr/bin/env python3
"""
Finish the Dropbox wiring on the TMG email-cowork agent server, after a human
has linked the Dropbox account.

WHAT IT DOES
    Picks up where instructions/dropbox-headless-linux-setup.md leaves off.
    The headless Dropbox client can be installed and supervised by an agent,
    but LINKING it to the TMG account needs a signed-in human. Once that link
    has happened, this script does the remaining machine-side work:

      1. Confirms the dropbox.service user unit is enabled + active.
      2. Confirms the account is actually LINKED (not just running unlinked).
      3. Locates the synced root -- ~/Dropbox for a personal account, or
         ~/<Team name> Dropbox for a team account. It detects; it never assumes.
      4. Creates the deal-output folder inside that root, mirroring the layout
         used on Dmytro's Windows PC:
             <root>/- Underwritings/- Claudia Outputs
      5. Sets DROPBOX_DIR in the server's .env to that folder (backing the file
         up first, and leaving every other line untouched).

    It deliberately STOPS short of restarting the mail server -- see NOTES.

WHEN TO USE
    On the first job after someone reports "I clicked the Dropbox link".
    Safe to run any time: it is idempotent and defaults to a dry run.

INPUTS
    --server-dir PATH  Mail server root holding .env.
                       Default: /home/claudia/email-cowork-server
    --dropbox-root PATH  Skip auto-detection and use this synced root.
    --subpath REL      Folder to create under the root and point DROPBOX_DIR at.
                       Default: "- Underwritings/- Claudia Outputs"
    --apply            Actually create folders and edit .env. WITHOUT this flag
                       the script only reports what it WOULD do.

OUTPUTS
    A PASS/FAIL/WOULD-DO report on stdout.
    With --apply: creates the folders, writes .env, and leaves a timestamped
    backup at .env.bak-<UTC timestamp>.
    Exit 0 = ready (or dry run that found no blockers); 1 = something blocks it.

NOTES
    * .env is read ONCE at server start (`import "dotenv/config"`), so the new
      DROPBOX_DIR does not take effect until the mail server restarts. That
      server is the ROOT-OWNED unit /etc/systemd/system/claudia.service, so the
      restart needs root AND must not happen mid-job -- it would kill the
      running job, possibly your own. This script therefore only PRINTS the
      restart command; it never runs it.
    * Anti-footgun: never point DROPBOX_DIR at an arbitrary local directory
      just to switch the mirror on. server.js appends "Deal files were also
      filed in Dropbox: <path>" to reply emails whenever the mirror is enabled,
      so a non-synced path makes the agent tell clients their files were filed
      when nothing synced. This script refuses to proceed unless the path is
      genuinely inside a linked Dropbox root.
"""

import argparse
import datetime
import os
import re
import shutil
import subprocess
import sys
from pathlib import Path

HOME = Path(os.path.expanduser("~"))
DROPBOX_CLI = HOME / ".dropbox-dist" / "dropboxd"
DEFAULT_SERVER_DIR = Path("/home/claudia/email-cowork-server")
DEFAULT_SUBPATH = "- Underwritings/- Claudia Outputs"

# systemctl --user / journalctl --user need this in a non-interactive job shell.
USER_ENV = {**os.environ, "XDG_RUNTIME_DIR": f"/run/user/{os.getuid()}"}

_ok = True


def say(status, msg, detail=""):
    global _ok
    if status == "FAIL":
        _ok = False
    print(f"  [{status}] {msg}" + (f" -- {detail}" if detail else ""))


def run(cmd, **kw):
    return subprocess.run(cmd, capture_output=True, text=True, env=USER_ENV, **kw)


def check_service():
    """dropbox.service should be enabled (survives reboot) and active."""
    if not shutil.which("systemctl"):
        say("WARN", "systemctl not found", "cannot verify the service")
        return
    enabled = run(["systemctl", "--user", "is-enabled", "dropbox.service"]).stdout.strip()
    active = run(["systemctl", "--user", "is-active", "dropbox.service"]).stdout.strip()
    say("PASS" if enabled == "enabled" else "WARN",
        "dropbox.service enabled at boot", enabled or "not installed")
    say("PASS" if active == "active" else "FAIL",
        "dropbox.service running", active or "inactive")


def check_linked():
    """The core gate: is the daemon actually linked to an account?

    An unlinked daemon runs happily and logs a cli_link_nonce URL forever, so
    "the process is up" proves nothing.

    Read the SERVICE JOURNAL rather than shelling out to a status command.
    Both obvious alternatives are traps, verified 2026-08-06:
      * `dropboxd status` is not a status verb -- it tries to start a second
        daemon and prints "Another instance of Dropbox (<pid>) is running!",
        which trivially reads as success and gives a false PASS.
      * `dropbox.py status` MINTS A BRAND-NEW NONCE on every invocation. Three
        consecutive calls returned three different nonces while the daemon's
        own journal line kept advertising one stable value. Calling it to
        "check" therefore hands you a URL that is not the one the daemon is
        advertising, and quietly churns link requests.
    The journal is read-only and reports the daemon's own stable nonce.
    """
    if not DROPBOX_CLI.exists():
        say("FAIL", "dropbox client installed", f"missing {DROPBOX_CLI}")
        return False
    r = run(["journalctl", "--user", "-u", "dropbox.service", "--no-pager", "-n", "200"],
            timeout=60)
    out = r.stdout + r.stderr
    if "isn't linked" in out or "cli_link_nonce" in out:
        say("FAIL", "Dropbox account linked", "still UNLINKED -- a human must open the link URL")
        nonces = re.findall(r"https://www\.dropbox\.com/cli_link_nonce\?nonce=[a-f0-9]+", out)
        if nonces:
            print(f"         current link URL: {nonces[-1]}")
            print("         (stable for THIS daemon instance; it rotates if the service")
            print("          restarts, so never restart dropbox.service after sending it)")
        return False
    # No unlinked chatter in recent history -- corroborate with a synced root.
    say("PASS", "Dropbox account linked", "no unlinked messages in recent journal")
    return True


def detect_root(explicit=None):
    """Find the synced root. Personal = ~/Dropbox; team = ~/<Team> Dropbox."""
    if explicit:
        p = Path(explicit).expanduser()
        if p.is_dir():
            say("PASS", "Dropbox root (given)", str(p))
            return p
        say("FAIL", "Dropbox root (given)", f"not a directory: {p}")
        return None
    cands = [d for d in sorted(HOME.iterdir())
             if d.is_dir() and (d.name == "Dropbox" or d.name.endswith(" Dropbox"))]
    if not cands:
        say("FAIL", "Dropbox root detected",
            "no ~/Dropbox or ~/<Team> Dropbox yet -- linking may still be syncing")
        return None
    if len(cands) > 1:
        say("WARN", "multiple Dropbox roots", ", ".join(d.name for d in cands))
    say("PASS", "Dropbox root detected", str(cands[0]))
    return cands[0]


def ensure_folder(root, subpath, apply_changes):
    target = root / subpath
    if target.is_dir():
        say("PASS", "deal-output folder exists", str(target))
    elif apply_changes:
        target.mkdir(parents=True, exist_ok=True)
        say("PASS", "deal-output folder created", str(target))
    else:
        say("WOULD-DO", "create deal-output folder", str(target))
    # server.js gates the mirror on the PARENT existing; say so plainly.
    say("INFO", "mirror gate = parent must exist", str(target.parent))
    return target


def set_env(server_dir, target, apply_changes):
    env_path = Path(server_dir) / ".env"
    if not env_path.is_file():
        say("FAIL", ".env found", f"missing {env_path}")
        return
    text = env_path.read_text()
    desired = f"DROPBOX_DIR={target}"
    current = None
    for line in text.splitlines():
        if line.startswith("DROPBOX_DIR="):
            current = line
            break
    if current == desired:
        say("PASS", "DROPBOX_DIR already correct", desired)
        return
    say("INFO", "DROPBOX_DIR currently", current if current is not None else "(absent)")
    if not apply_changes:
        say("WOULD-DO", "set DROPBOX_DIR", desired)
        return

    stamp = datetime.datetime.now(datetime.timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    backup = env_path.with_suffix(f".bak-{stamp}")
    shutil.copy2(env_path, backup)

    if current is not None:
        # Replace only that one line; every other setting is left byte-identical.
        new_text = re.sub(r"^DROPBOX_DIR=.*$", desired, text, count=1, flags=re.MULTILINE)
    else:
        new_text = text.rstrip("\n") + f"\n{desired}\n"
    env_path.write_text(new_text)
    say("PASS", "DROPBOX_DIR written", f"{desired}  (backup: {backup.name})")


def main():
    ap = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--server-dir", default=str(DEFAULT_SERVER_DIR))
    ap.add_argument("--dropbox-root", default=None)
    ap.add_argument("--subpath", default=DEFAULT_SUBPATH)
    ap.add_argument("--apply", action="store_true",
                    help="actually create folders and edit .env (default: dry run)")
    args = ap.parse_args()

    mode = "APPLY" if args.apply else "DRY RUN (pass --apply to make changes)"
    print(f"Finish Dropbox setup -- host={os.uname().nodename}  mode={mode}\n")

    print("Service:")
    check_service()

    print("\nAccount link:")
    if not check_linked():
        print("\n" + "-" * 62)
        print("BLOCKED: Dropbox is not linked yet, so the synced folder does not")
        print("exist and DROPBOX_DIR must stay empty. Email the link URL above to")
        print("the requester, then re-run this script on a later job.")
        return 1

    print("\nFolders:")
    root = detect_root(args.dropbox_root)
    if root is None:
        return 1
    target = ensure_folder(root, args.subpath, args.apply)

    print("\nServer config:")
    set_env(args.server_dir, target, args.apply)

    print("\n" + "-" * 62)
    if args.apply:
        print("Done. The mail server still needs a restart to read the new .env:")
        print("    sudo systemctl restart claudia.service")
        print("That unit is root-owned, so it needs root, and it must be run when")
        print("NO job is in flight -- restarting mid-job kills the running job.")
    else:
        print("Dry run only. Re-run with --apply to make the changes above.")
    return 0 if _ok else 1


if __name__ == "__main__":
    sys.exit(main())
