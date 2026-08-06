# Merging toolkit changes back to the tmg-plugins GitHub repo

How to land a new parser (or any edit to `process_rent_roll.py` /
`process_t12.py`) in `github.com/claudia763/tmg-plugins` from a TMG job
machine. Read this before touching git — three of the steps below are traps
that cost a full session the first time round (8/5/2026, Eclipse of White
Rock parsers).

## 1. git is not on PATH

There is no standalone Git install and no `gh` CLI. A working git ships
inside GitHub Desktop:

    C:\Users\dmytr\AppData\Local\GitHubDesktop\app-<ver>\resources\app\git\cmd

Glob the version — it moves (`app-3.6.3` at 8/2026). Prepend that directory
to `$env:PATH` at the top of every PowerShell call; PowerShell tool calls do
not share shell state, so it has to be re-set each time.

## 2. Set `core.autocrlf false` BEFORE the first checkout

The bundled git's system config sets `core.autocrlf=true`. The repo stores
these files LF, so a normal clone writes CRLF into the worktree. Consequences
if you miss it:

- Diffing the checkout against the LF files in `resources/tmg-plugins` shows
  every line as changed.
- `git am` fails with `does not match index`.
- Committing rewrites the whole file, turning a 700-line diff into a 5,000-
  line one that is impossible to review.

Do this instead:

    git clone --no-checkout <url> repo
    cd repo
    git config core.autocrlf false
    git checkout main

(If you already cloned: `git config core.autocrlf false` then
`git rm --cached -r . --quiet` and `git reset --hard`.)

Sanity check before committing — the blob size must equal the LF file size:

    git cat-file -s (git rev-parse "HEAD:multifamily-brokerage/skills/process_t12.py")

## 3. Credentials

Check `$env:GITHUB_TOKEN` first — the job harness may inject a fine-grained
PAT (`github_pat_...`) for the `claudia763` account with push rights. Confirm
before relying on it, and never print the token:

    $h=@{ Authorization="Bearer $env:GITHUB_TOKEN"; "User-Agent"="tmg" }
    (Invoke-WebRequest "https://api.github.com/repos/claudia763/tmg-plugins" `
        -Headers $h -UseBasicParsing).Content | ConvertFrom-Json |
        Select-Object -ExpandProperty permissions

Push with the token supplied through a credential helper, so it never reaches
the command line, the reflog or `.git/config`:

    git -c credential.helper='!f() { echo username=x-access-token; echo "password=$GITHUB_TOKEN"; }; f' push origin main

If there is no `GITHUB_TOKEN`, do NOT go hunting for workarounds. The
machine's stored GitHub Desktop credential belongs to a different account
(`Yacotli`) and is read-only on this repo; the `manager` and `desktop`
credential helpers both fail headlessly (the latter needs GitHub Desktop
actually running, for `DESKTOP_PORT`). Deliver the merge as the changed files
plus a `git format-patch` and say what is blocking. Do not fork the repo to a
personal account to raise a PR without asking — it puts a second public copy
of TMG's toolkit under someone's personal account.

## 4. Always rebase before pushing

The repo is also edited through the GitHub web UI ("Add files via upload")
and by the automatic `library-additions/` commits, so `origin/main` moves
between sessions. Fetch, check for overlap with your own change, and rebase —
never force-push:

    git fetch origin
    git rev-list --left-right --count origin/main...HEAD    # behind / ahead
    git diff --name-only <base> origin/main                 # what they touched
    git rebase origin/main
    git diff --stat origin/main HEAD                        # same stat as before?

## 5. Verify the landed artifact, not just the local one

A push reporting success is not proof the right bytes landed. Download the
files back from `raw.githubusercontent.com` at the pushed SHA, hash-compare
them against local, and re-run the toolkit from the downloaded copies against
the source files — the reconciliation block tying out is the real proof.

Also run `scripts/parser_detection_regression.py` (in this folder) whenever a
new registered parser is part of the change: it offers every XLSX you have to
every registered detector on both sides and fails if any detector claims a
file that is not its own. A too-loose detector is a silent bug — the wrong
parser runs and still prints a reconciliation block.

## 6. Housekeeping

- Keep the working clone out of `outbox/` (only outbox files get emailed).
- Write commit messages with `git commit -F <file>`. A `@'...'@` here-string
  passed straight to `git commit -m` gets mangled by PowerShell 5.1 argument
  parsing and git reads the fragments as pathspecs.
- Never modify anything under `resources/tmg-plugins` — it is read-only. Do
  the work in a clone inside the job folder.
