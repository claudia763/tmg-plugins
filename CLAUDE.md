# Agent instructions — tmg-plugins resource library

This repository is the shared resource library for TMG's email work agents ("Claudia"). It contains processing scripts, instruction files, and formatting templates for broker tasks (BOVs, OMs, rent rolls, T-12s, financials, and related work).

## Using the library

- Explore this library at the start of every job and read any README / CLAUDE.md relevant to the task before starting work.
- Prefer the library's scripts and templates over writing your own from scratch: run existing scripts on input files, follow instruction files, and base deliverables on the templates here.
- Copy templates into your job folder to fill in — treat this library's own files as read-only during a job.

## Keeping the library up to date — YOUR responsibility

1. **Update the library yourself.** When you create a reusable Python script, instruction/reference `.md` file, or template while completing a job, save a clean, documented copy into the library's writable additions directory (the `LIBRARY ADDITIONS` path given in your job instructions — organized under `scripts/`, `instructions/`, or `templates/`). The server automatically commits those files to this repository (under `library-additions/`), so agents on every machine receive them.
   - Every script gets a header comment stating what it does, its inputs, and its outputs.
   - Every `.md` file gets a one-line intro stating what it covers and when to use it.
   - If a similar resource already exists, improve it or version it — do not create near-duplicates.
2. **NEVER ask the email sender to manually upload files** to this repository, to the library, or anywhere else — not as a request in your reply, not as a "next step" suggestion. The sender cannot do this. You have write access to the additions directory; contributing new resources is part of completing the job, not something to delegate.
3. **Check the additions directory** alongside this repository on every job — earlier agents may have already built what you need.
