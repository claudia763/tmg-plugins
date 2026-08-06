# Library Additions — agent-contributed resources

This folder is the **writable companion** to the tmg-plugins resource library. It persists across jobs and is available to every future agent run. Files saved here are **automatically committed to the tmg-plugins GitHub repo** (under `library-additions/`) after each job, so agents on every machine share them — one more reason to contribute here and never ask a human to upload anything.

## Rules for agents

1. **Contribute reusable work back.** If, while completing a job, you write a Python script, an instruction/reference `.md` file, or a document template that would help future tasks (processing rent rolls, building BOVs/OMs, formatting T-12s, etc.), save a clean copy into this folder before finishing the job.
   - Organize sensibly: `scripts/` for Python, `instructions/` for .md guidance, `templates/` for document/spreadsheet templates.
   - Give every script a header comment and every .md file a one-line intro stating what it does and when to use it.
   - If a similar resource already exists here or in the main library, improve/version it rather than duplicating it.
2. **NEVER ask the email sender to manually upload files** to this library, to GitHub, or anywhere else. The sender cannot do that. Updating the library is your responsibility and you have write access to this folder — just save the files here yourself as part of the job.
3. **Check this folder at the start of every job**, alongside the main library — previous agents may have already built exactly what you need.
