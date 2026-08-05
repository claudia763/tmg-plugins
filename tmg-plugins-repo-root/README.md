# TMG Plugins — Claude Code marketplace

Install in Claude Code:

    /plugin marketplace add /path/to/tmg-plugins
    /plugin install multifamily-brokerage@tmg-plugins

(Or push this folder to a git repo and `/plugin marketplace add owner/repo`.)

Local-machine dependencies the skills expect (install once):
    npm install -g docx playwright && npx playwright install chromium
    pip install pymupdf

Skills can also be used without the plugin system by copying
multifamily-brokerage/skills/* into ~/.claude/skills/.
