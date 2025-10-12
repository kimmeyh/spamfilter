# CLI usage

Default mode: REGEX files

Developer setup
- See `memory-bank/dev-environment.md` to create/activate a Python venv before running commands below.

- Interactive updates
  - python withOutlookRulesYAML.py -u
- Force legacy files
  - python withOutlookRulesYAML.py --use-legacy-files
- Explicit regex files (default)
  - python withOutlookRulesYAML.py --use-regex-files
- Conversions
  - python withOutlookRulesYAML.py --convert-rules-to-regex
  - python withOutlookRulesYAML.py --convert-safe-senders-to-regex

Entrypoints
- main()
- OutlookSecurityAgent.set_active_mode()
