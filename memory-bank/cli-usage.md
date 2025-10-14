# CLI usage

Default mode: REGEX files

Developer setup
- See `memory-bank/dev-environment.md` to create/activate a Python venv before running commands below.

- Interactive updates
  - python withOutlookRulesYAML.py -u
  - During prompts: enter d/e/s/sd/?
    - d  - Add sender domain regex to SpamAutoDeleteHeader (blocks by domain)
    - e  - Add full sender email to SpamAutoDeleteHeader (blocks this email)
    - s  - Add literal address/domain to safe_senders (never block)
    - sd - Add sender-domain regex to safe_senders (never block any subdomain)
    - ?  - Show help for options
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
