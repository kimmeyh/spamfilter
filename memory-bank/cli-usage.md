# CLI usage

Default mode: REGEX files (only supported mode as of 10/14/2025)

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
- DEPRECATED 10/14/2025: Legacy file support removed
  # - Force legacy files
  #   - python withOutlookRulesYAML.py --use-legacy-files
- Explicit regex files (default and only supported mode)
  - python withOutlookRulesYAML.py --use-regex-files
  - python withOutlookRulesYAML.py  # same as above (default)
- Conversions
  - python withOutlookRulesYAML.py --convert-rules-to-regex
  - python withOutlookRulesYAML.py --convert-safe-senders-to-regex

Entrypoints
- main()
- OutlookSecurityAgent.set_active_mode()
