# CLI usage

Default mode: REGEX patterns in consolidated YAML files (only supported mode as of 10/14/2025)

Developer setup
- See `memory-bank/dev-environment.md` to create/activate a Python venv before running commands below.

- Standard processing (no interactive updates)
cd D:\Data\Harold\github\OutlookMailSpamFilter && ./.venv/Scripts/Activate.ps1 && python withOutlookRulesYAML.py

- Interactive mode processing (prompts to add rules)
cd D:\Data\Harold\github\OutlookMailSpamFilter && ./.venv/Scripts/Activate.ps1 && python withOutlookRulesYAML.py -u

- Interactive updates
  - python withOutlookRulesYAML.py -u
  - During prompts: enter d/e/s/sd/?
    - d  - Add sender domain regex to SpamAutoDeleteHeader (blocks by domain)
    - e  - Add full sender email to SpamAutoDeleteHeader (blocks this email)
    - s  - Add literal address/domain to safe_senders (never block)
    - sd - Add sender-domain regex to safe_senders (never block any subdomain)
    - ?  - Show help for options

Entrypoints
- main()
- OutlookSecurityAgent.set_active_mode()

File Structure
- rules.yaml - Spam filtering rules (regex patterns)
- rules_safe_senders.yaml - Trusted sender whitelist (regex patterns)
- archive/ - Timestamped backups of YAML files
