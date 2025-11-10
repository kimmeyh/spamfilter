# CLI usage

Default mode: REGEX patterns in consolidated YAML files (only supported mode as of 10/14/2025)

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

# DEPRECATED 10/14/2025: Legacy file support removed
# DEPRECATED 10/18/2025: Regex-specific filename variants removed - consolidated to single filenames
# Legacy mode logic commented out - keeping for reference
# - Force legacy files
#   - python withOutlookRulesYAML.py --use-legacy-files
# - Explicit regex files (default and only supported mode)
#   - python withOutlookRulesYAML.py --use-regex-files
#   - python withOutlookRulesYAML.py  # same as above (default)

- Standard processing (no interactive updates)
  - python withOutlookRulesYAML.py
  
# DEPRECATED 10/18/2025: Conversion utilities no longer needed - files already use consolidated names
# - Conversions
#   - python withOutlookRulesYAML.py --convert-rules-to-regex
#   - python withOutlookRulesYAML.py --convert-safe-senders-to-regex

Entrypoints
- main()
- OutlookSecurityAgent.set_active_mode()

File Structure
- rules.yaml - Spam filtering rules (regex patterns)
- rules_safe_senders.yaml - Trusted sender whitelist (regex patterns)
- archive/ - Timestamped backups of YAML files
