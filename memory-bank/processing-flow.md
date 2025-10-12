# Processing flow (current)

- Load mode and active files
  - OutlookSecurityAgent.set_active_mode() picks regex vs legacy files
  - OutlookSecurityAgent.get_rules() returns (rules_json, safe_senders)
- Primary processing
  - OutlookSecurityAgent.process_emails()
    - Safe senders are checked first
    - Rule evaluation honors regex patterns in regex mode
    - Two-pass: reprocess after interactive updates
- Interactive updates (optional)
  - Enabled with -u/--update_rules
  - OutlookSecurityAgent.prompt_update_rules()
    - Suggests domain-anchored header regex via build_domain_regex_from_address()
    - Can add to SpamAutoDeleteHeader.header or safe_senders
    - Persists immediately via export_rules_to_yaml() and export_safe_senders_to_yaml()
- End-of-run persistence
  - Always exports active structures:
    - export_rules_to_yaml()
    - export_safe_senders_to_yaml()

Diagnostics/invariants
- Exporters lower-case, trim, de-dupe, sort list fields
- Regex YAML uses single quotes for pattern stability
- Backups saved in archive/ with timestamp before overwrite
