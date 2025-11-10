# Quality checks and invariants

Exporters
- export_rules_to_yaml(): lowercases, strips, de-dupes, sorts all list fields; uses single quotes for regex stability
- export_safe_senders_to_yaml(): lowercases, strips, de-dupes, sorts safe_senders; uses single quotes for regex stability
- Both create timestamped backups in archive/ before overwrite

Processing
- Safe senders checked before rule evaluation on both passes
- Second-pass reprocessing runs after interactive updates
- All patterns treated as regex (legacy wildcard mode removed 10/14/2025)

YAML
- Consolidated filenames use single quotes consistently for regex patterns (10/18/2025)
- rules.yaml and rules_safe_senders.yaml contain regex patterns only

Historical Notes
- DEPRECATED 10/18/2025: Separate regex filename variants removed
- Files consolidated to: rules.yaml and rules_safe_senders.yaml (contain regex patterns)
