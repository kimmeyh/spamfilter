# Quality checks and invariants

Exporters
- export_rules_to_yaml(): lowercases, strips, de-dupes, sorts all list fields; uses single quotes for regex file
- export_safe_senders_to_yaml(): lowercases, strips, de-dupes, sorts safe_senders; uses single quotes for regex file
- Both create timestamped backups in archive/ before overwrite

Processing
- Safe senders checked before rule evaluation on both passes
- Second-pass reprocessing runs after interactive updates

YAML
- Regex files use single quotes consistently for patterns
