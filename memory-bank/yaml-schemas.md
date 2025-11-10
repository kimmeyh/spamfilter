# YAML schemas (effective)

## rules.yaml (regex mode - only supported format)
- version: string
- settings.default_execution_order_increment: int
- rules: list[Rule]

Rule
- name: string
- enabled: 'True'|'False' (string)
- isLocal: 'True'|'False'
- executionOrder: string|int
- conditions:
  - type: 'OR'|'AND'
  - from|header|subject|body: list[string] (regex)
- actions: object (assign_to_category/delete/move_to_folder/etc.)
- exceptions:
  - from|header|subject|body: list[string] (regex)
- metadata: optional

## rules_safe_senders.yaml (regex mode - only supported format)
- safe_senders: list[string] (regex patterns, commonly anchored with ^...$ for full address)

Conventions
- Lowercase and trimmed entries
- Single quotes in YAML to avoid backslash churn
- Wildcards in legacy '*' become '.*' via converters (deprecated - files already converted)
- All patterns are regex (no legacy wildcard-only mode)

Historical Notes
- DEPRECATED 10/18/2025: Regex-specific filename suffixes removed (_regex)
- Consolidated to single filenames: rules.yaml and rules_safe_senders.yaml
- These files contain regex patterns (legacy wildcard mode removed 10/14/2025)
