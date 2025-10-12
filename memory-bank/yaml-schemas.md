# YAML schemas (effective)

## rulesregex.yaml (regex mode default)
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

## rules_safe_sendersregex.yaml
- safe_senders: list[string] (regex patterns, commonly anchored with ^...$ for full address)

Conventions
- Lowercase and trimmed entries
- Single quotes in YAML to avoid backslash churn
- Wildcards in legacy '*' become '.*' via converters
