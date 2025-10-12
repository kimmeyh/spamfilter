# Regex conventions

- All patterns lowercased and trimmed on export
- YAML uses single quotes to avoid escape churn for backslashes
- Glob semantics: '*' in legacy becomes '.*' in regex converters
  - convert_rules_yaml_to_regex()
  - convert_safe_senders_yaml_to_regex()

Domain header patterns
- build_domain_regex_from_address(addr_or_domain)
  - Produces '@(?:[a-z0-9-]+\.)*<anchor>\.[a-z0-9.-]+$'
  - Anchor chosen from first meaningful subdomain left of TLD
  - Fallback: '@(?:[a-z0-9-]+\.)*[a-z0-9-]+\.[a-z0-9.-]+$'

Sender domain safe-senders patterns
- build_sender_domain_safe_regex(addr_or_domain)
  - Produces '^[^@\s]+@(?:[a-z0-9-]+\.)*<domain>$'
  - Matches any local part at the exact domain and any subdomains

Examples
- 'mailer\-daemon@aol\.com'
- '@(?:[a-z0-9-]+\.)*example\.com$'
- '@(?:[a-z0-9-]+\.)*example\.[a-z0-9.-]+$' (generic TLD)
- '^[^@\s]+@(?:[a-z0-9-]+\.)*lifeway\.com$' (safe-senders sender-domain)
