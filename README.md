# Bank Statement to Tally
A framework that extracts data from raw bank statements and excel files and transforms them to import ready data for Tally.

## Multi-client setup

Clients and their bank accounts are configured in `config/clients.json`.

Each client can define:

- `bank_categories`: the menu structure shown in the launcher.
- `banks`: account-specific settings, including the Tally bank ledger name and rule file.
- `duplicate_json_path`: the Tally export JSON used for duplicate checks.
- `ignore_json_path`: the statement description ignore rules.
- `sales.enabled_bank_codes` and `purchases.enabled_bank_codes`: bank codes allowed for those modules.

To add another client, add a new entry under `clients`, then add that client's bank accounts and rule files. Example:

```json
"New Client": {
  "duplicate_json_path": "./exports/new_client/Transactions.json",
  "ignore_json_path": "./rules/ignored_descriptions.json",
  "bank_categories": {
    "HDFC": ["HDFC 1234"]
  },
  "banks": {
    "HDFC 1234": {
      "ledger": "HDFC Bank 1234",
      "rule_path": "./rules/new_client/description_rules_hdfc_1234.json"
    }
  }
}
```
