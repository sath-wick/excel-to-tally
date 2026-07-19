import json
import os


_ROOT_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
CONFIG_PATH = os.path.join(_ROOT_DIR, "config", "clients.json")
DEFAULT_RULE_PATH = os.path.join(_ROOT_DIR, "rules", "description_rules_494.json")
DEFAULT_DUPLICATE_JSON_PATH = os.path.join(_ROOT_DIR, "exports", "Transactions.json")
DEFAULT_IGNORE_JSON_PATH = os.path.join(_ROOT_DIR, "rules", "ignored_descriptions.json")


def load_config(config_path=CONFIG_PATH):
    with open(config_path, "r", encoding="utf-8") as file_obj:
        return json.load(file_obj)


def get_default_client_name():
    config = load_config()
    clients = config.get("clients", {})
    configured_default = config.get("default_client")

    if configured_default in clients:
        return configured_default

    if clients:
        return next(iter(clients))

    return ""


def get_client_names():
    config = load_config()
    return list(config.get("clients", {}).keys())


def get_client(client_name=None):
    config = load_config()
    clients = config.get("clients", {})
    resolved_name = client_name or config.get("default_client") or get_default_client_name()

    if resolved_name not in clients:
        available = ", ".join(clients.keys()) or "none"
        raise ValueError(f"Unknown client '{resolved_name}'. Available clients: {available}")

    client = dict(clients[resolved_name])
    client["name"] = resolved_name
    return client


def get_client_bank_categories(client_name=None):
    return get_client(client_name).get("bank_categories", {})


def extract_bank_code(bank_ledger):
    ledger_text = str(bank_ledger).strip()
    digits_only = "".join(ch for ch in ledger_text if ch.isdigit())
    return digits_only if digits_only else ledger_text


def _slug(value):
    return "_".join(str(value).strip().lower().split())


def _candidate_keys(bank_ledger):
    ledger_text = str(bank_ledger).strip()
    digits_only = extract_bank_code(ledger_text)
    slug = _slug(ledger_text)

    candidates = []
    for candidate in (ledger_text, slug, digits_only):
        if candidate and candidate not in candidates:
            candidates.append(candidate)

    return candidates


def resolve_bank_account(client_name, bank_ledger):
    client = get_client(client_name)
    banks = client.get("banks", {})

    for candidate in _candidate_keys(bank_ledger):
        if candidate in banks:
            bank = dict(banks[candidate])
            bank.setdefault("name", candidate)
            bank.setdefault("ledger", bank_ledger)
            return bank

    return {
        "name": str(bank_ledger).strip(),
        "ledger": bank_ledger,
    }


def _normalize_path(path_str):
    if not path_str:
        return path_str
    if not os.path.isabs(path_str):
        return os.path.abspath(os.path.join(_ROOT_DIR, path_str))
    return path_str


def resolve_rule_path(bank_ledger, client_name=None, default_rule_path=DEFAULT_RULE_PATH):
    bank = resolve_bank_account(client_name, bank_ledger)
    configured_rule_path = bank.get("rule_path")

    if configured_rule_path:
        norm_path = _normalize_path(configured_rule_path)
        if os.path.exists(norm_path):
            return norm_path

    digits_only = extract_bank_code(bank_ledger)
    if digits_only == "494":
        norm_default = _normalize_path(default_rule_path)
        if os.path.exists(norm_default):
            return norm_default

    for candidate in _candidate_keys(bank_ledger):
        bank_specific_rule_path = os.path.join(_ROOT_DIR, "rules", f"description_rules_{candidate}.json")
        if os.path.exists(bank_specific_rule_path):
            return bank_specific_rule_path

    return _normalize_path(default_rule_path)


def resolve_duplicate_json_path(client_name=None):
    raw_path = get_client(client_name).get("duplicate_json_path", DEFAULT_DUPLICATE_JSON_PATH)
    return _normalize_path(raw_path)


def resolve_ignore_json_path(client_name=None):
    raw_path = get_client(client_name).get("ignore_json_path", DEFAULT_IGNORE_JSON_PATH)
    return _normalize_path(raw_path)


def resolve_module_dr_ledger(client_name, module_name, default_ledger):
    client = get_client(client_name)
    module_config = client.get(module_name.lower(), {})
    return module_config.get("dr_ledger", default_ledger)


def resolve_module_setting(client_name, module_name, setting_name, default_value=None):
    client = get_client(client_name)
    module_config = client.get(module_name.lower(), {})
    return module_config.get(setting_name, default_value)


def is_import_enabled(client_name, module_name, bank_ledger):
    client = get_client(client_name)
    module_config = client.get(module_name.lower(), {})
    enabled_codes = module_config.get("enabled_bank_codes")

    if not enabled_codes:
        return True

    return extract_bank_code(bank_ledger) in {str(code) for code in enabled_codes}
