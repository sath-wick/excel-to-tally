# Plug-and-Play AI Rule Generation & Management Architecture
**Project:** Tally Framework  
**Location:** `d:/tally-framework/plans/ai_rule_generation_and_management_plan.md`  
**Status:** Proposed / Review Required  
**Author:** Antigravity AI  

---

## 1. Executive Summary & Objectives

The goal is to eliminate manual JSON editing and mental mapping when creating description rules (`description_rules_<bank_code>.json`), while enabling complex **amount-dependent rule splits** without hardcoded Python overrides (`bank_3332_rules.py`).

We will implement a **loosely coupled, plug-and-play architecture** consisting of:
1. **Declarative Amount Schema Extension** (`core/rule_engine.py` & `core/engine.py`): Non-breaking extension allowing JSON rules to define `min_amount`, `max_amount`, and `amount_tiers`.
2. **Modular Rule Automation Engine (`tools/rule_generator/`)**: A self-contained package providing:
   - **Chart of Accounts Harvesting (`chart_of_accounts.py`)**: Whitelists existing Tally ledgers from rules & `Transactions.json` to prevent LLM hallucinations.
   - **Gemini Pro AI Engine (`gemini_engine.py`)**: Uses `google-genai` with Structured JSON mode to auto-classify unclassified descriptions and detect amount tiers.
   - **Historical Data Miner (`tally_miner.py`)**: Mines historical `Transactions.json` against statements to extract baseline rules at zero API cost.
   - **Safe Rule Applier (`rule_applier.py`)**: Merges, prioritizes, and writes rules while preserving UTF-8 BOM encoding.
3. **Plug-and-Play Interfaces**:
   - Standalone CLI utility (`tools/ai_rule_generator.py`)
   - Interactive Tkinter **Rule Studio GUI** (`tools/rule_studio_gui.py`) with live match preview and one-click AI generation.

---

## 2. Architectural Design & Loose Coupling Guarantee

To ensure this system acts as a **plug-and-play component** with **loose coupling** and **maximum efficiency**:

```
+-----------------------------------------------------------------------------------+
|                           EXISTING PIPELINE (Unchanged)                           |
|  launcher.py  --->  main.py  --->  VoucherEngine  --->  Statement_Import.xlsx    |
+-----------------------------------------------------------------------------------+
                                       |
                                       v  (Reads Unclassified Rows & Existing Ledgers)
+-----------------------------------------------------------------------------------+
|                   PLUG-AND-PLAY RULE AUTOMATION LAYER                             |
|                                                                                   |
|   +--------------------------+          +-------------------------------------+   |
|   |  chart_of_accounts.py    |          |  gemini_engine.py (Gemini Pro API)  |   |
|   |  (Whitelists COA Ledgers)|          |  - Structured JSON Prompting        |   |
|   +--------------------------+          |  - Amount Slab Detection            |   |
|                |                        +-------------------------------------+   |
|                +---------------------------+               |                      |
|                                            v               v                      |
|   +--------------------------+          +-------------------------------------+   |
|   |  tally_miner.py          | -------> |  rule_applier.py                    |   |
|   |  (Historical Mining)     |          |  - Deduplication & Priority Merge   |   |
|   +--------------------------+          +-------------------------------------+   |
+-----------------------------------------------------------------------------------+
                                                             |
                                                             v  (Updates UTF-8-BOM JSON)
+-----------------------------------------------------------------------------------+
|               rules/<Client>/description_rules_<bank_code>.json                   |
+-----------------------------------------------------------------------------------+
                                                             |
                                                             v  (Declarative Amount Support)
+-----------------------------------------------------------------------------------+
|                   core/rule_engine.py & core/engine.py                            |
|             (Evaluates min_amount, max_amount, & amount_tiers)                    |
+-----------------------------------------------------------------------------------+
```

### Loose Coupling Principles:
1. **Zero Impact on Core Execution Pipeline:** `main.py`, `sales_main.py`, and `purchases_main.py` do not depend on `google-genai` or the `tools/rule_generator/` module. If the internet is down or API keys are missing, standard statement extraction and import work 100% identically.
2. **100% Backwards Compatible Rule Schema:** Existing JSON rules that do not contain `min_amount`, `max_amount`, or `amount_tiers` are evaluated with zero performance penalty.
3. **Stand-Alone Package Structure:** All AI and mining logic resides cleanly under `tools/rule_generator/`, keeping `core/` strictly focused on deterministic voucher processing.

---

## 3. Detailed Component Specifications

### 3.1. Core Rule Engine Extension (`core/rule_engine.py` & `core/engine.py`)
We extend `RuleEngine.match(transaction)` (`core/rule_engine.py`) to verify amount boundaries when a regex pattern matches:

```python
# In core/rule_engine.py -> RuleEngine.match(transaction)
for rule in self.rules:
    patterns = rule["pattern"]
    if isinstance(patterns, str):
        patterns = [patterns]

    for pattern in patterns:
        if re.search(pattern, description, re.IGNORECASE):
            # Evaluate declarative amount filters if present
            txn_amount = getattr(transaction, "amount", 0.0)
            if "min_amount" in rule and rule["min_amount"] is not None and txn_amount < float(rule["min_amount"]):
                continue
            if "max_amount" in rule and rule["max_amount"] is not None and txn_amount > float(rule["max_amount"]):
                continue
            return rule
return None
```

We extend `_resolve_directional_rule(rule, transaction)` (`core/engine.py`) to evaluate `amount_tiers` if present:
```python
# In core/engine.py -> _resolve_directional_rule
if "amount_tiers" in resolved_rule and isinstance(resolved_rule["amount_tiers"], list):
    txn_amount = getattr(transaction, "amount", 0.0)
    for tier in resolved_rule["amount_tiers"]:
        t_min = float(tier.get("min", float("-inf")))
        t_max = float(tier.get("max", float("inf")))
        if t_min <= txn_amount <= t_max:
            if "ledger" in tier:
                resolved_rule["ledger"] = tier["ledger"]
            if "voucher_type" in tier:
                resolved_rule["voucher_type"] = tier["voucher_type"]
            break
```

---

### 3.2. Modular Rule Generator Package (`tools/rule_generator/`)

#### 3.2.1. `chart_of_accounts.py` (COA Harvester)
- **Purpose:** Extracts the exact whitelist of valid Tally ledgers for a given client to eliminate AI hallucinations.
- **Sources:**
  1. All `ledger`, `payment_ledger`, and `receipt_ledger` values from `rules/<Client>/description_rules_*.json`.
  2. All `dspvchledaccount` values from `exports/<Client>/Transactions.json`.
- **Output:** A deduplicated, sorted Python list of strings representing the exact Chart of Accounts (`e.g., ["Bank Charges & Commission", "Sundry Debtors - Chandra Enterprises", "Staff Welfare", ...].`)

#### 3.2.2. `gemini_engine.py` (Gemini Pro AI Classifier)
- **Purpose:** Calls Gemini Pro (`google-genai` Python SDK) using Structured JSON mode.
- **Input:**
  - `unclassified_rows`: List of dicts `[{"date": "12-04-2026", "description": "NEFT/AXIS/CHANDRA ENT/INV09", "amount": 15000.0, "direction": "OUT"}, ...]`
  - `chart_of_accounts`: Whitelisted ledger names.
  - `existing_rules_sample`: 10 representative existing rules from the client for pattern style matching.
- **Structured JSON Schema Prompt:**
  Instructs Gemini to return a strict JSON object:
  ```json
  {
    "rules": [
      {
        "pattern": ["CHANDRA ENT"],
        "ledger": "Sundry Debtors - Chandra Enterprises",
        "priority": 100,
        "min_amount": null,
        "max_amount": null,
        "amount_tiers": null,
        "confidence": "HIGH",
        "reasoning": "Exact match to whitelisted debtor ledger from NEFT payment description."
      }
    ]
  }
  ```
- **Amount Slab Intelligence:** Prompt explicitly directs Gemini: *"If you observe that transactions for the same keyword map to different natures based on amount (e.g. `ATM WITHDRAWAL` > 50,000 vs < 5,000), generate either separate rules with `min_amount`/`max_amount` or a single rule with `amount_tiers`."*

#### 3.2.3. `tally_miner.py` (Historical Data Miner)
- **Purpose:** Cross-references historical statements against `exports/<Client>/Transactions.json` (`dspvchdetail`) to mine high-confidence correlation rules at zero API cost.
- **Algorithm:**
  1. Joins statement rows and Tally JSON entries on `(Date, Amount, Direction)`.
  2. Calculates frequency distribution of `Description Substring -> Ledger`.
  3. Where correlation frequency $\ge 85\%$, generates a baseline rule (`priority: 50`).
  4. Where correlation splits cleanly by amount thresholds (using 1D decision-tree splitting on `amount`), generates `amount_tiers` rules automatically.

#### 3.2.4. `rule_applier.py` (Safe Applier & Deduplicator)
- **Purpose:** Safely merges candidate rules into `description_rules_<code>.json`.
- **Key Protections:**
  - Preserves `utf-8-sig` (BOM) encoding required by `RuleEngine`.
  - Checks if pattern already exists (`case-insensitive` check) to prevent duplicates.
  - Sorts existing + new rules by `priority` descending.
  - Backs up the original JSON file to `.bak` before writing.

---

### 3.3. User Interfaces

#### 3.3.1. CLI Utility (`tools/ai_rule_generator.py`)
A fast, non-interactive or semi-interactive terminal utility:
```bash
# Generate rules using Gemini Pro from the Unclassified tab of a processed Excel import
python tools/ai_rule_generator.py output/Statement_Import.xlsx "YES 2085" "Arjun Rao" --mode gemini --apply-high-confidence

# Or run Historical Mining across all clients
python tools/ai_rule_generator.py --mode mine --client "Arjun Rao"
```

#### 3.3.2. Interactive Rule Studio GUI (`tools/rule_studio_gui.py`)
A Tkinter desktop window designed to pair with `launcher.py`:
- **Top Panel (`Client & Bank Selection`)**: Dropdown to select Client and Bank Ledger. Auto-loads current rules and unclassified rows from the last run or selected Excel/PDF.
- **Middle Panel (`Unclassified Clusters & AI Suggestions Table`)**:
  - Groups unclassified rows by token similarity.
  - Shows `Confidence Badge` (`HIGH [Green]`, `MEDIUM [Yellow]`, `LOW [Red]`), `Suggested Pattern`, and `Suggested Ledger` (with autocomplete dropdown tied to `chart_of_accounts`).
- **Action Toolbar**:
  - `[ ✨ Ask Gemini Pro to Classify Remaining ]`: Triggers `gemini_engine.py` on selected/unclassified rows.
  - `[ + Add Amount Condition / Slabs ]`: Opens visual `min_amount` / `max_amount` / `amount_tiers` slab editor.
  - `[ Live Preview Matches ]`: Highlights statement rows that will be instantly resolved.
  - `[ Save & Apply to Rules File ]`: Calls `rule_applier.py` and updates `description_rules_<code>.json`.

---

## 4. Implementation & File Change Plan

### Component Group 1: Core Framework Extension
#### [MODIFY] `core/rule_engine.py`
- Add check for `min_amount` and `max_amount` inside `RuleEngine.match(transaction)`.

#### [MODIFY] `core/engine.py`
- Add check and ledger/voucher routing for `amount_tiers` inside `_resolve_directional_rule(rule, transaction)`.

---

### Component Group 2: Plug-and-Play Rule Generator Package
#### [NEW] `tools/rule_generator/__init__.py`
#### [NEW] `tools/rule_generator/chart_of_accounts.py`
#### [NEW] `tools/rule_generator/gemini_engine.py`
#### [NEW] `tools/rule_generator/tally_miner.py`
#### [NEW] `tools/rule_generator/rule_applier.py`

---

### Component Group 3: CLI & Desktop UI
#### [NEW] `tools/ai_rule_generator.py`
- Standalone CLI wrapper routing `--mode gemini` or `--mode mine` to `rule_generator` module.
#### [NEW] `tools/rule_studio_gui.py`
- Tkinter GUI desktop app for interactive clustering, Gemini generation, amount slab creation, and live review.

---

## 5. Verification & Testing Strategy

### 5.1. Automated / Unit Verification
1. **RuleEngine Amount Filtering Test:**
   - Create a test transaction object `Txn(description="CASH WITHDRAWAL", amount=65000)`.
   - Verify that when evaluated against rules with `min_amount: 50000` vs `max_amount: 49999`, exact rule selection and priority ordering succeed.
2. **Amount Tier Evaluation Test:**
   - Verify `_resolve_directional_rule` correctly overrides `ledger` based on `amount_tiers` definitions for Bank 3332 `Labour Charges` test cases (`>=2000`, `500..1999`, `<500`).
3. **COA Harvesting & Deduplication Test:**
   - Verify `chart_of_accounts.extract_chart_of_accounts("Arjun Rao")` returns clean, sorted ledger strings from both rules and `exports/Arjun Rao/Transactions.json`.
4. **Safe Rule Applier Test:**
   - Verify `rule_applier.append_rules()` preserves `utf-8-sig` BOM header, creates `.bak` file, and prevents duplicate patterns.

### 5.2. Integration & End-to-End Verification
1. Run `python tools/ai_rule_generator.py output/Statement_Import.xlsx "YES 2085" "Arjun Rao" --mode gemini` with sample unclassified rows and verify structured JSON return from Gemini Pro.
2. Launch `python tools/rule_studio_gui.py`, verify smooth Tkinter rendering, test live preview updates, and confirm safe writing to `rules/Arjun Rao/description_rules_2085.json`.
