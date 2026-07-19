import os
import json

def get_gemini_client(api_key):
    # Try importing new google-genai SDK first
    try:
        from google import genai
        from google.genai import types
        return "genai", genai.Client(api_key=api_key)
    except ImportError:
        pass

    # Fallback to legacy google-generativeai SDK
    try:
        import google.generativeai as legacy_genai
        legacy_genai.configure(api_key=api_key)
        return "legacy", legacy_genai
    except ImportError:
        raise ImportError(
            "Neither modern 'google-genai' nor legacy 'google-generativeai' SDK is installed. "
            "Please run: pip install google-genai"
        )

def generate_rules_with_gemini(unclassified_txns, chart_of_accounts, api_key=None, model_name=None, feedback_reasonings=None):
    """
    Sends unclassified transactions to Gemini to propose ledger mapping rules.
    """
    api_key = api_key or os.environ.get("GEMINI_API_KEY")
    if not api_key:
        raise ValueError("GEMINI_API_KEY environment variable or argument is missing.")

    # Determine default model name
    if not model_name:
        model_name = "gemini-3.5-flash"  # fast, smart default. User can override to gemini-3.5-pro

    sdk_type, client = get_gemini_client(api_key)

    # Format feedback section if user-provided corrections exist
    feedback_section = ""
    if feedback_reasonings:
        feedback_section = f"\n### User-Provided Past Mappings & Reasonings (MUST RESPECT):\nUse the following past corrections as strong guidance when mapping descriptions with similar patterns. Map similar transactions to these correct ledgers and incorporate/respect the reasoning provided:\n{json.dumps(feedback_reasonings, indent=2)}\n"

    # Format the prompt
    prompt = f"""
You are an expert accountant automating Tally ERP data imports.
We have a set of unclassified bank transaction rows and a strict whitelist of existing Tally Ledgers (Chart of Accounts).
{feedback_section}
Your goal is to analyze the unclassified descriptions, group similar transactions together, and generate classification rules mapping description patterns to correct Tally Ledgers.

### Rules JSON Schema Format:
Your output must be a single JSON object containing a "rules" list of rule objects. Each rule object must adhere to:
{{
  "pattern": ["regex_or_string"],          // Minimal matching regex/substring patterns, case-insensitive. Strip UTR, dates, numbers.
  "ledger": "Tally Ledger Name",          // MUST be selected exactly from the Whitelisted Ledgers below.
  "payment_ledger": "Optional Override",  // Use if the ledger is different for OUT (withdrawals) transactions.
  "receipt_ledger": "Optional Override",  // Use if the ledger is different for IN (deposits) transactions.
  "priority": 100,                        // Integer. Higher means evaluated first. Specific matches get 100, generic gets 10.
  "min_amount": null,                     // Optional: only match if transaction amount is >= this value.
  "max_amount": null,                     // Optional: only match if transaction amount is <= this value.
  "amount_tiers": [                       // Optional: Use if mapping splits into different ledgers depending on amounts.
     {{
       "min": 500,
       "max": 1999.99,
       "ledger": "Vehicle Maintainance"
     }}
  ],
  "confidence": "HIGH" | "MEDIUM",        // Your confidence level.
  "reasoning": "Reason for match"         // Brief explanation.
}}

### Whitelisted Tally Ledgers (Chart of Accounts):
{json.dumps(chart_of_accounts, indent=2)}

### Unclassified Transactions to Classify:
{json.dumps(unclassified_txns, indent=2)}

### Critical Rules:
1. ONLY map to ledger names listed in the Whitelisted Tally Ledgers above. If a transaction has no suitable ledger in the whitelist, DO NOT map it, or map it to a generic fallback like "Suspense A/c" or "Unclassified Receipts/Payments" only if they exist in the whitelist.
2. In pattern strings, clean them. E.g., change "UPI-ZOMATO-28374-PAY" to "ZOMATO" or "UPI-ZOMATO". Avoid putting specific txn IDs, dates, or variable reference numbers in the patterns.
3. If similar descriptions route to different ledgers based on amounts, use "min_amount", "max_amount", or "amount_tiers" to separate them.
4. Output valid, parseable JSON only.
"""

    if sdk_type == "genai":
        from google.genai import types
        response = client.models.generate_content(
            model=model_name,
            contents=prompt,
            config=types.GenerateContentConfig(
                response_mime_type="application/json"
            )
        )
        response_text = response.text
    else:
        # Legacy SDK
        model = client.GenerativeModel(model_name)
        response = model.generate_content(
            prompt,
            generation_config={"response_mime_type": "application/json"}
        )
        response_text = response.text

    # Parse and clean output
    try:
        # Strip codeblock wrappers if LLM returned them despite JSON mode
        cleaned_text = response_text.strip()
        if cleaned_text.startswith("```json"):
            cleaned_text = cleaned_text[7:]
        if cleaned_text.endswith("```"):
            cleaned_text = cleaned_text[:-3]
        cleaned_text = cleaned_text.strip()
        
        parsed = json.loads(cleaned_text)
        return parsed.get("rules", [])
    except Exception as e:
        print(f"Error parsing Gemini response: {e}")
        print("Raw response was:")
        print(response_text)
        return []
