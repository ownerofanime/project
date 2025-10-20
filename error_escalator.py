import os
import re
import json
import csv
import datetime
import requests

# --- Real-World Service Connectors ---
try:
    from openai import OpenAI
    import docx
    import pdfplumber
    import pandas as pd
except ImportError as e:
    print(f"Error: A required library is missing. {e}. Please run 'pip install openai python-docx pdfplumber pandas requests openpyxl'")
    exit()

# ==============================================================================
# 1. SECURE CONFIGURATION
# Loads credentials and settings from environment variables.
# ==============================================================================
class Config:
    def __init__(self):
        # self.openai_api_key = os.getenv("OPENAI_API_KEY")
        # self.zapier_webhook_url = os.getenv("ZAPIER_WEBHOOK_URL")
        self.openai_api_key = "df35b053783046adbedc8e423c841443"
        self.zapier_webhook_url = "https://hooks.zapier.com/hooks/catch/25037671/ur6czf1/"

        
        if not all([self.openai_api_key, self.zapier_webhook_url]):
            print("CRITICAL ERROR: Missing one or more environment variables (OPENAI_API_KEY, ZAPIER_WEBHOOK_URL).")
            exit()

# ==============================================================================
# 2. REAL-WORLD SERVICE MANAGERS
# Class to handle interactions with the Zapier webhook.
# ==============================================================================
class ZapierWebhookNotifier:
    """Handles sending notifications via a Zapier Webhook."""
    def __init__(self, config: Config):
        self.config = config

    def send_escalation_notification(self, recipient_email: str, recipient_name: str, subject: str, summary: str, steps: str):
        if not self.config.zapier_webhook_url:
            print("Zapier Webhook URL not configured. Skipping notification.")
            return

        payload = {
            "recipient_email": recipient_email,
            "recipient_name": recipient_name,
            "subject": subject,
            "problem_summary": summary,
            "recommended_steps": steps
        }

        try:
            print(f"Sending notification to Zapier for {recipient_email}...")
            response = requests.post(self.config.zapier_webhook_url, json=payload, timeout=10)
            response.raise_for_status()  # Raises an HTTPError for bad responses (4xx or 5xx)
            print("Successfully sent notification to Zapier.")
        except requests.exceptions.RequestException as e:
            print(f"Failed to send notification to Zapier: {e}")

# ==============================================================================
# 3. KNOWLEDGE AND DATA GATHERING
# ==============================================================================
def load_knowledge_base(filepath="Knowledge Base.docx") -> str:
    """Loads the entire text content from the DOCX knowledge base."""
    try:
        doc = docx.Document(filepath)
        return "\n".join([para.text for para in doc.paragraphs if para.text])
    except Exception as e:
        print(f"Error loading Knowledge Base: {e}")
        return ""

def load_case_log_as_dataframe(filepath="Case Log.xlsx") -> pd.DataFrame:
    """Loads the case log Excel file into a pandas DataFrame."""
    try:
        # Use pandas' read_excel function for .xlsx files
        df = pd.read_excel(filepath)
        print("Successfully loaded Case Log into DataFrame from Excel file.")
        return df
    except FileNotFoundError:
        print(f"Warning: Case log Excel file not found at '{filepath}'.")
        return pd.DataFrame()
    except Exception as e:
        print(f"Error loading Case Log into DataFrame from Excel file: {e}")
        return pd.DataFrame()

def load_escalation_contacts(filepath="Product Team Escalation Contacts.pdf") -> dict:
    """
    Parses the PDF to dynamically extract escalation contacts for each module.
    """
    contacts = {}
    print(f"Dynamically loading contacts from '{filepath}'...")
    try:
        with pdfplumber.open(filepath) as pdf:
            text = pdf.pages[0].extract_text()

        module_markers = {
            'CNTR': 'Container (CNTR)', 'VSL': 'Vessel (VS)',
            'EDI/API': 'EDI/API (EA)', 'INFRA/SRE': 'Others'
        }
        module_indices = {code: text.find(marker) for code, marker in module_markers.items() if marker in text}
        
        sorted_modules = sorted(module_indices.items(), key=lambda item: item[1])

        for i, (code, start_index) in enumerate(sorted_modules):
            end_index = sorted_modules[i+1][1] if i + 1 < len(sorted_modules) else len(text)
            section_text = text[start_index:end_index]
            
            name_match = re.search(r'([A-Z][a-z]+\s[A-Z][a-z]+)\s*-', section_text)
            email_match = re.search(r'([\w.-]+@[\w.-]+)', section_text)
            
            if name_match and email_match:
                contacts[code] = {"name": name_match.group(1).strip(), "email": email_match.group(0).strip()}

    except Exception as e:
        print(f"CRITICAL: Could not parse contacts PDF dynamically: {e}. Using fallback contacts.")
        contacts = {
            'CNTR': {"name": "Mark Lee", "email": "mark.lee@psa123.com"},
            'VSL': {"name": "Jaden Smith", "email": "jaden.smith@psa123.com"},
            'EDI/API': {"name": "Tom Tan", "email": "tom.tan@psa123.com"},
            'INFRA/SRE': {"name": "Jacky Chan", "email": "jacky.chan@psa123.com"}
        }
    
    print(f"Loaded contacts: {contacts}")
    return contacts


def gather_log_context(identifiers: dict, log_directory=".") -> str:
    """Searches all .log files for lines containing any of the identifiers."""
    context = ""
    log_files = [f for f in os.listdir(log_directory) if f.endswith(".log")]
    search_terms = [v for v in identifiers.values() if v]
    if not search_terms: return "No identifiers found to search logs."

    for log_file in log_files:
        try:
            with open(os.path.join(log_directory, log_file), 'r', encoding='utf-8') as f:
                relevant_lines = [line.strip() for line in f if any(term in line for term in search_terms)]
                if relevant_lines:
                    context += f"\n--- Relevant entries from {log_file} ---\n" + "\n".join(relevant_lines)
        except Exception as e:
            context += f"\nCould not read {log_file}: {e}\n"
    return context if context else "No relevant entries found in logs."

# ==============================================================================
# 4. AI ASSISTANT (Using OpenAI API)
# ==============================================================================
class AIAssistant:
    def __init__(self, config: Config):
        self.config = config
        self.client = OpenAI(api_key=self.config.openai_api_key)

    def analyze_issue(self, alert_text: str, log_context: str, knowledge_base: str, case_log_df: pd.DataFrame) -> dict:
        """Asks the OpenAI API to analyze the issue and decide on escalation."""
        
        case_log_summary = ""
        if not case_log_df.empty:
            past_problems = case_log_df['Problem Statements'].dropna().unique().tolist()
            case_log_summary = "Previously escalated problems:\n- " + "\n- ".join(past_problems)

        system_prompt = """
        You are an expert Site Reliability Engineer (SRE). Your task is to analyze an incoming alert and decide if it needs escalation based on a strict set of rules. You MUST respond with a valid JSON object.
        """

        user_prompt = f"""
        **Reference Documents:**
        --- KNOWLEDGE BASE (For solution steps) ---
        {knowledge_base}
        --- END KNOWLEDGE BASE ---

        --- HISTORICAL CASE LOG (For escalation history) ---
        {case_log_summary}
        --- END HISTORICAL CASE LOG ---

        **Current Incident Data:**
        --- INCOMING ALERT ---
        {alert_text}
        --- END ALERT ---

        --- RELEVANT LOG SNIPPETS ---
        {log_context}
        --- END LOG SNIPPETS ---

        **Your Task and Escalation Rules:**
        1.  Analyze the incoming alert.
        2.  Compare with the Historical Case Log. Does the current problem closely match any of the "Previously escalated problems"?
        3.  Apply Escalation Logic:
            * IF the problem matches a previously escalated problem, THEN you MUST escalate.
            * IF the problem does NOT match any previously escalated problem (it is new), THEN you MUST escalate.
            * ELSE (known but not escalated before), DO NOT escalate.
        4.  If Escalating, determine the responsible module (CNTR, VSL, EDI/API, or INFRA/SRE), summarize the problem, and recommend steps from the Knowledge Base.

        **Output Format:**
        Respond with a valid JSON object with this structure:
        {{
          "escalate": boolean,
          "problem_summary": "string or null",
          "recommended_steps": "string or null",
          "responsible_module": "string or null"
        }}
        """

        try:
            print("Sending context to OpenAI for analysis...")
            response = self.client.chat.completions.create(
                model="gpt-4o",
                response_format={"type": "json_object"},
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_prompt}
                ]
            )
            return json.loads(response.choices[0].message.content)
        except Exception as e:
            print(f"OpenAI analysis failed: {e}")
            return {"escalate": True, "problem_summary": "AI Analysis Failed", "recommended_steps": f"The AI model could not process the request. Error: {e}", "responsible_module": "INFRA/SRE"}

# ==============================================================================
# 5. MAIN ORCHESTRATOR
# ==============================================================================
def process_alert(alert_text: str, ai_assistant: AIAssistant, knowledge_base: str, case_log_df: pd.DataFrame, contacts: dict, notifier: ZapierWebhookNotifier):
    """Main function to process a single alert."""
    print(f"\n{'='*30}\nProcessing New Alert: {alert_text.split('|')[0].strip()}\n{'='*30}")
    
    identifiers = {}
    patterns = {
        'container_no': r'[A-Z]{4}[0-9]{7}', 'vessel_name': r'MV\s+[A-Za-z0-9\s]+',
        'message_ref': r'REF-IFT-[0-9]{4}', 'vessel_id': r'MV\s+[A-Z\s]+/[0-9A-Z]{3}',
        'booking_no': r'BK-[A-Z0-9]+'
    }
    for key, pattern in patterns.items():
        match = re.search(pattern, alert_text)
        if match: identifiers[key] = match.group(0).strip()
    print(f"Extracted Identifiers: {identifiers}")

    log_context = gather_log_context(identifiers)
    print("--- Log Context ---\n" + log_context + "\n-------------------")

    analysis = ai_assistant.analyze_issue(alert_text, log_context, knowledge_base, case_log_df)

    if not analysis:
        print("Could not get AI analysis. Manually escalating.")
        notifier.send_escalation_notification("jacky.chan@psa123.com", "Jacky", "Manual Escalation: AI System Failure", "AI analysis failed.", f"Alert: {alert_text}")
        return

    print(f"AI Decision: Escalate -> {analysis.get('escalate')}")

    if analysis.get("escalate"):
        module = analysis.get("responsible_module", "INFRA/SRE")
        contact_info = contacts.get(module, contacts["INFRA/SRE"])

        subject = f"AI Escalation: {analysis.get('problem_summary', 'Action Required')[:40]}"
        
        notifier.send_escalation_notification(
            contact_info['email'], contact_info['name'], subject,
            analysis.get('problem_summary', 'N/A'),
            analysis.get('recommended_steps', 'N/A')
        )
    else:
        print("AI determined the issue does not require escalation. Logging for review.")

if __name__ == "__main__":
    config = Config()
    notifier = ZapierWebhookNotifier(config)
    ai_assistant = AIAssistant(config)

    knowledge_base_text = load_knowledge_base()
    case_log_df = load_case_log_as_dataframe()
    escalation_contacts = load_escalation_contacts()

    if not knowledge_base_text or not escalation_contacts:
        print("Could not load knowledge base or contacts. Exiting.")
        exit()

    test_cases = [
        "RE: Email ALR-861600 | CMAU0000020 - Duplicate Container information received",
        "RE: Email ALR-861631 | VESSEL_ERR_4 - Customer reported unable to create vessel advice for MV Lion City 07 and hit error VESSEL_ERR_4.",
        "Alert: SMS INC-154599 | Issue: EDI message REF-IFT-0007 stuck in ERROR status.",
        "CRITICAL: Rate limiter throttled legitimate traffic after load surge. Hot keys on '/auth/token' created cache stampede; error budget burned for the window."
    ]

    for alert in test_cases:
        process_alert(alert, ai_assistant, knowledge_base_text, case_log_df, escalation_contacts, notifier)
