import os
import os.path
from dotenv import load_dotenv
from openai import OpenAI
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from openai import APIConnectionError
import json
import re
import datetime
import time

load_dotenv()

# OpenAI
client = OpenAI()
MODEL_NAME = "gpt-5-mini"
LOG_FILE = "agent_log.json"

# Google Sheets
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
SPREADSHEET_ID = os.getenv("SPREADSHEET_ID")

# System prompt
prompt_path = os.path.join(os.path.dirname(__file__), "instructions.txt")
with open(prompt_path, encoding="utf-8") as f:
    SYSTEM_PROMPT = f.read()


def get_credentials():
    # Google atuhentication
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
     # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                "credentials.json", SCOPES
            )
            flow.redirect_uri = "http://localhost:8080"
            auth_url, _ = flow.authorization_url()
            print("Please go to this URL and authorize the application:")
            print(auth_url)
            code = input("Enter the authorization code: ")
            flow.fetch_token(code=code)
            creds = flow.credentials

        # Save the credentials for the next run
        with open("token.json", "w") as token:
            token.write(creds.to_json())

    return creds

def log_action(user_input, ai_command, result):
    log_entry = {
        "timestamp": datetime.datetime.now().isoformat(),
        "user_input": user_input,
        "ai_command": ai_command,
        "result": str(result),
    }

    if os.path.exists(LOG_FILE):
        try:
            with open(LOG_FILE, "r") as f:
                log = json.load(f)
        except json.JSONDecodeError:
            log = []
    else:
        log = []

    log.append(log_entry)

    with open(LOG_FILE, "w") as f:
        json.dump(log, f, indent=2)

    print("Action logged.")

def parse_json_response(response_text):
    # Extract JSON from AI response
    if not response_text or not response_text.strip():
        print("Error: AI response is empty.")
        return None

    try:
        parsed = json.loads(response_text)
        if isinstance(parsed, dict):
            return [parsed]
        elif isinstance(parsed, list):
            return parsed
    except json.JSONDecodeError:
        pass

    match = re.search(r'\[.*\]', response_text, re.DOTALL)
    if match:
        try:
            parsed = json.loads(match.group())
            if isinstance(parsed, list):
                return parsed
        except json.JSONDecodeError:
            pass

    match = re.search(r'\{.*\}', response_text, re.DOTALL)
    if match:
        try:
            parsed = json.loads(match.group())
            if isinstance(parsed, dict):
                return [parsed]
        except json.JSONDecodeError:
            pass

    print("Error: Could not parse JSON from AI response")
    return None

def execute_single_command(service, command):
    action = command.get("action")
    range_ = command.get("range")

    if not action:
        print("Invalid command: Missing 'action'.")
        return None

    if action in ("READ", "WRITE") and not range_:
        print("Invalid command: Missing range for READ/WRITE action.")
        return None

    sheet = service.spreadsheets()

    # ACTIONS
    try:
        if action == "READ":
            result = (
                sheet.values()
                .get(spreadsheetId=SPREADSHEET_ID, range=range_)
                .execute()
            )
            values = result.get("values", [])
            print(f"Read {len(values)} rows:")
            for row in values:
                print(row)
            return values

        elif action == "WRITE":
            values = command.get("values")
            if not values:
                print("Error: Missing 'values' for WRITE action.")
                return None

            body = {"values": values}
            result = (
                sheet.values()
                .update(
                    spreadsheetId=SPREADSHEET_ID,
                    range=range_,
                    valueInputOption="RAW",
                    body=body,
                )
                .execute()
            )
            print(f"Updated {result.get('updatedCells')} cells.")
            return result

        elif action == "CREATESHEET":
            values = command.get("values")
            if not values or not values[0] or not values[0][0]:
                print("Error: Missing or invalid 'values' for CREATESHEET action.")
                return None

            sheet_title = values[0][0]

            request_body = {
                "requests": [
                    {
                        "addSheet": {
                            "properties": {
                                "title": sheet_title
                            }
                        }
                    }
                ]
            }
            result = (
                sheet.batchUpdate(
                    spreadsheetId=SPREADSHEET_ID,
                    body=request_body,
                )
                .execute()
            )
            print(f"Created new sheet '{sheet_title}'.")
            return result

        elif action == "LIST":
            result = sheet.get(spreadsheetId=SPREADSHEET_ID).execute()
            sheets = result.get("sheets", [])

            sheet_list = []
            print("\nSheets in spreadsheet:")
            for s in sheets:
                props = s["properties"]
                title = props["title"]
                rows = props["gridProperties"]["rowCount"]
                cols = props["gridProperties"]["columnCount"]
                sheet_list.append(title)
                print(f"  • {title} ({rows} rows X {cols} cols)")

            return sheet_list
        
        elif action == "CLEAR":
            result = (
                sheet.values()
                .clear(
                    spreadsheetId=SPREADSHEET_ID,
                    range=range_,
                )
                .execute()
            )
            print(f"Cleared range {range_}.")
            return result

        elif action == "DELETESHEET":
            values = command.get("values")
            if not values or not values[0] or not values[0][0]:
                print("Error: Missing sheet name for DELETESHEET.")
                return None

            sheet_name = values[0][0]

            # Finn sheet ID fra navnet
            spreadsheet = sheet.get(spreadsheetId=SPREADSHEET_ID).execute()
            sheet_id = None
            for s in spreadsheet.get("sheets", []):
                if s["properties"]["title"] == sheet_name:
                    sheet_id = s["properties"]["sheetId"]
                    break

            if sheet_id is None:
                print(f"Error: Sheet '{sheet_name}' not found.")
                return None

            request_body = {
                "requests": [
                    {"deleteSheet": {"sheetId": sheet_id}}
                ]
            }
            result = (
                sheet.batchUpdate(
                    spreadsheetId=SPREADSHEET_ID,
                    body=request_body,
                )
                .execute()
            )
            print(f"Deleted sheet '{sheet_name}'.")
            return result

        else:
            print(f"Unknown action: {action}")
            return None

    except HttpError as err:
        print(f"Google Sheets API error: {err}")
        return None

def execute_commands(service, response_text):
    # Parse AI response and execute one or more commands.
    commands = parse_json_response(response_text)
    if not commands:
        return None

    results = []
    total = len(commands)

    print(f"\nExecuting {total} command(s)...\n")

    for i, command in enumerate(commands, 1):
        print(f"--- Command {i}/{total}: {command.get('action')} ---")
        result = execute_single_command(service, command)
        results.append(result)

        # if command is not READ and result is None, stop execution
        if result is None and command.get("action") != "READ":
            print(f"Command {i} failed. Stopping execution.")
            break

        if i < total:
            time.sleep(1)

    print(f"\nCompleted {len(results)}/{total} commands.")
    return results

def main():
    creds = get_credentials()
    service = build("sheets", "v4", credentials=creds)

    print("Welcome to the AI Spreadsheet Agent!")
    print("Type 'quit' or 'exit' to stop. '--help' for available commands.\n")

    conversation_history = []

    while True:
        user_input = input("\nYou: ")
        if user_input.lower() in ["quit", "exit"]:
            break

        if user_input == "--help":
            print("\nAvailable commands:")
            print("  - READ range: Read data from a range (e.g. 'READ Sheet1!A1:B2')")
            print("  - WRITE range values: Write values to a range (e.g. 'WRITE Sheet1!A1 [[\"Hello\", \"World\"]]'")
            print("  - CREATESHEET sheetname: Create a new sheet (e.g. 'CREATESHEET NewSheet')")
            print("  - LIST: List all sheets in the spreadsheet")
            print("  - CLEAR range: Clear values in a range (e.g. 'CLEAR Sheet1!A1:B2')")
            print("  - DELETESHEET sheetname: Delete a sheet (e.g. 'DELETESHEET OldSheet')")
            continue

        conversation_history.append({
            "role": "user",
            "content": user_input
        })

        try:
            response = client.responses.create(
                model=MODEL_NAME,
                instructions=SYSTEM_PROMPT,
                input=conversation_history,
            )

            ai_output = response.output_text

            conversation_history.append({
                "role": "assistant",
                "content": ai_output
            })

            print(f"AI command: {ai_output}")
            results = execute_commands(service, ai_output)
            log_action(user_input, ai_output, str(results))

            if results:
                result_summary = str(results)
                conversation_history.append({
                    "role": "user",
                    "content": f"[SYSTEM] Command executed. Result: {result_summary}"
                })

            commands = parse_json_response(ai_output)
            wants_summary = False
            if commands:
                last_command = commands[-1]
                wants_summary = last_command.get("summary", False)

            # Summary
            if wants_summary:
                list_results = [r for r in (results or []) if isinstance(r, list)]
                if list_results:
                    summary = client.responses.create(
                        model=MODEL_NAME,
                        instructions="You are a helpful assistant. Answer the user's question based on the data. Be concise and short. If you don't know the answer, say you don't know. Only use the provided data to answer. DO NOT ask any follow up questions.",
                        input=f"User asked: {user_input}\n\nData:\n{list_results}",
                    )
                    print(f"\nAI: {summary.output_text}")

                    conversation_history.append({
                        "role": "assistant",
                        "content": summary.output_text
                    })

            MAX_HISTORY = 10
            if len(conversation_history) > MAX_HISTORY:
                conversation_history = conversation_history[-MAX_HISTORY:]

        except APIConnectionError:
            print("Failed to connect to OpenAI API.")
            log_action(user_input, "N/A", "APIConnectionError")
        except Exception as e:
            print(f"Error: {e}")
            log_action(user_input, "N/A", str(e))


if __name__ == "__main__":
    main()