# AI Spreadsheet Agent

A simple command-line tool that allows an AI to interact with a Google Sheet using natural language.  The project uses OpenAI's `gpt-5-mini` model to parse user input and convert it into structured JSON commands that are executed against the Google Sheets API.  It's designed for experimentation and automation of common spreadsheet tasks.

---

## Features

- **Read, write, clear** ranges of cells
- **Create and delete sheets** within a spreadsheet
- **List all sheets** (tabs) in a spreadsheet
- Support for **multiple commands** in one request
- Optional **summary responses** when the user asks questions about the data
- Maintains a local log (`agent_log.json`) of all interactions
- Simple conversational CLI interface

All the AI behavior is defined in `instructions.txt`, which is loaded as the system prompt when the agent starts.

---

## Prerequisites

- Python 3.8 or later
- A Google Cloud project with the **Google Sheets API** enabled
- OAuth 2.0 credentials (`credentials.json`) downloaded from the Google Cloud Console
- An OpenAI API key

---

## Setup

1. **Clone or download** this repository.
2. Create a virtual environment and activate it (recommended):
   ```powershell
   python -m venv .venv
   .\.venv\Scripts\Activate.ps1   # PowerShell
   # or on Bash: source .venv/bin/activate
   ```
3. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
4. Copy the example environment file and fill in your keys:
   ```bash
   cp .env.example .env
   # then edit .env and set OPENAI_API_KEY and SPREADSHEET_ID
   ```
5. Place the downloaded `credentials.json` from Google in the project root.
   The first time you run the program it will open a browser to authorize and will
   save an OAuth token to `token.json` (this file is ignored by git).

---

## Running the Agent

Run the main script:

```bash
python main.py
```

You'll be greeted with a prompt. Type natural language commands or queries such as:

```
Read Sheet1 and explain it to me
Write "Hello" to C3
Create a sheet named inventory and add a header row
List all sheets
How many rows are in Sheet1?
```

Typing `--help` will display a summary of available commands; `quit` or `exit`
will terminate the session.

---

## Command Structure

The OpenAI model responds with JSON objects that look like this:

```json
{ "action": "READ", "range": "Sheet1!A1:A5" }
```

For multiple actions or summaries the agent may return arrays or include a
`"summary": true` flag. See `instructions.txt` for full details on the
structured language expected by the model.

---

## Logging

All user inputs, AI commands, and results are appended to `agent_log.json`.
This file can be inspected or deleted between runs.

---

Made this as a fun project to experiment the use of agentic AI.