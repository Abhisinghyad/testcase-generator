# 🧪 Test Case Generator

A local web app that reads a **BRD** (`.pdf`, `.docx`, `.xlsx`, `.csv`) and uses an LLM to generate
**User Stories**, **Test Scenarios**, and **Test Cases**, exported as a **downloadable Excel**.
(The file is only offered for download — nothing is written to disk on the server.)

## Provider

Content is generated through the **OpenAI API**. Configure it in `.env`:

- `OPENAI_API_KEY` — your `sk-proj-...` (or `sk-...`) key
- `OPENAI_MODEL` — default `gpt-4o`
- `OPENAI_BASE_URL` — defaults to `https://api.openai.com/v1`
  (any OpenAI-compatible endpoint also works if you change this + key + model)

## Output

One Excel file with a sheet per selected item:

- **User Stories** — S.no, Module, Sub Module, User Story ID, Title, User Story, Acceptance Criteria, Priority
- **Test Scenarios** — S.no, Module, Sub Module, Scenario ID, Scenario Description, Type, Priority
- **Test Cases** — the full 21-column template (S.no … BugID). Your **Preconditions** and
  **Default Test Steps** are prepended to every test case; execution columns are left blank;
  Created On / Created By are auto-filled.

## Setup

1. Install [Node.js](https://nodejs.org) 18+ (v22 works).
2. `npm install`
3. Copy `.env.example` to `.env` and set `OPENAI_API_KEY` (and, if you like, `OPENAI_MODEL`).
4. `npm start`
5. Open **http://localhost:3000**.

## Using it

1. Click **Get Started**.
2. **Input Source** — *Upload File* (pick any file from your PC) or *File Path* (any full path on this PC).
3. **What to Generate** — tick any of User Stories / Scenarios / Test Cases.
4. If Test Cases is ticked, fill **Preconditions** and **Default Test Steps** (required).
5. **Generate** — click **Download** to save the Excel wherever you like.

## Quick test

A sample BRD is included at [`sample-brd/Sample-BRD.xlsx`](sample-brd/Sample-BRD.xlsx). Upload it,
tick all three items, add a precondition + a default step, and Generate.

## Notes

- Legacy `.doc` / `.xls` aren't supported — save as `.docx` / `.xlsx` (or `.pdf`) first.
- Scanned/image-only PDFs have no extractable text and won't work.
- File Path mode can read a file from any location on this PC.
