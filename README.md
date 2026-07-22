# 🧪 Test Case Creation Portal

> Turn requirements into a **professional, ready-to-use Excel test-case workbook** — automatically.

A local web app that generates QA **User Stories, Test Scenarios, and Test Cases** from your
requirements using an LLM, and exports a polished, multi-sheet Excel workbook with **traceability,
live coverage dashboards, and a quality scorecard**.

![status](https://img.shields.io/badge/status-v1-blue)
![node](https://img.shields.io/badge/node-18%2B-brightgreen)
![license](https://img.shields.io/badge/license-MIT-lightgrey)

---

## ✨ Features

- **Two ways to create test cases**
  - **From Scenarios** — type plain-language scenarios and AI expands them into complete test cases.
  - **From Documents** — upload a BRD (`.pdf`, `.docx`, `.xlsx`, `.csv`) and AI produces User Stories,
    Test Scenarios and Test Cases.
- **Professional Excel workbook** (7 sheets, correct tab order):
  1. **Project Summary** — cover page with document control & workbook contents
  2. **Test Execution Summary** — live `COUNTIF` execution metrics, status breakdown & sign-off
  3. **Testcases** — 24-column template with dropdown validation & live status colour-fills
  4. **Dashboard** — live counts by Type / Severity / Sub-Module with in-cell bars
  5. **Traceability Matrix (RTM)** — requirement → test case coverage with live counts
  6. **Solution Doc Coverage** — user-story / acceptance-criterion coverage
  7. **Quality Scorecard** — per-case quality scoring with a suite average
- **House-style test steps** — numbered, imperative, one action per line, with concrete quoted test data.
- **Traceability & gap analysis** — every requirement mapped to covering test cases; gaps flagged.
- **Provider-agnostic** — works with **OpenRouter**, **OpenAI**, or any OpenAI-compatible endpoint.
- **Attractive, dynamic UI** — animated gradient background, light/dark theme, responsive.
- **Download-only** — nothing is written to the server; you download the workbook directly.

---

## 🛠 Tech stack

| Layer | Tech |
|---|---|
| Server | Node.js + Express |
| Excel | ExcelJS (formulas, data-validation, conditional formatting) |
| Document parsing | `pdf-parse` (PDF), `mammoth` (DOCX), ExcelJS (XLSX) |
| AI | OpenAI-compatible Chat Completions API (OpenRouter / OpenAI / …) |
| Frontend | Vanilla HTML/CSS/JS (no build step) |

---

## 🚀 Getting started

### Prerequisites
- [Node.js](https://nodejs.org) 18+ (v22 recommended)
- An API key for an OpenAI-compatible provider (e.g. [OpenRouter](https://openrouter.ai))

### Install
```bash
npm install
```

### Configure
Copy `.env.example` to `.env` and set your key:
```env
OPENROUTER_API_KEY=sk-or-v1-your-key
OPENROUTER_BASE_URL=https://openrouter.ai/api/v1
OPENROUTER_MODEL=openai/gpt-4o-mini
OPENROUTER_MAX_TOKENS=4000
```
> To use OpenAI directly instead, set `OPENAI_API_KEY`, `OPENAI_BASE_URL=https://api.openai.com/v1`,
> and `OPENAI_MODEL=gpt-4o`.

### Run
```bash
npm start
```
Open **http://localhost:3000**.

---

## 📖 Usage

1. On the landing page choose **Create Manually** (scenarios) or **Create using Documents** (BRD upload).
2. **Scenarios:** type one scenario per line (or click *Example Scenarios*) → **Generate**.
   **Documents:** pick the file, tick what to generate, optionally add a **Description** to scope the AI.
3. Click **Download** to save the `.xlsx`.

Live formulas (execution metrics, dashboards, RTM counts, scorecard) recalculate in **Microsoft Excel**
when the workbook is opened and as you fill in the Execution Status column.

---

## ⚙️ Configuration reference

| Variable | Purpose | Default |
|---|---|---|
| `OPENROUTER_API_KEY` / `OPENAI_API_KEY` | Provider API key | — (required) |
| `OPENROUTER_BASE_URL` / `OPENAI_BASE_URL` | API base URL | OpenRouter / OpenAI |
| `OPENROUTER_MODEL` / `OPENAI_MODEL` | Model slug | `openai/gpt-4o-mini` / `gpt-4o` |
| `OPENROUTER_MAX_TOKENS` / `AI_MAX_TOKENS` | Max output tokens per request | `8000` |
| `PORT` | Server port | `3000` |

---

## ☁️ Deploy to the public internet (Render — free)

> ⚠️ **This app spends your API key.** Always set `APP_PASSWORD` before hosting publicly, or anyone
> who finds the URL can run generations on your account.

1. Push this repo to GitHub (already done if you're reading this there).
2. Go to **https://render.com** → sign up (free) → **New +** → **Blueprint**.
3. Connect this repository — Render reads [`render.yaml`](render.yaml) automatically.
4. In the service's **Environment** tab, set the two secrets:
   - `OPENROUTER_API_KEY` = your `sk-or-v1-…` key
   - `APP_PASSWORD` = a password of your choice
5. Click **Deploy**. In ~2 minutes you get a public URL like
   `https://testcase-generator.onrender.com`.

Alternatively, any Docker host (Railway, Fly.io, Google Cloud Run) can build the included
[`Dockerfile`](Dockerfile) — set the same environment variables there.

> Render's **free** web service sleeps after ~15 min of inactivity and takes a few seconds to wake on
> the next request. Upgrade the plan for always-on.

## 📝 Notes & limitations

- Legacy `.doc` / `.xls` are not supported — save as `.docx` / `.xlsx` (or `.pdf`) first.
- Scanned/image-only PDFs have no extractable text and won't work.
- The app currently accepts a single requirements document. Separate BRD + Solution-Document upload
  (with a conflicts sheet) is on the roadmap.
- Richer output uses more tokens per run — keep the scope tight on low-balance accounts.

---

## 📄 License

MIT
