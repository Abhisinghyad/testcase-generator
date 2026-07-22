import 'dotenv/config';
import express from 'express';
import multer from 'multer';
import path from 'node:path';
import fs from 'node:fs';
import { fileURLToPath } from 'node:url';
import mammoth from 'mammoth';
import ExcelJS from 'exceljs';
import pdfParse from 'pdf-parse/lib/pdf-parse.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));

// ---------------------------------------------------------------------------
// Configuration (all overridable via .env)
// ---------------------------------------------------------------------------
const AI_API_KEY = process.env.OPENAI_API_KEY || process.env.OPENROUTER_API_KEY || '';
const AI_BASE_URL = process.env.OPENAI_BASE_URL || process.env.OPENROUTER_BASE_URL || 'https://api.openai.com/v1';
const AI_MODEL = process.env.OPENAI_MODEL || process.env.OPENROUTER_MODEL || 'gpt-4o';
const MAX_OUTPUT_TOKENS = Number(process.env.AI_MAX_TOKENS || process.env.OPENROUTER_MAX_TOKENS || 8000);

const app = express();
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 50 * 1024 * 1024 } });
app.use(express.json({ limit: '4mb' }));
app.use(express.static(path.join(__dirname, 'public')));

// ---------------------------------------------------------------------------
// BRD text extraction
// ---------------------------------------------------------------------------
async function extractText(originalname, buf) {
  const ext = path.extname(originalname).toLowerCase();
  if (ext === '.pdf') return (await pdfParse(buf)).text;
  if (ext === '.docx') return (await mammoth.extractRawText({ buffer: buf })).value;
  if (ext === '.csv' || ext === '.txt') return buf.toString('utf8');
  if (ext === '.xlsx') {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buf);
    const parts = [];
    wb.eachSheet((sheet) => {
      parts.push(`--- Sheet: ${sheet.name} ---`);
      sheet.eachRow({ includeEmpty: false }, (row) => {
        const cells = row.values.slice(1).map((v) => {
          if (v == null) return '';
          if (typeof v === 'object') return v.text ?? v.result ?? v.richText?.map((r) => r.text).join('') ?? '';
          return String(v);
        });
        parts.push(cells.join(' | '));
      });
    });
    return parts.join('\n');
  }
  if (ext === '.doc' || ext === '.xls') {
    throw new Error(`Legacy ${ext} format is not supported. Please save as .docx / .xlsx (or .pdf) and try again.`);
  }
  throw new Error(`Unsupported file type "${ext}". Use .pdf, .docx, .xlsx or .csv.`);
}

// ---------------------------------------------------------------------------
// LLM call (OpenAI-compatible → OpenRouter)
// ---------------------------------------------------------------------------
async function llmChat(system, user) {
  if (!AI_API_KEY) throw new Error('No API key configured. Set OPENAI_API_KEY in .env.');
  const url = AI_BASE_URL.replace(/\/+$/, '') + '/chat/completions';
  const messages = [{ role: 'system', content: system }, { role: 'user', content: user }];

  // Try a few parameter shapes so we work across models that differ on
  // max_tokens vs max_completion_tokens and response_format support.
  const attempts = [
    { max_tokens: MAX_OUTPUT_TOKENS, response_format: { type: 'json_object' } },
    { max_completion_tokens: MAX_OUTPUT_TOKENS, response_format: { type: 'json_object' } },
    { max_tokens: MAX_OUTPUT_TOKENS },
    { max_completion_tokens: MAX_OUTPUT_TOKENS },
  ];

  let lastErr = 'Unknown error';
  for (const extra of attempts) {
    let res;
    try {
      res = await fetch(url, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          Authorization: `Bearer ${AI_API_KEY}`,
        },
        body: JSON.stringify({ model: AI_MODEL, messages, ...extra }),
      });
    } catch (e) {
      throw new Error(`Could not reach ${url}: ${e.message}`);
    }
    if (res.ok) {
      const data = await res.json();
      return data.choices?.[0]?.message?.content ?? '';
    }
    lastErr = await res.text().catch(() => res.statusText);
    if (res.status !== 400) break; // only param problems are worth retrying
  }
  throw new Error(`AI provider error: ${lastErr}`);
}

function parseJsonObject(text) {
  let s = (text || '').trim();
  const fence = s.match(/```(?:json)?\s*([\s\S]*?)```/i);
  if (fence) s = fence[1].trim();
  if (!/^[[{]/.test(s)) {
    const a = s.indexOf('{');
    const b = s.lastIndexOf('}');
    if (a !== -1 && b !== -1 && b > a) s = s.slice(a, b + 1);
  }
  try {
    return JSON.parse(s);
  } catch {
    throw new Error('Could not parse the AI response as JSON. Try again or use a smaller BRD.');
  }
}

const BASE_SYSTEM =
  'You are a senior QA engineer analysing a Business Requirements Document (BRD). ' +
  'Be specific, thorough, and grounded strictly in the BRD content. Do not invent requirements. ' +
  'Respond with ONLY a valid JSON object — no markdown fences, no commentary.';

const brdBlock = (brdText) => '=== BRD START ===\n' + brdText + '\n=== BRD END ===';

// Build the user prompt, folding in the user's free-text description/conditions.
function userPrompt(instruction, brdText, description) {
  const d = (description || '').trim();
  const extra = d
    ? `\n\nAdditional instructions and conditions from the user — treat these as high-priority requirements and follow them closely when generating:\n${d}`
    : '';
  return instruction + extra + '\n\n' + brdBlock(brdText);
}

// --- User Stories ---
async function generateUserStories(brdText, description) {
  const system = BASE_SYSTEM +
    '\nReturn shape: {"user_stories":[{"module":"","sub_module":"","title":"",' +
    '"story":"As a <role>, I want <goal>, so that <benefit>","acceptance_criteria":"bullet points, one per line","priority":"High|Medium|Low"}]}';
  const data = parseJsonObject(await llmChat(system, userPrompt('Generate user stories from this BRD.', brdText, description)));
  const arr = Array.isArray(data) ? data : (data.user_stories || data.userStories || []);
  return arr.map((s) => ({
    module: s.module ?? '', sub_module: s.sub_module ?? s.subModule ?? '',
    title: s.title ?? '', story: s.story ?? s.user_story ?? '',
    acceptance_criteria: s.acceptance_criteria ?? s.acceptanceCriteria ?? '',
    priority: s.priority ?? '',
  }));
}

// --- Test Scenarios ---
async function generateScenarios(brdText, description) {
  const system = BASE_SYSTEM +
    '\nReturn shape: {"scenarios":[{"module":"","sub_module":"","description":"a testable scenario",' +
    '"type":"Positive|Negative|Boundary|Security|Integration|UI","priority":"High|Medium|Low"}]}';
  const data = parseJsonObject(await llmChat(system, userPrompt('Generate test scenarios from this BRD.', brdText, description)));
  const arr = Array.isArray(data) ? data : (data.scenarios || data.test_scenarios || []);
  return arr.map((s) => ({
    module: s.module ?? '', sub_module: s.sub_module ?? s.subModule ?? '',
    description: s.description ?? s.scenario ?? '', type: s.type ?? '', priority: s.priority ?? '',
  }));
}

// --- Test Cases ---
async function generateTestCases(brdText, description) {
  const system = BASE_SYSTEM + `

You are a senior QA / test analyst producing a detailed, TRACEABLE test-case suite from the BRD.

Before writing: read the in-scope requirements fully; split any compound requirement ("the system must validate X and Y") into ATOMIC requirements; note every business rule, validation, field, status transition, role and integration; identify the actors/roles and the status lifecycle. If something is ambiguous, state the assumption inside the relevant test case rather than guessing silently.

Coverage — for EACH atomic requirement include: at least one POSITIVE (happy-path) case; NEGATIVE cases for every error/invalid path; BOUNDARY cases for every limit (dates, amounts, field max length, whole-number rules); BUSINESS-RULE VARIATIONS (e.g. NBO vs Non-NBO, per entity); INTEGRATION cases for external systems (success, timeout/failure, retry, idempotency / no duplicate postings); ROLE/ACCESS cases (authorised vs unauthorised); and STATUS-TRANSITION cases for each state change in the lifecycle. Do not leave any in-scope requirement without at least one test.

Quality bar (every case): traceable (has a BRD Req ID) · atomic (tests one thing) · clear numbered steps runnable by any tester with no prior context · deterministic expected result (exact status/message/value/field state — never "works correctly" or "as per business rule") · concrete test data (real RRNs, amounts, GL/account numbers, dates — not "a valid value") · correct Type & Severity · independent (own pre-requisites, no reliance on another case's leftover state).

Field rules:
- module: the product/module name, constant across the suite.
- sub_module: the requirement area / workflow phase (e.g. Creation, Authorisation, Rework & Resubmission, CBS Integration, Good Faith – Processing, View).
- type: one of Functional | Negative | Validation | UI | Integration.
- severity: Critical (end-to-end money-movement / posting-to-core paths) | High (status changes, amount/date/duplicate validations, queue routing, notifications) | Medium (field/UI/narration/export) | Low (minor display).
- description: one line starting with "Verify ..." stating exactly what is checked.
- pre_requisites: role assigned + system state + concrete data needed to run it; "None" if truly none.
- roles: the actor performing the test.
- test_steps: written EXACTLY per the "TEST STEPS — HOUSE STYLE" rules below.
- test_data: ALWAYS provide concrete example input values for this test — invent realistic sample values when the scenario does not specify them (e.g. username "john.doe@test.com", password "P@ssw0rd1", RRN 123456789012, amount 500.00, date 2026-01-15). Use "—" only when the test genuinely needs no input data.
- expected_result: ONE unambiguous, verifiable outcome (exact status, message, value or field state).
- brd_req_id: the BRD requirement ID(s) this case verifies; if the BRD has no explicit IDs, cite the section number/heading.
- solution_doc_ref: the user story / acceptance-criterion / section reference in the Solution Document this case verifies; "—" if there is no Solution Document or it does not apply.
- sol_doc_alignment: one of "Aligned" | "New (gap)" | "Review – not in Solution Doc". Use "Aligned" when only a single requirements document is provided.

TEST STEPS — HOUSE STYLE (apply to every test_steps value, follow precisely):
1. A NUMBERED list: "1. ", "2. ", "3. " … one action per line.
2. Each step is a short IMPERATIVE instruction starting with a verb (Login, Navigate, Select, Enter, Click, Observe, Verify, Wait…).
3. One user action per step — never combine two actions in one line.
4. Begin with entry/navigation (Login → Navigate to the screen → open the page) unless the pre-requisite already places the user there.
5. Use CONCRETE test data in quotes, not placeholders: RRN → a 12-digit value e.g. '123456789012'; Amounts → e.g. Refund Amount '100'; Dropdowns → the exact option e.g. Cardholder Type 'NBO' / 'Non-NBO'; Dates, GL/account numbers and comments → realistic sample values.
6. For negative/validation cases, enter the specific invalid value (e.g. Invalid RRN '000000000000', decimal Refund Amount '100.50', transaction older than 180 days).
7. End with the action that triggers the outcome, then an "Observe/Verify" step if the case checks a displayed value, status, message, or field state.
8. Keep steps atomic, ordered, and runnable by any tester with no prior context.
9. Do NOT restate the expected result inside the steps — steps are actions only.

Scope: if the user's additional instructions name specific sections/requirements or product context, cover ONLY those; otherwise cover the whole in-scope BRD.

Also output the atomic requirement list you extracted (STEP 1) as "requirements" — each with a stable "req_id" (reuse the BRD's own IDs or section numbers) and a short "requirement" text. Use those SAME ids in each test case's brd_req_id so every requirement is traceable and any gap is visible.

Prioritise the highest-value cases and keep wording tight. IMPORTANT: return COMPLETE, valid JSON — never stop mid-object.

Return shape: {"test_cases":[{"module":"","sub_module":"","type":"","severity":"Critical|High|Medium|Low","description":"","pre_requisites":"","roles":"","test_steps":"1. ...\\n2. ...","test_data":"","expected_result":"","brd_req_id":"","solution_doc_ref":"","sol_doc_alignment":"Aligned"}],"requirements":[{"req_id":"","requirement":"short text"}]}`;
  const data = parseJsonObject(await llmChat(system,
    userPrompt('Generate a thorough, prioritised, traceable set of test cases from this BRD, following every field rule, and the atomic requirement list.', brdText, description)));
  const arr = Array.isArray(data) ? data : (data.test_cases || data.testCases || []);
  const cases = arr.map((tc) => ({
    module: tc.module ?? '', sub_module: tc.sub_module ?? tc.subModule ?? '',
    type: tc.type ?? '', severity: tc.severity ?? '', description: tc.description ?? '',
    pre_requisites: tc.pre_requisites ?? tc.preRequisites ?? tc.preconditions ?? '',
    roles: tc.roles ?? '', test_steps: tc.test_steps ?? tc.testSteps ?? tc.steps ?? '',
    test_data: tc.test_data ?? tc.testData ?? '',
    expected_result: tc.expected_result ?? tc.expectedResult ?? '',
    brd_req_id: tc.brd_req_id ?? tc.brdReqId ?? tc.req_id ?? tc.requirement_id ?? '',
    solution_doc_ref: tc.solution_doc_ref ?? tc.solutionDocRef ?? tc.sol_doc_ref ?? '',
    sol_doc_alignment: tc.sol_doc_alignment ?? tc.solDocAlignment ?? tc.alignment ?? 'Aligned',
  }));
  const requirements = (data.requirements || data.requirement_list || []).map((r) => ({
    req_id: r.req_id ?? r.id ?? r.brd_req_id ?? '',
    requirement: r.requirement ?? r.text ?? r.description ?? '',
  }));
  return { cases, requirements };
}

// ---------------------------------------------------------------------------
// Excel building
// ---------------------------------------------------------------------------
function todayStr() {
  const d = new Date();
  return `${String(d.getDate()).padStart(2, '0')}-${String(d.getMonth() + 1).padStart(2, '0')}-${d.getFullYear()}`;
}
function prefix(base, extra) {
  base = (base || '').trim(); extra = (extra || '').trim();
  if (!base) return extra;
  if (!extra) return base;
  return `${base}\n${extra}`;
}
function styleHeader(row) {
  row.font = { bold: true, color: { argb: 'FFFFFFFF' } };
  row.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
  row.eachCell((c) => {
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2F5496' } };
    c.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
  });
}
function applyWidths(ws, widths, defWidth = 16) {
  ws.columns.forEach((col) => {
    col.width = widths[col.key] || defWidth;
    col.alignment = { vertical: 'top', wrapText: true };
  });
  styleHeader(ws.getRow(1));
}

const NAVY = 'FF2F5496';
function colLetter(n) { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }

const TC_SHEET = 'Testcases';
const TC_COLUMNS = [
  'S.no', 'Module', 'Sub Module', 'Test Case ID', 'Type', 'Severity',
  'Test Case Description', 'Pre-Requisites', 'Roles', 'Test Steps', 'Expected Result',
  'Actual Result', 'Execution Status', 'Date of Execution', 'Executed By', 'Created On',
  'Created By', 'Modified On', 'Modified By', 'Remarks', 'BugID',
  'BRD Req ID', 'Solution Doc Ref', 'Sol.Doc Alignment',
];
function tcCol(name) { return colLetter(TC_COLUMNS.indexOf(name) + 1); }
// Full-column formula ranges over the Testcases sheet (auto-update as rows change).
function tcRange(name) { return `${TC_SHEET}!$${tcCol(name)}$2:$${tcCol(name)}$1000`; }

function titleBand(ws, span, title, subtitle) {
  ws.mergeCells(1, 1, 1, span);
  const t = ws.getCell(1, 1);
  t.value = title;
  t.font = { bold: true, size: 16, color: { argb: 'FFFFFFFF' } };
  t.alignment = { vertical: 'middle', horizontal: 'center' };
  t.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: NAVY } };
  ws.getRow(1).height = 30;
  if (subtitle) {
    ws.mergeCells(2, 1, 2, span);
    const s = ws.getCell(2, 1);
    s.value = subtitle;
    s.alignment = { horizontal: 'center' };
    s.font = { italic: true, size: 10, color: { argb: 'FF7F7F7F' } };
  }
}
function sectionHead(ws, r, span, label) {
  ws.mergeCells(r, 1, r, span);
  const c = ws.getCell(r, 1);
  c.value = label; c.font = { bold: true, size: 12, color: { argb: NAVY } };
}
function tableHead(ws, r, headers) {
  headers.forEach((h, i) => {
    const c = ws.getRow(r).getCell(i + 1);
    c.value = h; c.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    c.alignment = { horizontal: 'center' };
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: NAVY } };
  });
}
function addListValidation(ws, name, lastRow, list) {
  const col = tcCol(name);
  const formulae = ['"' + list.join(',') + '"'];
  for (let r = 2; r <= lastRow; r++) {
    ws.getCell(`${col}${r}`).dataValidation = { type: 'list', allowBlank: true, showErrorMessage: true, formulae };
  }
}

function buildMeta(cases, description, sourceName) {
  const module = (cases.find((c) => c.module) || {}).module || 'Test Suite';
  const subs = [...new Set(cases.map((c) => (c.sub_module || '').trim()).filter(Boolean))];
  return {
    project: module,
    referenceDoc: sourceName || 'BRD / Solution Document',
    scope: (description || '').trim() || 'As per the attached requirements document',
    module,
    qaOwner: '—', deliveryPartner: '—', preparedFor: '—',
    date: todayStr(), version: '1.0', status: 'Draft',
    overview: `This workbook contains the QA test-case suite for ${module}. It was generated from the supplied requirements and covers functional, negative, validation, integration and role-based scenarios with full requirement traceability, live execution metrics and quality scoring.`,
    inScopeAreas: subs.length ? subs : [module],
  };
}

// ---- 3. Testcases sheet ----
function addTestCasesSheet(wb, testCases, { basicPreReq, basicSteps, createdBy }) {
  const ws = wb.addWorksheet(TC_SHEET, { views: [{ state: 'frozen', ySplit: 1 }] });
  ws.columns = TC_COLUMNS.map((h) => ({ header: h, key: h }));
  const created = todayStr();
  testCases.forEach((tc, i) => {
    ws.addRow({
      'S.no': i + 1, 'Module': tc.module, 'Sub Module': tc.sub_module,
      'Test Case ID': `TC_${String(i + 1).padStart(4, '0')}`, 'Type': tc.type, 'Severity': tc.severity,
      'Test Case Description': tc.description, 'Pre-Requisites': prefix(basicPreReq, tc.pre_requisites),
      'Roles': tc.roles, 'Test Steps': prefix(basicSteps, tc.test_steps), 'Expected Result': tc.expected_result,
      'Actual Result': '', 'Execution Status': 'Not Executed', 'Date of Execution': '', 'Executed By': '',
      'Created On': created, 'Created By': createdBy || 'AI Generated',
      'Modified On': '', 'Modified By': '', 'Remarks': '', 'BugID': '',
      'BRD Req ID': tc.brd_req_id || '', 'Solution Doc Ref': tc.solution_doc_ref || '—',
      'Sol.Doc Alignment': tc.sol_doc_alignment || 'Aligned',
    });
  });
  applyWidths(ws, {
    'S.no': 6, 'Module': 18, 'Sub Module': 18, 'Test Case ID': 14, 'Type': 14, 'Severity': 10,
    'Test Case Description': 40, 'Pre-Requisites': 30, 'Roles': 16, 'Test Steps': 46, 'Expected Result': 40,
    'Actual Result': 20, 'Execution Status': 16, 'Date of Execution': 16, 'Executed By': 14,
    'Created On': 14, 'Created By': 14, 'Modified On': 14, 'Modified By': 14, 'Remarks': 20, 'BugID': 12,
    'BRD Req ID': 16, 'Solution Doc Ref': 18, 'Sol.Doc Alignment': 22,
  });
  // Dropdown data-validation
  const lastVal = testCases.length + 30;
  addListValidation(ws, 'Type', lastVal, ['Functional', 'Negative', 'Validation', 'UI', 'Integration']);
  addListValidation(ws, 'Severity', lastVal, ['Critical', 'High', 'Medium', 'Low']);
  addListValidation(ws, 'Execution Status', lastVal, ['Passed', 'Failed', 'Blocked', 'In Progress', 'Not Executed']);
  addListValidation(ws, 'Sol.Doc Alignment', lastVal, ['Aligned', 'New (gap)', 'Review – not in Solution Doc']);
  // Status colour fills (live) — colour reserved for status only; body text stays black.
  const stCol = tcCol('Execution Status');
  const fills = { Passed: 'FFE2EFDA', Failed: 'FFFCE4E4', Blocked: 'FFFFF2CC', 'In Progress': 'FFDEEBF7', 'Not Executed': 'FFF2F2F2' };
  ws.addConditionalFormatting({
    ref: `${stCol}2:${stCol}${testCases.length + 1}`,
    rules: Object.entries(fills).map(([text, argb], idx) => ({
      type: 'containsText', operator: 'containsText', text, priority: idx + 1,
      style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb } } },
    })),
  });
  return ws;
}

// ---- 1. Project Summary (cover) ----
function addProjectSummarySheet(wb, meta) {
  const ws = wb.addWorksheet('Project Summary');
  ws.getColumn(1).width = 28; ws.getColumn(2).width = 78;
  titleBand(ws, 2, 'TEST CASE WORKBOOK', meta.module);
  let r = 4;
  sectionHead(ws, r, 2, 'Document Control'); r++;
  const kv = (k, v) => { const rr = ws.getRow(r); rr.getCell(1).value = k; rr.getCell(1).font = { bold: true }; rr.getCell(2).value = v; rr.getCell(2).alignment = { wrapText: true }; r++; };
  kv('Project', meta.project);
  kv('Reference Document(s)', meta.referenceDoc);
  kv('Scope', meta.scope);
  kv('Product / Module', meta.module);
  kv('QA Owner', meta.qaOwner);
  kv('Delivery Partner', meta.deliveryPartner);
  kv('Prepared For', meta.preparedFor);
  kv('Date', meta.date);
  kv('Version', meta.version);
  kv('Status', meta.status);
  r++;
  sectionHead(ws, r, 2, 'Project Overview'); r++;
  ws.mergeCells(r, 1, r, 2);
  const ov = ws.getCell(r, 1); ov.value = meta.overview; ov.alignment = { wrapText: true, vertical: 'top' }; ws.getRow(r).height = 64; r += 2;
  sectionHead(ws, r, 2, 'In-Scope Functional Areas'); r++;
  meta.inScopeAreas.forEach((a) => { ws.mergeCells(r, 1, r, 2); ws.getCell(r, 1).value = '•  ' + a; r++; });
  r++;
  sectionHead(ws, r, 2, 'Workbook Contents'); r++;
  [['Test Execution Summary', 'Live execution metrics & sign-off'],
   ['Testcases', 'The full test-case suite'],
   ['Dashboard', 'Counts by Sub-Module / Type / Severity'],
   ['Traceability Matrix (RTM)', 'Requirement → test case coverage'],
   ['Solution Doc Coverage', 'User story / acceptance-criterion coverage'],
   ['Quality Scorecard', 'Per-case quality scoring']].forEach(([n, d]) => {
    const rr = ws.getRow(r); rr.getCell(1).value = n; rr.getCell(1).font = { bold: true }; rr.getCell(2).value = d; r++;
  });
}

// ---- 2. Test Execution Summary (live formulas) ----
function addExecutionSummarySheet(wb) {
  const ws = wb.addWorksheet('Test Execution Summary');
  ws.getColumn(1).width = 24; ws.getColumn(2).width = 12; ws.getColumn(3).width = 10; ws.getColumn(4).width = 30;
  titleBand(ws, 4, 'TEST EXECUTION SUMMARY', 'Live — updates as the Execution Status column in Testcases is filled in');
  const idR = tcRange('Test Case ID');
  const stR = tcRange('Execution Status');
  let r = 4;
  sectionHead(ws, r, 4, 'Execution Metrics'); r++;
  const metric = (label, formula, pct) => {
    const rr = ws.getRow(r); rr.getCell(1).value = label; rr.getCell(1).font = { bold: true };
    const c = rr.getCell(2); c.value = { formula }; if (pct) c.numFmt = '0.0%';
    const wrote = r; r++; return wrote;
  };
  const totalRow = metric('Total Test Cases', `COUNTA(${idR})`);
  const passedRow = metric('Passed', `COUNTIF(${stR},"Passed")`);
  const failedRow = metric('Failed', `COUNTIF(${stR},"Failed")`);
  const blockedRow = metric('Blocked', `COUNTIF(${stR},"Blocked")`);
  metric('In Progress', `COUNTIF(${stR},"In Progress")`);
  metric('Not Executed', `COUNTIF(${stR},"Not Executed")`);
  const execRow = metric('Executed', `B${passedRow}+B${failedRow}+B${blockedRow}`);
  metric('Execution Rate', `IF(B${totalRow}=0,0,B${execRow}/B${totalRow})`, true);
  metric('Pass Rate', `IF((B${passedRow}+B${failedRow})=0,0,B${passedRow}/(B${passedRow}+B${failedRow}))`, true);
  metric('Fail Rate', `IF((B${passedRow}+B${failedRow})=0,0,B${failedRow}/(B${passedRow}+B${failedRow}))`, true);
  r++;
  sectionHead(ws, r, 4, 'Status Breakdown'); r++;
  tableHead(ws, r, ['Status', 'Count', '%', 'Bar']); r++;
  ['Passed', 'Failed', 'Blocked', 'In Progress', 'Not Executed'].forEach((s) => {
    const rr = ws.getRow(r);
    rr.getCell(1).value = s;
    rr.getCell(2).value = { formula: `COUNTIF(${stR},"${s}")` };
    rr.getCell(3).value = { formula: `IF(B${totalRow}=0,0,B${r}/B${totalRow})` }; rr.getCell(3).numFmt = '0.0%';
    rr.getCell(4).value = { formula: `REPT("█",ROUND(IF(B${totalRow}=0,0,B${r}/B${totalRow})*20,0))` };
    rr.getCell(4).font = { color: { argb: NAVY } };
    r++;
  });
  r++;
  sectionHead(ws, r, 4, 'Sign-off'); r++;
  [['Prepared By'], ['Reviewed By'], ['Approved By'], ['Date']].forEach(([k]) => {
    const rr = ws.getRow(r); rr.getCell(1).value = k; rr.getCell(1).font = { bold: true };
    rr.getCell(2).value = '__________________________'; r++;
  });
}

// ---- 4. Dashboard (live counts + in-cell bars) ----
function addDashboardSheet(wb, testCases) {
  const ws = wb.addWorksheet('Dashboard');
  ws.getColumn(1).width = 28; ws.getColumn(2).width = 10; ws.getColumn(3).width = 34;
  titleBand(ws, 3, 'DASHBOARD', 'Live counts — update as the Testcases sheet changes');
  let r = 4;
  const block = (title, col, values) => {
    sectionHead(ws, r, 3, title); r++;
    tableHead(ws, r, ['Category', 'Count', 'Bar']); r++;
    values.forEach((v) => {
      const rr = ws.getRow(r);
      const safe = String(v).replace(/"/g, '""');
      rr.getCell(1).value = v;
      rr.getCell(2).value = { formula: `COUNTIF(${tcRange(col)},"${safe}")` };
      rr.getCell(3).value = { formula: `REPT("█",MIN(B${r},40))` };
      rr.getCell(3).font = { color: { argb: NAVY } };
      r++;
    });
    r++;
  };
  block('Test Cases by Type', 'Type', ['Functional', 'Negative', 'Validation', 'UI', 'Integration']);
  block('Test Cases by Severity', 'Severity', ['Critical', 'High', 'Medium', 'Low']);
  const subs = [...new Set(testCases.map((t) => (t.sub_module || '').trim()).filter(Boolean))];
  block('Test Cases by Sub-Module', 'Sub Module', subs.length ? subs : ['(none)']);
}

const US_COLUMNS = ['S.no', 'Module', 'Sub Module', 'User Story ID', 'Title', 'User Story', 'Acceptance Criteria', 'Priority'];
function addUserStoriesSheet(wb, stories) {
  const ws = wb.addWorksheet('User Stories');
  ws.columns = US_COLUMNS.map((h) => ({ header: h, key: h }));
  stories.forEach((s, i) => ws.addRow({
    'S.no': i + 1, 'Module': s.module, 'Sub Module': s.sub_module,
    'User Story ID': `US_${String(i + 1).padStart(4, '0')}`, 'Title': s.title, 'User Story': s.story,
    'Acceptance Criteria': s.acceptance_criteria, 'Priority': s.priority,
  }));
  applyWidths(ws, { 'S.no': 6, 'Module': 18, 'Sub Module': 18, 'User Story ID': 14, 'Title': 28, 'User Story': 50, 'Acceptance Criteria': 45, 'Priority': 10 });
}

const SC_COLUMNS = ['S.no', 'Module', 'Sub Module', 'Scenario ID', 'Scenario Description', 'Type', 'Priority'];
function addScenariosSheet(wb, scenarios) {
  const ws = wb.addWorksheet('Test Scenarios');
  ws.columns = SC_COLUMNS.map((h) => ({ header: h, key: h }));
  scenarios.forEach((s, i) => ws.addRow({
    'S.no': i + 1, 'Module': s.module, 'Sub Module': s.sub_module,
    'Scenario ID': `SC_${String(i + 1).padStart(4, '0')}`, 'Scenario Description': s.description,
    'Type': s.type, 'Priority': s.priority,
  }));
  applyWidths(ws, { 'S.no': 6, 'Module': 18, 'Sub Module': 18, 'Scenario ID': 14, 'Scenario Description': 55, 'Type': 16, 'Priority': 10 });
}

// --- Traceability + coverage ---
function norm(s) { return String(s || '').toLowerCase().replace(/[^a-z0-9.]/g, ''); }
function tokenize(s) { return String(s || '').split(/[,;/|]+/).map((t) => norm(t)).filter(Boolean); }

// Compute per-requirement coverage from the actual generated test cases.
function computeCoverage(requirements, cases) {
  const caseList = cases.map((c, i) => ({ id: `TC_${String(i + 1).padStart(4, '0')}`, tokens: tokenize(c.brd_req_id) }));

  let reqs = (requirements || []).filter((r) => (r.req_id || '').trim() || (r.requirement || '').trim());
  if (!reqs.length) {
    // Fall back to distinct brd_req_ids found on the test cases.
    const seen = new Map();
    cases.forEach((c) => {
      const k = (c.brd_req_id || '').trim();
      if (k && !seen.has(k)) seen.set(k, { req_id: k, requirement: '' });
    });
    reqs = [...seen.values()];
  }

  const rows = reqs.map((r, i) => {
    const reqId = (r.req_id || '').trim() || `R${i + 1}`;
    const rn = norm(reqId);
    const covering = rn
      ? caseList.filter((c) => c.tokens.some((t) => t === rn || t.includes(rn) || rn.includes(t))).map((c) => c.id)
      : [];
    return { req_id: reqId, requirement: r.requirement || '', covering, status: covering.length ? 'Covered' : 'Gap' };
  });

  const covered = rows.filter((r) => r.status === 'Covered').length;
  return { rows, total: rows.length, covered, gaps: rows.filter((r) => r.status === 'Gap') };
}

// ---- 5. Traceability Matrix (RTM) ----
const RTM_COLUMNS = ['S.no', 'BRD Req ID', 'Requirement', 'Covered By (Test Case IDs)', 'Coverage', 'Live Test Count'];
function addTraceabilitySheet(wb, cov) {
  const ws = wb.addWorksheet('Traceability Matrix (RTM)');
  ws.columns = RTM_COLUMNS.map((h) => ({ header: h, key: h }));
  cov.rows.forEach((r, i) => {
    const row = ws.addRow({
      'S.no': i + 1, 'BRD Req ID': r.req_id, 'Requirement': r.requirement,
      'Covered By (Test Case IDs)': r.covering.join(', '), 'Coverage': r.status,
    });
    // Live count: how many Testcases rows reference this requirement id.
    row.getCell('Live Test Count').value = { formula: `COUNTIF(${tcRange('BRD Req ID')},"*"&B${i + 2}&"*")` };
    if (r.status === 'Gap') row.getCell('Coverage').font = { color: { argb: 'FFC00000' }, bold: true };
  });
  applyWidths(ws, { 'S.no': 6, 'BRD Req ID': 16, 'Requirement': 50, 'Covered By (Test Case IDs)': 34, 'Coverage': 12, 'Live Test Count': 14 });
}

// ---- 6. Solution Doc Coverage ----
const SDC_COLUMNS = ['S.no', 'Solution Doc Ref', 'Covered By (Test Case IDs)', 'Coverage', 'Alignment / Notes'];
function addSolutionDocCoverageSheet(wb, testCases) {
  const ws = wb.addWorksheet('Solution Doc Coverage');
  ws.columns = SDC_COLUMNS.map((h) => ({ header: h, key: h }));
  const map = new Map();
  testCases.forEach((tc, i) => {
    const ref = (tc.solution_doc_ref || '').trim();
    if (!ref || ref === '—') return;
    if (!map.has(ref)) map.set(ref, { ids: [], align: tc.sol_doc_alignment || 'Aligned' });
    map.get(ref).ids.push(`TC_${String(i + 1).padStart(4, '0')}`);
  });
  if (!map.size) {
    ws.addRow({ 'S.no': 1, 'Solution Doc Ref': '—', 'Covered By (Test Case IDs)': '—', 'Coverage': 'N/A',
      'Alignment / Notes': 'No Solution Document reference provided — see the Traceability Matrix (RTM) for requirement-level coverage.' });
  } else {
    let i = 1;
    map.forEach((v, ref) => { ws.addRow({ 'S.no': i, 'Solution Doc Ref': ref, 'Covered By (Test Case IDs)': v.ids.join(', '), 'Coverage': 'Covered', 'Alignment / Notes': v.align }); i++; });
  }
  applyWidths(ws, { 'S.no': 6, 'Solution Doc Ref': 24, 'Covered By (Test Case IDs)': 34, 'Coverage': 12, 'Alignment / Notes': 40 });
}

// ---- 7. Quality Scorecard (formula-linked to Testcases) ----
const QS_COLUMNS = ['S.no', 'Test Case ID', 'Traceable', 'Steps', 'Expected', 'Type & Severity', 'Pre-Req', 'Score %'];
function addQualityScorecardSheet(wb, testCases) {
  const ws = wb.addWorksheet('Quality Scorecard');
  ws.columns = QS_COLUMNS.map((h) => ({ header: h, key: h }));
  const D = tcCol('Test Case ID'), V = tcCol('BRD Req ID'), J = tcCol('Test Steps'),
    K = tcCol('Expected Result'), E = tcCol('Type'), F = tcCol('Severity'), H = tcCol('Pre-Requisites');
  testCases.forEach((tc, i) => {
    const tr = i + 2; // Testcases data row
    const r = i + 2;  // Scorecard data row
    const row = ws.addRow({});
    row.getCell(1).value = i + 1;
    row.getCell(2).value = { formula: `${TC_SHEET}!${D}${tr}` };
    row.getCell(3).value = { formula: `IF(${TC_SHEET}!${V}${tr}<>"",1,0)` };
    row.getCell(4).value = { formula: `IF(${TC_SHEET}!${J}${tr}<>"",1,0)` };
    row.getCell(5).value = { formula: `IF(${TC_SHEET}!${K}${tr}<>"",1,0)` };
    row.getCell(6).value = { formula: `IF(AND(${TC_SHEET}!${E}${tr}<>"",${TC_SHEET}!${F}${tr}<>""),1,0)` };
    row.getCell(7).value = { formula: `IF(${TC_SHEET}!${H}${tr}<>"",1,0)` };
    const sc = row.getCell(8); sc.value = { formula: `AVERAGE(C${r}:G${r})` }; sc.numFmt = '0%';
  });
  const last = testCases.length + 1;
  const avg = ws.addRow({});
  avg.getCell(2).value = 'Suite Average'; avg.getCell(2).font = { bold: true };
  const ac = avg.getCell(8); ac.value = { formula: `IFERROR(AVERAGE(H2:H${last}),0)` }; ac.numFmt = '0%'; ac.font = { bold: true };
  applyWidths(ws, { 'S.no': 6, 'Test Case ID': 14, 'Traceable': 11, 'Steps': 9, 'Expected': 10, 'Type & Severity': 15, 'Pre-Req': 10, 'Score %': 10 });
}

// Assemble the full professional workbook in the required tab order.
function buildTestcaseWorkbook(wb, cases, cov, { meta, basicPreReq, basicSteps }) {
  addProjectSummarySheet(wb, meta);          // 1
  addExecutionSummarySheet(wb);              // 2
  addTestCasesSheet(wb, cases, { basicPreReq, basicSteps, createdBy: '' }); // 3
  addDashboardSheet(wb, cases);              // 4
  addTraceabilitySheet(wb, cov);             // 5
  addSolutionDocCoverageSheet(wb, cases);    // 6
  addQualityScorecardSheet(wb, cases);       // 7
}

// ---------------------------------------------------------------------------
// Route: the page posts here
// ---------------------------------------------------------------------------
app.post('/generate-testcases', upload.single('brdFile'), async (req, res) => {
  try {
    const inputMethod = req.body.inputMethod || 'Upload File';
    const outputs = (req.body.outputs || '').split(',').map((s) => s.trim()).filter(Boolean);
    const preConditions = req.body.preConditions || '';
    const testStepsDefault = req.body.testSteps || '';
    const description = req.body.description || '';
    if (!outputs.length) return res.status(400).json({ message: 'Select at least one item to generate.' });

    // Resolve the BRD file (upload or path).
    let originalname, buf;
    if (inputMethod === 'File Path') {
      const p = (req.body.brdFilePath || '').trim();
      if (!p) return res.status(400).json({ message: 'Please enter the BRD file path.' });
      const resolved = path.resolve(p);
      if (!fs.existsSync(resolved)) return res.status(400).json({ message: `File not found: ${resolved}` });
      if (!fs.statSync(resolved).isFile()) return res.status(400).json({ message: `Not a file: ${resolved}` });
      originalname = path.basename(resolved);
      buf = fs.readFileSync(resolved);
    } else {
      if (!req.file) return res.status(400).json({ message: 'No BRD file uploaded.' });
      originalname = req.file.originalname;
      buf = req.file.buffer;
    }

    const brdText = (await extractText(originalname, buf)).trim();
    if (!brdText) return res.status(400).json({ message: 'No readable text found in the file (is it a scanned image?).' });

    // Generate the requested data first, then build the workbook in professional tab order.
    const wb = new ExcelJS.Workbook();
    wb.calcProperties = { fullCalcOnLoad: true };
    let total = 0;

    const stories = outputs.includes('UserStories') ? await generateUserStories(brdText, description) : null;
    const scenariosData = outputs.includes('Scenarios') ? await generateScenarios(brdText, description) : null;
    const tc = outputs.includes('TestCases') ? await generateTestCases(brdText, description) : null;

    if (tc) {
      const cov = computeCoverage(tc.requirements, tc.cases);
      buildTestcaseWorkbook(wb, tc.cases, cov, {
        meta: buildMeta(tc.cases, description, originalname),
        basicPreReq: preConditions, basicSteps: testStepsDefault,
      });
      total += tc.cases.length;
    }
    if (stories) { addUserStoriesSheet(wb, stories); total += stories.length; }
    if (scenariosData) { addScenariosSheet(wb, scenariosData); total += scenariosData.length; }
    if (total === 0) return res.status(422).json({ message: 'The AI returned no results. Try a more detailed BRD.' });

    // Stream back for download only (nothing written to disk).
    const base = path.parse(originalname).name.replace(/[^\w.-]+/g, '_') || 'Output';
    const stamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
    const label = outputs.includes('TestCases') ? 'TestCases' : (outputs.includes('Scenarios') ? 'Scenarios' : 'UserStories');
    const fileName = `${label}_${base}_${stamp}.xlsx`;

    const buffer = await wb.xlsx.writeBuffer();
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${fileName}"`);
    res.setHeader('X-Row-Count', String(total));
    res.send(Buffer.from(buffer));
  } catch (err) {
    console.error(err);
    if (!res.headersSent) res.status(500).json({ message: err.message || 'Generation failed.' });
  }
});

// Manual test-case creation — no AI, just build the Excel from user-entered rows.
app.post('/generate-manual', async (req, res) => {
  try {
    const rows = Array.isArray(req.body.testCases) ? req.body.testCases : [];
    const preConditions = req.body.preConditions || '';
    const testStepsDefault = req.body.testSteps || '';
    const cases = rows
      .map((tc) => ({
        module: tc.module || '', sub_module: tc.sub_module || '', type: tc.type || '', severity: tc.severity || '',
        description: tc.description || '', pre_requisites: tc.pre_requisites || '', roles: tc.roles || '',
        test_steps: tc.test_steps || '', test_data: tc.test_data || '', expected_result: tc.expected_result || '', brd_req_id: tc.brd_req_id || '',
      }))
      // keep rows that have at least a description or steps
      .filter((tc) => tc.description.trim() || tc.test_steps.trim());

    if (!cases.length) return res.status(400).json({ message: 'Add at least one test case with a description or steps.' });

    const wb = new ExcelJS.Workbook();
    addTestCasesSheet(wb, cases, { basicPreReq: preConditions, basicSteps: testStepsDefault, createdBy: '' });

    const stamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
    const fileName = `TestCases_Manual_${stamp}.xlsx`;
    const buffer = await wb.xlsx.writeBuffer();
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${fileName}"`);
    res.setHeader('X-Row-Count', String(cases.length));
    res.send(Buffer.from(buffer));
  } catch (err) {
    console.error(err);
    if (!res.headersSent) res.status(500).json({ message: err.message || 'Generation failed.' });
  }
});

// Scenario-driven creation — user types scenarios, AI expands them into full test cases.
app.post('/generate-scenarios', async (req, res) => {
  try {
    const scenarios = (req.body.scenarios || '').trim();
    if (!scenarios) return res.status(400).json({ message: 'Please enter at least one scenario.' });

    const { cases, requirements } = await generateTestCases(scenarios, '');
    if (!cases.length) return res.status(422).json({ message: 'No test cases were generated. Add more detail to your scenarios.' });

    const wb = new ExcelJS.Workbook();
    wb.calcProperties = { fullCalcOnLoad: true };
    const cov = computeCoverage(requirements, cases);
    buildTestcaseWorkbook(wb, cases, cov, {
      meta: buildMeta(cases, scenarios, 'Typed scenarios'),
      basicPreReq: '', basicSteps: '',
    });

    const stamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
    const fileName = `TestCases_Scenarios_${stamp}.xlsx`;
    const buffer = await wb.xlsx.writeBuffer();
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${fileName}"`);
    res.setHeader('X-Row-Count', String(cases.length));
    res.send(Buffer.from(buffer));
  } catch (err) {
    console.error(err);
    if (!res.headersSent) res.status(500).json({ message: err.message || 'Generation failed.' });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`\n  Test Case Generator running at http://localhost:${PORT}`);
  console.log(`  Provider : ${AI_BASE_URL}  (${AI_MODEL})`);
  console.log(`  Output   : download only (no files written to disk)\n`);
});
