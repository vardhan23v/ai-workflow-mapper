# FL-01 AI Workflow Audit Tool

Build a small web app in this TanStack Start project that guides through and exports the FL-01 deliverables.

## Goals
- Replace the placeholder `src/routes/index.tsx` with an interactive AI workflow audit page.
- Let the user input 10–15 real recurring tasks, choose a classification, and add a one-line rationale.
- Let the user define three target tasks with measurable success definitions.
- Provide an upload slot / preview area for the Claude Project screenshot.
- Generate a print-friendly audit document and export it so it can be uploaded as the submission file.

## Page structure

```text
/ (home)
├── Intro & instructions
├── Tool-setup checklist (Claude, ChatGPT, Anthropic Academy)
├── Task audit table
│   ├── 15 rows: task name, classification dropdown, rationale input
│   └── Validation summary (counts + "just me" count)
├── Target tasks
│   └── 3 rows: task name + "done well" definition
├── Claude Project screenshot preview (file upload -> local preview)
└── Export panel
    ├── Print-friendly preview of the full audit
    └── Download as DOCX / print to PDF
```

## Data model
Keep in browser state only (no backend needed unless the user later wants persistence).

- `TaskAuditItem`: `{ id, name, category: 'just-me' | 'delegate' | 'collaborate' | 'automate', rationale }`
- `TargetTask`: `{ id, name, successDefinition }`
- `ScreenshotFile`: File reference for local preview

## UI/UX
- Use Tailwind v4 and the project’s existing shadcn-style tokens.
- Inline validation: highlight empty required fields, ensure at least two rows are classified as `just-me`.
- Classification is shown as a simple segmented control per row.
- Screenshot preview via URL.createObjectURL.

## Export
- Provide a styled printable view (`@media print`) that renders the audit table, target tasks, and screenshot.
- Add a "Download DOCX" button that uses `docx-js` to produce a 1–2 page deliverable.
- Keep the DOCX under 2 pages, table format, matching the assignment rubric.

## Metadata
Add app-specific `head()` to `src/routes/index.tsx` with title, description, og/twitter tags.

## Out of scope
- No backend database (the assignment data is personal and temporary).
- No auth.
- No file storage; screenshots stay client-side for preview / print.

## Deliverables for the user
After the build, the user will open the app in preview, fill in their personal tasks, paste any tool-account/Academy evidence in notes, upload their Claude Project screenshot, and export the DOCX/PDF for the FL-01 submission.
