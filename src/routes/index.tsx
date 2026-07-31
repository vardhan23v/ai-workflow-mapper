import { createFileRoute } from "@tanstack/react-router";
import {
  ChangeEvent,
  useEffect,
  useMemo,
  useRef,
  useState,
} from "react";
import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  Table,
  TableRow,
  TableCell,
  AlignmentType,
  WidthType,
  BorderStyle,
  HeadingLevel,
  ImageRun,
} from "docx";

const CATEGORIES = [
  { id: "just-me", label: "Just me", short: "JM" },
  { id: "delegate", label: "Delegate to AI (review)", short: "DG" },
  { id: "collaborate", label: "Collaborate with AI", short: "CL" },
  { id: "automate", label: "Fully automate", short: "AU" },
] as const;

type CategoryId = (typeof CATEGORIES)[number]["id"];

interface AuditItem {
  id: string;
  name: string;
  category: CategoryId | "";
  rationale: string;
}

interface TargetTask {
  id: string;
  name: string;
  successDefinition: string;
}

function makeId() {
  return Math.random().toString(36).slice(2, 9);
}

function createDefaultItems(count = 15): AuditItem[] {
  return Array.from({ length: count }, () => ({
    id: makeId(),
    name: "",
    category: "",
    rationale: "",
  }));
}

function createDefaultTargets(count = 3): TargetTask[] {
  return Array.from({ length: count }, () => ({
    id: makeId(),
    name: "",
    successDefinition: "",
  }));
}

const CHECKLIST = [
  { key: "claude", label: "Claude account created", link: "https://claude.ai" },
  { key: "chatgpt", label: "ChatGPT account created", link: "https://chat.openai.com" },
  {
    key: "academy",
    label: "Enrolled in AI Fluency: Framework & Foundations",
    link: "https://www.anthropic.com/academy",
  },
  {
    key: "claude-project",
    label: "Created a Claude Project with custom instructions",
    link: "https://support.anthropic.com/en/articles/9367035-what-are-projects",
  },
];

export const Route = createFileRoute("/")({
  head: () => ({
    meta: [
      { title: "AI Workflow Audit — FL-01" },
      {
        name: "description",
        content:
          "Map your recurring tasks, classify how AI fits in, and export the FL-01 deliverable.",
      },
      { property: "og:title", content: "AI Workflow Audit — FL-01" },
      {
        property: "og:description",
        content:
          "Map your recurring tasks, classify how AI fits in, and export the FL-01 deliverable.",
      },
      { property: "og:type", content: "website" },
      { name: "twitter:card", content: "summary_large_image" },
    ],
    links: [{ rel: "canonical", href: "/" }],
  }),
  component: Index,
});

function Index() {
  const [items, setItems] = useState<AuditItem[]>(() => createDefaultItems());
  const [targets, setTargets] = useState<TargetTask[]>(() =>
    createDefaultTargets()
  );
  const [checklist, setChecklist] = useState<Record<string, boolean>>({});
  const [screenshot, setScreenshot] = useState<File | null>(null);
  const [screenshotPreview, setScreenshotPreview] = useState<string | null>(
    null
  );
  const [notes, setNotes] = useState("");
  const printRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    if (!screenshot) {
      if (screenshotPreview) URL.revokeObjectURL(screenshotPreview);
      setScreenshotPreview(null);
      return;
    }
    const url = URL.createObjectURL(screenshot);
    setScreenshotPreview(url);
    return () => URL.revokeObjectURL(url);
  }, [screenshot]);

  const counts = useMemo(() => {
    const out: Record<string, number> = {
      "just-me": 0,
      delegate: 0,
      collaborate: 0,
      automate: 0,
      "": 0,
    };
    items.forEach((i) => {
      const key = i.category || "";
      out[key] = (out[key] ?? 0) + 1;
    });
    return out;
  }, [items]);

  const filledCount = items.filter((i) => i.name.trim()).length;
  const justMeCount = counts["just-me"];
  const missingCount = items.filter(
    (i) => i.name.trim() && (!i.category || !i.rationale.trim())
  ).length;

  function updateItem(id: string, patch: Partial<AuditItem>) {
    setItems((prev) => prev.map((i) => (i.id === id ? { ...i, ...patch } : i)));
  }

  function updateTarget(id: string, patch: Partial<TargetTask>) {
    setTargets((prev) =>
      prev.map((t) => (t.id === id ? { ...t, ...patch } : t))
    );
  }

  function handleScreenshot(e: ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0] ?? null;
    setScreenshot(file);
  }

  function triggerPrint() {
    if (typeof window !== "undefined") {
      window.print();
    }
  }

  async function downloadDocx() {
    if (typeof window === "undefined") return;

    const border = {
      top: { style: BorderStyle.SINGLE, size: 1, color: "AAAAAA" },
      bottom: { style: BorderStyle.SINGLE, size: 1, color: "AAAAAA" },
      left: { style: BorderStyle.SINGLE, size: 1, color: "AAAAAA" },
      right: { style: BorderStyle.SINGLE, size: 1, color: "AAAAAA" },
    };

    const headerCell = (text: string, width: number) =>
      new TableCell({
        width: { size: width, type: WidthType.DXA },
        shading: { fill: "E5E7EB", type: "clear" as any },
        children: [
          new Paragraph({
            children: [new TextRun({ text, bold: true, size: 20 })],
          }),
        ],
        borders: border,
      });

    const bodyCell = (text: string, width: number) =>
      new TableCell({
        width: { size: width, type: WidthType.DXA },
        children: [new Paragraph({ children: [new TextRun({ text, size: 20 })] })],
        borders: border,
      });

    const completedChecklist = CHECKLIST.filter((c) => checklist[c.key]);

    const children: (Paragraph | Table)[] = [
      new Paragraph({
        heading: HeadingLevel.HEADING_1,
        alignment: AlignmentType.CENTER,
        children: [
          new TextRun({ text: "AI Workflow Audit — FL-01", bold: true, size: 32 }),
        ],
      }),
      new Paragraph({
        children: [
          new TextRun({
            text: `Completed tools checklist: ${completedChecklist
              .map((c) => c.label)
              .join(", ") || "Not recorded"}`,
            size: 20,
          }),
        ],
      }),
      new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("Task Audit")] }),
      new Paragraph({
        children: [new TextRun({ text: `${filledCount} tasks documented`, size: 20 })],
      }),
    ];

    const auditColWidths: [number, number, number] = [3600, 2200, 3560];

    const auditRows = items
      .filter((i) => i.name.trim())
      .map((i) => {
        const label = CATEGORIES.find((c) => c.id === i.category)?.label || i.category || "";
        return new TableRow({
          children: [
            bodyCell(i.name, auditColWidths[0]),
            bodyCell(label, auditColWidths[1]),
            bodyCell(i.rationale, auditColWidths[2]),
          ],
        });
      });

    children.push(
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: auditColWidths,
        rows: [
          new TableRow({
            children: [
              headerCell("Task", auditColWidths[0]),
              headerCell("Classification", auditColWidths[1]),
              headerCell("Rationale", auditColWidths[2]),
            ],
          }),
          ...(auditRows.length
            ? auditRows
            : [
                new TableRow({
                  children: [
                    bodyCell("No tasks entered", auditColWidths[0]),
                    bodyCell("-", auditColWidths[1]),
                    bodyCell("-", auditColWidths[2]),
                  ],
                }),
              ]),
        ],
      })
    );

    children.push(
      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun("Target Tasks for FL-02–FL-04")],
      })
    );

    const targetColWidths: [number, number] = [4200, 5160];

    const targetRows = targets.map((t) => {
      return new TableRow({
        children: [
          bodyCell(t.name.trim() || "(not set)", targetColWidths[0]),
          bodyCell(t.successDefinition.trim() || "(not set)", targetColWidths[1]),
        ],
      });
    });

    children.push(
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: targetColWidths,
        rows: [
          new TableRow({
            children: [headerCell("Target task", targetColWidths[0]), headerCell('What "done well" means', targetColWidths[1])],
          }),
          ...targetRows,
        ],
      })
    );

    if (screenshot) {
      const arrayBuffer = await screenshot.arrayBuffer();
      const imageRun = new ImageRun({
        data: new Uint8Array(arrayBuffer),
        transformation: { width: 450, height: 300 },
        type: screenshot.type.includes("png") ? "png" : "jpg",
      });
      children.push(
        new Paragraph({
          heading: HeadingLevel.HEADING_2,
          children: [new TextRun("Claude Project Screenshot")],
        }),
        new Paragraph({ children: [imageRun] })
      );
    }

    if (notes.trim()) {
      children.push(
        new Paragraph({
          heading: HeadingLevel.HEADING_2,
          children: [new TextRun("Notes")],
        }),
        new Paragraph({ children: [new TextRun({ text: notes, size: 20 })] })
      );
    }

    const doc = new Document({
      sections: [
        {
          properties: {
            page: {
              size: { width: 12240, height: 15840 },
              margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
            },
          },
          children,
        },
      ],
    });

    const blob = await Packer.toBlob(doc);
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "FL-01-AI-Workflow-Audit.docx";
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  }

  return (
    <div className="min-h-screen bg-background text-foreground">
      <div className="mx-auto max-w-5xl px-4 py-10 md:py-14">
        <header className="mb-10">
          <h1 className="text-3xl font-bold tracking-tight md:text-4xl">
            AI Workflow Audit
          </h1>
          <p className="mt-2 text-muted-foreground">
            FL-01 assignment helper — map your recurring tasks, choose how AI
            fits, and export your deliverable.
          </p>
        </header>

        <section className="mb-10 rounded-xl border border-border bg-card p-6 shadow-sm">
          <h2 className="mb-4 text-lg font-semibold">Tool setup checklist</h2>
          <p className="mb-4 text-sm text-muted-foreground">
            Mark each item after you complete it. These appear in your exported
            document.
          </p>
          <div className="grid gap-3 sm:grid-cols-2">
            {CHECKLIST.map((item) => (
              <label
                key={item.key}
                className="flex cursor-pointer items-start gap-3 rounded-lg border border-border p-3 transition-colors hover:bg-accent"
              >
                <input
                  type="checkbox"
                  className="mt-1 h-4 w-4 accent-primary"
                  checked={!!checklist[item.key]}
                  onChange={(e) =>
                    setChecklist((prev) => ({
                      ...prev,
                      [item.key]: e.target.checked,
                    }))
                  }
                />
                <span className="text-sm leading-relaxed">
                  <a
                    href={item.link}
                    target="_blank"
                    rel="noreferrer"
                    className="font-medium underline-offset-2 hover:underline"
                    onClick={(e) => e.stopPropagation()}
                  >
                    {item.label}
                  </a>
                </span>
              </label>
            ))}
          </div>
        </section>

        <section className="mb-10 rounded-xl border border-border bg-card p-6 shadow-sm">
          <div className="mb-4 flex flex-wrap items-center justify-between gap-3">
            <h2 className="text-lg font-semibold">Task audit</h2>
            <div className="flex flex-wrap gap-2 text-xs">
              <span className="rounded-full bg-secondary px-2.5 py-1">
                {filledCount}/15 tasks
              </span>
              <span className="rounded-full bg-secondary px-2.5 py-1">
                {justMeCount} just me
              </span>
              <span className="rounded-full bg-secondary px-2.5 py-1">
                {counts["delegate"]} delegate
              </span>
              <span className="rounded-full bg-secondary px-2.5 py-1">
                {counts["collaborate"]} collaborate
              </span>
              <span className="rounded-full bg-secondary px-2.5 py-1">
                {counts["automate"]} automate
              </span>
            </div>
          </div>

          {missingCount > 0 && (
            <p className="mb-4 text-sm text-destructive">
              {missingCount} filled task{missingCount > 1 ? "s" : ""} still
              need a classification and rationale.
            </p>
          )}

          <div className="overflow-x-auto rounded-lg border border-border">
            <table className="w-full text-left text-sm">
              <thead className="bg-muted">
                <tr>
                  <th className="px-4 py-3 font-semibold">#</th>
                  <th className="min-w-[220px] px-4 py-3 font-semibold">
                    Recurring task
                  </th>
                  <th className="min-w-[260px] px-4 py-3 font-semibold">
                    Classification
                  </th>
                  <th className="min-w-[280px] px-4 py-3 font-semibold">
                    Rationale
                  </th>
                </tr>
              </thead>
              <tbody className="divide-y divide-border">
                {items.map((item, idx) => (
                  <tr key={item.id} className="bg-card hover:bg-accent/50">
                    <td className="px-4 py-3 align-top text-muted-foreground">
                      {idx + 1}
                    </td>
                    <td className="px-2 py-2 align-top">
                      <input
                        type="text"
                        value={item.name}
                        onChange={(e) =>
                          updateItem(item.id, { name: e.target.value })
                        }
                        placeholder="e.g. weekly meeting notes"
                        className="w-full rounded-md border border-input bg-background px-3 py-2 text-sm outline-none ring-offset-background focus-visible:ring-2 focus-visible:ring-ring"
                      />
                    </td>
                    <td className="px-2 py-2 align-top">
                      <div className="flex flex-wrap gap-1">
                        {CATEGORIES.map((cat) => {
                          const active = item.category === cat.id;
                          return (
                            <button
                              key={cat.id}
                              type="button"
                              onClick={() =>
                                updateItem(item.id, { category: cat.id })
                              }
                              className={[
                                "rounded-md border px-2 py-1.5 text-xs font-medium transition-colors",
                                active
                                  ? "border-primary bg-primary text-primary-foreground"
                                  : "border-input bg-background text-foreground hover:bg-accent",
                              ].join(" ")}
                              title={cat.label}
                            >
                              {cat.short} — {cat.label}
                            </button>
                          );
                        })}
                      </div>
                    </td>
                    <td className="px-2 py-2 align-top">
                      <input
                        type="text"
                        value={item.rationale}
                        onChange={(e) =>
                          updateItem(item.id, { rationale: e.target.value })
                        }
                        placeholder="Why this classification?"
                        className="w-full rounded-md border border-input bg-background px-3 py-2 text-sm outline-none ring-offset-background focus-visible:ring-2 focus-visible:ring-ring"
                      />
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </section>

        <section className="mb-10 rounded-xl border border-border bg-card p-6 shadow-sm">
          <h2 className="mb-4 text-lg font-semibold">Target tasks for FL-02–FL-04</h2>
          <p className="mb-4 text-sm text-muted-foreground">
            Pick the three audit tasks you will reuse. Define what "done well"
            means for each.
          </p>
          <div className="space-y-4">
            {targets.map((target, idx) => (
              <div key={target.id} className="grid gap-4 md:grid-cols-2">
                <div>
                  <label className="mb-1.5 block text-xs font-medium">
                    Target task {idx + 1}
                  </label>
                  <input
                    type="text"
                    value={target.name}
                    onChange={(e) =>
                      updateTarget(target.id, { name: e.target.value })
                    }
                    placeholder="Task name"
                    className="w-full rounded-md border border-input bg-background px-3 py-2 text-sm outline-none ring-offset-background focus-visible:ring-2 focus-visible:ring-ring"
                  />
                </div>
                <div>
                  <label className="mb-1.5 block text-xs font-medium">
                    What "done well" means
                  </label>
                  <input
                    type="text"
                    value={target.successDefinition}
                    onChange={(e) =>
                      updateTarget(target.id, {
                        successDefinition: e.target.value,
                      })
                    }
                    placeholder="Specific, measurable success definition"
                    className="w-full rounded-md border border-input bg-background px-3 py-2 text-sm outline-none ring-offset-background focus-visible:ring-2 focus-visible:ring-ring"
                  />
                </div>
              </div>
            ))}
          </div>
        </section>

        <section className="mb-10 rounded-xl border border-border bg-card p-6 shadow-sm">
          <h2 className="mb-4 text-lg font-semibold">Claude Project screenshot</h2>
          <p className="mb-4 text-sm text-muted-foreground">
            Upload the screenshot of your configured Claude Project. It will be
            included in the exported document.
          </p>
          <input
            type="file"
            accept="image/png,image/jpeg,image/jpg"
            onChange={handleScreenshot}
            className="block w-full cursor-pointer text-sm file:mr-4 file:rounded-md file:border-0 file:bg-secondary file:px-4 file:py-2 file:text-sm file:font-medium hover:file:bg-secondary/80"
          />
          {screenshotPreview && (
            <div className="mt-4">
              <img
                src={screenshotPreview}
                alt="Claude Project screenshot preview"
                className="max-h-80 rounded-lg border border-border object-contain"
              />
            </div>
          )}
        </section>

        <section className="mb-10 rounded-xl border border-border bg-card p-6 shadow-sm">
          <h2 className="mb-4 text-lg font-semibold">Notes for reviewers</h2>
          <textarea
            value={notes}
            onChange={(e) => setNotes(e.target.value)}
            rows={4}
            placeholder="Optional context, links to tool accounts, Academy progress, etc."
            className="w-full rounded-md border border-input bg-background px-3 py-2 text-sm outline-none ring-offset-background focus-visible:ring-2 focus-visible:ring-ring"
          />
        </section>

        <section className="mb-10 rounded-xl border border-border bg-card p-6 shadow-sm">
          <h2 className="mb-4 text-lg font-semibold">Export deliverable</h2>
          <div className="flex flex-wrap gap-3">
            <button
              onClick={triggerPrint}
              className="inline-flex items-center justify-center rounded-md border border-input bg-background px-4 py-2 text-sm font-medium transition-colors hover:bg-accent"
            >
              Print / save as PDF
            </button>
            <button
              onClick={downloadDocx}
              className="inline-flex items-center justify-center rounded-md bg-primary px-4 py-2 text-sm font-medium text-primary-foreground transition-colors hover:bg-primary/90"
            >
              Download DOCX
            </button>
          </div>
        </section>

        {/* Hidden print area */}
        <div ref={printRef} className="print-only hidden">
          <div className="print-header">
            <h1>AI Workflow Audit — FL-01</h1>
            <p>
              Tools checklist:{" "}
              {CHECKLIST.filter((c) => checklist[c.key])
                .map((c) => c.label)
                .join(", ") || "Not completed"}
            </p>
          </div>

          <h2>Task Audit</h2>
          <p>{filledCount} tasks documented</p>
          <table className="print-table">
            <thead>
              <tr>
                <th>Task</th>
                <th>Classification</th>
                <th>Rationale</th>
              </tr>
            </thead>
            <tbody>
              {items
                .filter((i) => i.name.trim())
                .map((i) => (
                  <tr key={i.id}>
                    <td>{i.name}</td>
                    <td>
                      {CATEGORIES.find((c) => c.id === i.category)?.label ||
                        i.category}
                    </td>
                    <td>{i.rationale}</td>
                  </tr>
                ))}
            </tbody>
          </table>

          <h2>Target Tasks</h2>
          <table className="print-table">
            <thead>
              <tr>
                <th>Target task</th>
                <th>What "done well" means</th>
              </tr>
            </thead>
            <tbody>
              {targets.map((t) => (
                <tr key={t.id}>
                  <td>{t.name || "(not set)"}</td>
                  <td>{t.successDefinition || "(not set)"}</td>
                </tr>
              ))}
            </tbody>
          </table>

          {screenshotPreview && (
            <>
              <h2>Claude Project Screenshot</h2>
              <img src={screenshotPreview} alt="Claude Project screenshot" />
            </>
          )}

          {notes && (
            <>
              <h2>Notes</h2>
              <p className="whitespace-pre-wrap">{notes}</p>
            </>
          )}
        </div>
      </div>

      <style>{`
        @media print {
          body * {
            visibility: hidden;
          }
          .print-only,
          .print-only * {
            visibility: visible;
          }
          .print-only {
            display: block !important;
            position: absolute;
            left: 0;
            top: 0;
            width: 100%;
            padding: 24px;
            background: white;
            color: black;
          }
          .print-only h1 {
            font-size: 22pt;
            font-weight: bold;
            margin-bottom: 8px;
          }
          .print-only h2 {
            font-size: 14pt;
            font-weight: bold;
            margin-top: 18px;
            margin-bottom: 8px;
          }
          .print-only p {
            font-size: 11pt;
            margin-bottom: 8px;
          }
          .print-only .print-table {
            width: 100%;
            border-collapse: collapse;
            margin-bottom: 16px;
          }
          .print-only .print-table th,
          .print-only .print-table td {
            border: 1px solid #999;
            padding: 6px;
            font-size: 10pt;
            text-align: left;
            vertical-align: top;
          }
          .print-only .print-table th {
            background: #eee;
          }
          .print-only img {
            max-width: 100%;
            max-height: 60vh;
            object-fit: contain;
            border: 1px solid #ccc;
          }
        }
      `}</style>
    </div>
  );
}
