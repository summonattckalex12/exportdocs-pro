import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  HeadingLevel,
  Table,
  TableRow,
  TableCell,
  WidthType,
  AlignmentType,
  BorderStyle,
  ShadingType,
  PageBreak,
  TableOfContents,
  LevelFormat,
  PageOrientation,
  VerticalAlign,
  Footer,
  PageNumber,
  ImageRun,
} from "docx";
import FileSaver from "file-saver";
const { saveAs } = FileSaver;

import type { ParsedPM, PMSection, StatusKind, SummaryRow } from "./pm-html-parser";
import { summarizePM } from "./pm-html-parser";

export interface CoverInput {
  reportTitle: string;
  subtitle: string;
  companyName: string;
  operatingSystem: string;
  pelaksana: string;
  tanggal: string;
  contractNo: string;
  periode: string;
  vendorName: string;
  reviewerClient: string;
  reviewerVendor: string;
  reviewerVendorRole: string;
  reviewerDate: string;
  executiveSummary: string;
  summaryConclusion: string;
  recommendation: string;
  /** Optional cover logo (left) — data URL (image/png|jpg|gif). */
  logoDataUrl?: string;
  /** Optional cover logo (right) — data URL. */
  logoRightDataUrl?: string;
}




// Decode a data URL into bytes + docx image type.
function decodeDataUrl(dataUrl: string): { data: Uint8Array; type: "png" | "jpg" | "gif" | "bmp" } | null {
  const m = /^data:image\/(png|jpe?g|gif|bmp);base64,(.+)$/i.exec(dataUrl.trim());
  if (!m) return null;
  const ext = m[1].toLowerCase();
  const type = (ext === "jpeg" ? "jpg" : ext) as "png" | "jpg" | "gif" | "bmp";
  const bin = atob(m[2]);
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  return { data: bytes, type };
}

// ---------- style helpers ----------
const COLOR_OK = "32CB00";
const COLOR_WARN = "FFCB2F";
const COLOR_CRIT = "FE0000";
const COLOR_HEADER = "9B9B9B";
const COLOR_TITLE_BG = "000000";
const COLOR_BRAND = "C8102E";

const CONTENT_WIDTH_DXA = 9360;

function statusFill(s: StatusKind): string | undefined {
  if (s === "OK") return COLOR_OK;
  if (s === "Warning") return COLOR_WARN;
  if (s === "Critical") return COLOR_CRIT;
  return undefined;
}

const border = { style: BorderStyle.SINGLE, size: 4, color: "BFBFBF" };
const cellBorders = { top: border, bottom: border, left: border, right: border };

function P(text: string, opts: { bold?: boolean; size?: number; color?: string; align?: (typeof AlignmentType)[keyof typeof AlignmentType] } = {}) {
  return new Paragraph({
    alignment: opts.align,
    children: [new TextRun({ text, bold: opts.bold, size: opts.size ?? 20, color: opts.color })],
  });
}

function multilineParas(text: string, opts: { bold?: boolean; size?: number; color?: string } = {}) {
  const lines = (text || "").split(/\r?\n/);
  return lines.map(
    (line) =>
      new Paragraph({
        children: [
          new TextRun({ text: line, bold: opts.bold, size: opts.size ?? 18, color: opts.color, font: "Consolas" }),
        ],
      }),
  );
}

function cell(children: Paragraph[], opts: { width?: number; fill?: string; bold?: boolean; align?: (typeof AlignmentType)[keyof typeof AlignmentType] } = {}) {
  return new TableCell({
    borders: cellBorders,
    verticalAlign: VerticalAlign.CENTER,
    width: opts.width ? { size: opts.width, type: WidthType.DXA } : undefined,
    shading: opts.fill ? { fill: opts.fill, type: ShadingType.CLEAR, color: "auto" } : undefined,
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    children,
  });
}

function textCell(text: string, opts: Parameters<typeof cell>[1] & { bold?: boolean; size?: number; color?: string } = {}) {
  return cell(
    [
      new Paragraph({
        alignment: opts.align,
        children: [
          new TextRun({
            text,
            bold: opts.bold,
            size: opts.size ?? 20,
            color: opts.color,
          }),
        ],
      }),
    ],
    opts,
  );
}

// ---------- cover ----------
// Sizes are in half-points (docx unit): 28pt=56, 26pt=52, 22pt=44, 20pt=40.
function buildCover(cover: CoverInput): (Paragraph | Table)[] {
  const out: (Paragraph | Table)[] = [];

  // ---- Optional dual logos (left / right) as a 2-col table ----
  const logoL = cover.logoDataUrl ? decodeDataUrl(cover.logoDataUrl) : null;
  const logoR = cover.logoRightDataUrl ? decodeDataUrl(cover.logoRightDataUrl) : null;

  const logoCell = (logo: ReturnType<typeof decodeDataUrl>, align: (typeof AlignmentType)[keyof typeof AlignmentType]) =>
    new TableCell({
      borders: {
        top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
        bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
        left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
        right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
      },
      width: { size: Math.floor(CONTENT_WIDTH_DXA / 2), type: WidthType.DXA },
      verticalAlign: VerticalAlign.CENTER,
      children: [
        new Paragraph({
          alignment: align,
          children: logo
            ? [
                new ImageRun({
                  type: logo.type,
                  data: logo.data,
                  transformation: { width: 120, height: 120 },
                  altText: { title: "Logo", description: "Cover logo", name: "logo" },
                }),
              ]
            : [new TextRun("")],
        }),
      ],
    });

  if (logoL || logoR) {
    out.push(
      new Table({
        width: { size: CONTENT_WIDTH_DXA, type: WidthType.DXA },
        columnWidths: [Math.floor(CONTENT_WIDTH_DXA / 2), Math.floor(CONTENT_WIDTH_DXA / 2)],
        rows: [
          new TableRow({
            children: [logoCell(logoL, AlignmentType.LEFT), logoCell(logoR, AlignmentType.RIGHT)],
          }),
        ],
      }),
    );
    out.push(new Paragraph({ spacing: { before: 200 }, children: [new TextRun("")] }));
  } else {
    out.push(new Paragraph({ spacing: { before: 400 }, children: [new TextRun("")] }));
  }

  // Brand micro-label
  out.push(
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 400 },
      children: [new TextRun({ text: "⟪ ExcportCuy ⟫", size: 22, bold: true, color: COLOR_BRAND })],
    }),
  );

  // Report title — 26pt
  out.push(
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 200, after: 120 },
      children: [
        new TextRun({
          text: (cover.reportTitle || "LAPORAN PREVENTIVE MAINTENANCE").toUpperCase(),
          size: 52,
          bold: true,
        }),
      ],
    }),
  );
  out.push(
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: cover.subtitle || "Perangkat Lunak", size: 28 })],
    }),
  );
  out.push(
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 400 },
      children: [
        new TextRun({
          text: `- ${cover.operatingSystem || "Red Hat Enterprise Linux"} -`,
          size: 24,
          italics: true,
        }),
      ],
    }),
  );

  // Company name — bold 28pt
  out.push(
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 400, after: 120 },
      children: [
        new TextRun({ text: cover.companyName || "", size: 56, bold: true, color: "0A0A0A" }),
      ],
    }),
  );
  out.push(
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: `Maintenance ${cover.operatingSystem || "OS"}`, size: 24 })],
    }),
  );
  out.push(
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: `No Contract : ${cover.contractNo || "-"}`, size: 22 })],
    }),
  );
  out.push(
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: `Periode ${cover.periode || "-"}`, size: 22 })],
    }),
  );

  // ---- Spacer to push confidentiality notice down ----
  out.push(new Paragraph({ spacing: { before: 2400 }, children: [new TextRun("")] }));

  // ---- Boxed confidentiality notice (bottom of cover) ----
  const boxBorder = { style: BorderStyle.SINGLE, size: 8, color: COLOR_BRAND };
  const confidentialBox = new Table({
    width: { size: CONTENT_WIDTH_DXA, type: WidthType.DXA },
    columnWidths: [CONTENT_WIDTH_DXA],
    rows: [
      new TableRow({
        children: [
          new TableCell({
            borders: { top: boxBorder, bottom: boxBorder, left: boxBorder, right: boxBorder },
            shading: { fill: "FFF7F5", type: ShadingType.CLEAR, color: "auto" },
            margins: { top: 240, bottom: 240, left: 300, right: 300 },
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { after: 160 },
                children: [
                  new TextRun({
                    text: "PEMBERITAHUAN KERAHASIAAN",
                    size: 24,
                    bold: true,
                    color: COLOR_BRAND,
                  }),
                ],
              }),
              new Paragraph({
                alignment: AlignmentType.JUSTIFIED,
                children: [
                  new TextRun({
                    text:
                      `Material dalam dokumen ini dimiliki oleh ${cover.vendorName || "vendor"}. Dokumen ini diajukan kepada "${cover.companyName || "klien"}" untuk tujuan laporan. ` +
                      `Dengan penerimaan dokumen ini "${cover.companyName || "klien"}" telah setuju untuk terikat dengan sifat kerahasiaan laporan ini. Reproduksi pada distribusi dari setiap bagian ` +
                      `dari dokumen ini tidak diperkenankan tanpa persetujuan tertulis sebelumnya dari ${cover.vendorName || "vendor"}.`,
                    size: 20,
                  }),
                ],
              }),
            ],
          }),
        ],
      }),
    ],
  });
  out.push(confidentialBox);

  out.push(new Paragraph({ children: [new PageBreak()] }));
  return out;
}


// ---------- Table of Contents ----------
function buildToc(): (Paragraph | TableOfContents)[] {
  return [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 200 },
      children: [new TextRun({ text: "Table of Contents", size: 32, bold: true })],
    }),
    new TableOfContents("Contents", {
      hyperlink: true,
      headingStyleRange: "1-3",
    }),
    new Paragraph({ children: [new PageBreak()] }),
  ];
}

// ---------- document control ----------
function docControlTable(cover: CoverInput): Table {
  const col1 = 3000;
  const col2 = CONTENT_WIDTH_DXA - col1;
  const rows: [string, string][] = [
    ["Nama Perusahaan", cover.companyName],
    ["Sistem Operasi", cover.operatingSystem],
    ["Pelaksana PM", cover.pelaksana],
    ["Tanggal dan waktu", cover.tanggal],
  ];
  return new Table({
    width: { size: CONTENT_WIDTH_DXA, type: WidthType.DXA },
    columnWidths: [col1, col2],
    rows: rows.map(
      ([k, v]) =>
        new TableRow({
          children: [textCell(k, { width: col1, fill: "F2F2F2", bold: true }), textCell(v || "-", { width: col2 })],
        }),
    ),
  });
}

function revisionTable(): Table {
  const w = [1800, 1800, CONTENT_WIDTH_DXA - 3600 - 2000, 2000];
  const headerRow = new TableRow({
    tableHeader: true,
    children: ["Date", "Revision", "Description", "Author"].map((h, i) =>
      textCell(h, { width: w[i], fill: COLOR_HEADER, bold: true }),
    ),
  });
  const empty = new TableRow({
    children: w.map((width) => textCell(" ", { width })),
  });
  return new Table({ width: { size: CONTENT_WIDTH_DXA, type: WidthType.DXA }, columnWidths: w, rows: [headerRow, empty] });
}

function activityTable(): Table {
  const w = [900, CONTENT_WIDTH_DXA - 900 - 1600, 1600];
  const acts: [number, string, string][] = [
    [1, "Menampilkan tanggal pengambilan PM, versi OS dan kernel", "OK"],
    [2, "Pemeriksaan CPU, disk dan memory usage", "OK"],
    [3, "Pemeriksaan uptime dan reboot terakhir", "OK"],
    [4, "Pemeriksaan error log message", "OK"],
    [5, "Pemeriksaan I/O log message", "OK"],
  ];
  const head = new TableRow({
    tableHeader: true,
    children: ["No", "Aktifitas", "Status"].map((h, i) => textCell(h, { width: w[i], fill: COLOR_HEADER, bold: true })),
  });
  const body = acts.map(
    ([no, act, st]) =>
      new TableRow({
        children: [
          textCell(String(no), { width: w[0], align: AlignmentType.CENTER }),
          textCell(act, { width: w[1] }),
          textCell(st, { width: w[2], fill: COLOR_OK, align: AlignmentType.CENTER, bold: true }),
        ],
      }),
  );
  const conclusionCell = new TableCell({
    borders: cellBorders,
    columnSpan: 3,
    shading: { fill: "F2F2F2", type: ShadingType.CLEAR, color: "auto" },
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    children: [
      new Paragraph({
        children: [
          new TextRun({
            text: "Hasil akhir Preventive Maintenance : Sistem Operasi dalam keadaan layak beroperasi dan berjalan dengan normal.",
            italics: true,
            size: 20,
          }),
        ],
      }),
    ],
  });

  return new Table({
    width: { size: CONTENT_WIDTH_DXA, type: WidthType.DXA },
    columnWidths: w,
    rows: [head, ...body, new TableRow({ children: [conclusionCell] })],
  });
}

function reviewerTable(cover: CoverInput): Table {
  const w = [3600, 2400, CONTENT_WIDTH_DXA - 6000];
  const hdr = (title: string) =>
    new TableRow({
      children: [
        textCell(title, { width: w[0], fill: "CBE6F7", bold: true }),
        textCell("Tanggal", { width: w[1], fill: "CBE6F7", bold: true }),
        textCell("Tanda tangan", { width: w[2], fill: "CBE6F7", bold: true }),
      ],
    });
  const row = (name: string, date: string) =>
    new TableRow({
      children: [textCell(name, { width: w[0] }), textCell(date, { width: w[1] }), textCell(" ", { width: w[2] })],
    });
  return new Table({
    width: { size: CONTENT_WIDTH_DXA, type: WidthType.DXA },
    columnWidths: w,
    rows: [
      hdr(cover.companyName || "Klien"),
      row(cover.reviewerClient || " ", " "),
      hdr(cover.vendorName || "Vendor"),
      row(`${cover.reviewerVendor || ""}${cover.reviewerVendorRole ? ` (${cover.reviewerVendorRole})` : ""}`, cover.reviewerDate || ""),
    ],
  });
}

// ---------- list of servers ----------
function listServerTable(pms: ParsedPM[]): Table {
  const w = [800, 3500, 2500, CONTENT_WIDTH_DXA - 6800];
  const head = new TableRow({
    tableHeader: true,
    children: ["No", "Hostname", "IP Address", "OS"].map((h, i) => textCell(h, { width: w[i], fill: COLOR_HEADER, bold: true })),
  });
  const body = pms.map(
    (p, i) =>
      new TableRow({
        children: [
          textCell(String(i + 1), { width: w[0], align: AlignmentType.CENTER }),
          textCell(p.hostname || "-", { width: w[1] }),
          textCell(p.ipAddress || "-", { width: w[2] }),
          textCell(p.osRelease || "-", { width: w[3] }),
        ],
      }),
  );
  return new Table({ width: { size: CONTENT_WIDTH_DXA, type: WidthType.DXA }, columnWidths: w, rows: [head, ...body] });
}

function resultSummaryTable(sums: SummaryRow[]): Table {
  const headers = ["Hostname", "Mountpoint", "Uptime", "Time & Date Sync", "CPU Usage", "Memory Usage", "Log Kill Memory", "Log Error"];
  const colBase = Math.floor(CONTENT_WIDTH_DXA / headers.length);
  const w = headers.map(() => colBase);
  const head = new TableRow({
    tableHeader: true,
    children: headers.map((h, i) => textCell(h, { width: w[i], fill: COLOR_HEADER, bold: true, align: AlignmentType.CENTER, size: 16 })),
  });
  const body = sums.map((s) => {
    const statuses: [StatusKind, StatusKind, StatusKind, StatusKind, StatusKind, StatusKind, StatusKind] = [
      s.mountpoint,
      s.uptime,
      s.timeSync,
      s.cpu,
      s.memory,
      s.logKillMemory,
      s.logError,
    ];
    return new TableRow({
      children: [
        textCell(s.hostname, { width: w[0], bold: true, size: 16 }),
        ...statuses.map((st, i) =>
          textCell(st || "-", {
            width: w[i + 1],
            fill: statusFill(st),
            align: AlignmentType.CENTER,
            bold: true,
            size: 16,
          }),
        ),
      ],
    });
  });
  return new Table({ width: { size: CONTENT_WIDTH_DXA, type: WidthType.DXA }, columnWidths: w, rows: [head, ...body] });
}

// ---------- per-host section table ----------
function sectionTable(section: PMSection): Table {
  const wLabel = 2400;
  const wSep = 400;
  const wValue = CONTENT_WIDTH_DXA - wLabel - wSep;

  const header = new TableRow({
    children: [
      new TableCell({
        borders: cellBorders,
        columnSpan: 3,
        shading: { fill: COLOR_HEADER, type: ShadingType.CLEAR, color: "auto" },
        margins: { top: 80, bottom: 80, left: 120, right: 120 },
        children: [new Paragraph({ children: [new TextRun({ text: section.title, bold: true, size: 22 })] })],
      }),
    ],
  });

  const body = section.rows.map((r) => {
    const fill = statusFill(r.status || "");
    const valueParas = r.value.includes("\n") ? multilineParas(r.value) : [new Paragraph({ children: [new TextRun({ text: r.value, size: 18, font: r.value.length > 40 ? "Consolas" : undefined })] })];
    return new TableRow({
      children: [
        textCell(r.label, { width: wLabel, bold: true, fill: "F7F7F7" }),
        textCell(":", { width: wSep, align: AlignmentType.CENTER }),
        cell(valueParas, { width: wValue, fill }),
      ],
    });
  });

  return new Table({
    width: { size: CONTENT_WIDTH_DXA, type: WidthType.DXA },
    columnWidths: [wLabel, wSep, wValue],
    rows: [header, ...body],
  });
}

function hostSection(pm: ParsedPM, index: number): (Paragraph | Table)[] {
  const out: (Paragraph | Table)[] = [];
  out.push(
    new Paragraph({
      heading: HeadingLevel.HEADING_2,
      children: [new TextRun({ text: `${index + 1}. ${pm.hostname}`, bold: true, size: 26 })],
      spacing: { before: 400, after: 200 },
    }),
  );
  // Big centered hostname banner
  out.push(
    new Table({
      width: { size: CONTENT_WIDTH_DXA, type: WidthType.DXA },
      columnWidths: [CONTENT_WIDTH_DXA],
      rows: [
        new TableRow({
          children: [
            new TableCell({
              borders: cellBorders,
              shading: { fill: COLOR_TITLE_BG, type: ShadingType.CLEAR, color: "auto" },
              margins: { top: 120, bottom: 120, left: 120, right: 120 },
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [new TextRun({ text: pm.hostname, bold: true, size: 26, color: "FFFFFF" })],
                }),
              ],
            }),
          ],
        }),
      ],
    }),
  );
  for (const s of pm.sections) {
    out.push(new Paragraph({ spacing: { before: 120 }, children: [new TextRun("")] }));
    out.push(sectionTable(s));
  }
  out.push(new Paragraph({ children: [new PageBreak()] }));
  return out;
}

// ---------- main entry ----------
async function _buildBlob(cover: CoverInput, pms: ParsedPM[]): Promise<Blob> {
  const summaries = pms.map(summarizePM);

  const heading = (text: string, level: (typeof HeadingLevel)[keyof typeof HeadingLevel] = HeadingLevel.HEADING_1) =>
    new Paragraph({
      heading: level,
      spacing: { before: 300, after: 200 },
      children: [new TextRun({ text, bold: true, size: level === HeadingLevel.HEADING_1 ? 32 : 26 })],
    });

  const bullet = (text: string) =>
    new Paragraph({
      bullet: { level: 0 },
      children: [new TextRun({ text, size: 20 })],
    });

  const bullets = (multiline: string) =>
    (multiline || "")
      .split(/\r?\n/)
      .map((l) => l.trim())
      .filter(Boolean)
      .map((l) => bullet(l.replace(/^[-•*]\s*/, "")));

  const children: (Paragraph | Table | TableOfContents)[] = [];

  buildCover(cover).forEach((p) => children.push(p));
  buildToc().forEach((c) => children.push(c));

  children.push(heading("Document Control"));
  children.push(docControlTable(cover));
  children.push(heading("Revision", HeadingLevel.HEADING_2));
  children.push(revisionTable());
  children.push(heading("List Of Activity", HeadingLevel.HEADING_2));
  children.push(activityTable());
  children.push(heading("Document Reviewer", HeadingLevel.HEADING_2));
  children.push(reviewerTable(cover));
  children.push(new Paragraph({ children: [new PageBreak()] }));

  children.push(heading("Overview"));
  children.push(heading("Executive Summary", HeadingLevel.HEADING_2));
  children.push(
    new Paragraph({
      alignment: AlignmentType.JUSTIFIED,
      children: [
        new TextRun({
          text:
            cover.executiveSummary ||
            `Dengan ini ${cover.vendorName || "vendor"} telah melaksanakan Kegiatan Perawatan Sistem Operasi ${cover.operatingSystem || ""} di ${cover.companyName || "klien"} yang berjumlah ${pms.length} Server.`,
          size: 22,
        }),
      ],
    }),
  );

  children.push(heading("List Server", HeadingLevel.HEADING_2));
  children.push(listServerTable(pms));

  children.push(heading("Ringkasan Hasil Preventive Maintenance", HeadingLevel.HEADING_2));
  children.push(resultSummaryTable(summaries));

  children.push(heading("Summary Conclusion", HeadingLevel.HEADING_2));
  bullets(cover.summaryConclusion).forEach((p) => children.push(p));

  children.push(heading("Recommendation", HeadingLevel.HEADING_2));
  bullets(cover.recommendation).forEach((p) => children.push(p));

  children.push(new Paragraph({ children: [new PageBreak()] }));

  children.push(heading("Lampiran Pekerjaan Preventive Maintenance"));
  pms.forEach((pm, i) => {
    hostSection(pm, i).forEach((n) => children.push(n));
  });

  const doc = new Document({
    creator: "ExcportCuy",
    title: cover.reportTitle,
    styles: {
      default: { document: { run: { font: "Calibri", size: 22 } } },
      paragraphStyles: [
        {
          id: "Heading1",
          name: "Heading 1",
          basedOn: "Normal",
          next: "Normal",
          quickFormat: true,
          run: { size: 32, bold: true, color: "1F1F1F" },
          paragraph: { spacing: { before: 360, after: 200 }, outlineLevel: 0 },
        },
        {
          id: "Heading2",
          name: "Heading 2",
          basedOn: "Normal",
          next: "Normal",
          quickFormat: true,
          run: { size: 26, bold: true, color: "1F1F1F" },
          paragraph: { spacing: { before: 240, after: 160 }, outlineLevel: 1 },
        },
        {
          id: "Heading3",
          name: "Heading 3",
          basedOn: "Normal",
          next: "Normal",
          quickFormat: true,
          run: { size: 22, bold: true, color: "1F1F1F" },
          paragraph: { spacing: { before: 200, after: 120 }, outlineLevel: 2 },
        },
      ],
    },
    numbering: {
      config: [
        {
          reference: "bullets",
          levels: [
            {
              level: 0,
              format: LevelFormat.BULLET,
              text: "•",
              alignment: AlignmentType.LEFT,
              style: { paragraph: { indent: { left: 720, hanging: 360 } } },
            },
          ],
        },
      ],
    },
    sections: [
      {
        properties: {
          page: {
            size: { width: 12240, height: 15840, orientation: PageOrientation.PORTRAIT },
            margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
          },
        },
        footers: {
          default: new Footer({
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [
                  new TextRun({ text: "Halaman ", size: 18 }),
                  new TextRun({ children: [PageNumber.CURRENT], size: 18 }),
                  new TextRun({ text: " dari ", size: 18 }),
                  new TextRun({ children: [PageNumber.TOTAL_PAGES], size: 18 }),
                ],
              }),
            ],
          }),
        },
        children: children as any,
      },
    ],
  });

  return Packer.toBlob(doc);
}

export async function buildDocxBlob(cover: CoverInput, pms: ParsedPM[]): Promise<Blob> {
  return _buildBlob(cover, pms);
}

export async function buildAndDownloadDocx(cover: CoverInput, pms: ParsedPM[], filename: string): Promise<Blob> {
  const blob = await _buildBlob(cover, pms);
  saveAs(blob, filename.endsWith(".docx") ? filename : `${filename}.docx`);
  return blob;
}

