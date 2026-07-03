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
  Header,
  Footer,
  PageNumber,
  ImageRun,
  HorizontalPositionAlign,
  VerticalPositionAlign,
  HorizontalPositionRelativeFrom,
  VerticalPositionRelativeFrom,
  TextWrappingType,
  TabStopType,
  TabStopPosition,
} from "docx";
import FileSaver from "file-saver";
const { saveAs } = FileSaver;

import coverBgUrl from "@/assets/cover-bg.jpeg?url";

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
  /** Optional cover logo (left / client) — data URL. */
  logoDataUrl?: string;
  /** Optional cover logo (right / vendor) — data URL. */
  logoRightDataUrl?: string;
  /** Optional multi-line client address for "Dibuat untuk". */
  clientAddress?: string;
  /** Optional multi-line vendor address for "Dibuat oleh". */
  vendorAddress?: string;
  /** Optional custom cover background image (data URL). If omitted, ships default. */
  coverBackgroundDataUrl?: string;
}

type ImgBytes = { data: Uint8Array; type: "png" | "jpg" | "gif" | "bmp" };
type ImgBytesSized = ImgBytes & { w: number; h: number };

// Decode a data URL into bytes + docx image type.
function decodeDataUrl(dataUrl: string): ImgBytes | null {
  const m = /^data:image\/(png|jpe?g|gif|bmp);base64,(.+)$/i.exec(dataUrl.trim());
  if (!m) return null;
  const ext = m[1].toLowerCase();
  const type = (ext === "jpeg" ? "jpg" : ext) as ImgBytes["type"];
  const bin = atob(m[2]);
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  return { data: bytes, type };
}

async function measureDataUrl(dataUrl: string): Promise<{ w: number; h: number }> {
  return new Promise((resolve) => {
    if (typeof Image === "undefined") return resolve({ w: 0, h: 0 });
    const im = new Image();
    im.onload = () => resolve({ w: im.naturalWidth, h: im.naturalHeight });
    im.onerror = () => resolve({ w: 0, h: 0 });
    im.src = dataUrl;
  });
}

async function decodeSized(dataUrl?: string): Promise<ImgBytesSized | null> {
  if (!dataUrl) return null;
  const raw = decodeDataUrl(dataUrl);
  if (!raw) return null;
  const dims = await measureDataUrl(dataUrl);
  return { ...raw, w: dims.w || 1, h: dims.h || 1 };
}

// Fit an image into a max box while preserving aspect ratio.
function fitBox(img: { w: number; h: number } | null, maxW: number, maxH: number) {
  if (!img || !img.w || !img.h) return { width: maxW, height: maxH };
  const r = Math.min(maxW / img.w, maxH / img.h);
  return { width: Math.max(1, Math.round(img.w * r)), height: Math.max(1, Math.round(img.h * r)) };
}

async function fetchBytes(url: string, type: ImgBytes["type"]): Promise<ImgBytes | null> {
  try {
    const res = await fetch(url);
    if (!res.ok) return null;
    return { data: new Uint8Array(await res.arrayBuffer()), type };
  } catch {
    return null;
  }
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
// Half-points: 28pt=56, 26pt=52, 22pt=44, 20pt=40, 18pt=36, 14pt=28, 11pt=22.
const A4_PX_W = 794; // ~8.27in @ 96dpi
const A4_PX_H = 1123; // ~11.69in @ 96dpi

function noBorders() {
  const b = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
  return { top: b, bottom: b, left: b, right: b };
}

function buildCover(
  cover: CoverInput,
  bg: ImgBytes | null,
  logoL: ImgBytesSized | null,
  logoR: ImgBytesSized | null,
): (Paragraph | Table)[] {
  const out: (Paragraph | Table)[] = [];

  // ---- Full-page background image, floating behind text ----
  if (bg) {
    out.push(
      new Paragraph({
        spacing: { before: 0, after: 0, line: 20 },
        children: [
          new ImageRun({
            type: bg.type,
            data: bg.data,
            transformation: { width: A4_PX_W, height: A4_PX_H },
            altText: { title: "Cover background", description: "Decorative cover background", name: "cover-bg" },
            floating: {
              horizontalPosition: {
                relative: HorizontalPositionRelativeFrom.PAGE,
                align: HorizontalPositionAlign.LEFT,
              },
              verticalPosition: {
                relative: VerticalPositionRelativeFrom.PAGE,
                align: VerticalPositionAlign.TOP,
              },
              behindDocument: true,
              zIndex: 0,
              allowOverlap: true,
              wrap: { type: TextWrappingType.NONE },
            },
          }),
        ],
      }),
    );
  }

  // ---- Header row: kop client logo (top-left) + kop vendor logo (top-right) ----
  const headerLogoCell = (logo: ImgBytesSized | null, align: (typeof AlignmentType)[keyof typeof AlignmentType]) =>
    new TableCell({
      borders: noBorders(),
      width: { size: Math.floor(CONTENT_WIDTH_DXA / 2), type: WidthType.DXA },
      verticalAlign: VerticalAlign.TOP,
      children: [
        new Paragraph({
          alignment: align,
          children: logo
            ? [
                new ImageRun({
                  type: logo.type,
                  data: logo.data,
                  transformation: fitBox(logo, 90, 70),
                  altText: { title: "Logo", description: "Header logo", name: "logo" },
                }),
              ]
            : [new TextRun("")],
        }),
      ],
    });

  out.push(
    new Table({
      width: { size: CONTENT_WIDTH_DXA, type: WidthType.DXA },
      columnWidths: [Math.floor(CONTENT_WIDTH_DXA / 2), Math.floor(CONTENT_WIDTH_DXA / 2)],
      rows: [
        new TableRow({
          children: [headerLogoCell(logoL, AlignmentType.LEFT), headerLogoCell(logoR, AlignmentType.RIGHT)],
        }),
      ],
    }),
  );

  // Spacer to push title block below the top decorative image area
  for (let i = 0; i < 12; i++) {
    out.push(new Paragraph({ children: [new TextRun("")] }));
  }

  // ---- Title block (right-aligned, white) ----
  const rightPar = (text: string, opts: { size: number; bold?: boolean; color?: string }) =>
    new Paragraph({
      alignment: AlignmentType.RIGHT,
      spacing: { before: 40, after: 40 },
      children: [new TextRun({ text, size: opts.size, bold: opts.bold, color: opts.color ?? "FFFFFF" })],
    });

  out.push(rightPar((cover.reportTitle || "LAPORAN PREVENTIVE MAINTENANCE").toUpperCase(), { size: 52, bold: true }));
  out.push(rightPar(`- ${cover.operatingSystem || "Red Hat Enterprise Linux"} -`, { size: 32 }));
  out.push(rightPar(`Periode ${cover.periode || "-"}`, { size: 32 }));
  out.push(rightPar(`No Contract : ${cover.contractNo || "-"}`, { size: 28 }));

  // Spacer before confidentiality box
  for (let i = 0; i < 4; i++) {
    out.push(new Paragraph({ children: [new TextRun("")] }));
  }

  // ---- Confidentiality box ----
  const boxBorder = { style: BorderStyle.SINGLE, size: 8, color: "2E7D32" };
  const confidentialBox = new Table({
    width: { size: CONTENT_WIDTH_DXA, type: WidthType.DXA },
    columnWidths: [CONTENT_WIDTH_DXA],
    rows: [
      new TableRow({
        children: [
          new TableCell({
            borders: { top: boxBorder, bottom: boxBorder, left: boxBorder, right: boxBorder },
            shading: { fill: "FFFFFF", type: ShadingType.CLEAR, color: "auto" },
            margins: { top: 200, bottom: 200, left: 260, right: 260 },
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { after: 160 },
                children: [new TextRun({ text: "PEMBERITAHUAN KERAHASIAAN", size: 22, bold: true, font: "Consolas" })],
              }),
              new Paragraph({
                alignment: AlignmentType.JUSTIFIED,
                children: [
                  new TextRun({
                    text:
                      `Material dalam dokumen ini dimiliki oleh ${cover.vendorName || "vendor"}. Dokumen ini diajukan kepada "${cover.companyName || "klien"}" untuk tujuan laporan. ` +
                      `Dengan penerimaan dokumen ini "${cover.companyName || "klien"}" telah setuju untuk terikat dengan sifat kerahasiaan laporan ini. Reproduksi pada distribusi dari setiap bagian ` +
                      `dari dokumen ini tidak diperkenankan tanpa persetujuan tertulis sebelumnya dari ${cover.vendorName || "vendor"}.`,
                    size: 18,
                    font: "Consolas",
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

  // Spacer before Dibuat untuk / Dibuat oleh
  for (let i = 0; i < 3; i++) {
    out.push(new Paragraph({ children: [new TextRun("")] }));
  }

  // ---- Dibuat untuk / Dibuat oleh two-column footer block ----
  const halfW = Math.floor(CONTENT_WIDTH_DXA / 2);
  const addressLines = (raw: string | undefined) =>
    (raw || "")
      .split(/\r?\n/)
      .map((l) => l.trim())
      .filter(Boolean);

  const madeCell = (
    heading: string,
    logo: ImgBytesSized | null,
    name: string,
    address: string[],
    align: (typeof AlignmentType)[keyof typeof AlignmentType],
  ) =>
    new TableCell({
      borders: noBorders(),
      width: { size: halfW, type: WidthType.DXA },
      verticalAlign: VerticalAlign.TOP,
      children: [
        new Paragraph({
          alignment: align,
          spacing: { after: 120 },
          children: [new TextRun({ text: heading, bold: true, size: 22 })],
        }),
        new Paragraph({
          alignment: align,
          spacing: { after: 120 },
          children: logo
            ? [
                new ImageRun({
                  type: logo.type,
                  data: logo.data,
                  transformation: fitBox(logo, 110, 90),
                  altText: { title: "Logo", description: "Party logo", name: "party-logo" },
                }),
              ]
            : [new TextRun("")],
        }),
        new Paragraph({
          alignment: align,
          spacing: { after: 60 },
          children: [new TextRun({ text: name, bold: true, size: 22 })],
        }),
        ...address.map(
          (line) =>
            new Paragraph({
              alignment: align,
              spacing: { after: 20 },
              children: [new TextRun({ text: line, size: 20 })],
            }),
        ),
      ],
    });

  out.push(
    new Table({
      width: { size: CONTENT_WIDTH_DXA, type: WidthType.DXA },
      columnWidths: [halfW, halfW],
      rows: [
        new TableRow({
          children: [
            madeCell(
              "Dibuat untuk",
              logoL,
              cover.companyName || "-",
              addressLines(cover.clientAddress),
              AlignmentType.LEFT,
            ),
            madeCell(
              "Dibuat oleh :",
              logoR,
              cover.vendorName || "-",
              addressLines(cover.vendorAddress),
              AlignmentType.RIGHT,
            ),
          ],
        }),
      ],
    }),
  );

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

  // Load default cover background (bundled) unless user provided one.
  const bg = cover.coverBackgroundDataUrl
    ? decodeDataUrl(cover.coverBackgroundDataUrl)
    : await fetchBytes(coverBgUrl, "jpg");

  buildCover(cover, bg).forEach((p) => children.push(p));
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
            size: { width: 11906, height: 16838, orientation: PageOrientation.PORTRAIT },
            margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
          },
          titlePage: true,
        },
        headers: {
          first: new Header({
            children: [new Paragraph({ children: [new TextRun("")] })],
          }),
          default: new Header({
            children: [
              new Paragraph({
                tabStops: [
                  { type: TabStopType.RIGHT, position: TabStopPosition.MAX },
                ],
                border: {
                  bottom: { style: BorderStyle.SINGLE, size: 6, color: "1F3864", space: 1 },
                },
                children: [
                  new TextRun({ text: cover.companyName || "", bold: true, size: 18, color: "1F3864" }),
                  new TextRun({ text: "\t" }),
                  new TextRun({ text: `${cover.reportTitle || "Laporan PM"} - ${cover.periode || ""}`, size: 18, color: "1F3864" }),
                ],
              }),
            ],
          }),
        },
        footers: {
          first: new Footer({
            children: [new Paragraph({ children: [new TextRun("")] })],
          }),
          default: new Footer({
            children: [
              new Paragraph({
                tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
                children: [
                  new TextRun({ text: `${cover.vendorName || "Vendor"} - PM ${cover.operatingSystem || ""}`, size: 18, color: "595959" }),
                  new TextRun({ text: "\t" }),
                  new TextRun({ children: [PageNumber.CURRENT], size: 18, bold: true, color: "1F3864" }),
                  new TextRun({ text: " / ", size: 18, color: "595959" }),
                  new TextRun({ children: [PageNumber.TOTAL_PAGES], size: 18, color: "595959" }),
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

