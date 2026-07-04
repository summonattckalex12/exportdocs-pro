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
  TableLayoutType,
} from "docx";



import coverBgUrl from "@/assets/cover-bg.jpeg?url";
import miiLogoUrl from "@/assets/mii-logo.png?url";

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

// Word documents are XML. PM logs can contain invisible terminal/audit control
// chars (ASCII 0x00-0x1F, 0x7F-0x9F) that are illegal/unsafe in XML and make
// Microsoft Word show "Word experienced an error trying to open the file".
function safeDocxText(value: unknown): string {
  const input = String(value ?? "");
  let out = "";
  for (let i = 0; i < input.length; i += 1) {
    const code = input.charCodeAt(i);

    if (code >= 0xd800 && code <= 0xdbff) {
      const next = input.charCodeAt(i + 1);
      if (next >= 0xdc00 && next <= 0xdfff) {
        out += input[i] + input[i + 1];
        i += 1;
      }
      continue;
    }
    if (code >= 0xdc00 && code <= 0xdfff) continue;

    const allowedXml = code === 0x09 || code === 0x0a || code === 0x0d || (code >= 0x20 && code <= 0xd7ff) || (code >= 0xe000 && code <= 0xfffd);
    const unsafeWordControl = code >= 0x7f && code <= 0x9f;
    if (allowedXml && !unsafeWordControl) out += input[i];
  }
  return out;
}

function P(text: string, opts: { bold?: boolean; size?: number; color?: string; align?: (typeof AlignmentType)[keyof typeof AlignmentType] } = {}) {
  return new Paragraph({
    alignment: opts.align,
    children: [new TextRun({ text: safeDocxText(text), bold: opts.bold, size: opts.size ?? 20, color: opts.color })],
  });
}

function multilineParas(text: string, opts: { bold?: boolean; size?: number; color?: string } = {}) {
  const lines = safeDocxText(text).split(/\r?\n/);
  return lines.map(
    (line) =>
      new Paragraph({
        children: [
          new TextRun({ text: safeDocxText(line), bold: opts.bold, size: opts.size ?? 18, color: opts.color, font: "Consolas" }),
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
            text: safeDocxText(text),
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
        spacing: { before: 0, after: 0 },
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

  // Spacer to push title block below the top decorative image area of the background.
  // Cover intentionally has NO logos at the top — logos live in the bottom block.
  // 10x enter agar posisi judul turun mengikuti template (blue panel).
  for (let i = 0; i < 10; i++) {
    out.push(new Paragraph({ children: [new TextRun("")] }));
  }


  // ---- Title block (right-aligned, white on the blue panel of the bg) ----
  const rightPar = (text: string, opts: { size: number; bold?: boolean; color?: string; spaceAfter?: number }) =>
    new Paragraph({
      alignment: AlignmentType.RIGHT,
      spacing: { before: 20, after: opts.spaceAfter ?? 40 },
      children: [new TextRun({ text, size: opts.size, bold: opts.bold, color: opts.color ?? "FFFFFF" })],
    });

  out.push(rightPar((cover.reportTitle || "LAPORAN PREVENTIVE MAINTENANCE").toUpperCase(), { size: 40, bold: true, spaceAfter: 80 }));
  out.push(rightPar(`- ${cover.operatingSystem || "Red Hat Enterprise Linux"} -`, { size: 24 }));
  out.push(rightPar(`Periode ${cover.periode || "-"}`, { size: 24 }));
  out.push(rightPar(`No Contract : ${cover.contractNo || "-"}`, { size: 20, spaceAfter: 80 }));

  // Spacer before confidentiality box (push it to middle-lower area)
  for (let i = 0; i < 6; i++) {
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
            margins: { top: 160, bottom: 160, left: 220, right: 220 },
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { after: 140 },
                children: [new TextRun({ text: "PEMBERITAHUAN KERAHASIAAN", size: 20, bold: true, font: "Courier New" })],
              }),
              new Paragraph({
                alignment: AlignmentType.JUSTIFIED,
                children: [
                  new TextRun({
                    text:
                      `Material dalam dokumen ini dimiliki oleh ${cover.vendorName || "vendor"}. Dokumen ini diajukan kepada "${cover.companyName || "klien"}" untuk tujuan laporan. ` +
                      `Dengan penerimaan dokumen ini "${cover.companyName || "klien"}" telah setuju untuk terikat dengan sifat kerahasiaan laporan ini. Reproduksi pada distribusi dari setiap bagian ` +
                      `dari dokumen ini tidak diperkenankan tanpa persetujuan tertulis sebelumnya dari ${cover.vendorName || "vendor"}.`,
                    size: 16,
                    font: "Courier New",
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
  out.push(new Paragraph({ children: [new TextRun("")] }));

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
          spacing: { after: 100 },
          children: [new TextRun({ text: heading, bold: true, size: 22 })],
        }),
        new Paragraph({
          alignment: align,
          spacing: { after: 100 },
          children: logo
            ? [
                new ImageRun({
                  type: logo.type,
                  data: logo.data,
                  transformation: fitBox(logo, 110, 80),
                  altText: { title: "Logo", description: "Party logo", name: "party-logo" },
                }),
              ]
            : [new TextRun("")],
        }),
        new Paragraph({
          alignment: align,
          spacing: { after: 40 },
          children: [new TextRun({ text: name, bold: true, size: 22 })],
        }),
        ...address.map(
          (line) =>
            new Paragraph({
              alignment: align,
              spacing: { after: 20 },
              children: [new TextRun({ text: line, size: 18 })],
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
function buildTocEntries(pms: ParsedPM[]) {
  const overviewPage = 4;
  const lampiranPage = 5;
  return [
    { title: "Document Control", level: 1, page: 3 },
    { title: "Revision", level: 2, page: 3 },
    { title: "List Of Activity", level: 2, page: 3 },
    { title: "Document Reviewer", level: 2, page: 3 },
    { title: "Overview", level: 1, page: overviewPage },
    { title: "Executive Summary", level: 2, page: overviewPage },
    { title: "List Server", level: 2, page: overviewPage },
    { title: "Ringkasan Hasil Preventive Maintenance", level: 2, page: overviewPage },
    { title: "Keterangan Threshold Status", level: 2, page: overviewPage },
    { title: "Summary Conclusion", level: 2, page: overviewPage },
    { title: "Recommendation", level: 2, page: overviewPage },

    { title: "Lampiran Pekerjaan Preventive Maintenance", level: 1, page: lampiranPage },
    ...pms.map((pm, i) => ({ title: `${i + 1}. ${safeDocxText(pm.hostname || `Server ${i + 1}`)}`, level: 2, page: lampiranPage + i })),
  ];
}

function buildToc(pms: ParsedPM[]): (Paragraph | TableOfContents)[] {
  const cachedEntries = buildTocEntries(pms);
  return [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 200 },
      children: [new TextRun({ text: "Table of Contents", size: 32, bold: true })],
    }),
    new TableOfContents("Contents", {
      hyperlink: true,
      headingStyleRange: "1-3",
      useAppliedParagraphOutlineLevel: true,
      beginDirty: true,
      cachedEntries,
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

// ---------- threshold explanation table ----------
function thresholdTable(): Table {
  const wNo = 800;
  const wItem = 2000;
  const wRest = CONTENT_WIDTH_DXA - wNo - wItem;
  const wWarn = Math.floor(wRest / 2);
  const wCrit = wRest - wWarn;
  const w = [wNo, wItem, wWarn, wCrit];
  const headerBg = "CBE6F7";

  const topRow = new TableRow({
    tableHeader: true,
    children: [
      textCell("No", { width: wNo, fill: headerBg, bold: true, align: AlignmentType.CENTER }),
      textCell("Item", { width: wItem, fill: headerBg, bold: true, align: AlignmentType.CENTER }),
      new TableCell({
        borders: cellBorders,
        columnSpan: 2,
        shading: { fill: headerBg, type: ShadingType.CLEAR, color: "auto" },
        margins: { top: 80, bottom: 80, left: 120, right: 120 },
        verticalAlign: VerticalAlign.CENTER,
        children: [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new TextRun({ text: "Threshold", bold: true, size: 20 })],
          }),
        ],
      }),
    ],
  });

  const subRow = new TableRow({
    tableHeader: true,
    children: [
      textCell("", { width: wNo, fill: headerBg }),
      textCell("", { width: wItem, fill: headerBg }),
      textCell("Warning", { width: wWarn, fill: headerBg, bold: true, align: AlignmentType.CENTER }),
      textCell("Critical", { width: wCrit, fill: headerBg, bold: true, align: AlignmentType.CENTER }),
    ],
  });

  type Row = { item: string; warn: string; warnFill?: string; crit: string; critFill?: string; critColor?: string };
  const rows: Row[] = [
    { item: "Disk", warn: "80%", warnFill: COLOR_WARN, crit: "90%", critFill: COLOR_CRIT, critColor: "FFFFFF" },
    { item: "CPU", warn: "80%", warnFill: COLOR_WARN, crit: "90%", critFill: COLOR_CRIT, critColor: "FFFFFF" },
    { item: "Memory", warn: "80%", warnFill: COLOR_WARN, crit: "90%", critFill: COLOR_CRIT, critColor: "FFFFFF" },
    { item: "Uptime", warn: "180 days", warnFill: COLOR_WARN, crit: "", critFill: "000000" },
  ];
  const body = rows.map(
    (r, i) =>
      new TableRow({
        children: [
          textCell(String(i + 1), { width: w[0], align: AlignmentType.CENTER }),
          textCell(r.item, { width: w[1] }),
          textCell(r.warn, { width: w[2], fill: r.warnFill, align: AlignmentType.CENTER, bold: true }),
          textCell(r.crit, { width: w[3], fill: r.critFill, align: AlignmentType.CENTER, bold: true, color: r.critColor }),
        ],
      }),
  );
  return new Table({
    width: { size: CONTENT_WIDTH_DXA, type: WidthType.DXA },
    columnWidths: w,
    rows: [topRow, subRow, ...body],
  });
}

// ---------- warning/critical detail report ----------
function warningCriticalReport(pms: ParsedPM[]): (Paragraph | Table)[] {
  const out: (Paragraph | Table)[] = [];
  const w = [800, 2400, 2600, CONTENT_WIDTH_DXA - 5800 - 1200, 1200];
  const head = new TableRow({
    tableHeader: true,
    children: ["No", "Hostname", "Section", "Temuan", "Status"].map((h, i) =>
      textCell(h, { width: w[i], fill: COLOR_HEADER, bold: true, align: AlignmentType.CENTER }),
    ),
  });

  const bodyRows: TableRow[] = [];
  let idx = 0;
  for (const pm of pms) {
    for (const sec of pm.sections) {
      for (const row of sec.rows) {
        if (row.status === "Warning" || row.status === "Critical") {
          idx += 1;
          const val = safeDocxText(row.value || "");
          const short = val.length > 200 ? val.slice(0, 200) + "…" : val;
          bodyRows.push(
            new TableRow({
              children: [
                textCell(String(idx), { width: w[0], align: AlignmentType.CENTER, size: 18 }),
                textCell(safeDocxText(pm.hostname), { width: w[1], bold: true, size: 18 }),
                textCell(safeDocxText(sec.title), { width: w[2], size: 18 }),
                cell(
                  [
                    new Paragraph({ children: [new TextRun({ text: safeDocxText(row.label), bold: true, size: 18 })] }),
                    new Paragraph({ children: [new TextRun({ text: short, size: 16, font: "Consolas" })] }),
                  ],
                  { width: w[3] },
                ),
                textCell(row.status || "", {
                  width: w[4],
                  fill: statusFill(row.status || ""),
                  align: AlignmentType.CENTER,
                  bold: true,
                  size: 18,
                  color: row.status === "Critical" ? "FFFFFF" : undefined,
                }),
              ],
            }),
          );
        }
      }
    }
  }

  if (bodyRows.length === 0) {
    out.push(
      new Paragraph({
        spacing: { before: 100, after: 100 },
        children: [
          new TextRun({
            text: "Tidak ditemukan temuan berstatus Warning atau Critical pada seluruh server. Semua parameter dalam kondisi OK.",
            italics: true,
            size: 20,
          }),
        ],
      }),
    );
    return out;
  }

  out.push(
    new Paragraph({
      spacing: { before: 80, after: 120 },
      children: [
        new TextRun({
          text: `Total temuan Warning/Critical: ${bodyRows.length} item. Berikut rinciannya:`,
          size: 20,
        }),
      ],
    }),
  );
  out.push(
    new Table({
      width: { size: CONTENT_WIDTH_DXA, type: WidthType.DXA },
      columnWidths: w,
      layout: TableLayoutType.FIXED,
      rows: [head, ...bodyRows],
    }),
  );
  return out;
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
        children: [new Paragraph({ children: [new TextRun({ text: safeDocxText(section.title), bold: true, size: 22 })] })],
      }),
    ],
  });

  const body = section.rows.map((r) => {
    const fill = statusFill(r.status || "");
    const val = safeDocxText(r.value || "");
    const lines = val.split(/\r?\n/);
    const maxLine = lines.reduce((m, l) => Math.max(m, l.length), 0);
    const isMultiline = lines.length > 1;
    const isWide = maxLine > 60;
    // Shrink monospace content that would otherwise overflow the fixed value column.
    const preSize = maxLine > 110 ? 12 : maxLine > 90 ? 13 : maxLine > 70 ? 14 : 16;
    const valueParas = isMultiline
      ? multilineParas(val, { size: preSize })
      : [new Paragraph({ children: [new TextRun({ text: val, size: isWide ? 14 : 18, font: isWide || val.length > 40 ? "Consolas" : undefined })] })];
    return new TableRow({
      children: [
        textCell(safeDocxText(r.label), { width: wLabel, bold: true, fill: "F7F7F7" }),
        textCell(":", { width: wSep, align: AlignmentType.CENTER }),
        cell(valueParas, { width: wValue, fill }),
      ],
    });
  });

  return new Table({
    width: { size: CONTENT_WIDTH_DXA, type: WidthType.DXA },
    columnWidths: [wLabel, wSep, wValue],
    layout: TableLayoutType.FIXED,
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

  // Vendor logo (MII) is bundled & fixed — user can override via logoRightDataUrl.
  const miiBundled = await fetchBytes(miiLogoUrl, "png");
  const miiSized: ImgBytesSized | null = miiBundled ? { ...miiBundled, w: 208, h: 156 } : null;

  const [logoLUser, logoRUser] = await Promise.all([
    decodeSized(cover.logoDataUrl),
    decodeSized(cover.logoRightDataUrl),
  ]);
  const logoL = logoLUser; // client (bottom-left only)
  const logoR = logoRUser ?? miiSized; // MII fixed unless overridden

  buildCover(cover, bg, logoL, logoR).forEach((p) => children.push(p));
  buildToc(pms).forEach((c) => children.push(c));


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

  children.push(heading("Keterangan Threshold Status", HeadingLevel.HEADING_2));
  children.push(
    new Paragraph({
      alignment: AlignmentType.JUSTIFIED,
      spacing: { after: 120 },
      children: [
        new TextRun({
          text:
            "Tabel di bawah ini menjelaskan ambang batas (threshold) yang digunakan untuk menentukan status OK, Warning, dan Critical pada masing-masing parameter pemeriksaan Preventive Maintenance.",
          size: 20,
        }),
      ],
    }),
  );
  children.push(thresholdTable());

  // Report Summary (Warning & Critical) sekarang di-generate sebagai file terpisah
  // via buildAndDownloadReportSummary() untuk memudahkan analisa data.



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
    features: { updateFields: true },
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
        {
          id: "TOC1",
          name: "TOC 1",
          basedOn: "Normal",
          next: "Normal",
          run: { size: 22 },
          paragraph: { spacing: { after: 80 } },
        },
        {
          id: "TOC2",
          name: "TOC 2",
          basedOn: "Normal",
          next: "Normal",
          run: { size: 21 },
          paragraph: { indent: { left: 360 }, spacing: { after: 60 } },
        },
        {
          id: "TOC3",
          name: "TOC 3",
          basedOn: "Normal",
          next: "Normal",
          run: { size: 20 },
          paragraph: { indent: { left: 720 }, spacing: { after: 40 } },
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
              new Table({
                width: { size: CONTENT_WIDTH_DXA, type: WidthType.DXA },
                columnWidths: [1400, CONTENT_WIDTH_DXA - 2800, 1400],
                rows: [
                  new TableRow({
                    children: [
                      new TableCell({
                        borders: noBorders(),
                        width: { size: 1400, type: WidthType.DXA },
                        verticalAlign: VerticalAlign.CENTER,
                        children: [
                          new Paragraph({
                            alignment: AlignmentType.LEFT,
                            children: logoL
                              ? [
                                  new ImageRun({
                                    type: logoL.type,
                                    data: logoL.data,
                                    transformation: fitBox(logoL, 60, 40),
                                    altText: { title: "Client", description: "Client logo", name: "client-logo" },
                                  }),
                                ]
                              : [new TextRun({ text: cover.companyName || "", bold: true, size: 16, color: "1F3864" })],
                          }),
                        ],
                      }),
                      new TableCell({
                        borders: noBorders(),
                        width: { size: CONTENT_WIDTH_DXA - 2800, type: WidthType.DXA },
                        verticalAlign: VerticalAlign.CENTER,
                        children: [
                          new Paragraph({
                            alignment: AlignmentType.CENTER,
                            children: [
                              new TextRun({
                                text: `${cover.reportTitle || "Laporan PM"} — ${cover.periode || ""}`,
                                size: 18,
                                color: "1F3864",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new TableCell({
                        borders: noBorders(),
                        width: { size: 1400, type: WidthType.DXA },
                        verticalAlign: VerticalAlign.CENTER,
                        children: [
                          new Paragraph({
                            alignment: AlignmentType.RIGHT,
                            children: logoR
                              ? [
                                  new ImageRun({
                                    type: logoR.type,
                                    data: logoR.data,
                                    transformation: fitBox(logoR, 60, 40),
                                    altText: { title: "Vendor", description: "Vendor logo", name: "vendor-logo" },
                                  }),
                                ]
                              : [new TextRun("")],
                          }),
                        ],
                      }),
                    ],
                  }),
                ],
              }),
              new Paragraph({
                spacing: { before: 40, after: 0 },
                border: {
                  bottom: { style: BorderStyle.SINGLE, size: 6, color: "1F3864", space: 1 },
                },
                children: [new TextRun("")],
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
                  new TextRun({ text: `${cover.companyName || "Customer"} - PM ${cover.operatingSystem || ""}`, size: 18, color: "1F1F1F", bold: true }),
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
  const raw = await _buildBlob(cover, pms);
  // Force the correct DOCX MIME type — strict readers (Word Mobile, WPS,
  // Google Docs, iOS Files) reject files served as octet-stream or zip.
  return new Blob([await raw.arrayBuffer()], {
    type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  });
}

// Sanitize filename for cross-OS compatibility (Windows/macOS/Linux/Android/iOS).
function safeFilename(name: string): string {
  const base = (name || "document").replace(/\.docx$/i, "");
  const cleaned =
    base
      .replace(/[\\/:*?"<>|\u0000-\u001F]/g, "_")
      .replace(/\s+/g, " ")
      .trim()
      .slice(0, 120) || "document";
  return `${cleaned}.docx`;
}

// Cross-device download: Chrome/Edge/Firefox/Safari desktop, Android, iOS Safari.
function downloadBlob(blob: Blob, filename: string) {
  const name = safeFilename(filename);
  const nav = typeof navigator !== "undefined" ? navigator : undefined;

  const anyNav = nav as unknown as { msSaveOrOpenBlob?: (b: Blob, n: string) => void };
  if (anyNav?.msSaveOrOpenBlob) {
    anyNav.msSaveOrOpenBlob(blob, name);
    return;
  }

  const url = URL.createObjectURL(blob);
  const ua = nav?.userAgent || "";
  const isIOS = /iPad|iPhone|iPod/.test(ua) && !(nav as unknown as { MSStream?: unknown })?.MSStream;

  const a = document.createElement("a");
  a.href = url;
  a.download = name;
  a.rel = "noopener";
  // iOS Safari ignores `download`; opening in a new tab lets the user "Save to Files".
  if (isIOS) a.target = "_blank";
  a.style.display = "none";
  document.body.appendChild(a);
  a.click();
  setTimeout(() => {
    try {
      document.body.removeChild(a);
    } catch {
      /* noop */
    }
    URL.revokeObjectURL(url);
  }, 4000);
}

export async function buildAndDownloadDocx(cover: CoverInput, pms: ParsedPM[], filename: string): Promise<Blob> {
  const blob = await buildDocxBlob(cover, pms);
  downloadBlob(blob, filename);
  return blob;
}




// ============================================================================
// Report Summary (Warning & Critical) — file DOCX terpisah untuk analisa data
// ============================================================================
function logErrorSections(pm: ParsedPM): PMSection[] {
  return pm.sections.filter((s) => {
    const t = s.title.toLowerCase();
    return t.includes("error log") || t.includes("log error") || t.includes("kill");
  });
}

async function _buildReportSummaryBlob(cover: CoverInput, pms: ParsedPM[]): Promise<Blob> {
  const children: (Paragraph | Table)[] = [];

  children.push(
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 120 },
      children: [new TextRun({ text: "REPORT SUMMARY — WARNING & CRITICAL", bold: true, size: 32, color: "1F1F1F" })],
    }),
  );
  children.push(
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 240 },
      children: [
        new TextRun({
          text: `${cover.companyName || "-"} · ${cover.operatingSystem || ""} · Periode ${cover.periode || "-"}`,
          size: 20,
          color: "595959",
        }),
      ],
    }),
  );

  children.push(
    new Paragraph({
      heading: HeadingLevel.HEADING_2,
      spacing: { before: 120, after: 120 },
      children: [new TextRun({ text: "Ringkasan Temuan (Warning & Critical)", bold: true, size: 26 })],
    }),
  );
  warningCriticalReport(pms).forEach((c) => children.push(c));

  children.push(new Paragraph({ children: [new PageBreak()] }));
  children.push(
    new Paragraph({
      heading: HeadingLevel.HEADING_2,
      spacing: { before: 120, after: 120 },
      children: [new TextRun({ text: "Isian Log Error per Server", bold: true, size: 26 })],
    }),
  );

  let hasAnyLog = false;
  for (const pm of pms) {
    const logs = logErrorSections(pm);
    if (logs.length === 0) continue;
    hasAnyLog = true;
    children.push(
      new Paragraph({
        heading: HeadingLevel.HEADING_3,
        spacing: { before: 200, after: 100 },
        children: [new TextRun({ text: safeDocxText(pm.hostname || "-"), bold: true, size: 22 })],
      }),
    );
    for (const s of logs) {
      children.push(new Paragraph({ spacing: { before: 60 }, children: [new TextRun("")] }));
      children.push(sectionTable(s));
    }
  }
  if (!hasAnyLog) {
    children.push(
      new Paragraph({
        spacing: { before: 100 },
        children: [new TextRun({ text: "Tidak ada bagian Log Error / Kill Memory yang tercatat.", italics: true, size: 20 })],
      }),
    );
  }

  const doc = new Document({
    creator: "ExcportCuy",
    title: `${cover.reportTitle || "Report Summary"} — Warning & Critical`,
    styles: {
      default: { document: { run: { font: "Calibri", size: 22 } } },
      paragraphStyles: [
        { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { size: 32, bold: true, color: "1F1F1F" },
          paragraph: { spacing: { before: 300, after: 200 }, outlineLevel: 0 } },
        { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { size: 26, bold: true, color: "1F1F1F" },
          paragraph: { spacing: { before: 240, after: 160 }, outlineLevel: 1 } },
        { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { size: 22, bold: true, color: "1F1F1F" },
          paragraph: { spacing: { before: 200, after: 120 }, outlineLevel: 2 } },
      ],
    },
    sections: [
      {
        properties: {
          page: {
            size: { width: 11906, height: 16838, orientation: PageOrientation.PORTRAIT },
            margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 },
          },
        },
        footers: {
          default: new Footer({
            children: [
              new Paragraph({
                tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
                children: [
                  new TextRun({ text: `${cover.companyName || "Customer"} — Report Summary`, size: 18, bold: true, color: "1F1F1F" }),
                  new TextRun({ text: "\t" }),
                  new TextRun({ children: [PageNumber.CURRENT], size: 18, bold: true }),
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

export async function buildReportSummaryBlob(cover: CoverInput, pms: ParsedPM[]): Promise<Blob> {
  const raw = await _buildReportSummaryBlob(cover, pms);
  return new Blob([await raw.arrayBuffer()], {
    type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  });
}

export async function buildAndDownloadReportSummary(cover: CoverInput, pms: ParsedPM[], filename: string): Promise<Blob> {
  const blob = await buildReportSummaryBlob(cover, pms);
  downloadBlob(blob, filename);
  return blob;
}



