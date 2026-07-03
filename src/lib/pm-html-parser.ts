// Parses a Preventive Maintenance HTML report into structured data.
// Supports the tg-* table format seen in PM_DEV-HR-*.html samples.

export type StatusKind = "OK" | "Warning" | "Critical" | "Info" | "";

export interface PMRow {
  label: string;
  value: string;
  status?: StatusKind;
}

export interface PMSection {
  title: string;
  rows: PMRow[];
  status?: StatusKind; // rolled up from any Status row in this section
}

export interface ParsedPM {
  hostname: string;
  ipAddress: string;
  osRelease: string;
  kernel: string;
  sections: PMSection[];
  raw: { general: Record<string, string> };
}

function classifyStatus(cell: HTMLElement): StatusKind {
  const cls = cell.className || "";
  if (cls.includes("tg-qd4f")) return "OK";
  if (cls.includes("tg-b9d2")) return "Warning";
  if (cls.includes("tg-xuni")) return "Critical";
  const txt = (cell.textContent || "").trim().toLowerCase();
  if (txt === "ok") return "OK";
  if (txt.startsWith("warn")) return "Warning";
  if (txt.startsWith("crit")) return "Critical";
  return "";
}

function cellText(cell: HTMLElement | null | undefined): string {
  if (!cell) return "";
  // Preserve line breaks from <pre> tags
  const pres = cell.querySelectorAll("pre");
  if (pres.length) {
    return Array.from(cell.childNodes)
      .map((n) => (n.nodeType === 3 ? (n.textContent ?? "") : (n as HTMLElement).textContent ?? ""))
      .join("")
      .replace(/\r/g, "")
      .trim();
  }
  return (cell.textContent ?? "").replace(/\s+/g, " ").trim();
}

export function parsePMHtml(html: string, fallbackName = "Unknown"): ParsedPM {
  const doc = new DOMParser().parseFromString(html, "text/html");
  const table = doc.querySelector("table");
  const result: ParsedPM = {
    hostname: fallbackName,
    ipAddress: "",
    osRelease: "",
    kernel: "",
    sections: [],
    raw: { general: {} },
  };
  if (!table) return result;

  // Hostname often in <th class="tg-ks93"> at top
  const th = table.querySelector("th");
  if (th) result.hostname = cellText(th as HTMLElement) || fallbackName;

  let current: PMSection | null = null;
  const rows = table.querySelectorAll("tr");
  rows.forEach((tr) => {
    const tds = Array.from(tr.querySelectorAll("td")) as HTMLElement[];
    if (tds.length === 0) return;

    // Section header: single cell with colspan=3, tg-k57h / tg-rujl / tg-lft
    const first = tds[0];
    const colspan = parseInt(first.getAttribute("colspan") || "1", 10);
    const cls = first.className || "";
    const isHeader =
      colspan >= 3 &&
      (cls.includes("tg-k57h") || cls.includes("tg-rujl") || cls.includes("tg-lft"));

    if (isHeader) {
      const title = cellText(first).replace(/\s*:\s*$/, "").trim();
      current = { title, rows: [] };
      result.sections.push(current);
      return;
    }

    if (tds.length >= 3) {
      const label = cellText(tds[0]);
      const value = cellText(tds[2]);
      const status = classifyStatus(tds[2]);
      const row: PMRow = { label, value, status };
      if (current) {
        current.rows.push(row);
        if (label.toLowerCase() === "status" && status) {
          current.status = status;
        }
      } else {
        // pre-first-section rows (rare)
      }

      // Capture common headline fields regardless of section
      const key = label.toLowerCase();
      if (key === "hostname") result.hostname = value || result.hostname;
      if (key === "ip address") result.ipAddress = value.split(/\s+/).find((t) => /\d+\.\d+\.\d+\.\d+/.test(t)) || value;
      if (key === "os release") result.osRelease = value;
      if (key === "kernel version") result.kernel = value;
      result.raw.general[label] = value;
    }
  });

  return result;
}

// Aggregate section statuses for the top-level "Result Preventive Maintenance" table.
export interface SummaryRow {
  hostname: string;
  mountpoint: StatusKind;
  uptime: StatusKind;
  timeSync: StatusKind;
  cpu: StatusKind;
  memory: StatusKind;
  logKillMemory: StatusKind;
  logError: StatusKind;
}

function findSectionStatus(pm: ParsedPM, matcher: (t: string) => boolean): StatusKind {
  const s = pm.sections.find((x) => matcher(x.title.toLowerCase()));
  return (s?.status || "") as StatusKind;
}

export function summarizePM(pm: ParsedPM): SummaryRow {
  return {
    hostname: pm.hostname,
    mountpoint: findSectionStatus(pm, (t) => t.includes("disk") || t.includes("mountpoint")),
    uptime: findSectionStatus(pm, (t) => t.includes("uptime")),
    timeSync: findSectionStatus(pm, (t) => t.includes("time") && t.includes("sync")),
    cpu: findSectionStatus(pm, (t) => t.includes("cpu")),
    memory: findSectionStatus(pm, (t) => t.includes("memory") && !t.includes("kill")),
    logKillMemory: findSectionStatus(pm, (t) => t.includes("kill")),
    logError: findSectionStatus(pm, (t) => t.includes("error log")),
  };
}
