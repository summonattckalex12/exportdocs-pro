import { createFileRoute } from "@tanstack/react-router";
import { useEffect, useMemo, useRef, useState } from "react";
import { toast } from "sonner";
import { format } from "date-fns";
import JSZip from "jszip";
import {
  DndContext,
  closestCenter,
  KeyboardSensor,
  PointerSensor,
  useSensor,
  useSensors,
  type DragEndEvent,
} from "@dnd-kit/core";
import {
  SortableContext,
  sortableKeyboardCoordinates,
  useSortable,
  verticalListSortingStrategy,
  arrayMove,
} from "@dnd-kit/sortable";
import { CSS } from "@dnd-kit/utilities";
import { DragonLogo } from "@/components/DragonLogo";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Textarea } from "@/components/ui/textarea";
import { Card, CardContent, CardHeader, CardTitle, CardDescription } from "@/components/ui/card";
import { Popover, PopoverContent, PopoverTrigger } from "@/components/ui/popover";
import { Calendar } from "@/components/ui/calendar";
import { cn } from "@/lib/utils";
import { parsePMHtml, summarizePM, type ParsedPM, type StatusKind } from "@/lib/pm-html-parser";
import { buildAndDownloadDocx, buildAndDownloadReportSummary, type CoverInput } from "@/lib/docx-exporter";
import { archiveExport } from "@/lib/exports.functions";
import { untar, ungzipToTar } from "@/lib/tar";
import {
  Upload,
  FileText,
  Trash2,
  Download,
  Flame,
  ImagePlus,
  X,
  CalendarIcon,
  GripVertical,
  FileArchive,
} from "lucide-react";

export const Route = createFileRoute("/")({
  head: () => ({
    meta: [
      { title: "ExcportCuy — Convert HTML PM Reports to Word" },
      { name: "description", content: "Ubah laporan HTML Preventive Maintenance menjadi dokumen Word yang rapi, dengan cover, TOC, dan lampiran per server." },
      { property: "og:title", content: "ExcportCuy — HTML → Word Exporter" },
      { property: "og:description", content: "Convert PM HTML reports to a beautiful, structured Word document." },
      { property: "og:type", content: "website" },
      { name: "twitter:card", content: "summary_large_image" },
    ],
  }),
  component: Home,
});

const today = new Date().toISOString().slice(0, 10);

const DEFAULT_COVER: CoverInput = {
  reportTitle: "LAPORAN PREVENTIVE MAINTENANCE",
  subtitle: "Perangkat Lunak",
  companyName: "PT. Bayan Resources TBK",
  operatingSystem: "Red Hat Enterprise Linux",
  pelaksana: "Ali Imran",
  tanggal: today,
  contractNo: "MCARE-3691-PO HO120 BR MII IMPLEMENTASI- Bayan",
  periode: "Juli 2025",
  vendorName: "PT. Mitra Integrasi Informatika (MII)",
  reviewerClient: "Bayu",
  reviewerVendor: "Ali Imran",
  reviewerVendorRole: "Technical Consultant",
  reviewerDate: today,
  executiveSummary: "",
  logoDataUrl: "",
  logoRightDataUrl: "",
  coverBackgroundDataUrl: "",
  clientAddress: "Jl. Gatot Subroto Kav 40-42\nJakarta 12190",
  vendorAddress: "APL Tower 37th Floor\nJl. Letjen S. Parman Kav 28\nJakarta Barat 11470",
  summaryConclusion:
    "Rata-rata pemakaian memory dan CPU masih normal.\nRata-rata time & date sync.\nRata-rata uptime diatas 180days.\nDitemukan informasi log error.\nDitemukan penggunaan disk diatas threshold.\nDitemukan informasi log kill memory.",
  recommendation:
    "Melakukan housekeeping pada server yang sudah diatas threshold.\nMenambahkan ram atau melakukan clear buff/cache pada memory server yang warning.\nMelakukan sinkronisasi ntp/chrony.\nPengecekan pada server dengan uptime diatas 180days.",
};

interface FileItem {
  id: string;
  name: string;
  pm: ParsedPM;
}

function StatusPill({ s }: { s: StatusKind }) {
  const map: Record<string, string> = {
    OK: "bg-emerald-500/15 text-emerald-400 border-emerald-500/30",
    Warning: "bg-amber-400/15 text-amber-300 border-amber-400/30",
    Critical: "bg-red-500/15 text-red-400 border-red-500/30",
  };
  return (
    <span className={`text-[10px] rounded border px-1.5 py-0.5 font-medium ${map[s] ?? "bg-muted text-muted-foreground border-border"}`}>
      {s || "—"}
    </span>
  );
}

// Read a File as data URL (safe for large files — avoids stack overflow from btoa+spread)
function readAsDataUrl(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const r = new FileReader();
    r.onload = () => resolve(String(r.result || ""));
    r.onerror = () => reject(r.error);
    r.readAsDataURL(file);
  });
}

function Home() {
  const [cover, setCover] = useState<CoverInput>(DEFAULT_COVER);
  const [files, setFiles] = useState<FileItem[]>([]);
  const [busy, setBusy] = useState(false);
  const [filename, setFilename] = useState("Laporan_PM.docx");
  const filenameEditedRef = useRef(false);

  const summaries = useMemo(() => files.map((f) => summarizePM(f.pm)), [files]);

  const set = <K extends keyof CoverInput>(k: K, v: CoverInput[K]) => setCover((c) => ({ ...c, [k]: v }));

  // Auto-derive nama file output: Laporan_PM_[Customer]_[bulan]_[tahun].docx
  // Berjalan otomatis selama user belum mengubah field manual atau load preset dengan fileName.
  useEffect(() => {
    if (filenameEditedRef.current) return;
    const resolved = resolveFileNameTemplate(
      "Laporan_PM_[Customer]_[bulan]_[tahun].docx",
      cover.periode,
      cover.companyName,
    );
    setFilename(resolved);
  }, [cover.companyName, cover.periode]);

  async function ingestHtml(name: string, text: string, items: FileItem[]) {
    const pm = parsePMHtml(text, name.replace(/\.html?$/i, ""));
    items.push({ id: `${name}-${Math.random().toString(36).slice(2, 8)}`, name, pm });
  }

  async function onFiles(list: FileList | null) {
    if (!list) return;
    const items: FileItem[] = [];
    let archCount = 0;
    for (const f of Array.from(list)) {
      const lname = f.name.toLowerCase();
      try {
        if (lname.endsWith(".zip")) {
          const zip = await JSZip.loadAsync(await f.arrayBuffer());
          const entries = Object.values(zip.files).filter((z) => !z.dir && /\.html?$/i.test(z.name));
          for (const entry of entries) {
            const text = await entry.async("string");
            await ingestHtml(entry.name.split("/").pop() || entry.name, text, items);
            archCount++;
          }
        } else if (lname.endsWith(".tar.gz") || lname.endsWith(".tgz") || lname.endsWith(".tar")) {
          const buf = new Uint8Array(await f.arrayBuffer());
          const tarBuf = lname.endsWith(".tar") ? buf : ungzipToTar(buf);
          const entries = untar(tarBuf).filter((e) => /\.html?$/i.test(e.name));
          for (const entry of entries) {
            const text = new TextDecoder().decode(entry.data);
            await ingestHtml(entry.name.split("/").pop() || entry.name, text, items);
            archCount++;
          }
        } else if (/\.html?$/i.test(f.name)) {
          await ingestHtml(f.name, await f.text(), items);
        }
      } catch (err) {
        console.error(err);
        toast.error(`Gagal baca: ${f.name}`);
      }
    }
    setFiles((prev) => [...prev, ...items]);
    if (items.length) toast.success(`${items.length} file HTML dimuat${archCount ? ` (${archCount} dari arsip)` : ""}`);
  }

  async function onLogo(fileList: FileList | null, side: "left" | "right" | "bg") {
    const f = fileList?.[0];
    if (!f) return;
    if (!/^image\//.test(f.type)) {
      toast.error("Hanya file gambar yang didukung");
      return;
    }
    try {
      const dataUrl = await readAsDataUrl(f);
      const key = side === "left" ? "logoDataUrl" : side === "right" ? "logoRightDataUrl" : "coverBackgroundDataUrl";
      set(key, dataUrl);
      const labelMap = { left: "Logo kiri", right: "Logo kanan", bg: "Background cover" } as const;
      toast.success(`${labelMap[side]} dimuat`);
    } catch (e) {
      console.error(e);
      toast.error("Gagal baca gambar");
    }
  }

  // Parse preset TXT: baris `key=value`, `key:value`, atau JSON.
  // Baris diawali `#` atau `;` diabaikan. Multiline via `\n` literal.
  function parsePreset(text: string): Partial<CoverInput> {
    const trimmed = text.trim();
    // Coba JSON dulu
    if (trimmed.startsWith("{")) {
      try {
        const obj = JSON.parse(trimmed);
        if (obj && typeof obj === "object") return obj as Partial<CoverInput>;
      } catch {
        /* fallthrough */
      }
    }
    const out: Record<string, string> = {};
    for (const raw of trimmed.split(/\r?\n/)) {
      const line = raw.trim();
      if (!line || line.startsWith("#") || line.startsWith(";")) continue;
      const m = line.match(/^([A-Za-z_][A-Za-z0-9_]*)\s*[:=]\s*(.*)$/);
      if (!m) continue;
      const key = m[1];
      let val = m[2];
      // Strip surrounding quotes
      if ((val.startsWith('"') && val.endsWith('"')) || (val.startsWith("'") && val.endsWith("'"))) {
        val = val.slice(1, -1);
      }
      // Convert literal \n → newline
      val = val.replace(/\\n/g, "\n");
      out[key] = val;
    }
    return out as Partial<CoverInput>;
  }

  // Placeholder: [bulan]/[tahun] diambil dari `cover.periode` (contoh "Juli 2026"),
  // [Customer] dari `cover.companyName`. Fallback ke tanggal hari ini bila periode kosong.
  function resolveFileNameTemplate(tpl: string, periode: string, customer: string): string {
    const monthsID = [
      "Januari","Februari","Maret","April","Mei","Juni",
      "Juli","Agustus","September","Oktober","November","Desember",
    ];
    let bulan = "";
    let tahun = "";
    const parts = (periode || "").trim().split(/\s+/);
    if (parts.length >= 2) {
      bulan = parts[0];
      tahun = parts[parts.length - 1];
    } else {
      const d = new Date();
      bulan = monthsID[d.getMonth()];
      tahun = String(d.getFullYear());
    }
    const safeCustomer = (customer || "Customer")
      .replace(/[\\/:*?"<>|]+/g, "")
      .replace(/[.,]+/g, "")
      .replace(/\s+/g, "_")
      .replace(/_+/g, "_")
      .replace(/^_|_$/g, "");
    const safeBulan = bulan.trim().replace(/\s+/g, "_");
    const safeTahun = tahun.trim().replace(/\s+/g, "_");
    const resolved = tpl
      .replace(/\[bulan\]/gi, safeBulan)
      .replace(/\[tahun\]/gi, safeTahun)
      .replace(/\[customer\]/gi, safeCustomer);
    // Rapikan: buang spasi berlebih & underscore ganda
    return resolved
      .replace(/\s+/g, "")
      .replace(/_+/g, "_")
      .replace(/_+(\.docx)$/i, "$1");
  }

  async function onPresetFile(list: FileList | null) {
    const f = list?.[0];
    if (!f) return;
    try {
      const text = await f.text();
      const preset = parsePreset(text);
      const keys = Object.keys(preset) as string[];
      const allowed: (keyof CoverInput)[] = [
        "reportTitle", "subtitle", "companyName", "operatingSystem", "pelaksana",
        "tanggal", "contractNo", "periode", "vendorName", "reviewerClient",
        "reviewerVendor", "reviewerVendorRole", "reviewerDate", "executiveSummary",
        "summaryConclusion", "recommendation", "clientAddress", "vendorAddress",
      ];
      const filtered = keys.filter((k) => (allowed as string[]).includes(k));
      const hasFileName = typeof (preset as Record<string, unknown>).fileName === "string";
      if (filtered.length === 0 && !hasFileName) {
        toast.error("Tidak ada field valid ditemukan di preset");
        return;
      }
      let nextPeriode = cover.periode;
      let nextCustomer = cover.companyName;
      setCover((c) => {
        const next = { ...c };
        for (const k of filtered) {
          const v = (preset as Record<string, unknown>)[k as string];
          if (typeof v === "string") (next as Record<string, unknown>)[k as string] = v;
        }
        nextPeriode = next.periode;
        nextCustomer = next.companyName;
        return next;
      });
      if (hasFileName) {
        const rawName = String((preset as Record<string, unknown>).fileName || "");
        const resolved = resolveFileNameTemplate(rawName, nextPeriode, nextCustomer);
        filenameEditedRef.current = true;
        setFilename(resolved.endsWith(".docx") ? resolved : `${resolved}.docx`);
      }
      toast.success(`Preset dimuat (${filtered.length + (hasFileName ? 1 : 0)} field)`);
    } catch (e) {
      console.error(e);
      toast.error("Gagal baca preset");
    }
  }



  const sensors = useSensors(
    useSensor(PointerSensor, { activationConstraint: { distance: 4 } }),
    useSensor(KeyboardSensor, { coordinateGetter: sortableKeyboardCoordinates }),
  );

  function onDragEnd(e: DragEndEvent) {
    const { active, over } = e;
    if (!over || active.id === over.id) return;
    setFiles((prev) => {
      const oldIdx = prev.findIndex((f) => f.id === active.id);
      const newIdx = prev.findIndex((f) => f.id === over.id);
      if (oldIdx < 0 || newIdx < 0) return prev;
      return arrayMove(prev, oldIdx, newIdx);
    });
  }

  async function handleExport() {
    if (files.length === 0) {
      toast.error("Upload minimal 1 file HTML dulu ya");
      return;
    }
    try {
      setBusy(true);
      const blob = await buildAndDownloadDocx(cover, files.map((f) => f.pm), filename);
      toast.success("Word berhasil dibuat 🔥");
      try {
        const buf = new Uint8Array(await blob.arrayBuffer());
        let bin = "";
        for (let i = 0; i < buf.length; i++) bin += String.fromCharCode(buf[i]);
        const base64 = btoa(bin);
        const res = await archiveExport({
          data: {
            filename: filename.endsWith(".docx") ? filename : `${filename}.docx`,
            base64,
            meta: {
              serverCount: files.length,
              companyName: cover.companyName,
              vendorName: cover.vendorName,
              periode: cover.periode,
              hosts: files.map((f) => ({
                hostname: f.pm.hostname,
                ipAddress: f.pm.ipAddress,
                osRelease: f.pm.osRelease,
              })),
            },
          },
        });
        if (res?.dbEnabled) toast.message("Riwayat tersimpan di database");
      } catch (archiveErr) {
        console.warn("archive skipped:", archiveErr);
      }
    } catch (e) {
      console.error(e);
      toast.error("Gagal membuat dokumen. Cek console.");
    } finally {
      setBusy(false);
    }
  }

  async function handleExportReportSummary() {
    if (files.length === 0) {
      toast.error("Upload minimal 1 file HTML dulu ya");
      return;
    }
    try {
      setBusy(true);
      const base = filename.replace(/\.docx$/i, "");
      await buildAndDownloadReportSummary(cover, files.map((f) => f.pm), `${base}_ReportSummary.docx`);
      toast.success("Report Summary berhasil dibuat 🔥");
    } catch (e) {
      console.error(e);
      toast.error("Gagal membuat Report Summary. Cek console.");
    } finally {
      setBusy(false);
    }
  }

  function downloadPresetTemplate() {
    const tpl = `# Preset Metadata ExcportCuy
# Baris diawali # atau ; = komentar. Gunakan \\n untuk baris baru pada alamat / paragraf.
# fileName mendukung placeholder [Customer], [bulan], [tahun] (Customer dari companyName; bulan/tahun dari periode).
reportTitle=LAPORAN PREVENTIVE MAINTENANCE
subtitle=Perangkat Lunak
companyName=PT. Contoh Customer
operatingSystem=Red Hat Enterprise Linux
pelaksana=Nama Engineer
tanggal=2026-07-04
contractNo=MCARE-XXXX-PO
periode=Juli 2026
vendorName=PT. Mitra Integrasi Informatika (MII)
reviewerClient=Nama Reviewer Klien
reviewerVendor=Nama Reviewer Vendor
reviewerVendorRole=Technical Consultant
reviewerDate=2026-07-04
clientAddress=Jl. Contoh No 1\\nJakarta 12190
vendorAddress=APL Tower 37th Floor\\nJl. Letjen S. Parman Kav 28\\nJakarta Barat 11470
executiveSummary=
summaryConclusion=Rata-rata pemakaian memory dan CPU masih normal.\\nRata-rata time & date sync.\\nRata-rata uptime diatas 180days.\\nDitemukan informasi log error.\\nDitemukan penggunaan disk diatas threshold.\\nDitemukan informasi log kill memory.
recommendation=Melakukan housekeeping pada server yang sudah diatas threshold.\\nMenambahkan ram atau melakukan clear buff/cache pada memory server yang warning.\\nMelakukan sinkronisasi ntp/chrony.\\nPengecekan pada server dengan uptime diatas 180days.
fileName=Laporan_PM_[Customer]_[bulan]_[tahun].docx
`;
    const blob = new Blob([tpl], { type: "text/plain;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "preset_template.txt";
    document.body.appendChild(a);
    a.click();
    setTimeout(() => {
      try { document.body.removeChild(a); } catch { /* noop */ }
      URL.revokeObjectURL(url);
    }, 2000);
  }

  return (
    <div className="min-h-screen bg-background text-foreground">
      <header className="sticky top-0 z-20 border-b border-border/60 bg-background/80 backdrop-blur">
        <div className="mx-auto flex max-w-6xl items-center justify-between px-6 py-4">
          <div className="flex items-center gap-3">
            <DragonLogo className="h-9 w-9" />
            <div>
              <div className="text-lg font-bold tracking-tight">
                Ex<span className="text-[hsl(var(--brand))]">cport</span>Cuy
              </div>
              <div className="text-[11px] text-muted-foreground -mt-0.5">HTML → Word, rapi &amp; berapi</div>
            </div>
          </div>
          <div className="flex items-center gap-2 text-xs text-muted-foreground">
            <Flame className="h-4 w-4 text-[hsl(var(--brand))]" />
            <span>Port 9812</span>
          </div>
        </div>
      </header>

      <section className="border-b border-border/60">
        <div className="mx-auto max-w-6xl px-6 py-10 grid gap-8 lg:grid-cols-[1.2fr,1fr] items-center">
          <div>
            <p className="text-xs font-mono uppercase tracking-widest text-[hsl(var(--brand))]">Dragon-forged exporter</p>
            <h1 className="mt-2 text-4xl font-bold tracking-tight leading-tight">
              Ubah HTML Preventive Maintenance jadi <span className="text-[hsl(var(--brand))]">Word</span> yang rapi.
            </h1>
            <p className="mt-3 text-sm text-muted-foreground max-w-xl">
              Isi cover, upload file HTML (atau ZIP berisi banyak HTML), atur urutan, lalu klik export. Otomatis cover, TOC, Document Control, Executive Summary, List Server, dan Lampiran per host.
            </p>
          </div>
          <div className="rounded-xl border border-border/70 bg-card p-6 flex items-center justify-center">
            <DragonLogo className="h-40 w-40" />
          </div>
        </div>
      </section>

      <main className="mx-auto max-w-6xl px-6 py-10 grid gap-8 lg:grid-cols-[1.1fr,1fr]">
        <Card>
          <CardHeader>
            <CardTitle className="text-base">Cover &amp; Analisis Dokumen</CardTitle>
            <CardDescription>Data ini akan tampil di halaman depan, Document Control, dan Overview.</CardDescription>
          </CardHeader>
          <CardContent className="space-y-4">
            {/* Logo uploads: client (kiri bawah cover + header). Vendor (MII) fixed dari asset. */}
            <div className="grid gap-3 sm:grid-cols-2">
              <LogoBox
                label="Logo Klien (kiri bawah cover & header)"
                dataUrl={cover.logoDataUrl || ""}
                onPick={(l) => onLogo(l, "left")}
                onClear={() => set("logoDataUrl", "")}
              />
              <LogoBox
                label="Background Cover (opsional)"
                dataUrl={cover.coverBackgroundDataUrl || ""}
                onPick={(l) => onLogo(l, "bg")}
                onClear={() => set("coverBackgroundDataUrl", "")}
              />
            </div>
            <p className="text-[11px] text-muted-foreground -mt-2">
              Logo vendor (MII) sudah fixed dari sistem — otomatis muncul di bawah kanan cover &amp; header konten.
            </p>

            <div className="rounded-lg border border-dashed border-border/70 bg-muted/20 p-3 flex items-center justify-between gap-3">
              <div className="min-w-0">
                <div className="text-xs font-medium">Upload Preset Metadata (.txt / .json)</div>
                <div className="text-[11px] text-muted-foreground">
                  Format: <code className="font-mono">key=value</code> per baris atau JSON. Contoh:{" "}
                  <code className="font-mono">companyName=PT ABC</code>. Gunakan <code>\n</code> untuk baris baru.
                </div>
              </div>
              <div className="flex shrink-0 items-center gap-2">
                <button
                  type="button"
                  onClick={downloadPresetTemplate}
                  className="inline-flex items-center gap-1.5 rounded-md border border-border bg-background px-3 py-1.5 text-xs cursor-pointer hover:bg-muted"
                >
                  <Download className="h-3 w-3" /> Template
                </button>
                <label
                  htmlFor="preset-upload"
                  className="inline-flex items-center gap-1.5 rounded-md border border-border bg-background px-3 py-1.5 text-xs cursor-pointer hover:bg-muted"
                >
                  <Upload className="h-3 w-3" /> Pilih file
                  <input
                    id="preset-upload"
                    type="file"
                    accept=".txt,.json,text/plain,application/json"
                    className="hidden"
                    onChange={(e) => onPresetFile(e.target.files)}
                  />
                </label>
              </div>
            </div>




            <div className="grid gap-4 sm:grid-cols-2">
              <Field label="Judul Laporan" v={cover.reportTitle} onChange={(v) => set("reportTitle", v)} />
              <Field label="Subjudul" v={cover.subtitle} onChange={(v) => set("subtitle", v)} />
              <Field label="Nama Perusahaan (Klien)" v={cover.companyName} onChange={(v) => set("companyName", v)} />
              <Field label="Sistem Operasi" v={cover.operatingSystem} onChange={(v) => set("operatingSystem", v)} />
              <Field label="Pelaksana PM" v={cover.pelaksana} onChange={(v) => set("pelaksana", v)} />
              <DateField label="Tanggal & Waktu" v={cover.tanggal} onChange={(v) => set("tanggal", v)} />
              <Field label="No Contract" v={cover.contractNo} onChange={(v) => set("contractNo", v)} />
              <Field label="Periode" v={cover.periode} onChange={(v) => set("periode", v)} />
              <Field label="Nama Vendor" v={cover.vendorName} onChange={(v) => set("vendorName", v)} />
              <Field label="Reviewer Klien" v={cover.reviewerClient} onChange={(v) => set("reviewerClient", v)} />
              <Field label="Reviewer Vendor" v={cover.reviewerVendor} onChange={(v) => set("reviewerVendor", v)} />
              <Field label="Jabatan Reviewer Vendor" v={cover.reviewerVendorRole} onChange={(v) => set("reviewerVendorRole", v)} />
              <DateField label="Tanggal Reviewer" v={cover.reviewerDate} onChange={(v) => set("reviewerDate", v)} />
              <Field label="Nama File Output" v={filename} onChange={(v) => { filenameEditedRef.current = true; setFilename(v); }} />
              <p className="text-[10px] text-muted-foreground">Otomatis: Laporan_PM_[Customer]_[bulan]_[tahun].docx (berdasarkan Nama Customer &amp; Periode).</p>
            </div>

            <div>
              <Label className="text-xs">Executive Summary <span className="text-muted-foreground">(opsional, auto jika kosong)</span></Label>
              <Textarea rows={3} value={cover.executiveSummary} onChange={(e) => set("executiveSummary", e.target.value)} className="mt-1" />
            </div>
            <div className="grid gap-4 sm:grid-cols-2">
              <div>
                <Label className="text-xs">Alamat Klien <span className="text-muted-foreground">(baris = enter)</span></Label>
                <Textarea rows={3} value={cover.clientAddress || ""} onChange={(e) => set("clientAddress", e.target.value)} className="mt-1 text-xs" />
              </div>
              <div>
                <Label className="text-xs">Alamat Vendor <span className="text-muted-foreground">(baris = enter)</span></Label>
                <Textarea rows={3} value={cover.vendorAddress || ""} onChange={(e) => set("vendorAddress", e.target.value)} className="mt-1 text-xs" />
              </div>
            </div>

            <div>
              <Label className="text-xs">Summary Conclusion <span className="text-muted-foreground">(satu poin per baris)</span></Label>
              <Textarea rows={4} value={cover.summaryConclusion} onChange={(e) => set("summaryConclusion", e.target.value)} className="mt-1 font-mono text-xs" />
            </div>
            <div>
              <Label className="text-xs">Recommendation <span className="text-muted-foreground">(satu poin per baris)</span></Label>
              <Textarea rows={4} value={cover.recommendation} onChange={(e) => set("recommendation", e.target.value)} className="mt-1 font-mono text-xs" />
            </div>
          </CardContent>
        </Card>

        <div className="space-y-6">
          <Card>
            <CardHeader>
              <CardTitle className="text-base">File HTML Preventive Maintenance</CardTitle>
              <CardDescription>Upload HTML, ZIP, atau .tar.gz berisi banyak HTML. Seret item untuk mengatur urutan — urutan file = urutan lampiran.</CardDescription>
            </CardHeader>
            <CardContent className="space-y-4">
              <label
                htmlFor="upload"
                className="flex flex-col items-center justify-center gap-2 rounded-lg border-2 border-dashed border-border p-8 text-center cursor-pointer hover:border-[hsl(var(--brand))]/60 hover:bg-muted/40 transition"
                onDragOver={(e) => e.preventDefault()}
                onDrop={(e) => {
                  e.preventDefault();
                  onFiles(e.dataTransfer.files);
                }}
              >
                <div className="flex items-center gap-2">
                  <Upload className="h-6 w-6 text-[hsl(var(--brand))]" />
                  <FileArchive className="h-6 w-6 text-[hsl(var(--brand))]/70" />
                </div>
                <div className="text-sm font-medium">Drop .html / .zip / .tar.gz / .tgz atau klik untuk memilih</div>
                <div className="text-xs text-muted-foreground">Multiple files &amp; arsip didukung</div>
                <input
                  id="upload"
                  type="file"
                  accept=".html,.htm,.zip,.tar,.tar.gz,.tgz,.gz"
                  multiple
                  className="hidden"
                  onChange={(e) => onFiles(e.target.files)}
                />
              </label>

              {files.length > 0 && (
                <DndContext sensors={sensors} collisionDetection={closestCenter} onDragEnd={onDragEnd}>
                  <SortableContext items={files.map((f) => f.id)} strategy={verticalListSortingStrategy}>
                    <ul className="space-y-2">
                      {files.map((f, i) => (
                        <SortableFileItem
                          key={f.id}
                          index={i}
                          file={f}
                          summary={summaries[i]}
                          onRemove={() => setFiles((p) => p.filter((x) => x.id !== f.id))}
                        />
                      ))}
                    </ul>
                  </SortableContext>
                </DndContext>
              )}
            </CardContent>
          </Card>


          <Card>
            <CardContent className="pt-6 space-y-3">
              <Button
                onClick={handleExport}
                disabled={busy}
                className="w-full h-11 bg-[hsl(var(--brand))] text-white hover:bg-[hsl(var(--brand))]/90"
              >
                <Download className="h-4 w-4 mr-2" />
                {busy ? "Membangun dokumen..." : "Export ke Word (.docx)"}
              </Button>
              <Button
                onClick={handleExportReportSummary}
                disabled={busy}
                variant="outline"
                className="w-full h-10"
              >
                <Download className="h-4 w-4 mr-2" />
                {busy ? "Sedang memproses..." : "Export Report Summary (Warning & Critical)"}
              </Button>
              <p className="text-[11px] text-muted-foreground text-center">
                Report Summary di-generate sebagai file .docx terpisah — berisi rincian temuan Warning/Critical beserta remark dan isian log error tiap server, untuk mempermudah analisa data.
              </p>
              <p className="text-[11px] text-muted-foreground text-center">
                Table of Contents otomatis terisi; di Microsoft Word tekan <kbd className="px-1 rounded bg-muted">Ctrl</kbd>+<kbd className="px-1 rounded bg-muted">A</kbd> lalu <kbd className="px-1 rounded bg-muted">F9</kbd> untuk refresh nomor halaman.
              </p>
            </CardContent>
          </Card>
        </div>
      </main>

      <footer className="border-t border-border/60 py-6 text-center text-xs text-muted-foreground">
        Made with 🐉 by ExcportCuy · running on port 9812
      </footer>
    </div>
  );
}

function LogoBox({
  label,
  dataUrl,
  onPick,
  onClear,
}: {
  label: string;
  dataUrl: string;
  onPick: (files: FileList | null) => void;
  onClear: () => void;
}) {
  const id = `logo-${label.replace(/\s+/g, "-").toLowerCase()}`;
  return (
    <div className="rounded-lg border border-border/70 bg-muted/30 p-3 flex items-center gap-3">
      <div className="h-16 w-16 rounded-md border border-border/60 bg-background flex items-center justify-center overflow-hidden shrink-0">
        {dataUrl ? (
          <img src={dataUrl} alt="logo" className="max-h-full max-w-full object-contain" />
        ) : (
          <ImagePlus className="h-6 w-6 text-muted-foreground" />
        )}
      </div>
      <div className="flex-1 min-w-0">
        <div className="text-xs font-medium">{label}</div>
        <div className="text-[11px] text-muted-foreground">PNG/JPG</div>
        <div className="mt-2 flex gap-2">
          <label htmlFor={id} className="inline-flex items-center gap-1.5 rounded-md border border-border bg-background px-2.5 py-1 text-xs cursor-pointer hover:bg-muted">
            <Upload className="h-3 w-3" /> Pilih
            <input id={id} type="file" accept="image/*" className="hidden" onChange={(e) => onPick(e.target.files)} />
          </label>
          {dataUrl && (
            <button type="button" onClick={onClear} className="inline-flex items-center gap-1 rounded-md border border-border bg-background px-2.5 py-1 text-xs text-muted-foreground hover:text-foreground">
              <X className="h-3 w-3" /> Hapus
            </button>
          )}
        </div>
      </div>
    </div>
  );
}

function Field({ label, v, onChange }: { label: string; v: string; onChange: (v: string) => void }) {
  return (
    <div>
      <Label className="text-xs" dangerouslySetInnerHTML={{ __html: label }} />
      <Input value={v} onChange={(e) => onChange(e.target.value)} className="mt-1" />
    </div>
  );
}

function DateField({ label, v, onChange }: { label: string; v: string; onChange: (v: string) => void }) {
  const parsed = v ? new Date(v) : undefined;
  const valid = parsed && !isNaN(parsed.getTime()) ? parsed : undefined;
  return (
    <div>
      <Label className="text-xs">{label}</Label>
      <Popover>
        <PopoverTrigger asChild>
          <Button
            variant="outline"
            className={cn(
              "mt-1 w-full justify-start text-left font-normal h-9",
              !valid && "text-muted-foreground",
            )}
          >
            <CalendarIcon className="h-4 w-4 mr-2" />
            {valid ? format(valid, "PPP") : <span>Pilih tanggal</span>}
          </Button>
        </PopoverTrigger>
        <PopoverContent className="w-auto p-0" align="start">
          <Calendar
            mode="single"
            selected={valid}
            onSelect={(d) => d && onChange(d.toISOString().slice(0, 10))}
            initialFocus
            className={cn("p-3 pointer-events-auto")}
          />
        </PopoverContent>
      </Popover>
    </div>
  );
}

function SortableFileItem({
  file,
  index,
  summary,
  onRemove,
}: {
  file: FileItem;
  index: number;
  summary: ReturnType<typeof summarizePM>;
  onRemove: () => void;
}) {
  const { attributes, listeners, setNodeRef, transform, transition, isDragging } = useSortable({ id: file.id });
  const style: React.CSSProperties = {
    transform: CSS.Transform.toString(transform),
    transition,
    opacity: isDragging ? 0.6 : 1,
    zIndex: isDragging ? 10 : "auto",
  };
  return (
    <li
      ref={setNodeRef}
      style={style}
      className={cn(
        "rounded-md border border-border/70 bg-card/40 p-3",
        isDragging && "ring-2 ring-[hsl(var(--brand))]/60 shadow-lg",
      )}
    >
      <div className="flex items-start justify-between gap-3">
        <div className="flex items-start gap-2 min-w-0">
          <button
            type="button"
            className="mt-0.5 flex flex-col items-center gap-1 cursor-grab active:cursor-grabbing text-muted-foreground hover:text-foreground touch-none"
            {...attributes}
            {...listeners}
            aria-label="Seret untuk mengurutkan"
          >
            <GripVertical className="h-4 w-4" />
            <span className="text-[10px] font-mono">#{index + 1}</span>
          </button>
          <FileText className="h-4 w-4 mt-1 text-muted-foreground shrink-0" />
          <div className="min-w-0">
            <div className="text-sm font-medium truncate">{file.pm.hostname}</div>
            <div className="text-[11px] text-muted-foreground truncate">
              {file.name} · {file.pm.ipAddress || "no ip"} · {file.pm.osRelease || "-"}
            </div>
            <div className="mt-2 flex flex-wrap gap-1">
              <StatusPill s={summary.mountpoint} /> <StatusPill s={summary.uptime} /> <StatusPill s={summary.timeSync} />{" "}
              <StatusPill s={summary.cpu} /> <StatusPill s={summary.memory} /> <StatusPill s={summary.logKillMemory} />{" "}
              <StatusPill s={summary.logError} />
            </div>
          </div>
        </div>
        <Button size="icon" variant="ghost" onClick={onRemove} className="shrink-0">
          <Trash2 className="h-4 w-4" />
        </Button>
      </div>
    </li>
  );
}
