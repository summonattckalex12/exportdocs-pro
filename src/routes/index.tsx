import { createFileRoute } from "@tanstack/react-router";
import { useMemo, useState } from "react";
import { toast } from "sonner";
import { DragonLogo } from "@/components/DragonLogo";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Textarea } from "@/components/ui/textarea";
import { Card, CardContent, CardHeader, CardTitle, CardDescription } from "@/components/ui/card";
import { parsePMHtml, summarizePM, type ParsedPM, type StatusKind } from "@/lib/pm-html-parser";
import { buildAndDownloadDocx, type CoverInput } from "@/lib/docx-exporter";
import { archiveExport } from "@/lib/exports.functions";
import { Upload, FileText, Trash2, Download, Flame, ImagePlus, X } from "lucide-react";

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
  summaryConclusion:
    "Rata-rata pemakaian memory dan CPU masih normal.\nRata-rata time & date sync dalam kondisi baik.\nDitemukan beberapa server dengan status Warning pada log error.",
  recommendation:
    "Melakukan housekeeping atau penambahan disk pada server yang mendekati threshold.\nMelakukan sinkronisasi ntp/chrony agar time & date pada server sesuai.\nMelakukan patching security atau update untuk menghilangkan temuan scan VA.",
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

function Home() {
  const [cover, setCover] = useState<CoverInput>(DEFAULT_COVER);
  const [files, setFiles] = useState<FileItem[]>([]);
  const [busy, setBusy] = useState(false);
  const [filename, setFilename] = useState("Laporan_PM.docx");

  const summaries = useMemo(() => files.map((f) => summarizePM(f.pm)), [files]);

  const set = <K extends keyof CoverInput>(k: K, v: CoverInput[K]) => setCover((c) => ({ ...c, [k]: v }));

  async function onFiles(list: FileList | null) {
    if (!list) return;
    const items: FileItem[] = [];
    for (const f of Array.from(list)) {
      if (!/\.html?$/i.test(f.name)) continue;
      const text = await f.text();
      const pm = parsePMHtml(text, f.name.replace(/\.html?$/i, ""));
      items.push({ id: `${f.name}-${Math.random().toString(36).slice(2, 8)}`, name: f.name, pm });
    }
    setFiles((prev) => [...prev, ...items]);
    if (items.length) toast.success(`${items.length} file HTML dimuat`);
  }

  async function onLogo(fileList: FileList | null) {
    const f = fileList?.[0];
    if (!f) return;
    if (!/^image\//.test(f.type)) {
      toast.error("Hanya file gambar yang didukung");
      return;
    }
    const buf = await f.arrayBuffer();
    const b64 = btoa(String.fromCharCode(...new Uint8Array(buf)));
    set("logoDataUrl", `data:${f.type};base64,${b64}`);
    toast.success("Logo cover dimuat");
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

      // Archive ke server (best-effort — kalau server function gagal, biarkan)
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

  return (
    <div className="min-h-screen bg-background text-foreground">
      {/* nav */}
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

      {/* hero */}
      <section className="border-b border-border/60">
        <div className="mx-auto max-w-6xl px-6 py-10 grid gap-8 lg:grid-cols-[1.2fr,1fr] items-center">
          <div>
            <p className="text-xs font-mono uppercase tracking-widest text-[hsl(var(--brand))]">Dragon-forged exporter</p>
            <h1 className="mt-2 text-4xl font-bold tracking-tight leading-tight">
              Ubah HTML Preventive Maintenance jadi <span className="text-[hsl(var(--brand))]">Word</span> yang rapi.
            </h1>
            <p className="mt-3 text-sm text-muted-foreground max-w-xl">
              Isi cover analisis dokumen, upload file HTML PM (mendukung banyak file), lalu klik export. Otomatis dibuat cover, Document Control, Table of Contents, Executive Summary, List Server, dan Lampiran per host lengkap dengan pewarnaan status.
            </p>
          </div>
          <div className="rounded-xl border border-border/70 bg-card p-6 flex items-center justify-center">
            <DragonLogo className="h-40 w-40" />
          </div>
        </div>
      </section>

      <main className="mx-auto max-w-6xl px-6 py-10 grid gap-8 lg:grid-cols-[1.1fr,1fr]">
        {/* Cover form */}
        <Card>
          <CardHeader>
            <CardTitle className="text-base">Cover &amp; Analisis Dokumen</CardTitle>
            <CardDescription>Data ini akan tampil di halaman depan, Document Control, dan Overview.</CardDescription>
          </CardHeader>
          <CardContent className="space-y-4">
            <div className="grid gap-4 sm:grid-cols-2">
              <Field label="Judul Laporan" v={cover.reportTitle} onChange={(v) => set("reportTitle", v)} />
              <Field label="Subjudul" v={cover.subtitle} onChange={(v) => set("subtitle", v)} />
              <Field label="Nama Perusahaan (Klien)" v={cover.companyName} onChange={(v) => set("companyName", v)} />
              <Field label="Sistem Operasi" v={cover.operatingSystem} onChange={(v) => set("operatingSystem", v)} />
              <Field label="Pelaksana PM" v={cover.pelaksana} onChange={(v) => set("pelaksana", v)} />
              <Field label="Tanggal &amp; Waktu" v={cover.tanggal} onChange={(v) => set("tanggal", v)} />
              <Field label="No Contract" v={cover.contractNo} onChange={(v) => set("contractNo", v)} />
              <Field label="Periode" v={cover.periode} onChange={(v) => set("periode", v)} />
              <Field label="Nama Vendor" v={cover.vendorName} onChange={(v) => set("vendorName", v)} />
              <Field label="Reviewer Klien" v={cover.reviewerClient} onChange={(v) => set("reviewerClient", v)} />
              <Field label="Reviewer Vendor" v={cover.reviewerVendor} onChange={(v) => set("reviewerVendor", v)} />
              <Field label="Jabatan Reviewer Vendor" v={cover.reviewerVendorRole} onChange={(v) => set("reviewerVendorRole", v)} />
              <Field label="Tanggal Reviewer" v={cover.reviewerDate} onChange={(v) => set("reviewerDate", v)} />
              <Field label="Nama File Output" v={filename} onChange={setFilename} />
            </div>

            <div>
              <Label className="text-xs">Executive Summary <span className="text-muted-foreground">(opsional, auto jika kosong)</span></Label>
              <Textarea rows={3} value={cover.executiveSummary} onChange={(e) => set("executiveSummary", e.target.value)} className="mt-1" />
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

        {/* Files + export */}
        <div className="space-y-6">
          <Card>
            <CardHeader>
              <CardTitle className="text-base">File HTML Preventive Maintenance</CardTitle>
              <CardDescription>Upload semua file HTML PM per server. Urutan file = urutan lampiran.</CardDescription>
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
                <Upload className="h-6 w-6 text-[hsl(var(--brand))]" />
                <div className="text-sm font-medium">Drop file .html atau klik untuk memilih</div>
                <div className="text-xs text-muted-foreground">Mendukung multiple files</div>
                <input
                  id="upload"
                  type="file"
                  accept=".html,.htm"
                  multiple
                  className="hidden"
                  onChange={(e) => onFiles(e.target.files)}
                />
              </label>

              {files.length > 0 && (
                <ul className="space-y-2">
                  {files.map((f, i) => {
                    const s = summaries[i];
                    return (
                      <li key={f.id} className="rounded-md border border-border/70 bg-card/40 p-3">
                        <div className="flex items-start justify-between gap-3">
                          <div className="flex items-start gap-2 min-w-0">
                            <FileText className="h-4 w-4 mt-0.5 text-muted-foreground shrink-0" />
                            <div className="min-w-0">
                              <div className="text-sm font-medium truncate">{f.pm.hostname}</div>
                              <div className="text-[11px] text-muted-foreground truncate">{f.name} · {f.pm.ipAddress || "no ip"} · {f.pm.osRelease || "-"}</div>
                              <div className="mt-2 flex flex-wrap gap-1">
                                <StatusPill s={s.mountpoint} /> <StatusPill s={s.uptime} /> <StatusPill s={s.timeSync} /> <StatusPill s={s.cpu} /> <StatusPill s={s.memory} /> <StatusPill s={s.logKillMemory} /> <StatusPill s={s.logError} />
                              </div>
                            </div>
                          </div>
                          <Button size="icon" variant="ghost" onClick={() => setFiles((p) => p.filter((x) => x.id !== f.id))}>
                            <Trash2 className="h-4 w-4" />
                          </Button>
                        </div>
                      </li>
                    );
                  })}
                </ul>
              )}
            </CardContent>
          </Card>

          <Card>
            <CardContent className="pt-6">
              <Button
                onClick={handleExport}
                disabled={busy}
                className="w-full h-11 bg-[hsl(var(--brand))] text-white hover:bg-[hsl(var(--brand))]/90"
              >
                <Download className="h-4 w-4 mr-2" />
                {busy ? "Membangun dokumen..." : "Export ke Word (.docx)"}
              </Button>
              <p className="mt-3 text-[11px] text-muted-foreground text-center">
                Setelah dibuka di Word, tekan <kbd className="px-1 rounded bg-muted">Ctrl</kbd>+<kbd className="px-1 rounded bg-muted">A</kbd> lalu <kbd className="px-1 rounded bg-muted">F9</kbd> untuk memperbarui Table of Contents.
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

function Field({ label, v, onChange }: { label: string; v: string; onChange: (v: string) => void }) {
  return (
    <div>
      <Label className="text-xs" dangerouslySetInnerHTML={{ __html: label }} />
      <Input value={v} onChange={(e) => onChange(e.target.value)} className="mt-1" />
    </div>
  );
}
