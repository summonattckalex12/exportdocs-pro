# ExcportCuy 🐉

Konverter laporan **HTML → Word (.docx)** untuk laporan Preventive Maintenance (PM). Dibangun dengan TanStack Start + React 19, dilengkapi cover kustom, header/footer otomatis, Table of Contents, dan penyimpanan riwayat export ke MySQL + storage lokal.

---

## ✨ Fitur

- **Upload fleksibel** — `.html`, `.zip`, `.tar`, `.tar.gz`, `.tgz` (batch banyak file sekaligus).
- **Drag & drop reorder** — atur urutan file sebelum export pakai handle grip.
- **Cover kustom** — background image, logo klien (kiri) & vendor (kanan), title block, kotak Pemberitahuan Kerahasiaan, blok "Dibuat untuk / Dibuat oleh".
- **Ukuran gambar konsisten** — semua logo di-scale ke bounding box tetap dengan aspect ratio dipertahankan.
- **Date picker** — kalender shadcn untuk Tanggal & Waktu serta Tanggal Reviewer.
- **Header & footer otomatis** — nama klien + judul di header, vendor + page number di footer (cover bebas header/footer).
- **Table of Contents rapi** dan appendix per-host dengan status berwarna.
- **Riwayat export** — metadata upload & hasil disimpan di MySQL, file DOCX diarsipkan ke `./data/exports`.
- **Docker-ready** — jalan di port `9812` via `docker compose`.

---

## 🚀 Quick Start

### Development
```bash
bun install
bun run dev
# http://localhost:8080
```

### Docker (produksi lokal)
```bash
./deploy.sh up              # app saja
./deploy.sh up --with-db    # app + MySQL
# akses: http://localhost:9812
```

---

## 🛠 deploy.sh

Script manajemen lifecycle container.

| Command | Deskripsi |
|---|---|
| `./deploy.sh up [--with-db]` | Build & start. Tambah `--with-db` untuk mengaktifkan MySQL. |
| `./deploy.sh down` | Stop & hapus container. |
| `./deploy.sh restart` | Restart tanpa rebuild. |
| `./deploy.sh rebuild` | Rebuild image dari nol. |
| `./deploy.sh logs` | Ikuti log app. |
| `./deploy.sh clean` | Bersihkan volume & data. |

---

## 📁 Struktur Data Lokal

```
./data/
├── exports/     # arsip DOCX hasil generate
└── mysql/       # data MySQL (jika --with-db)
```

---

## 🗄 Database Schema

Tabel dibuat otomatis lewat `db/init.sql`:

- `pm_uploads` — metadata file HTML yang di-upload (nama, ukuran, timestamp).
- `exports` — riwayat generate DOCX (nama file, path, jumlah host, status).

---

## 📝 Cara Pakai

1. Isi form cover: nama PT klien & vendor, alamat, OS, periode, no kontrak, tanggal.
2. Upload logo klien (kiri) & logo vendor (kanan). Opsional: background cover.
3. Upload file laporan HTML — bisa single file, ZIP, atau tar.gz.
4. Drag file untuk atur urutan appendix.
5. Klik **Generate DOCX**. File otomatis diunduh + diarsipkan di `./data/exports`.

---

## 🧱 Tech Stack

- **Framework**: TanStack Start v1 + React 19 + Vite 7
- **Styling**: Tailwind CSS v4 + shadcn/ui
- **DOCX**: [`docx`](https://www.npmjs.com/package/docx)
- **Arsip**: JSZip + pako (gunzip) + tar parser custom
- **DB**: MySQL 8 (via `mysql2`)
- **Deploy**: Docker + docker-compose (port `9812`)

---

## ⚙️ Environment

Variabel penting (di-set otomatis oleh `deploy.sh` saat `--with-db`):

```env
DB_HOST=db
DB_PORT=3306
DB_USER=excportcuy
DB_PASSWORD=excportcuy
DB_NAME=excportcuy
EXPORTS_DIR=/app/data/exports
```

---

## 📄 License

Internal tool. Gunakan sesuai kebutuhan tim.
