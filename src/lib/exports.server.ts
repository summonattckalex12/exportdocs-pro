// Server-only helpers for archiving exported DOCX files.
// Loaded lazily from the server function so mysql2 is never bundled to the browser.

import { promises as fs } from "node:fs";
import path from "node:path";

const EXPORT_DIR = process.env.EXPORT_DIR || "/app/data/exports";

export interface ArchivePayload {
  filename: string;
  base64: string;
  meta: {
    serverCount: number;
    companyName?: string;
    vendorName?: string;
    periode?: string;
    hosts: { hostname: string; ipAddress?: string; osRelease?: string }[];
  };
}

export interface ArchiveResult {
  id: number | null;
  storagePath: string;
  sizeBytes: number;
  dbEnabled: boolean;
  dbError?: string;
}

function sanitize(name: string): string {
  return name.replace(/[^\w.\-]+/g, "_").slice(0, 180) || "export.docx";
}

async function ensureDir(dir: string) {
  await fs.mkdir(dir, { recursive: true });
}

async function getPool() {
  const url = process.env.DATABASE_URL;
  if (!url) return null;
  try {
    const mysql = await import("mysql2/promise");
    return mysql.createPool(url + (url.includes("?") ? "&" : "?") + "connectionLimit=4");
  } catch (e) {
    console.warn("[archive] mysql2 unavailable:", (e as Error).message);
    return null;
  }
}

export async function archiveDocx(payload: ArchivePayload): Promise<ArchiveResult> {
  await ensureDir(EXPORT_DIR);
  const stamp = new Date().toISOString().replace(/[:.]/g, "-");
  const safe = sanitize(payload.filename);
  const finalName = `${stamp}__${safe}`;
  const fullPath = path.join(EXPORT_DIR, finalName);

  const buf = Buffer.from(payload.base64, "base64");
  await fs.writeFile(fullPath, buf);

  const result: ArchiveResult = {
    id: null,
    storagePath: fullPath,
    sizeBytes: buf.byteLength,
    dbEnabled: false,
  };

  const pool = await getPool();
  if (!pool) return result;

  try {
    result.dbEnabled = true;
    const [insert]: any = await pool.execute(
      `INSERT INTO exports
       (filename, storage_path, size_bytes, server_count, company_name, vendor_name, periode, status)
       VALUES (?, ?, ?, ?, ?, ?, ?, 'success')`,
      [
        finalName,
        fullPath,
        buf.byteLength,
        payload.meta.serverCount,
        payload.meta.companyName ?? null,
        payload.meta.vendorName ?? null,
        payload.meta.periode ?? null,
      ],
    );
    const exportId = insert.insertId as number;
    result.id = exportId;

    if (payload.meta.hosts.length) {
      const values = payload.meta.hosts.map((h) => [
        exportId,
        sanitize(h.hostname || "unknown") + ".html",
        h.hostname ?? null,
        h.ipAddress ?? null,
        h.osRelease ?? null,
      ]);
      await pool.query(
        `INSERT INTO pm_uploads (export_id, filename, hostname, ip_address, os_release) VALUES ?`,
        [values],
      );
    }
  } catch (e) {
    result.dbError = (e as Error).message;
    console.warn("[archive] DB insert failed:", result.dbError);
  } finally {
    await pool.end().catch(() => undefined);
  }

  return result;
}

export async function listExports(limit = 50) {
  const pool = await getPool();
  if (!pool) return [];
  try {
    const [rows]: any = await pool.query(
      `SELECT id, filename, size_bytes, server_count, company_name, vendor_name, periode, status, created_at
       FROM exports ORDER BY created_at DESC LIMIT ?`,
      [limit],
    );
    return rows;
  } catch (e) {
    console.warn("[history] list failed:", (e as Error).message);
    return [];
  } finally {
    await pool.end().catch(() => undefined);
  }
}
