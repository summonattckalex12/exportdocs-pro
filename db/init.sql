-- ExcportCuy — MySQL schema
-- Dijalankan otomatis saat container MySQL pertama kali start.

CREATE DATABASE IF NOT EXISTS excportcuy
  CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci;

USE excportcuy;

-- Metadata tiap file HTML PM yang di-upload dan ikut di-export.
CREATE TABLE IF NOT EXISTS pm_uploads (
  id            BIGINT AUTO_INCREMENT PRIMARY KEY,
  export_id     BIGINT NULL,
  filename      VARCHAR(255) NOT NULL,
  hostname      VARCHAR(255) NULL,
  ip_address    VARCHAR(64)  NULL,
  os_release    VARCHAR(255) NULL,
  size_bytes    INT UNSIGNED NULL,
  created_at    DATETIME DEFAULT CURRENT_TIMESTAMP,
  KEY idx_uploads_export (export_id),
  KEY idx_uploads_host   (hostname)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- Riwayat proses export DOCX.
CREATE TABLE IF NOT EXISTS exports (
  id             BIGINT AUTO_INCREMENT PRIMARY KEY,
  filename       VARCHAR(255) NOT NULL,
  storage_path   VARCHAR(512) NOT NULL,
  size_bytes     INT UNSIGNED NOT NULL,
  server_count   INT UNSIGNED NOT NULL DEFAULT 0,
  company_name   VARCHAR(255) NULL,
  vendor_name    VARCHAR(255) NULL,
  periode        VARCHAR(128) NULL,
  status         ENUM('pending','success','failed') NOT NULL DEFAULT 'success',
  error_message  TEXT NULL,
  created_at     DATETIME DEFAULT CURRENT_TIMESTAMP,
  KEY idx_exports_created (created_at)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

GRANT ALL PRIVILEGES ON excportcuy.* TO 'excport'@'%';
FLUSH PRIVILEGES;
