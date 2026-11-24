CREATE TABLE IF NOT EXISTS license_requests (
  id TEXT PRIMARY KEY,
  device_id TEXT NOT NULL UNIQUE,
  request_code TEXT NOT NULL,
  status TEXT NOT NULL CHECK (status IN ('pending', 'approved', 'rejected')),
  license_token TEXT,
  created_at TEXT NOT NULL,
  updated_at TEXT NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_license_requests_status_created
ON license_requests (status, created_at);
