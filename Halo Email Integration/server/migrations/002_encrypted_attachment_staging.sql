ALTER TABLE email_attachment_prefetch
  ADD COLUMN staging_version INTEGER NOT NULL DEFAULT 1,
  ADD COLUMN draft_item_id TEXT NOT NULL DEFAULT '',
  ADD COLUMN halo_action_id TEXT NOT NULL DEFAULT '',
  ADD COLUMN action_created_at BIGINT;

ALTER TABLE email_attachment_prefetch_items
  ADD COLUMN content_ciphertext BYTEA,
  ADD COLUMN content_iv BYTEA,
  ADD COLUMN content_tag BYTEA,
  ADD COLUMN content_key_id TEXT,
  ADD COLUMN content_sha256 TEXT,
  ADD COLUMN prepared_at BIGINT;

CREATE INDEX idx_email_attachment_prefetch_staging_expiry
  ON email_attachment_prefetch(staging_version, status, expires_at);
