ALTER TABLE email_attachment_prefetch_items
  ADD COLUMN expected_content_sha256 TEXT;

UPDATE email_attachment_prefetch_items
SET expected_content_sha256 = content_sha256
WHERE content_sha256 IS NOT NULL;
