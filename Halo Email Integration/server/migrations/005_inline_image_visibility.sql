ALTER TABLE inline_image_cache
  ADD COLUMN show_for_users BOOLEAN NOT NULL DEFAULT TRUE;

ALTER TABLE inline_image_cache
  DROP CONSTRAINT inline_image_cache_pkey;

ALTER TABLE inline_image_cache
  ADD PRIMARY KEY (halo_tenant, sha256, show_for_users);
