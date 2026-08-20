CREATE TABLE organisations (
  id UUID PRIMARY KEY,
  microsoft_tenant_id TEXT NOT NULL UNIQUE,
  slug TEXT NOT NULL UNIQUE,
  name TEXT NOT NULL,
  halo_url TEXT NOT NULL,
  halo_client_id TEXT NOT NULL,
  status TEXT NOT NULL DEFAULT 'active' CHECK (status IN ('active', 'suspended')),
  created_at BIGINT NOT NULL,
  updated_at BIGINT NOT NULL
);

CREATE TABLE users (
  id UUID NOT NULL,
  organisation_id UUID NOT NULL,
  object_id TEXT NOT NULL,
  email TEXT NOT NULL DEFAULT '',
  display_name TEXT NOT NULL DEFAULT '',
  role TEXT NOT NULL DEFAULT 'member' CHECK (role IN ('owner', 'admin', 'member')),
  created_at BIGINT NOT NULL,
  updated_at BIGINT NOT NULL,
  PRIMARY KEY (organisation_id, id),
  UNIQUE (organisation_id, object_id),
  FOREIGN KEY (organisation_id) REFERENCES organisations(id) ON DELETE CASCADE
);

CREATE TABLE halo_grants (
  id UUID NOT NULL,
  organisation_id UUID NOT NULL,
  user_id UUID NOT NULL,
  halo_url TEXT NOT NULL,
  client_id TEXT NOT NULL,
  scope TEXT NOT NULL,
  encrypted_token_json JSONB NOT NULL,
  invalidated_at BIGINT,
  created_at BIGINT NOT NULL,
  updated_at BIGINT NOT NULL,
  PRIMARY KEY (organisation_id, id),
  UNIQUE (organisation_id, user_id),
  FOREIGN KEY (organisation_id, user_id)
    REFERENCES users(organisation_id, id) ON DELETE CASCADE
);

CREATE TABLE sessions (
  organisation_id UUID NOT NULL,
  session_hash TEXT NOT NULL,
  user_id UUID NOT NULL,
  expires_at BIGINT NOT NULL,
  created_at BIGINT NOT NULL,
  updated_at BIGINT NOT NULL,
  PRIMARY KEY (organisation_id, session_hash),
  FOREIGN KEY (organisation_id, user_id)
    REFERENCES users(organisation_id, id) ON DELETE CASCADE
);

CREATE TABLE background_sessions (
  organisation_id UUID NOT NULL,
  background_session_hash TEXT NOT NULL,
  session_hash TEXT NOT NULL,
  expires_at BIGINT NOT NULL,
  created_at BIGINT NOT NULL,
  updated_at BIGINT NOT NULL,
  PRIMARY KEY (organisation_id, background_session_hash),
  FOREIGN KEY (organisation_id, session_hash)
    REFERENCES sessions(organisation_id, session_hash) ON DELETE CASCADE
);

CREATE TABLE bug_report_sessions (
  organisation_id UUID NOT NULL,
  session_hash TEXT NOT NULL,
  user_id UUID NOT NULL,
  diagnostics_json JSONB NOT NULL DEFAULT '{}'::jsonb,
  expires_at BIGINT NOT NULL,
  claimed_at BIGINT,
  consumed_at BIGINT,
  created_at BIGINT NOT NULL,
  PRIMARY KEY (organisation_id, session_hash),
  FOREIGN KEY (organisation_id, user_id)
    REFERENCES users(organisation_id, id) ON DELETE CASCADE
);

CREATE TABLE conversation_mappings (
  organisation_id UUID NOT NULL,
  id UUID NOT NULL,
  mailbox_email TEXT NOT NULL,
  ticket_id BIGINT NOT NULL,
  ticket_number TEXT NOT NULL,
  conversation_id TEXT NOT NULL DEFAULT '',
  normalized_subject TEXT NOT NULL DEFAULT '',
  created_at BIGINT NOT NULL,
  updated_at BIGINT NOT NULL,
  PRIMARY KEY (organisation_id, id)
);

CREATE TABLE message_mappings (
  organisation_id UUID NOT NULL,
  mailbox_email TEXT NOT NULL,
  message_id_key TEXT NOT NULL,
  mapping_id UUID NOT NULL,
  created_at BIGINT NOT NULL,
  PRIMARY KEY (organisation_id, mailbox_email, message_id_key),
  FOREIGN KEY (organisation_id, mapping_id)
    REFERENCES conversation_mappings(organisation_id, id) ON DELETE CASCADE
);

CREATE INDEX idx_sessions_user_id ON sessions(organisation_id, user_id);
CREATE INDEX idx_sessions_expires_at ON sessions(organisation_id, expires_at);
CREATE INDEX idx_background_sessions_expires_at
  ON background_sessions(organisation_id, expires_at);
CREATE INDEX idx_bug_report_sessions_expires_at
  ON bug_report_sessions(organisation_id, expires_at);
CREATE INDEX idx_conversation_mappings_conversation
  ON conversation_mappings(organisation_id, mailbox_email, conversation_id);
CREATE INDEX idx_message_mappings_mapping_id
  ON message_mappings(organisation_id, mapping_id);

ALTER TABLE organisations ENABLE ROW LEVEL SECURITY;
ALTER TABLE organisations FORCE ROW LEVEL SECURITY;
CREATE POLICY organisations_tenant_isolation ON organisations
  USING (
    microsoft_tenant_id = current_setting('app.current_microsoft_tenant_id', true)
    OR id::text = current_setting('app.current_organisation_id', true)
  )
  WITH CHECK (
    microsoft_tenant_id = current_setting('app.current_microsoft_tenant_id', true)
    OR id::text = current_setting('app.current_organisation_id', true)
  );

ALTER TABLE users ENABLE ROW LEVEL SECURITY;
ALTER TABLE users FORCE ROW LEVEL SECURITY;
CREATE POLICY users_tenant_isolation ON users
  USING (organisation_id::text = current_setting('app.current_organisation_id', true))
  WITH CHECK (organisation_id::text = current_setting('app.current_organisation_id', true));

ALTER TABLE halo_grants ENABLE ROW LEVEL SECURITY;
ALTER TABLE halo_grants FORCE ROW LEVEL SECURITY;
CREATE POLICY halo_grants_tenant_isolation ON halo_grants
  USING (organisation_id::text = current_setting('app.current_organisation_id', true))
  WITH CHECK (organisation_id::text = current_setting('app.current_organisation_id', true));

ALTER TABLE sessions ENABLE ROW LEVEL SECURITY;
ALTER TABLE sessions FORCE ROW LEVEL SECURITY;
CREATE POLICY sessions_tenant_isolation ON sessions
  USING (organisation_id::text = current_setting('app.current_organisation_id', true))
  WITH CHECK (organisation_id::text = current_setting('app.current_organisation_id', true));

ALTER TABLE background_sessions ENABLE ROW LEVEL SECURITY;
ALTER TABLE background_sessions FORCE ROW LEVEL SECURITY;
CREATE POLICY background_sessions_tenant_isolation ON background_sessions
  USING (organisation_id::text = current_setting('app.current_organisation_id', true))
  WITH CHECK (organisation_id::text = current_setting('app.current_organisation_id', true));

ALTER TABLE bug_report_sessions ENABLE ROW LEVEL SECURITY;
ALTER TABLE bug_report_sessions FORCE ROW LEVEL SECURITY;
CREATE POLICY bug_report_sessions_tenant_isolation ON bug_report_sessions
  USING (organisation_id::text = current_setting('app.current_organisation_id', true))
  WITH CHECK (organisation_id::text = current_setting('app.current_organisation_id', true));

ALTER TABLE conversation_mappings ENABLE ROW LEVEL SECURITY;
ALTER TABLE conversation_mappings FORCE ROW LEVEL SECURITY;
CREATE POLICY conversation_mappings_tenant_isolation ON conversation_mappings
  USING (organisation_id::text = current_setting('app.current_organisation_id', true))
  WITH CHECK (organisation_id::text = current_setting('app.current_organisation_id', true));

ALTER TABLE message_mappings ENABLE ROW LEVEL SECURITY;
ALTER TABLE message_mappings FORCE ROW LEVEL SECURITY;
CREATE POLICY message_mappings_tenant_isolation ON message_mappings
  USING (organisation_id::text = current_setting('app.current_organisation_id', true))
  WITH CHECK (organisation_id::text = current_setting('app.current_organisation_id', true));
