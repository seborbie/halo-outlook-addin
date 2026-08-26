ALTER TABLE conversation_mappings
  ADD COLUMN action_mode TEXT NOT NULL DEFAULT 'email';

ALTER TABLE conversation_mappings
  ADD CONSTRAINT conversation_mappings_action_mode_check
  CHECK (action_mode IN ('email', 'private-note'));
