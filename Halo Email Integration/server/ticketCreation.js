const crypto = require("crypto");

const SUPPORTED_CORE_PROPERTIES = new Set([
  "agent_id",
  "asset_id",
  "category_1",
  "category_2",
  "category_3",
  "category_4",
  "client_id",
  "details",
  "estimate",
  "impact",
  "priority_id",
  "site_id",
  "status_id",
  "summary",
  "team_id",
  "urgency",
  "user_id",
]);

const FIELD_TYPE_NAMES = {
  bool: "boolean",
  boolean: "boolean",
  checkbox: "boolean",
  currency: "number",
  date: "date",
  datetime: "datetime",
  decimal: "number",
  dropdown: "select",
  duration: "duration",
  integer: "number",
  memo: "multiline",
  multiselect: "multiselect",
  multipleselection: "multiselect",
  number: "number",
  numeric: "number",
  select: "select",
  singleselection: "select",
  rich: "multiline",
  text: "text",
  textarea: "multiline",
  time: "time",
};

const ENTITY_TYPE_NAMES = {
  agent: "agent",
  asset: "asset",
  client: "client",
  company: "client",
  contact: "user",
  site: "site",
  team: "team",
  user: "user",
};

// These numeric values are Halo custom-field types. They must not be applied to
// standard ticket fields, whose internal type numbers have different meanings.
const CUSTOM_FIELD_TYPE_NUMBERS = {
  0: "text",
  1: "multiline",
  2: "select",
  3: "multiselect",
  4: "date",
  5: "time",
  6: "boolean",
  7: "unsupported",
  8: "multiline",
};

const STANDARD_FIELD_DEFINITIONS = {
  2: { managed: true, property: "summary", type: "text" },
  3: { property: "details", type: "multiline" },
  4: { entity: "asset", property: "asset_id", type: "asset" },
  13: { entity: "team", property: "team_id", type: "team" },
  14: { entity: "agent", property: "agent_id", required: true, type: "agent" },
  16: { property: "estimate", type: "duration" },
  27: { optionSource: "severity", property: "impact", type: "select" },
  28: { optionSource: "severity", property: "urgency", type: "select" },
};

function normalizeTicketTypes(payload) {
  return extractRecords(payload, ["tickettypes", "requesttypes", "types", "records"])
    .map((value) => normalizeTicketType(value))
    .filter((value) => value.id && value.name && value.active && value.canCreate)
    .sort((left, right) =>
      `${left.group}\u0000${left.name}`.localeCompare(
        `${right.group}\u0000${right.name}`,
        undefined,
        {
          sensitivity: "base",
        }
      )
    );
}

function normalizeTicketType(value) {
  const record = asObject(value);
  const id = positiveInteger(
    first(record, ["id", "tickettype_id", "tickettypeid", "requesttype_id"])
  );
  const inactive = booleanField(first(record, ["inactive", "is_inactive", "isinactive"]), false);
  const activeValue = first(record, ["active", "isactive", "is_active"]);
  const canCreateValue = first(record, [
    "cancreate",
    "can_create",
    "canagentscreate",
    "can_agents_create",
  ]);
  return {
    id: id ? String(id) : "",
    name: text(first(record, ["name", "tickettypename", "tickettype_name", "requesttypename"])),
    group: text(
      first(record, ["group_name", "groupname", "tickettypegroup_name", "requesttypegroup_name"])
    ),
    active: activeValue === undefined ? !inactive : booleanField(activeValue, true),
    canCreate: canCreateValue === undefined ? true : booleanField(canCreateValue, true),
  };
}

function normalizeTicketTypeSchema(typeId, typeName, payload, extraFieldsPayload) {
  const record = firstRecord(payload);
  const rawFields = uniqueRawFields([
    ...extractNestedFields(record),
    ...extractRecords(extraFieldsPayload, [
      "tickettypefields",
      "requesttypefields",
      "fields",
      "records",
    ]),
  ]);
  const fields = rawFields
    .map((field, index) => normalizeField(field, index))
    .filter((field) => field.key && field.visibleOnCreate)
    .sort((left, right) => left.order - right.order || left.label.localeCompare(right.label));
  const defaults = normalizeTypeDefaults(record);
  fields.forEach((field) => {
    if (isEmptyValue(field.defaultValue) && !isEmptyValue(defaults[field.key])) {
      field.defaultValue = defaults[field.key];
    }
  });
  return finalizeTicketTypeSchema({
    typeId: String(typeId),
    typeName: typeName || text(first(record, ["name", "tickettypename", "requesttypename"])),
    fields,
    defaults,
  });
}

function hydrateTicketCreationFieldOptions(schema, optionPayloads = {}) {
  const fields = schema.fields.map((field) => {
    if (!field.optionSource || field.options.length) {
      return { ...field };
    }
    const payload = optionPayloads[field.key] ?? optionPayloads[field.id];
    const options = normalizeOptionsPayload(payload);
    return {
      ...field,
      options,
      supported: options.length ? field.supported : false,
    };
  });
  return finalizeTicketTypeSchema({
    typeId: schema.typeId,
    typeName: schema.typeName,
    fields,
    defaults: schema.defaults,
  });
}

function finalizeTicketTypeSchema({ typeId, typeName, fields, defaults }) {
  const unsupportedRequired = fields.filter(
    (field) => !field.managed && field.required && !field.supported
  );
  const schema = {
    typeId,
    typeName,
    fields,
    defaults,
    available: unsupportedRequired.length === 0,
    unavailableReason: unsupportedRequired.length
      ? `Halo requires unsupported field${unsupportedRequired.length === 1 ? "" : "s"}: ${unsupportedRequired
          .map((field) => field.label)
          .join(", ")}.`
      : "",
    warnings: fields
      .filter((field) => !field.managed && !field.required && !field.supported)
      .map((field) => `${field.label} is not available in Outlook.`)
      .slice(0, 10),
  };
  schema.revision = hashJson(schema);
  return schema;
}

function normalizeField(value, index) {
  const placement = asObject(value);
  const fieldInfo = asObject(
    first(placement, ["fieldinfo", "field_info", "fielddefinition", "field_definition"])
  );
  // TicketType fields are placement/configuration rows in Halo. The descriptive
  // metadata used by Halo's own form (label, type, choices, etc.) lives in the
  // nested fieldinfo object on many Halo versions.
  const record = { ...fieldInfo, ...placement };
  const customFieldId =
    positiveInteger(
      first(placement, [
        "customfield_id",
        "customfieldid",
        "fieldinfo_id",
        "fieldinfoid",
        "field_id",
        "fieldid",
      ])
    ) ||
    positiveInteger(first(fieldInfo, ["id", "field_id", "fieldid"])) ||
    positiveInteger(first(placement, ["id"]));
  const rawName =
    text(first(fieldInfo, ["fieldname", "field_name", "name", "property", "database_name"])) ||
    text(first(placement, ["fieldname", "field_name", "name", "property", "database_name"]));
  const rawProperty = normalizeProperty(rawName);
  const standardDefinitionById = STANDARD_FIELD_DEFINITIONS[customFieldId] || null;
  const standardDefinitionByName = getStandardFieldDefinitionByProperty(rawProperty);
  const isCustom = booleanField(
    first(record, ["iscustom", "is_custom", "customfield", "custom_field", "custom"]),
    Boolean(
      customFieldId &&
      !standardDefinitionById &&
      !standardDefinitionByName &&
      !SUPPORTED_CORE_PROPERTIES.has(rawProperty)
    )
  );
  const standardDefinition = isCustom ? null : standardDefinitionByName || standardDefinitionById;
  const property = standardDefinition?.property || rawProperty;
  const key =
    isCustom && customFieldId ? `custom:${customFieldId}` : property ? `core:${property}` : "";
  const options = normalizeOptions(
    first(record, ["options", "values", "choices", "lookupvalues", "selection_values"])
  );
  const coreEntityByProperty = {
    agent_id: "agent",
    asset_id: "asset",
    client_id: "client",
    site_id: "site",
    team_id: "team",
    user_id: "user",
  };
  const entity =
    normalizeEntityType(
      first(record, ["entity", "lookupentity", "lookup_type", "entity_type", "table_name"])
    ) ||
    standardDefinition?.entity ||
    coreEntityByProperty[property] ||
    "";
  const rawFieldType = first(record, ["fieldtype", "field_type", "type"]);
  let fieldType =
    standardDefinition?.type || normalizeFieldType(rawFieldType, options, entity, isCustom);
  if (
    isCustom &&
    fieldType === "text" &&
    Number(first(record, ["inputtype", "input_type"])) === 1
  ) {
    fieldType = "number";
  }
  const agentNewVisibility = finiteNumber(
    first(placement, ["technew", "agentnew", "agent_new", "newticketvisibility"])
  );
  const explicitlyRequired = booleanField(
    first(record, ["mandatory", "required", "isrequired", "is_required"]),
    false
  );
  // Halo stores Agent New Ticket Screen visibility per ticket type. Level 0 is
  // hidden, 1 is optional, 2 is recommended, and 3 is mandatory.
  const required =
    explicitlyRequired || Boolean(standardDefinition?.required) || agentNewVisibility === 3;
  const supported = Boolean(
    key &&
    (isCustom || SUPPORTED_CORE_PROPERTIES.has(property)) &&
    [
      "agent",
      "asset",
      "boolean",
      "client",
      "date",
      "datetime",
      "duration",
      "multiline",
      "multiselect",
      "number",
      "select",
      "site",
      "team",
      "text",
      "time",
      "user",
    ].includes(fieldType)
  );
  return {
    key,
    id: customFieldId ? String(customFieldId) : "",
    property: isCustom ? "" : property,
    label:
      text(
        first(placement, ["label", "displayname", "display_name", "fieldlabel", "field_label"])
      ) ||
      text(
        first(fieldInfo, [
          "label",
          "displayname",
          "display_name",
          "fieldlabel",
          "field_label",
          "labellong",
          "label_long",
        ])
      ) ||
      text(first(fieldInfo, ["name", "fieldname", "field_name"])) ||
      rawName ||
      `Field ${index + 1}`,
    type: fieldType,
    entity,
    required,
    supported,
    defaultValue: normalizeDefaultValue(
      first(record, ["defaultvalue", "default_value", "default", "value"]),
      fieldType
    ),
    options,
    optionSource:
      standardDefinition?.optionSource ||
      (["select", "multiselect"].includes(fieldType) && !options.length && customFieldId
        ? "fieldinfo"
        : ""),
    order:
      finiteNumber(first(record, ["sequence", "seq", "order", "displayorder", "display_order"])) ??
      index,
    core: !isCustom,
    managed: Boolean(standardDefinition?.managed),
    recommended: !required && agentNewVisibility === 2,
    visibleOnCreate: agentNewVisibility === null || agentNewVisibility > 0,
  };
}

function validateCreationInput(schema, input) {
  const source = asObject(input);
  const summary = text(source.summary).slice(0, 500);
  if (!summary) {
    throw requestError("A ticket summary is required.", 400, "summary");
  }
  if (!schema.available) {
    throw requestError(schema.unavailableReason || "This ticket type is unavailable.", 409);
  }
  if (source.schemaRevision && source.schemaRevision !== schema.revision) {
    throw requestError("Halo ticket fields changed. Reopen the creation form and review it.", 409);
  }
  const submitted = asObject(source.values);
  const values = {};
  for (const field of schema.fields) {
    if (field.managed || !field.supported) {
      continue;
    }
    let value = Object.prototype.hasOwnProperty.call(submitted, field.key)
      ? submitted[field.key]
      : field.defaultValue;
    value = normalizeSubmittedValue(value, field);
    if (field.required && isEmptyValue(value)) {
      throw requestError(`${field.label} is required by Halo.`, 400, field.key);
    }
    if (!isEmptyValue(value)) {
      values[field.key] = value;
    }
  }
  return { summary, values };
}

function buildTicketPayload({ schema, summary, values, requester }) {
  const originNote =
    "Created from an Outlook email. The source email is recorded in the following Email action.";
  const payload = {
    tickettype_id: Number(schema.typeId),
    summary,
    details: originNote,
  };
  const customfields = [];
  for (const field of schema.fields) {
    if (!Object.prototype.hasOwnProperty.call(values, field.key)) {
      continue;
    }
    if (field.core && field.property) {
      if (field.property === "details") {
        payload.details = `${values[field.key]}\n\n${originNote}`;
      } else if (field.property !== "summary") {
        payload[field.property] = values[field.key];
      }
    } else if (field.id) {
      customfields.push({ id: Number(field.id), value: values[field.key] });
    }
  }
  const normalizedRequester = normalizeRequester(requester);
  if (normalizedRequester.id) {
    payload.user_id = Number(normalizedRequester.id);
  }
  if (normalizedRequester.clientId) {
    payload.client_id = Number(normalizedRequester.clientId);
  }
  if (normalizedRequester.siteId) {
    payload.site_id = Number(normalizedRequester.siteId);
  }
  if (customfields.length) {
    payload.customfields = customfields;
  }
  return payload;
}

function normalizeRequesters(payload) {
  return extractRecords(payload, ["users", "requesters", "records"])
    .map((value) => normalizeRequester(value))
    .filter((value) => value.id && (value.emailAddress || value.name));
}

function normalizeLookupResults(payload, entity) {
  const keysByEntity = {
    agent: ["agents", "records"],
    asset: ["assets", "records"],
    client: ["clients", "records"],
    site: ["sites", "records"],
    team: ["teams", "records"],
    user: ["users", "records"],
  };
  return extractRecords(payload, keysByEntity[entity] || ["records"])
    .map((value) => {
      const record = asObject(value);
      const id = positiveInteger(
        first(record, ["id", `${entity}_id`, `${entity}id`, "user_id", "client_id"])
      );
      return {
        id: id ? String(id) : "",
        label:
          text(
            first(record, ["name", "display_name", "displayname", "asset_name", "inventory_number"])
          ) || (id ? `${entity} ${id}` : ""),
        secondary: text(
          first(record, ["email", "emailaddress", "client_name", "site_name", "inventory_number"])
        ),
      };
    })
    .filter((value) => value.id && value.label)
    .slice(0, 20);
}

function normalizeRequester(value) {
  const record = asObject(value);
  const id = positiveInteger(first(record, ["id", "user_id", "userid"]));
  const clientId = positiveInteger(first(record, ["client_id", "clientid"]));
  const siteId = positiveInteger(first(record, ["site_id", "siteid"]));
  return {
    id: id ? String(id) : "",
    name: text(first(record, ["name", "display_name", "displayname"])),
    emailAddress: text(first(record, ["emailaddress", "email_address", "email"])).toLowerCase(),
    clientId: clientId ? String(clientId) : "",
    clientName: text(first(record, ["client_name", "clientname", "company_name"])),
    siteId: siteId ? String(siteId) : "",
    siteName: text(first(record, ["site_name", "sitename"])),
  };
}

function getCreatedTicket(payload) {
  const record = extractRecords(payload, ["tickets", "records"])[0] || {};
  const id = positiveInteger(first(record, ["id", "ticket_id", "ticketid"]));
  if (!id) {
    return null;
  }
  return {
    id: String(id),
    ticketNumber:
      text(first(record, ["ticket_number", "ticketnumber", "faultid", "reference"])) || String(id),
  };
}

function normalizeTypeDefaults(record) {
  const defaults = {};
  const mappings = {
    "core:agent_id": [
      "default_agent_id",
      "defaultagent_id",
      "defaultagentid",
      "default_agent",
      "defaultagent",
    ],
    "core:impact": ["default_impact", "defaultimpact"],
    "core:priority_id": ["default_priority_id", "defaultpriority_id", "defaultpriorityid"],
    "core:status_id": ["default_status_id", "defaultstatus_id", "defaultstatusid"],
    "core:team_id": ["default_team_id", "defaultteam_id", "defaultteamid"],
    "core:urgency": ["default_urgency", "defaulturgency"],
  };
  for (const [key, aliases] of Object.entries(mappings)) {
    const value = first(record, aliases);
    if (!isEmptyValue(value)) {
      defaults[key] = value;
    }
  }
  return defaults;
}

function extractNestedFields(record) {
  for (const key of [
    "fields",
    "tickettypefields",
    "requesttypefields",
    "fieldlist",
    "customfields",
  ]) {
    if (Array.isArray(record[key])) {
      return record[key];
    }
  }
  return [];
}

function uniqueRawFields(values) {
  const seen = new Set();
  return values.filter((value, index) => {
    const record = asObject(value);
    const fieldInfo = asObject(
      first(record, ["fieldinfo", "field_info", "fielddefinition", "field_definition"])
    );
    const key =
      text(
        first(record, [
          "customfield_id",
          "customfieldid",
          "fieldinfo_id",
          "fieldinfoid",
          "field_id",
          "fieldid",
        ])
      ) ||
      text(first(fieldInfo, ["id", "field_id", "fieldid", "fieldname", "name"])) ||
      text(first(record, ["id", "fieldname", "name"])) ||
      String(index);
    if (seen.has(key)) {
      return false;
    }
    seen.add(key);
    return true;
  });
}

function normalizeOptions(value) {
  const values = Array.isArray(value) ? value : [];
  return values
    .map((entry) => {
      if (entry === null || entry === undefined) {
        return null;
      }
      if (typeof entry !== "object") {
        return { value: String(entry), label: String(entry) };
      }
      const record = asObject(entry);
      const optionValue = first(record, ["id", "value", "key", "code"]);
      const label = text(first(record, ["name", "label", "display", "text", "value"]));
      return isEmptyValue(optionValue)
        ? null
        : { value: String(optionValue), label: label || String(optionValue) };
    })
    .filter(Boolean);
}

function normalizeOptionsPayload(payload) {
  if (payload === undefined || payload === null) {
    return [];
  }
  if (Array.isArray(payload)) {
    return normalizeOptions(payload);
  }
  const record = asObject(payload);
  return normalizeOptions(
    first(record, ["options", "values", "choices", "lookupvalues", "selection_values"]) ||
      extractRecords(payload, ["records"])
  );
}

function normalizeFieldType(value, options, entity, allowNumericTypes = true) {
  if (entity) {
    return entity;
  }
  const normalized = text(value)
    .toLowerCase()
    .replace(/[\s_-]+/g, "");
  if (normalized && FIELD_TYPE_NAMES[normalized]) {
    return FIELD_TYPE_NAMES[normalized];
  }
  if (
    allowNumericTypes &&
    Object.prototype.hasOwnProperty.call(CUSTOM_FIELD_TYPE_NUMBERS, normalized)
  ) {
    return CUSTOM_FIELD_TYPE_NUMBERS[normalized];
  }
  if (options.length) {
    return "select";
  }
  if (value === undefined || value === null || value === "") {
    return "text";
  }
  return "unsupported";
}

function normalizeEntityType(value) {
  const normalized = text(value)
    .toLowerCase()
    .replace(/[\s_-]+/g, "");
  return ENTITY_TYPE_NAMES[normalized] || "";
}

function normalizeProperty(value) {
  const normalized = text(value)
    .trim()
    .replace(/([a-z0-9])([A-Z])/g, "$1_$2")
    .replace(/[^A-Za-z0-9]+/g, "_")
    .replace(/^_+|_+$/g, "")
    .toLowerCase();
  const aliases = {
    agent: "agent_id",
    assignedtoint: "agent_id",
    asset: "asset_id",
    client: "client_id",
    company: "client_id",
    priority: "priority_id",
    estimate: "estimate",
    sectio: "team_id",
    sectio_: "team_id",
    site: "site_id",
    status: "status_id",
    symptom: "summary",
    symptom2: "details",
    team: "team_id",
    user: "user_id",
  };
  return aliases[normalized] || normalized;
}

function getStandardFieldDefinitionByProperty(property) {
  const definitionByProperty = {
    agent_id: { entity: "agent", property: "agent_id", required: true, type: "agent" },
    asset_id: { entity: "asset", property: "asset_id", type: "asset" },
    details: { property: "details", type: "multiline" },
    estimate: { property: "estimate", type: "duration" },
    impact: { optionSource: "severity", property: "impact", type: "select" },
    summary: { managed: true, property: "summary", type: "text" },
    team_id: { entity: "team", property: "team_id", type: "team" },
    urgency: { optionSource: "severity", property: "urgency", type: "select" },
  };
  return definitionByProperty[property] || null;
}

function normalizeSubmittedValue(value, field) {
  if (isEmptyValue(value)) {
    return field.type === "multiselect" ? [] : "";
  }
  if (field.type === "boolean") {
    return booleanField(value, false);
  }
  if (field.type === "number") {
    const number = finiteNumber(value);
    if (number === null) {
      throw requestError(`${field.label} must be a number.`, 400, field.key);
    }
    return number;
  }
  if (field.type === "duration") {
    if (typeof value === "number" && Number.isFinite(value) && value >= 0) {
      return value;
    }
    const match = /^(\d{1,4}):([0-5]\d)$/.exec(String(value).trim());
    if (!match) {
      throw requestError(`${field.label} must use hours and minutes (HH:MM).`, 400, field.key);
    }
    return Number(match[1]) + Number(match[2]) / 60;
  }
  if (field.type === "multiselect") {
    const selected = Array.isArray(value) ? value.map(String) : [String(value)];
    validateOptions(field, selected);
    return selected;
  }
  if (field.type === "select") {
    const selected = String(value);
    validateOptions(field, [selected]);
    return selected;
  }
  if (["agent", "asset", "client", "site", "team", "user"].includes(field.type)) {
    const id = positiveInteger(value);
    if (!id) {
      throw requestError(`${field.label} must reference a valid Halo record.`, 400, field.key);
    }
    return id;
  }
  if (field.type === "date" || field.type === "datetime") {
    const date = new Date(value);
    if (Number.isNaN(date.getTime())) {
      throw requestError(`${field.label} must be a valid date.`, 400, field.key);
    }
    return field.type === "date" ? String(value).slice(0, 10) : date.toISOString();
  }
  if (field.type === "time") {
    const time = String(value).trim();
    if (!/^([01]\d|2[0-3]):[0-5]\d(?::[0-5]\d)?$/.test(time)) {
      throw requestError(`${field.label} must be a valid time.`, 400, field.key);
    }
    return time;
  }
  return text(value).slice(0, field.type === "multiline" ? 10000 : 2000);
}

function validateOptions(field, values) {
  if (!field.options.length) {
    return;
  }
  const allowed = new Set(field.options.map((option) => String(option.value)));
  if (values.some((value) => !allowed.has(String(value)))) {
    throw requestError(`${field.label} contains an invalid Halo option.`, 400, field.key);
  }
}

function normalizeDefaultValue(value, fieldType) {
  if (value === undefined || value === null) {
    return fieldType === "multiselect" ? [] : "";
  }
  if (fieldType === "boolean") {
    return booleanField(value, false);
  }
  if (fieldType === "multiselect") {
    return Array.isArray(value) ? value : [value];
  }
  return value;
}

function extractRecords(payload, keys) {
  if (Array.isArray(payload)) {
    return payload;
  }
  const record = asObject(payload);
  for (const key of [...keys, "data", "results"]) {
    const value = first(record, [key]);
    if (Array.isArray(value)) {
      return value;
    }
    if (value && typeof value === "object") {
      const nested = extractRecords(value, keys);
      if (nested.length) {
        return nested;
      }
    }
  }
  return Object.keys(record).length ? [record] : [];
}

function firstRecord(payload) {
  return extractRecords(payload, ["tickettypes", "requesttypes", "types", "records"])[0] || {};
}

function first(record, keys) {
  for (const key of keys) {
    if (Object.prototype.hasOwnProperty.call(record, key)) {
      return record[key];
    }
    const found = Object.keys(record).find(
      (candidate) => candidate.toLowerCase() === key.toLowerCase()
    );
    if (found) {
      return record[found];
    }
  }
  return undefined;
}

function asObject(value) {
  return value && typeof value === "object" && !Array.isArray(value) ? value : {};
}

function positiveInteger(value) {
  const number = Number.parseInt(String(value || ""), 10);
  return Number.isSafeInteger(number) && number > 0 ? number : 0;
}

function finiteNumber(value) {
  if (value === "" || value === null || value === undefined) {
    return null;
  }
  const number = Number(value);
  return Number.isFinite(number) ? number : null;
}

function booleanField(value, fallback) {
  if (value === undefined || value === null || value === "") {
    return fallback;
  }
  if (typeof value === "string") {
    return !["0", "false", "no", "off"].includes(value.toLowerCase());
  }
  return Boolean(value);
}

function text(value) {
  return value === undefined || value === null ? "" : String(value).trim();
}

function isEmptyValue(value) {
  return (
    value === undefined || value === null || value === "" || (Array.isArray(value) && !value.length)
  );
}

function hashJson(value) {
  return crypto.createHash("sha256").update(JSON.stringify(value)).digest("hex");
}

function requestError(message, status = 400, fieldKey = "") {
  const error = new Error(message);
  error.status = status;
  error.isRequestError = true;
  error.fieldKey = fieldKey;
  return error;
}

module.exports = {
  buildTicketPayload,
  getCreatedTicket,
  hydrateTicketCreationFieldOptions,
  normalizeRequesters,
  normalizeLookupResults,
  normalizeTicketTypeSchema,
  normalizeTicketTypes,
  validateCreationInput,
};
