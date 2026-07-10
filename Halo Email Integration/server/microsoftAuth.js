const { createRemoteJWKSet, jwtVerify } = require("jose");

const DEFAULT_AUTHORITY = "https://login.microsoftonline.com/common";

function createMicrosoftAuthVerifier(options = {}) {
  const clientId = options.clientId || process.env.ADDIN_CLIENT_ID || "";
  const authority = normalizeAuthority(options.authority || process.env.ADDIN_AUTHORITY);
  const audience = normalizeAudience(
    options.audience || process.env.ADDIN_API_AUDIENCE || getApiAudience(clientId)
  );
  const tokenAudiences = getTokenAudiences(
    clientId,
    audience,
    options.apiClientId || process.env.ADDIN_API_CLIENT_ID
  );
  const requiredScope =
    options.requiredScope || process.env.ADDIN_REQUIRED_SCOPE || "access_as_user";

  if (!clientId) {
    return {
      enabled: false,
      async verify() {
        throw new Error("Microsoft add-in authentication is not configured.");
      },
    };
  }

  const jwks = createRemoteJWKSet(new URL(`${authority}/discovery/v2.0/keys`));

  return {
    enabled: true,
    async verify(token) {
      const result = await jwtVerify(token, jwks, {
        audience: tokenAudiences,
        clockTolerance: 30,
      });
      const claims = result.payload;

      validateMicrosoftClaims(claims, requiredScope);
      return claims;
    },
  };
}

function getMicrosoftAuthConfig(options = {}) {
  const clientId = options.clientId || process.env.ADDIN_CLIENT_ID || "";
  const authority = normalizeAuthority(options.authority || process.env.ADDIN_AUTHORITY);
  const audience = normalizeAudience(
    options.audience || process.env.ADDIN_API_AUDIENCE || getApiAudience(clientId)
  );
  const scopes = getAuthScopes(audience, options.scopes || process.env.ADDIN_AUTH_SCOPES);

  return {
    authority,
    clientId,
    scopes,
    ssoEnabled: Boolean(clientId),
  };
}

function getAuthScopes(audience, configuredScopes) {
  if (configuredScopes) {
    const scopes = configuredScopes
      .split(/[,\s]+/)
      .map((scope) => scope.trim())
      .filter(Boolean);

    if (audience && !scopes.some((scope) => scope.startsWith(`${audience}/`))) {
      throw new Error(
        `ADDIN_AUTH_SCOPES must request a delegated scope for the add-in API (${audience}/...).`
      );
    }

    return scopes;
  }

  return audience ? [`${audience}/access_as_user`] : [];
}

function getApiAudience(clientId) {
  return clientId ? `api://${clientId}` : "";
}

function getTokenAudiences(clientId, audience, configuredApiClientId) {
  return Array.from(
    new Set([configuredApiClientId || clientId, normalizeAudience(audience)].filter(Boolean))
  );
}

function normalizeAudience(value) {
  return String(value || "").replace(/\/+$/, "");
}

function normalizeAuthority(value) {
  return String(value || DEFAULT_AUTHORITY).replace(/\/+$/, "");
}

function validateMicrosoftClaims(claims, requiredScope = "access_as_user") {
  if (!claims || typeof claims !== "object") {
    throw new Error("Microsoft authentication token was invalid.");
  }

  if (!claims.tid || !(claims.oid || claims.sub)) {
    throw new Error("Microsoft authentication token did not include a stable user identity.");
  }

  const expectedIssuers = new Set([
    `https://login.microsoftonline.com/${claims.tid}/v2.0`,
    `https://sts.windows.net/${claims.tid}/`,
  ]);

  if (claims.iss && !expectedIssuers.has(claims.iss)) {
    throw new Error("Microsoft authentication token issuer was invalid.");
  }

  const scopes = new Set(String(claims.scp || "").split(/\s+/).filter(Boolean));
  if (requiredScope && !scopes.has(requiredScope)) {
    throw new Error(`Microsoft authentication token did not include the ${requiredScope} scope.`);
  }
}

module.exports = {
  createMicrosoftAuthVerifier,
  getApiAudience,
  getAuthScopes,
  getMicrosoftAuthConfig,
  getTokenAudiences,
  validateMicrosoftClaims,
};
