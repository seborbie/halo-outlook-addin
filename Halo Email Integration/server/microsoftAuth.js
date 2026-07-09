const { createRemoteJWKSet, jwtVerify } = require("jose");

const DEFAULT_AUTHORITY = "https://login.microsoftonline.com/common";

function createMicrosoftAuthVerifier(options = {}) {
  const clientId = options.clientId || process.env.ADDIN_CLIENT_ID || "";
  const authority = normalizeAuthority(options.authority || process.env.ADDIN_AUTHORITY);

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
        audience: clientId,
        clockTolerance: 30,
      });
      const claims = result.payload;

      validateMicrosoftClaims(claims);
      return claims;
    },
  };
}

function getMicrosoftAuthConfig(options = {}) {
  const clientId = options.clientId || process.env.ADDIN_CLIENT_ID || "";
  const authority = normalizeAuthority(options.authority || process.env.ADDIN_AUTHORITY);
  const scopes = getAuthScopes(clientId, options.scopes || process.env.ADDIN_AUTH_SCOPES);

  return {
    authority,
    clientId,
    scopes,
    ssoEnabled: Boolean(clientId),
  };
}

function getAuthScopes(clientId, configuredScopes) {
  if (configuredScopes) {
    return configuredScopes
      .split(/[,\s]+/)
      .map((scope) => scope.trim())
      .filter(Boolean);
  }

  return clientId ? ["openid", "profile", "email", "User.Read"] : [];
}

function normalizeAuthority(value) {
  return String(value || DEFAULT_AUTHORITY).replace(/\/+$/, "");
}

function validateMicrosoftClaims(claims) {
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
}

module.exports = {
  createMicrosoftAuthVerifier,
  getMicrosoftAuthConfig,
  validateMicrosoftClaims,
};
