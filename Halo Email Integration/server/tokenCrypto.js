const crypto = require("crypto");

const TOKEN_KEY_ID = "env-v1";

function createTokenCrypto(env = process.env) {
  const key = decodeEncryptionKey(env.HALO_TOKEN_ENCRYPTION_KEY, env);

  return {
    decryptJson(value) {
      const decipher = crypto.createDecipheriv(
        "aes-256-gcm",
        key,
        Buffer.from(value.iv, "base64")
      );
      decipher.setAuthTag(Buffer.from(value.tag, "base64"));

      const plaintext = Buffer.concat([
        decipher.update(Buffer.from(value.ciphertext, "base64")),
        decipher.final(),
      ]).toString("utf8");

      return JSON.parse(plaintext);
    },

    encryptJson(value) {
      const iv = crypto.randomBytes(12);
      const cipher = crypto.createCipheriv("aes-256-gcm", key, iv);
      const ciphertext = Buffer.concat([
        cipher.update(JSON.stringify(value), "utf8"),
        cipher.final(),
      ]);

      return {
        keyId: TOKEN_KEY_ID,
        iv: iv.toString("base64"),
        ciphertext: ciphertext.toString("base64"),
        tag: cipher.getAuthTag().toString("base64"),
      };
    },
  };
}

function decodeEncryptionKey(value, env = process.env) {
  if (!value) {
    if (env.NODE_ENV === "test" || env.HALO_AUTH_TEST_MODE === "1") {
      return Buffer.alloc(32, 7);
    }

    throw new Error("HALO_TOKEN_ENCRYPTION_KEY must be set to a 32-byte base64 or base64url value.");
  }

  const normalized = value.replace(/-/g, "+").replace(/_/g, "/");
  const padded = normalized.padEnd(normalized.length + ((4 - (normalized.length % 4)) % 4), "=");
  const key = Buffer.from(padded, "base64");

  if (key.length !== 32) {
    throw new Error("HALO_TOKEN_ENCRYPTION_KEY must decode to exactly 32 bytes.");
  }

  return key;
}

module.exports = {
  createTokenCrypto,
  decodeEncryptionKey,
};
