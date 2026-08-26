function registerStatusRoute(app, options = {}) {
  app.get("/", (_req, res) => {
    res.status(200).type("text/plain").send("Halo Outlook add-in is up.");
  });

  app.get("/health/ready", async (_req, res) => {
    const checkReady = options.checkReady;
    try {
      const ready = typeof checkReady === "function" && (await checkReady());
      if (!ready) {
        res.status(503).json({ ready: false });
        return;
      }
      res.status(200).json({ ready: true });
    } catch {
      res.status(503).json({ ready: false });
    }
  });
}

module.exports = { registerStatusRoute };
