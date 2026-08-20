function registerStatusRoute(app) {
  app.get("/api/health", (_req, res) => {
    res.status(200).json({ ok: true, service: "inboxlink" });
  });
}

module.exports = { registerStatusRoute };
