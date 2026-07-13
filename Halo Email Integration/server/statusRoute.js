function registerStatusRoute(app) {
  app.get("/", (_req, res) => {
    res.status(200).type("text/plain").send("Halo Outlook add-in is up.");
  });
}

module.exports = { registerStatusRoute };
