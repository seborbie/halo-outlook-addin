const fs = require("fs");
const path = require("path");

require("dotenv").config({ path: path.join(__dirname, "..", ".env") });

const express = require("express");
const { registerHaloAuthRoutes } = require("./haloAuth");
const { registerStatusRoute } = require("./statusRoute");

const app = express();
const port = Number(process.env.PORT || 3000);
const distPath = path.join(__dirname, "..", "dist");
const stagedWebPath = path.join(__dirname, "..", "web");
const workspaceWebPath = path.join(__dirname, "..", "..", "apps", "web", "build");
const webPath = fs.existsSync(stagedWebPath) ? stagedWebPath : workspaceWebPath;

app.disable("x-powered-by");
app.use(express.json({ limit: "2mb" }));

registerHaloAuthRoutes(app);
registerStatusRoute(app);

app.get("/bugreport", (_req, res) => {
  res.sendFile(path.join(distPath, "bugreport.html"));
});

app.use(express.static(distPath));
app.use("/public", express.static(distPath));
app.use(express.static(webPath, { extensions: ["html"] }));

app.get(/.*/, (req, res) => {
  res.sendFile(path.join(webPath, "200.html"));
});

app.listen(port, () => {
  console.log(`Halo Outlook add-in server listening on port ${port}`);
});
