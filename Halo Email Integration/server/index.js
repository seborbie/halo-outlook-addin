const path = require("path");
const express = require("express");
const { registerHaloAuthRoutes } = require("./haloAuth");

const app = express();
const port = Number(process.env.PORT || 3000);
const distPath = path.join(__dirname, "..", "dist");

app.disable("x-powered-by");
app.use(express.json({ limit: "2mb" }));

registerHaloAuthRoutes(app);

app.use(express.static(distPath));
app.use("/public", express.static(distPath));

app.get(/.*/, (req, res) => {
  res.sendFile(path.join(distPath, "taskpane.html"));
});

app.listen(port, () => {
  console.log(`Halo Outlook add-in server listening on port ${port}`);
});
