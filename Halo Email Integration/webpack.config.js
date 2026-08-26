/* eslint-disable no-undef */

const devCerts = require("office-addin-dev-certs");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const fs = require("fs");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const path = require("path");
const webpack = require("webpack");
const packageInfo = require("./package.json");

require("dotenv").config({ path: path.resolve(__dirname, ".env") });

const { registerHaloAuthRoutes } = require("./server/haloAuth");
const { createHaloStore } = require("./server/haloStore");
const { registerStatusRoute } = require("./server/statusRoute");

const urlDevOrigin = "https://localhost:3000";
const urlDev = `${urlDevOrigin}/`;
const developmentManifestId = "e3c74ceb-c7d1-4264-a732-749a0d34c412";
const developmentManifestVersion = `${packageInfo.version}.2`;
const developmentRuntimeCacheToken = `${developmentManifestVersion}-references-1`;
const productionRuntimeCacheToken = `${packageInfo.version}.1`;

function createDevelopmentManifest(content) {
  const source = content.toString();
  const manifest = source
    .replace(/<Id>[^<]+<\/Id>/, `<Id>${developmentManifestId}</Id>`)
    .replace(/<Version>[^<]+<\/Version>/, `<Version>${developmentManifestVersion}</Version>`)
    .replace(
      /<DisplayName DefaultValue="[^"]+"\/>/,
      '<DisplayName DefaultValue="LOCAL DIAGNOSTICS - HaloPSA Outlook Add-in"/>'
    )
    .replace(
      /<Description DefaultValue="[^"]+"\/>/,
      '<Description DefaultValue="Local HaloPSA Outlook add-in with Smart Alerts diagnostics."/>'
    )
    .replace(/\?v=[^"<]+/g, `?v=${developmentRuntimeCacheToken}`);

  if (!manifest.includes(`<Id>${developmentManifestId}</Id>`)) {
    throw new Error("The local diagnostics manifest ID could not be generated.");
  }
  return manifest;
}

function getProductionBaseUrl(value) {
  if (!value) {
    throw new Error(
      "PUBLIC_BASE_URL must be set to the deployed HTTPS origin for production builds."
    );
  }

  let url;
  try {
    url = new URL(value);
  } catch {
    throw new Error("PUBLIC_BASE_URL must be a valid absolute HTTPS URL.");
  }

  if (url.protocol !== "https:") {
    throw new Error("PUBLIC_BASE_URL must use https://.");
  }

  if (url.username || url.password || url.pathname !== "/" || url.search || url.hash) {
    throw new Error(
      "PUBLIC_BASE_URL must contain only the HTTPS origin, without credentials or a path."
    );
  }

  return `${url.origin}/`;
}

async function getHttpsOptions() {
  const httpsOptions = await devCerts.getHttpsServerOptions();
  return { ca: httpsOptions.ca, key: httpsOptions.key, cert: httpsOptions.cert };
}

module.exports = async (env, options) => {
  const dev = options.mode === "development";
  const serving = Boolean(env && env.WEBPACK_SERVE);
  const store = serving ? createHaloStore() : null;
  if (store) {
    await store.initialize();
  }
  const publicBaseUrl = dev ? urlDev : getProductionBaseUrl(process.env.PUBLIC_BASE_URL);
  const publicOrigin = publicBaseUrl.replace(/\/+$/, "");
  const config = {
    devtool: "source-map",
    entry: {
      bugreport: "./src/bugreport/bugreport.ts",
      polyfill: ["core-js/stable", "regenerator-runtime/runtime"],
      taskpane: ["./src/taskpane/taskpane.ts", "./src/taskpane/taskpane.html"],
    },
    output: {
      clean: true,
    },
    resolve: {
      extensions: [".ts", ".html", ".js"],
    },
    module: {
      rules: [
        {
          test: /\.ts$/,
          exclude: /node_modules/,
          use: {
            loader: "babel-loader",
          },
        },
        {
          test: /\.html$/,
          exclude: /node_modules/,
          use: "html-loader",
        },
        {
          test: /\.(png|jpg|jpeg|gif|ico)$/,
          type: "asset/resource",
          generator: {
            filename: "assets/[name][ext][query]",
          },
        },
      ],
    },
    plugins: [
      new webpack.DefinePlugin({
        __HALO_ADD_IN_VERSION__: JSON.stringify(packageInfo.version),
      }),
      new HtmlWebpackPlugin({
        filename: "bugreport.html",
        template: "./src/bugreport/bugreport.html",
        chunks: ["bugreport"],
      }),
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["polyfill", "taskpane"],
      }),
      new CopyWebpackPlugin({
        patterns: [
          {
            from: "assets/*",
            to: "assets/[name][ext][query]",
          },
          {
            from: "src/commands/classic-send-runtime.js",
            to: "classic-send-runtime.js",
            info: { minimized: true },
            transform(content) {
              return content
                .toString()
                .replace(new RegExp("__HALO_PUBLIC_BASE_URL__", "g"), publicOrigin);
            },
          },
          {
            from: "manifest*.xml",
            to: "[name]" + "[ext]",
            transform(content) {
              if (dev) {
                return content;
              } else {
                return content.toString().split(urlDevOrigin).join(publicOrigin);
              }
            },
          },
          ...(dev
            ? [
                {
                  from: "manifest.xml",
                  to: "manifest.debug.xml",
                  transform(content) {
                    return createDevelopmentManifest(content);
                  },
                },
              ]
            : []),
        ],
      }),
      new HtmlWebpackPlugin({
        filename: "commands.html",
        inject: false,
        templateContent: () =>
          fs
            .readFileSync(path.resolve(__dirname, "src", "commands", "commands.html"), "utf8")
            .replace(
              /__HALO_RUNTIME_CACHE_TOKEN__/g,
              dev ? developmentRuntimeCacheToken : productionRuntimeCacheToken
            ),
      }),
    ],
    devServer: {
      headers: {
        "Access-Control-Allow-Origin": "*",
        "Cache-Control": "no-store, no-cache, must-revalidate",
        Expires: "0",
        Pragma: "no-cache",
      },
      static: {
        directory: path.join(__dirname, "dist"),
        publicPath: "/public",
      },
      historyApiFallback: {
        rewrites: [{ from: /^\/bugreport\/?$/, to: "/bugreport.html" }],
      },
      setupMiddlewares: (middlewares, devServer) => {
        if (!devServer || !devServer.app) {
          throw new Error("webpack-dev-server is not available for Halo auth routes.");
        }

        registerHaloAuthRoutes(devServer.app, { store });
        registerStatusRoute(devServer.app, { checkReady: () => store.isReady() });
        console.info(
          "[halo-dev] Local server ready. Sideload dist/manifest.debug.xml with npm start; " +
            "running npm run dev-server alone does not update Outlook's add-in registration."
        );
        return middlewares;
      },
      server: {
        type: "https",
        options:
          env.WEBPACK_BUILD || options.https !== undefined
            ? options.https
            : await getHttpsOptions(),
      },
      port: 3000,
    },
  };

  return config;
};

module.exports._test = {
  createDevelopmentManifest,
  developmentManifestId,
  developmentManifestVersion,
  developmentRuntimeCacheToken,
};
