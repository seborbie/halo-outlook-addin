/* eslint-disable no-undef */

const devCerts = require("office-addin-dev-certs");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const path = require("path");

require("dotenv").config({ path: path.resolve(__dirname, ".env") });

const { registerHaloAuthRoutes } = require("./server/haloAuth");

const urlDevOrigin = "https://localhost:3000";
const urlDev = `${urlDevOrigin}/`;

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
  const publicBaseUrl = dev ? urlDev : getProductionBaseUrl(process.env.PUBLIC_BASE_URL);
  const publicOrigin = publicBaseUrl.replace(/\/+$/, "");
  const config = {
    devtool: "source-map",
    entry: {
      polyfill: ["core-js/stable", "regenerator-runtime/runtime"],
      taskpane: ["./src/taskpane/taskpane.ts", "./src/taskpane/taskpane.html"],
    },
    output: {
      clean: false,
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
        ],
      }),
      new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        inject: false,
      }),
    ],
    devServer: {
      headers: {
        "Access-Control-Allow-Origin": "*",
      },
      static: {
        directory: path.join(__dirname, "dist"),
        publicPath: "/public",
      },
      setupMiddlewares: (middlewares, devServer) => {
        if (!devServer || !devServer.app) {
          throw new Error("webpack-dev-server is not available for Halo auth routes.");
        }

        registerHaloAuthRoutes(devServer.app);
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
