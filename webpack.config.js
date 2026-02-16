/* eslint-disable no-undef */

const devCerts = require("office-addin-dev-certs");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const webpack = require("webpack");
const dotenv = require("dotenv");

const urlDev = "https://localhost:3000/";
const urlProd = "https://www.contoso.com/";

async function getHttpsOptions() {
  const httpsOptions = await devCerts.getHttpsServerOptions();
  return { ca: httpsOptions.ca, key: httpsOptions.key, cert: httpsOptions.cert };
}

module.exports = async (env, options) => {
  dotenv.config();

  const dev = options.mode === "development";

  return {
    devtool: "source-map",
    experiments: {
      css: true,
    },
    entry: {
      polyfill: ["core-js/stable", "regenerator-runtime/runtime"],
      react: ["react", "react-dom"],
      taskpane: {
        import: ["./src/taskpane/index.tsx", "./src/taskpane/taskpane.html"],
        dependOn: "react",
      },
      dialog: {
        import: ["./src/dialog/index.tsx", "./src/dialog/dialog.html"],
        dependOn: "react",
      },
      commands: "./src/commands/commands.ts",
    },
    output: {
      clean: true,
    },
    resolve: {
      extensions: [".ts", ".tsx", ".html", ".js"],
    },
    module: {
      rules: [
        {
          test: /\.ts$/,
          exclude: /node_modules/,
          use: { loader: "babel-loader" },
        },
        {
          test: /\.tsx?$/,
          exclude: /node_modules/,
          use: ["ts-loader"],
        },
        {
          test: /\.html$/,
          exclude: /node_modules/,
          use: "html-loader",
        },
        {
          test: /\.(png|jpg|jpeg|ttf|woff|woff2|gif|ico)$/,
          type: "asset/resource",
          generator: { filename: "assets/[name][ext][query]" },
        },
        {
          test: /\.css$/,
          type: "css",
        },
      ],
    },
    plugins: [
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["polyfill", "taskpane", "react"],
      }),
      new HtmlWebpackPlugin({
        filename: "dialog.html",
        template: "./src/dialog/dialog.html",
        chunks: ["polyfill", "dialog", "react"],
      }),
      new CopyWebpackPlugin({
        patterns: [
          { from: "assets/*", to: "assets/[name][ext][query]" },
          {
            from: "manifest*.xml",
            to: "[name]" + "[ext]",
            transform(content) {
              if (dev) return content;
              return content.toString().replace(new RegExp(urlDev, "g"), urlProd);
            },
          },
        ],
      }),
      new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["polyfill", "commands"],
      }),
      new webpack.ProvidePlugin({
        Promise: ["es6-promise", "Promise"],
      }),
    ],

    devServer: {
      // HMR breaks inside Outlook webview and triggers:
      // "undefined is not an object (evaluating 'currentUpdateRuntime.push')"
      hot: false,
      liveReload: true,
      client: {
        overlay: false,
      },
    
      headers: {
        "Access-Control-Allow-Origin": "*",
      },
    
      setupMiddlewares: (middlewares, devServer) => {
        if (!devServer || !devServer.app) return middlewares;

        console.log("[webpack-config] Setting up relay middleware");

        // Register relay middleware on Express app (runs before webpack's routing)
        // Relay for SingleCase API to avoid CORS and support dynamic workspace hosts.
        // Frontend calls: /singlecase/{host}/publicapi/v1/...
        // Relay forwards to: https://{host}/publicapi/v1/...
        devServer.app.use("/singlecase", async (req, res, next) => {
          console.log(`[webpack-relay] ${req.method} ${req.url}`);
          try {
            const originalUrl = String(req.originalUrl || req.url || "");
            const m = originalUrl.match(/^\/singlecase\/([^/?#]+)(\/.*)?$/);

            const host = m ? decodeURIComponent(m[1]) : "";
            const restPath = m && m[2] ? m[2] : "";

            if (!host) {
              console.error("[webpack-relay] Missing host in URL:", originalUrl);
              res.status(400).send("Missing host in /singlecase/{host}/...");
              return;
            }
            if (!restPath) {
              console.error("[webpack-relay] Missing path after host:", originalUrl);
              res.status(400).send("Missing path after host.");
              return;
            }

            const upstreamUrl = `https://${host}${restPath}`;
            console.log(`[webpack-relay] Proxying to: ${upstreamUrl}`);

            const headers = {};
            for (const [k, v] of Object.entries(req.headers || {})) {
              if (!v) continue;
              if (k.toLowerCase() === "host") continue;
              headers[k] = v;
            }

            const method = req.method || "GET";

            const body =
              method === "GET" || method === "HEAD"
                ? undefined
                : await new Promise((resolve, reject) => {
                    const chunks = [];
                    req.on("data", (c) => chunks.push(c));
                    req.on("end", () => resolve(Buffer.concat(chunks)));
                    req.on("error", reject);
                  });

            const upstreamRes = await fetch(upstreamUrl, {
              method,
              headers,
              body,
            });

            console.log(`[webpack-relay] Upstream response: ${upstreamRes.status}`);
            res.status(upstreamRes.status);

            upstreamRes.headers.forEach((value, key) => {
              // Do not pass Set-Cookie through dev relay
              if (key.toLowerCase() === "set-cookie") return;
              res.setHeader(key, value);
            });

            const buf = Buffer.from(await upstreamRes.arrayBuffer());
            res.send(buf);
          } catch (e) {
            console.error("[webpack-relay] Error:", e);
            res.status(500).send(String(e));
            next(e);
          }
        });

        // Diagnostic endpoint
        devServer.app.get("/__singlecasecheck", async (req, res) => {
          try {
            const token = req.headers["authentication"];
            if (!token) {
              res.status(400).json({ ok: false, error: "Missing Authentication header" });
              return;
            }

            const hostRaw = String(req.query.host || "");
            const host = hostRaw.replace(/^https?:\/\//i, "").split("/")[0];

            if (!host) {
              res
                .status(400)
                .json({ ok: false, error: "Missing host query param, use ?host=valfor-demo.singlecase.ch" });
              return;
            }

            const r = await fetch(`https://${host}/publicapi/v1/cases`, {
              headers: { Authentication: String(token) },
            });

            const text = await r.text();
            res.status(r.status).json({
              ok: r.ok,
              status: r.status,
              contentType: r.headers.get("content-type"),
              firstBytes: text.slice(0, 200),
            });
          } catch (e) {
            res.status(500).json({ ok: false, error: String(e) });
          }
        });

        return middlewares;
      },
    
      // Important: remove devServer.proxy completely (we are relaying via middleware)
      proxy: undefined,
    
      server: {
        type: "https",
        options: env.WEBPACK_BUILD || options.https !== undefined ? options.https : await getHttpsOptions(),
      },
      port: process.env.npm_package_config_dev_server_port || 3000,
    },
  };
};
