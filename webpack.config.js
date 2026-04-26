const path = require("path");
const fs = require("fs");
const crypto = require("crypto");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyPlugin = require("copy-webpack-plugin");
const devCerts = require("office-addin-dev-certs");
const http = require("http");

const ARTIFACT_ROOT = path.resolve(__dirname, "debug-artifacts");

function safePathPart(value) {
  return String(value ?? "")
    .replace(/[^a-zA-Z0-9._-]/g, "_")
    .replace(/^_+|_+$/g, "")
    .slice(0, 120) || "artifact";
}

function uniqueArtifactName(name) {
  const parsed = path.parse(safePathPart(name));
  const ext = parsed.ext || ".bin";
  const base = parsed.name || "artifact";
  return `${base}-${crypto.randomBytes(3).toString("hex")}${ext}`;
}

module.exports = async (env, argv) => {
  const isDev = argv.mode !== "production";
  const httpsOptions = isDev ? await devCerts.getHttpsServerOptions() : undefined;

  const handleLogRequest = (req, res) => {
    res.setHeader("Access-Control-Allow-Origin", "*");
    res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
    res.setHeader("Access-Control-Allow-Headers", "Content-Type");

    if (req.method === "OPTIONS") {
      res.writeHead(204).end();
      return;
    }

    if (req.url === "/__terminal-log" && req.method === "POST") {
      let body = "";
      req.on("data", (chunk) => (body += chunk));
      req.on("end", () => {
        try {
          const { level, msg } = JSON.parse(body);
          const prefix = level === "error" ? "❌" : level === "warn" ? "⚠️" : "📋";
          console.log(`${prefix} [Browser] ${msg}`);
        } catch {
          console.log(`📋 [Browser] ${body}`);
        }
        res.end("ok");
      });
    } else if (req.url === "/__debug-artifact" && req.method === "POST") {
      let body = "";
      req.on("data", (chunk) => (body += chunk));
      req.on("end", () => {
        try {
          const payload = JSON.parse(body);
          const kind = safePathPart(payload.kind ?? "misc");
          const dir = path.join(ARTIFACT_ROOT, kind);
          fs.mkdirSync(dir, { recursive: true });

          const filename = uniqueArtifactName(payload.filename ?? "artifact.bin");
          const filePath = path.join(dir, filename);
          const metaPath = filePath.replace(/\.[^.]+$/, ".json");
          fs.writeFileSync(filePath, Buffer.from(String(payload.base64 ?? ""), "base64"));
          fs.writeFileSync(metaPath, JSON.stringify(payload.metadata ?? {}, null, 2));
          console.log(`📸 [Browser] saved artifact ${filePath}`);
          res.setHeader("Content-Type", "application/json");
          res.end(JSON.stringify({ ok: true, path: filePath, metadataPath: metaPath }));
        } catch (err) {
          const message = err instanceof Error ? err.message : String(err);
          console.log(`❌ [Browser] save artifact failed: ${message}`);
          res.writeHead(500, { "Content-Type": "application/json" });
          res.end(JSON.stringify({ ok: false, error: message }));
        }
      });
    } else {
      res.writeHead(404).end();
    }
  };
  const logHttpsServer = isDev
    ? require("https").createServer(httpsOptions, handleLogRequest)
    : http.createServer(handleLogRequest);
  await new Promise((resolve) => logHttpsServer.listen(3001, resolve));
  console.log(`📋 Terminal log server running on ${isDev ? "https" : "http"}://localhost:3001`);

  return {
    mode: isDev ? "development" : "production",
    devtool: isDev ? "inline-source-map" : false,
    entry: {
      taskpane: "./src/taskpane/taskpane.ts"
    },
    output: {
      path: path.resolve(__dirname, "dist"),
      filename: "[name].js",
      clean: true
    },
    resolve: {
      extensions: [".ts", ".js"]
    },
    module: {
      rules: [
        { test: /\.ts$/, use: "ts-loader", exclude: /node_modules/ },
        { test: /\.css$/, use: ["style-loader", "css-loader"] }
      ]
    },
    plugins: [
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["taskpane"],
        inject: "body",
        scriptLoading: "blocking"
      }),
      new CopyPlugin({
        patterns: [
          { from: "assets", to: "assets", noErrorOnMissing: true },
          { from: "config", to: "config", noErrorOnMissing: true },
          { from: "manifest.xml", to: "manifest.xml", noErrorOnMissing: true }
        ]
      })
    ],
    devServer: {
      static: [
        { directory: path.resolve(__dirname, "dist") },
        { directory: path.resolve(__dirname, "config"), publicPath: "/config" }
      ],
      hot: true,
      port: 3000,
      server: httpsOptions
        ? { type: "https", options: httpsOptions }
        : "https",
      headers: {
        "Access-Control-Allow-Origin": "*"
      }
    }
  };
};
