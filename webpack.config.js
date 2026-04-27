const path = require("path");
const fs = require("fs");
const crypto = require("crypto");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyPlugin = require("copy-webpack-plugin");
const devCerts = require("office-addin-dev-certs");

const ARTIFACT_ROOT = path.resolve(__dirname, "debug-artifacts");
const SKILLS_ROOT = path.resolve(__dirname, "skills");

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

function parseJsonBody(req) {
  return new Promise((resolve, reject) => {
    let body = "";
    req.on("data", (chunk) => (body += chunk));
    req.on("end", () => {
      try {
        resolve(body ? JSON.parse(body) : {});
      } catch (err) {
        reject(err);
      }
    });
    req.on("error", reject);
  });
}

function sendJson(res, status, payload) {
  res.writeHead(status, { "Content-Type": "application/json" });
  res.end(JSON.stringify(payload));
}

function safeSkillName(value) {
  const name = String(value ?? "").trim();
  if (!/^[a-z0-9][a-z0-9-]{0,63}$/i.test(name)) {
    throw new Error("skill name 只能包含字母、数字、连字符，且最长 64 字符");
  }
  return name.toLowerCase();
}

function flattenDuckDuckGoTopics(topics, out = []) {
  for (const item of topics || []) {
    if (item.Topics) {
      flattenDuckDuckGoTopics(item.Topics, out);
    } else if (item.Text || item.FirstURL) {
      out.push({
        title: item.Text ? String(item.Text).split(" - ")[0] : String(item.FirstURL),
        url: item.FirstURL || "",
        snippet: item.Text || ""
      });
    }
  }
  return out;
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
      parseJsonBody(req).then((payload) => {
        try {
          const { level, msg } = payload;
          const prefix = level === "error" ? "❌" : level === "warn" ? "⚠️" : "📋";
          console.log(`${prefix} [Browser] ${msg}`);
        } catch {
          console.log("📋 [Browser] malformed terminal-log payload");
        }
        res.end("ok");
      }).catch((err) => {
        console.log(`📋 [Browser] ${err instanceof Error ? err.message : String(err)}`);
        res.end("ok");
      });
    } else if (req.url === "/__debug-artifact" && req.method === "POST") {
      parseJsonBody(req).then((payload) => {
        try {
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
      }).catch((err) => sendJson(res, 400, { ok: false, error: err instanceof Error ? err.message : String(err) }));
    } else if (req.url === "/__web-search" && req.method === "POST") {
      parseJsonBody(req).then(async (payload) => {
        const query = String(payload.query ?? "").trim();
        const top = Math.max(1, Math.min(10, Number(payload.top ?? 5) || 5));
        if (!query) {
          sendJson(res, 200, { ok: false, error: "query is required", results: [] });
          return;
        }
        try {
          const url = `https://api.duckduckgo.com/?q=${encodeURIComponent(query)}&format=json&no_redirect=1&no_html=1`;
          const searchRes = await fetch(url);
          if (!searchRes.ok) {
            sendJson(res, 200, { ok: false, error: `search endpoint HTTP ${searchRes.status}`, results: [] });
            return;
          }
          const data = await searchRes.json();
          const results = [];
          if (data.AbstractText || data.AbstractURL) {
            results.push({
              title: data.Heading || query,
              url: data.AbstractURL || "",
              snippet: data.AbstractText || ""
            });
          }
          results.push(...flattenDuckDuckGoTopics(data.RelatedTopics));
          sendJson(res, 200, { ok: true, query, results: results.slice(0, top) });
        } catch (err) {
          sendJson(res, 200, { ok: false, error: err instanceof Error ? err.message : String(err), results: [] });
        }
      }).catch((err) => sendJson(res, 400, { ok: false, error: err instanceof Error ? err.message : String(err) }));
    } else if (req.url === "/__skill-file" && req.method === "POST") {
      parseJsonBody(req).then((payload) => {
        try {
          const action = String(payload.action ?? "");
          const name = safeSkillName(payload.name);
          const filePath = path.join(SKILLS_ROOT, `${name}.md`);
          if (!filePath.startsWith(SKILLS_ROOT)) throw new Error("invalid skill path");
          if (action === "read") {
            if (!fs.existsSync(filePath)) {
              sendJson(res, 404, { ok: false, error: `skill not found: ${name}` });
              return;
            }
            sendJson(res, 200, { ok: true, name, path: filePath, content: fs.readFileSync(filePath, "utf8") });
            return;
          }
          if (action === "create") {
            fs.mkdirSync(SKILLS_ROOT, { recursive: true });
            const description = String(payload.description ?? "").trim();
            const instructions = String(payload.instructions ?? "").trim();
            const content = `---\nname: ${name}\ndescription: ${description}\n---\n\n${instructions}\n`;
            fs.writeFileSync(filePath, content);
            sendJson(res, 200, { ok: true, name, path: filePath });
            return;
          }
          sendJson(res, 400, { ok: false, error: `unsupported action: ${action}` });
        } catch (err) {
          sendJson(res, 400, { ok: false, error: err instanceof Error ? err.message : String(err) });
        }
      }).catch((err) => sendJson(res, 400, { ok: false, error: err instanceof Error ? err.message : String(err) }));
    } else {
      res.writeHead(404).end();
    }
  };
  if (isDev) {
    const logHttpsServer = require("https").createServer(httpsOptions, handleLogRequest);
    await new Promise((resolve) => logHttpsServer.listen(3001, resolve));
    console.log("📋 Terminal log server running on https://localhost:3001");
  }

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
