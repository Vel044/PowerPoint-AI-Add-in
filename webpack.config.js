const path = require("path");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyPlugin = require("copy-webpack-plugin");
const devCerts = require("office-addin-dev-certs");
const http = require("http");

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
