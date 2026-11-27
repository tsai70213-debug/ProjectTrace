"use strict";

const { spawn } = require("node:child_process");
const { appendFile, mkdir, readFile } = require("node:fs/promises");
const { createServer, get } = require("node:http");
const { join } = require("node:path");
const { setTimeout: delay } = require("node:timers/promises");

const port = 8764;
const origin = `http://localhost:${port}`;
const hostname = "localhost:" + port;
const service = "wiring-demo-local-v1";
const logDirectory = join(__dirname, "logs");
const logPath = join(logDirectory, "tracking.log");
const trackedActions = new Set(["login", "operator_switch", "workbook_import", "workbook_import_failed", "review_changed", "report_export"]);
const maxEventBytes = 32 * 1024;
let logWrites = Promise.resolve();
const recordedEvents = new Set();
const assets = new Map([
  ["/", ["index.html", "text/html; charset=utf-8"]],
  ["/index.html", ["index.html", "text/html; charset=utf-8"]],
  ["/styles.css", ["styles.css", "text/css; charset=utf-8"]],
  ["/app.js", ["app.js", "text/javascript; charset=utf-8"]],
  ["/demo-data.js", ["demo-data.js", "text/javascript; charset=utf-8"]],
  ["/vendor/xlsx.full.min.js", ["vendor/xlsx.full.min.js", "text/javascript; charset=utf-8"]],
]);

function appendTracking(event) {
  const writing = logWrites.then(async () => {
    if (event.id && recordedEvents.has(event.id)) return;
    await mkdir(logDirectory, { recursive: true });
    await appendFile(logPath, JSON.stringify({ recordedAt: new Date().toISOString(), ...event }) + "\n", "utf8");
    if (event.id) recordedEvents.add(event.id);
  });
  logWrites = writing.catch(() => {});
  return writing;
}

async function trackRequest(request, response) {
  if (request.method !== "POST") return response.writeHead(405, { Allow: "POST" }).end();
  if (request.headers.origin !== origin) return response.writeHead(403).end();
  if (request.headers["content-type"]?.split(";")[0] !== "application/json") return response.writeHead(415).end();
  if (Number(request.headers["content-length"]) > maxEventBytes) return response.writeHead(413).end();
  let event;
  try {
    const chunks = [];
    let size = 0;
    for await (const chunk of request) {
      size += chunk.length;
      if (size > maxEventBytes) return response.writeHead(413).end();
      chunks.push(chunk);
    }
    event = JSON.parse(Buffer.concat(chunks).toString("utf8"));
  } catch {
    return response.writeHead(400).end();
  }
  if (!event || !trackedActions.has(event.action)
    || typeof event.id !== "string" || !/^[\da-f-]{36}$/i.test(event.id)
    || typeof event.employeeId !== "string" || !event.employeeId.trim() || event.employeeId.length > 40
    || typeof event.time !== "string" || event.time.length !== 24 || !Number.isFinite(Date.parse(event.time))
    || !event.details || typeof event.details !== "object" || Array.isArray(event.details)) {
    return response.writeHead(400).end();
  }
  try {
    const { id, time, action, employeeId, details } = event;
    await appendTracking({ id, time, action, employeeId, details });
    return response.writeHead(204).end();
  } catch (error) {
    console.error(`Could not write ${logPath}: ${error.message}`);
    return response.writeHead(503).end();
  }
}

function probeServer() {
  return new Promise((resolve) => {
    const request = get({
      host: "127.0.0.1", port, path: "/_health", headers: { Host: hostname }, timeout: 1000,
    }, (response) => {
      let body = "";
      response.setEncoding("utf8");
      response.on("data", (part) => { body = (body + part).slice(0, 128); });
      response.on("end", () => resolve(response.statusCode === 200 && body === service ? "ready" : "occupied"));
      response.on("error", () => resolve("occupied"));
    });
    request.on("timeout", () => request.destroy(new Error("Timed out checking port")));
    request.on("error", (error) => resolve(error.code === "ECONNREFUSED" ? "stopped" : "occupied"));
  });
}

async function ensureServer() {
  let state = await probeServer();
  if (state === "stopped") {
    const child = spawn(process.execPath, [__filename], { detached: true, stdio: "ignore", windowsHide: true });
    let launchError;
    child.on("error", (error) => { launchError = error; });
    child.unref();
    for (let attempt = 0; attempt < 30 && state === "stopped"; attempt++) {
      await delay(150);
      if (launchError) throw launchError;
      state = await probeServer();
    }
  }
  if (state === "occupied") throw new Error(`Port ${port} is in use by a different application.`);
  if (state !== "ready") throw new Error(`Could not start the service at ${origin}.`);
  console.log(`Wiring demo: ${origin}`);
}

const server = createServer(async (request, response) => {
  response.setHeader("Cache-Control", "no-store");
  response.setHeader("X-Content-Type-Options", "nosniff");
  if (request.headers.host !== hostname) return response.writeHead(421).end();
  const path = request.url.split("?", 1)[0];
  if (path === "/_tracking") return trackRequest(request, response);
  if (request.method !== "GET" && request.method !== "HEAD") {
    return response.writeHead(405, { Allow: "GET, HEAD" }).end();
  }
  if (path === "/_health") return response.writeHead(200, { "Content-Type": "text/plain" }).end(service);
  const asset = assets.get(path);
  if (!asset) return response.writeHead(404).end();
  try {
    const content = await readFile(join(__dirname, asset[0]));
    response.writeHead(200, { "Content-Type": asset[1], "Content-Length": content.length });
    response.end(request.method === "HEAD" ? undefined : content);
  } catch {
    response.writeHead(500).end();
  }
});

if (process.argv.includes("--start")) {
  ensureServer().catch(async (error) => {
    console.error(error.message);
    process.exitCode = 1;
    await appendTracking({ action: "launch_failed", details: { message: error.message } }).catch(() => {});
  });
} else {
  server.on("error", (error) => {
    console.error(error.message);
    process.exitCode = 1;
  });
  server.listen(port, "127.0.0.1", () => {
    console.log(`Wiring demo: ${origin}`);
    appendTracking({ action: "server_started", details: { origin, pid: process.pid } })
      .catch((error) => console.error(`Could not write ${logPath}: ${error.message}`));
  });
}
