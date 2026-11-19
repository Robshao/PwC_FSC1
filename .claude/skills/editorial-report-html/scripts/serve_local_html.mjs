#!/usr/bin/env node

import { createReadStream } from "node:fs";
import { realpath, stat } from "node:fs/promises";
import { createServer } from "node:http";
import { dirname, extname, resolve } from "node:path";
import process from "node:process";

function usage() {
  return [
    "Usage: node scripts/serve_local_html.mjs <file.html> [options]",
    "",
    "Options:",
    "  --port <number>        Loopback port; 0 selects a free port (default: 0)",
    "  --ttl-seconds <number> Stop automatically after this many seconds",
    "                         (default: 1800; 0 disables the TTL)",
    "  --help                 Show this help",
  ].join("\n");
}

function fail(message) {
  process.stderr.write(`serve_local_html: ${message}\n`);
  process.exitCode = 1;
}

function parseInteger(value, label, min, max) {
  if (!/^\d+$/.test(value ?? "")) {
    throw new Error(`${label} must be an integer`);
  }
  const parsed = Number(value);
  if (!Number.isSafeInteger(parsed) || parsed < min || parsed > max) {
    throw new Error(`${label} must be between ${min} and ${max}`);
  }
  return parsed;
}

function parseArgs(argv) {
  let htmlPath;
  let port = 0;
  let ttlSeconds = 1800;

  for (let index = 0; index < argv.length; index += 1) {
    const arg = argv[index];
    if (arg === "--help") {
      return { help: true };
    }
    if (arg === "--port") {
      port = parseInteger(argv[++index], "--port", 0, 65535);
      continue;
    }
    if (arg === "--ttl-seconds") {
      ttlSeconds = parseInteger(argv[++index], "--ttl-seconds", 0, 86400);
      continue;
    }
    if (arg.startsWith("-")) {
      throw new Error(`unknown option: ${arg}`);
    }
    if (htmlPath) {
      throw new Error("provide exactly one HTML file");
    }
    htmlPath = arg;
  }

  if (!htmlPath) {
    throw new Error("missing HTML file");
  }
  return { help: false, htmlPath, port, ttlSeconds };
}

async function main() {
  let options;
  try {
    options = parseArgs(process.argv.slice(2));
  } catch (error) {
    fail(error.message);
    process.stderr.write(`${usage()}\n`);
    return;
  }

  if (options.help) {
    process.stdout.write(`${usage()}\n`);
    return;
  }

  let targetPath;
  try {
    targetPath = await realpath(resolve(options.htmlPath));
    const targetStat = await stat(targetPath);
    if (!targetStat.isFile() || ![".html", ".htm"].includes(extname(targetPath).toLowerCase())) {
      throw new Error("target must be an existing .html or .htm file");
    }
  } catch (error) {
    fail(error.message);
    return;
  }

  const rootPath = dirname(targetPath);
  const targetRoute = `/${encodeURIComponent(targetPath.slice(rootPath.length + 1))}`;

  const server = createServer(async (request, response) => {
    response.setHeader("Cache-Control", "no-store");
    response.setHeader("X-Content-Type-Options", "nosniff");

    if (!["GET", "HEAD"].includes(request.method ?? "")) {
      response.writeHead(405, { Allow: "GET, HEAD" });
      response.end("Method Not Allowed\n");
      return;
    }

    let pathname;
    try {
      pathname = decodeURIComponent(new URL(request.url ?? "/", "http://127.0.0.1").pathname);
    } catch {
      response.writeHead(400);
      response.end("Bad Request\n");
      return;
    }

    if (pathname !== "/" && pathname !== decodeURIComponent(targetRoute)) {
      response.writeHead(404);
      response.end("Not Found\n");
      return;
    }

    try {
      const fileStat = await stat(targetPath);
      response.setHeader("Content-Length", fileStat.size);
      response.setHeader("Content-Type", "text/html; charset=utf-8");
      response.writeHead(200);
      if (request.method === "HEAD") {
        response.end();
        return;
      }
      const stream = createReadStream(targetPath);
      stream.on("error", () => response.destroy());
      stream.pipe(response);
    } catch {
      response.writeHead(500);
      response.end("Internal Server Error\n");
    }
  });

  let ttl;
  const stop = (reason) => {
    if (ttl) {
      clearTimeout(ttl);
    }
    server.close(() => {
      if (reason) {
        process.stderr.write(`serve_local_html: stopped (${reason})\n`);
      }
    });
  };

  server.on("error", (error) => {
    fail(error.message);
  });
  process.once("SIGINT", () => stop("SIGINT"));
  process.once("SIGTERM", () => stop("SIGTERM"));

  server.listen(options.port, "127.0.0.1", () => {
    const address = server.address();
    if (!address || typeof address === "string") {
      fail("could not determine the loopback address");
      stop();
      return;
    }

    const url = `http://127.0.0.1:${address.port}${targetRoute}`;
    process.stdout.write(
      `${JSON.stringify({
        url,
        host: "127.0.0.1",
        port: address.port,
        root: rootPath,
        ttlSeconds: options.ttlSeconds,
      })}\n`,
    );

    if (options.ttlSeconds > 0) {
      ttl = setTimeout(() => stop("TTL expired"), options.ttlSeconds * 1000);
      ttl.unref();
    }
  });
}

await main();
