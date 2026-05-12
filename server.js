const http = require("node:http");
const fsp = require("node:fs/promises");
const path = require("node:path");
const { spawn } = require("node:child_process");

const PORT = 8000;
const REPO_ROOT = __dirname;
const ALLOWED_ORIGIN = "http://localhost:3000";
const ASSET_ROOT = path.join(REPO_ROOT, ".openxml-assets");

const filesByNumber = new Map([
  [1, "Supermarket-Sales-Sample-Data.xlsx"],
  [2, "Project-Management-Sample-Data.xlsx"],
  [3, "IC-Budget-Dashboard-Template-Example.xlsx"],
  [4, "FSI-2023-DOWNLOAD.xlsx"],
  [5, "fsi-2021.xlsx"],
  [6, "Formula Excel Template.xlsx"],
  [7, "ONLYOFFICE Spreadsheet Sample (1).xlsx"],
  [8, "IC-Data-Quality-Dashboard-Template-Example.xlsx"],
]);

const contentTypesByExtension = new Map([
  [".bmp", "image/bmp"],
  [".bin", "application/octet-stream"],
  [".emf", "application/octet-stream"],
  [".gif", "image/gif"],
  [".jpg", "image/jpeg"],
  [".jpeg", "image/jpeg"],
  [".json", "application/json"],
  [".png", "image/png"],
  [".svg", "image/svg+xml"],
  [".tif", "image/tiff"],
  [".tiff", "image/tiff"],
  [".wmf", "application/octet-stream"],
  [".xml", "application/xml"],
]);

function sendJson(response, statusCode, payload) {
  response.writeHead(statusCode, {
    "Access-Control-Allow-Origin": ALLOWED_ORIGIN,
    "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type",
    "Content-Type": "application/json",
  });
  response.end(JSON.stringify(payload, null, 2));
}

function sendRaw(response, statusCode, contentType, content) {
  response.writeHead(statusCode, {
    "Access-Control-Allow-Origin": ALLOWED_ORIGIN,
    "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type",
    "Content-Type": contentType,
  });
  response.end(content);
}

function validFilesPayload() {
  return Object.fromEntries(
    [...filesByNumber.entries()].map(([number, fileName]) => [
      number,
      path.join("demo-files", fileName),
    ]),
  );
}

function parseRequestedNumber(body) {
  const trimmedBody = body.trim();

  if (/^\d+$/.test(trimmedBody)) {
    return Number(trimmedBody);
  }

  const parsed = JSON.parse(trimmedBody || "{}");
  const value = parsed.number ?? parsed.id ?? parsed.fileNumber;
  return typeof value === "string" ? Number(value) : value;
}

function readRequestBody(request) {
  return new Promise((resolve, reject) => {
    let body = "";

    request.on("data", (chunk) => {
      body += chunk;

      if (body.length > 1024) {
        request.destroy();
        reject(new Error("Request body is too large."));
      }
    });

    request.on("end", () => resolve(body));
    request.on("error", reject);
  });
}

function assetUrlForPath(requestedNumber, assetPath) {
  return `http://localhost:${PORT}/xlsx-assets/${requestedNumber}/${assetPath
    .split("/")
    .map(encodeURIComponent)
    .join("/")}`;
}

function attachAssetUrls(workbook, requestedNumber) {
  for (const sheet of workbook.workbook?.sheets || []) {
    for (const drawing of sheet.drawings || []) {
      for (const chart of drawing.charts || []) {
        stripChartAssetFields(chart);
      }

      for (const image of drawing.images || []) {
        if (image.assetPath) {
          image.assetUrl = assetUrlForPath(requestedNumber, image.assetPath);
        }
      }
    }

    for (const chart of sheet.charts || []) {
      stripChartAssetFields(chart);
    }

    for (const image of sheet.images || []) {
      if (image.assetPath) {
        image.assetUrl = assetUrlForPath(requestedNumber, image.assetPath);
      }
    }
  }

  return workbook;
}

function stripChartAssetFields(chart) {
  delete chart.assetPath;
  delete chart.assetUrl;
  delete chart.renderedAssetPath;
  delete chart.renderedAssetUrl;
}

async function prepareAssetDirectory(requestedNumber) {
  const outputDirectory = path.join(ASSET_ROOT, String(requestedNumber));

  await fsp.rm(outputDirectory, {
    force: true,
    recursive: true,
  });
  await fsp.mkdir(outputDirectory, {
    recursive: true,
  });

  return outputDirectory;
}

function runCliForFile(filePath, outputDirectory) {
  return new Promise((resolve, reject) => {
    const child = spawn(
      "dotnet",
      [
        "run",
        "--project",
        "src/OpenXml.Cli",
        "--",
        "xlsx-assets",
        filePath,
        outputDirectory,
      ],
      {
        cwd: REPO_ROOT,
        stdio: ["ignore", "pipe", "pipe"],
      },
    );

    let stdout = "";
    let stderr = "";

    child.stdout.on("data", (chunk) => {
      stdout += chunk;
    });

    child.stderr.on("data", (chunk) => {
      stderr += chunk;
    });

    child.on("error", reject);
    child.on("close", (code) => {
      resolve({ code, stdout, stderr });
    });
  });
}

async function readWorkbookJson(outputDirectory, requestedNumber) {
  const workbookJsonPath = path.join(outputDirectory, "workbook.json");
  const workbookJson = await fsp.readFile(workbookJsonPath, "utf8");
  const workbook = JSON.parse(workbookJson);

  return JSON.stringify(attachAssetUrls(workbook, requestedNumber), null, 2);
}

function safeResolveAssetPath(requestedNumber, relativePath) {
  const baseDirectory = path.join(ASSET_ROOT, String(requestedNumber));
  const resolvedPath = path.resolve(baseDirectory, relativePath);
  const normalizedBase = path.resolve(baseDirectory);

  if (resolvedPath !== normalizedBase && !resolvedPath.startsWith(`${normalizedBase}${path.sep}`)) {
    return null;
  }

  return resolvedPath;
}

async function handleAssetGet(response, requestedNumber, relativePath) {
  if (!Number.isInteger(requestedNumber) || !filesByNumber.has(requestedNumber)) {
    sendJson(response, 404, {
      error: "Unknown asset file number.",
    });
    return;
  }

  const assetPath = safeResolveAssetPath(requestedNumber, relativePath);

  if (!assetPath) {
    sendJson(response, 400, {
      error: "Invalid asset path.",
    });
    return;
  }

  try {
    const content = await fsp.readFile(assetPath);
    const contentType =
      contentTypesByExtension.get(path.extname(assetPath).toLowerCase()) ||
      "application/octet-stream";

    sendRaw(response, 200, contentType, content);
  } catch (error) {
    sendJson(response, 404, {
      error: "Asset not found.",
      details: error.message,
    });
  }
}

async function handlePost(request, response, numberFromPath) {
  let requestedNumber;

  if (numberFromPath !== undefined) {
    requestedNumber = Number(numberFromPath);
  } else {
    try {
      const body = await readRequestBody(request);
      requestedNumber = parseRequestedNumber(body);
    } catch (error) {
      sendJson(response, 400, {
        error: "Expected a raw number or JSON body like { \"number\": 1 }.",
        details: error.message,
        files: validFilesPayload(),
      });
      return;
    }
  }

  if (!Number.isInteger(requestedNumber) || !filesByNumber.has(requestedNumber)) {
    sendJson(response, 400, {
      error: "Unknown file number.",
      requestedNumber,
      files: validFilesPayload(),
    });
    return;
  }

  const filePath = path.join(
    REPO_ROOT,
    "demo-files",
    filesByNumber.get(requestedNumber),
  );

  try {
    const outputDirectory = await prepareAssetDirectory(requestedNumber);
    const result = await runCliForFile(filePath, outputDirectory);

    if (result.code !== 0) {
      sendJson(response, 500, {
        error: "OpenXml.Cli failed.",
        exitCode: result.code,
        stderr: result.stderr.trim(),
        stdout: result.stdout.trim(),
      });
      return;
    }

    const workbookJson = await readWorkbookJson(outputDirectory, requestedNumber);
    sendRaw(response, 200, "application/json", workbookJson);
  } catch (error) {
    sendJson(response, 500, {
      error: "Unable to run OpenXml.Cli.",
      details: error.message,
    });
  }
}

const server = http.createServer((request, response) => {
  const url = new URL(request.url, `http://${request.headers.host}`);

  if (request.method === "OPTIONS") {
    response.writeHead(204, {
      "Access-Control-Allow-Origin": ALLOWED_ORIGIN,
      "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
      "Access-Control-Allow-Headers": "Content-Type",
    });
    response.end();
    return;
  }

  if (request.method === "GET" && url.pathname === "/") {
    sendJson(response, 200, {
      usage:
        "POST /xlsx/1, or POST /xlsx with a raw number or JSON body like { \"number\": 1 }.",
      assets:
        "Workbook requests extract assets automatically. Asset files are served from /xlsx-assets/<number>/assets/...",
      files: validFilesPayload(),
    });
    return;
  }

  const assetMatch = url.pathname.match(/^\/xlsx-assets\/(\d+)\/(.+)$/);

  if (request.method === "GET" && assetMatch) {
    handleAssetGet(response, Number(assetMatch[1]), decodeURIComponent(assetMatch[2]));
    return;
  }

  const pathMatch = url.pathname.match(/^\/xlsx(?:\/(\d+))?$/);

  if (request.method === "POST" && pathMatch) {
    handlePost(request, response, pathMatch[1]);
    return;
  }

  sendJson(response, 404, {
    error: "Not found.",
    usage:
      "POST /xlsx/1, or POST /xlsx with a raw number or JSON body like { \"number\": 1 }.",
  });
});

server.listen(PORT, () => {
  console.log(`OpenXML demo server listening at http://localhost:${PORT}`);
});
