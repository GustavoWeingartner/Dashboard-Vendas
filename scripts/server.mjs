import { createReadStream, existsSync, statSync } from "node:fs";
import { createServer } from "node:http";
import { extname, join, normalize, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = fileURLToPath(new URL(".", import.meta.url));
const root = resolve(__dirname, "..");
const port = Number(process.env.PORT || 4173);
const googleSheetsSpreadsheetId = "1JldFrcw8oaVAWXXhFyJCMVvm_Be9IXju90UAqpzSLXM";
const googleSheetsGid = "27856229";

const mimeTypes = {
  ".html": "text/html; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".mjs": "text/javascript; charset=utf-8",
  ".tsx": "text/plain; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".csv": "text/csv; charset=utf-8",
  ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  ".xls": "application/vnd.ms-excel",
};

function sendFile(response, filePath) {
  const type = mimeTypes[extname(filePath).toLowerCase()] || "application/octet-stream";
  response.writeHead(200, { "Content-Type": type });
  createReadStream(filePath).pipe(response);
}

function googleSheetsCsvUrl() {
  return `https://docs.google.com/spreadsheets/d/${googleSheetsSpreadsheetId}/gviz/tq?tqx=out:csv&gid=${googleSheetsGid}&cacheBust=${Date.now()}`;
}

async function proxyGoogleSheets(response) {
  try {
    const upstream = await fetch(googleSheetsCsvUrl(), { cache: "no-store" });
    if (!upstream.ok) {
      response.writeHead(upstream.status, { "Content-Type": "text/plain; charset=utf-8" });
      response.end(`Google Sheets returned HTTP ${upstream.status}`);
      return;
    }

    const csvText = await upstream.text();
    response.writeHead(200, {
      "Content-Type": "text/csv; charset=utf-8",
      "Cache-Control": "no-store, max-age=0",
    });
    response.end(csvText);
  } catch (error) {
    response.writeHead(502, { "Content-Type": "text/plain; charset=utf-8" });
    response.end(`Unable to fetch Google Sheets CSV: ${error?.message || "unknown error"}`);
  }
}

createServer(async (request, response) => {
  const requestUrl = new URL(request.url || "/", `http://localhost:${port}`);

  if (requestUrl.pathname === "/api/google-sheets") {
    await proxyGoogleSheets(response);
    return;
  }

  const decodedPath = decodeURIComponent(requestUrl.pathname);
  const safePath = normalize(decodedPath).replace(/^(\.\.[/\\])+/, "");
  let filePath = join(root, safePath === "/" ? "index.html" : safePath);

  if (!filePath.startsWith(root)) {
    response.writeHead(403);
    response.end("Forbidden");
    return;
  }

  if (!existsSync(filePath) || statSync(filePath).isDirectory()) {
    filePath = join(root, "index.html");
  }

  sendFile(response, filePath);
}).listen(port, () => {
  console.log(`Dashboard running at http://localhost:${port}`);
});
