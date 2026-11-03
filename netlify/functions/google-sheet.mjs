const MAX_BYTES = 6 * 1024 * 1024;
const SHEET_ID = /\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/;
const GID = /[?&#]gid=(\d+)/;

function sheetExportUrl(raw) {
  let parsed;
  try {
    parsed = new URL(String(raw || "").trim());
  } catch (err) {
    return "";
  }
  if (parsed.protocol !== "https:" || parsed.hostname !== "docs.google.com") return "";
  const match = parsed.pathname.match(SHEET_ID);
  if (!match) return "";
  const gid = String(raw).match(GID);
  let url = `https://docs.google.com/spreadsheets/d/${match[1]}/export?format=xlsx`;
  if (gid) url += `&gid=${gid[1]}`;
  return url;
}

function allowedHost(hostname) {
  const host = String(hostname || "").toLowerCase();
  return host === "docs.google.com" || host.endsWith(".googleusercontent.com");
}

function filenameFromHeader(header) {
  const value = String(header || "");
  const starred = value.match(/filename\*=UTF-8''([^;]+)/i);
  let name = "google-sheet.xlsx";
  if (starred) name = decodeURIComponent(starred[1]);
  else {
    const plain = value.match(/filename="?([^";]+)"?/i);
    if (plain) name = plain[1];
  }
  name = name.replace(/[^\w.\- ()]+/g, "_").replace(/^[\s._]+|[\s._]+$/g, "") || "google-sheet";
  if (!name.toLowerCase().endsWith(".xlsx")) name += ".xlsx";
  return name;
}

async function fetchSheet(startUrl) {
  let current = startUrl;
  for (let hop = 0; hop < 5; hop += 1) {
    const response = await fetch(current, {
      redirect: "manual",
      headers: { "User-Agent": "CleanSheet/1.0" },
    });
    if (response.status >= 300 && response.status < 400) {
      const location = response.headers.get("location");
      if (!location) throw new Error("redirect");
      const next = new URL(location, current);
      if (!allowedHost(next.hostname)) throw new Error("redirect");
      current = next.toString();
      continue;
    }
    return response;
  }
  throw new Error("redirect");
}

function text(status, message) {
  return new Response(message, {
    status,
    headers: { "Content-Type": "text/plain; charset=utf-8", "Cache-Control": "no-store" },
  });
}

export default async (request) => {
  const source = new URL(request.url).searchParams.get("url") || "";
  const exportUrl = sheetExportUrl(source);
  if (!exportUrl) {
    return text(400, "That is not a Google Sheet link. Paste the full docs.google.com link.");
  }
  try {
    const response = await fetchSheet(exportUrl);
    if (!response.ok) {
      return text(403, "The sheet did not open. Set sharing to Anyone with the link, then try again.");
    }
    const bytes = new Uint8Array(await response.arrayBuffer());
    if (bytes.byteLength > MAX_BYTES) {
      return text(413, "The sheet is larger than 6 MB. Try a smaller copy.");
    }
    if (bytes[0] !== 0x50 || bytes[1] !== 0x4b) {
      return text(403, "The sheet did not open. Set sharing to Anyone with the link, then try again.");
    }
    const name = filenameFromHeader(response.headers.get("content-disposition"));
    return new Response(bytes, {
      status: 200,
      headers: {
        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "X-Filename": encodeURIComponent(name),
        "Cache-Control": "no-store",
      },
    });
  } catch (err) {
    return text(502, "Could not connect to Google Sheets. Check your internet connection and try again.");
  }
};
