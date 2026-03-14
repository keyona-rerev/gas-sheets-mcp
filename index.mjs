import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { SSEServerTransport } from "@modelcontextprotocol/sdk/server/sse.js";
import { z } from "zod";
import fetch from "node-fetch";
import http from "http";

const GAS_URL   = process.env.SHEETS_GAS_URL;
const GAS_TOKEN = process.env.SHEETS_GAS_TOKEN;
const PORT      = process.env.PORT || 3000;

if (!GAS_URL || !GAS_TOKEN) {
  console.error("FATAL: SHEETS_GAS_URL and SHEETS_GAS_TOKEN must be set");
  process.exit(1);
}

async function callGAS(action, params = {}) {
  const res = await fetch(GAS_URL, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ token: GAS_TOKEN, action, ...params }),
  });
  const json = await res.json();
  if (!json.success) throw new Error(json.error || "GAS error");
  return json.data;
}

const transports = new Map();

function buildServer() {
  const server = new McpServer({ name: "google-sheets", version: "1.0.0" });

  server.tool(
    "sheets_list_spreadsheets",
    "Search your Google Drive for spreadsheets. Returns id, name, and url. Optionally filter by name.",
    { query: z.string().optional().describe("Optional name filter") },
    async (p) => {
      const data = await callGAS("listSpreadsheets", p);
      return { content: [{ type: "text", text: JSON.stringify(data, null, 2) }] };
    }
  );

  server.tool(
    "sheets_create_spreadsheet",
    "Create a new Google Spreadsheet. Returns the spreadsheetId, name, and url.",
    { name: z.string().describe("Name of the new spreadsheet") },
    async (p) => {
      const data = await callGAS("createSpreadsheet", p);
      return { content: [{ type: "text", text: JSON.stringify(data, null, 2) }] };
    }
  );

  server.tool(
    "sheets_list_sheets",
    "List all sheets (tabs) in a spreadsheet. Returns name, index, rowCount, colCount.",
    { spreadsheetId: z.string() },
    async (p) => {
      const data = await callGAS("listSheets", p);
      return { content: [{ type: "text", text: JSON.stringify(data, null, 2) }] };
    }
  );

  server.tool(
    "sheets_create_sheet",
    "Create a new sheet (tab) in a spreadsheet. Optionally set column headers.",
    {
      spreadsheetId: z.string(),
      sheetName: z.string(),
      headers: z.array(z.string()).optional().describe("Optional column headers for row 1"),
    },
    async (p) => {
      const data = await callGAS("createSheet", p);
      return { content: [{ type: "text", text: JSON.stringify(data, null, 2) }] };
    }
  );

  server.tool(
    "sheets_delete_sheet",
    "Delete a sheet (tab) from a spreadsheet. Permanent — cannot be undone.",
    {
      spreadsheetId: z.string(),
      sheetName: z.string(),
    },
    async (p) => {
      const data = await callGAS("deleteSheet", p);
      return { content: [{ type: "text", text: JSON.stringify(data, null, 2) }] };
    }
  );

  server.tool(
    "sheets_clear_sheet",
    "Clear all data from a sheet. Optionally preserve the header row.",
    {
      spreadsheetId: z.string(),
      sheetName: z.string(),
      keepHeaders: z.boolean().optional().describe("If true, row 1 headers are preserved"),
    },
    async (p) => {
      const data = await callGAS("clearSheet", p);
      return { content: [{ type: "text", text: JSON.stringify(data, null, 2) }] };
    }
  );

  server.tool(
    "sheets_read_rows",
    "Read all rows from a sheet as an array of objects keyed by header row. Returns headers and rows.",
    {
      spreadsheetId: z.string(),
      sheetName: z.string(),
      limit: z.number().optional().describe("Max rows to return"),
    },
    async (p) => {
      const data = await callGAS("readRows", p);
      return { content: [{ type: "text", text: JSON.stringify(data, null, 2) }] };
    }
  );

  server.tool(
    "sheets_append_rows",
    "Append one or more rows to a sheet. Pass rows as an array of objects (keys = column headers) or 2D arrays.",
    {
      spreadsheetId: z.string(),
      sheetName: z.string(),
      rows: z.array(z.any()).describe("Array of row objects or 2D array"),
    },
    async (p) => {
      const data = await callGAS("appendRows", p);
      return { content: [{ type: "text", text: JSON.stringify(data, null, 2) }] };
    }
  );

  server.tool(
    "sheets_update_rows",
    "Update rows where a column matches a value. Pass updates as {columnName: newValue}.",
    {
      spreadsheetId: z.string(),
      sheetName: z.string(),
      matchColumn: z.string().describe("Column header to match on"),
      matchValue: z.string().describe("Value to find in matchColumn"),
      updates: z.record(z.any()).describe("Object of { columnName: newValue } to apply"),
    },
    async (p) => {
      const data = await callGAS("updateRows", p);
      return { content: [{ type: "text", text: JSON.stringify(data, null, 2) }] };
    }
  );

  server.tool(
    "sheets_delete_rows",
    "Delete rows where a column matches a specific value.",
    {
      spreadsheetId: z.string(),
      sheetName: z.string(),
      matchColumn: z.string(),
      matchValue: z.string(),
    },
    async (p) => {
      const data = await callGAS("deleteRows", p);
      return { content: [{ type: "text", text: JSON.stringify(data, null, 2) }] };
    }
  );

  server.tool(
    "sheets_read_range",
    "Read a specific A1-notation range from a sheet. Returns a 2D array of values.",
    {
      spreadsheetId: z.string(),
      sheetName: z.string(),
      range: z.string().describe("A1 notation, e.g. 'A1:D10' or 'B2:F20'"),
    },
    async (p) => {
      const data = await callGAS("readRange", p);
      return { content: [{ type: "text", text: JSON.stringify(data, null, 2) }] };
    }
  );

  server.tool(
    "sheets_write_range",
    "Write a 2D array of values to a specific A1-notation range.",
    {
      spreadsheetId: z.string(),
      sheetName: z.string(),
      range: z.string().describe("A1 notation target range"),
      values: z.array(z.array(z.any())).describe("2D array of values to write"),
    },
    async (p) => {
      const data = await callGAS("writeRange", p);
      return { content: [{ type: "text", text: JSON.stringify(data, null, 2) }] };
    }
  );

  server.tool(
    "sheets_find_rows",
    "Find rows where column values match a query object. All conditions must match (AND logic).",
    {
      spreadsheetId: z.string(),
      sheetName: z.string(),
      query: z.record(z.string()).describe("Object of { columnName: valueToMatch }"),
    },
    async (p) => {
      const data = await callGAS("findRows", p);
      return { content: [{ type: "text", text: JSON.stringify(data, null, 2) }] };
    }
  );

  return server;
}

const httpServer = http.createServer(async (req, res) => {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, Accept");

  if (req.method === "OPTIONS") {
    res.writeHead(204); res.end(); return;
  }

  if (req.method === "GET" && req.url === "/health") {
    res.writeHead(200, { "Content-Type": "application/json" });
    res.end(JSON.stringify({ status: "ok", service: "google-sheets-mcp" }));
    return;
  }

  if (req.method === "GET" && req.url === "/mcp") {
    const server = buildServer();
    const transport = new SSEServerTransport("/messages", res);
    transports.set(transport.sessionId, transport);
    res.on("close", () => transports.delete(transport.sessionId));
    await server.connect(transport);
    return;
  }

  if (req.method === "POST" && req.url.startsWith("/messages")) {
    const url = new URL(req.url, "http://localhost");
    const sessionId = url.searchParams.get("sessionId");
    const transport = transports.get(sessionId);
    if (!transport) {
      res.writeHead(404);
      res.end(JSON.stringify({ error: "Session not found" }));
      return;
    }
    await transport.handlePostMessage(req, res);
    return;
  }

  res.writeHead(404);
  res.end("Not found");
});

httpServer.listen(PORT, () => {
  console.log(`Google Sheets MCP (SSE) running on port ${PORT}`);
});