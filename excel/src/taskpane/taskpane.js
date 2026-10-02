/* global Office, Excel, document */
import { askHermes } from "../shared/hermes";

const MAX_ROWS = 500;

const SYSTEM = `You are Hermes, embedded in an Excel task pane. Chat naturally and concisely.

When a message starts with a bracketed [Active sheet ...] note, that is the current sheet's data — use it. The data is only re-sent when it changes, so for follow-up questions rely on the data already shown earlier in the conversation.

When the user wants to modify the workbook, append EXACTLY ONE fenced block at the very end of your reply:
\`\`\`json
{"actions":[ ... ]}
\`\`\`
Supported action types:
- {"type":"setCell","cell":"B2","old":"<current>","new":"<new value>"}
- {"type":"setCells","range":"A1:B3","values":[["x","y"],["x","y"]]}
- {"type":"format","range":"C2:C50","numberFormat":"$#,##0.00","bold":false,"fill":"#FFF2CC"}
- {"type":"createTable","range":"A1:F200","name":"Leads","hasHeaders":true}
- {"type":"createChart","chartType":"ColumnClustered","dataRange":"A1:B20","title":"By status"}   // chartType: ColumnClustered | Bar | Line | Pie | XYScatter
- {"type":"newSheet","name":"Dashboard"}
- {"type":"renameSheet","to":"Acquisitions 2021"}
Always also give a short natural-language reply. Use absolute A1 refs; qualify cross-sheet refs as Sheet!A1. Omit the block entirely when no change is requested.`;

const history = [{ role: "system", content: SYSTEM }];
let pendingActions = [];
let lastSig = null;
let busy = false;

Office.onReady(() => {
  document.getElementById("ask").onclick = ask;
  document.getElementById("apply").onclick = apply;
  document.getElementById("newchat").onclick = newChat;
  document.getElementById("prompt").addEventListener("keydown", (e) => {
    if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); ask(); }
  });
});

// ---- conversation ----------------------------------------------------------

async function ask() {
  const input = document.getElementById("prompt");
  const prompt = input.value.trim();
  if (!prompt || busy) return;
  input.value = "";
  addBubble("user", prompt);
  setBusy(true, "Reading sheet…");
  try {
    const snap = await getSnapshot();
    const sig = signature(snap);
    let content = prompt;
    if (sig !== lastSig) {
      lastSig = sig;
      content = `${dataNote(snap)}\n\n${prompt}`;
    }
    history.push({ role: "user", content });

    setBusy(true, "Hermes is thinking…");
    const raw = await askHermes(history);
    history.push({ role: "assistant", content: raw });

    const { prose, actions } = splitReply(raw);
    addBubble("bot", prose);
    pendingActions = actions;
    renderActions(actions);
    setStatus(actions.length ? `${actions.length} proposed action(s) — review, then Apply.` : "Ready.");
  } catch (e) {
    addBubble("bot", "⚠ " + e.message);
    setStatus("Error.");
  } finally {
    setBusy(false);
  }
}

function newChat() {
  history.length = 1; // keep system message
  lastSig = null;
  clearPending();
  document.getElementById("log").innerHTML = "";
  setStatus("New chat. Open your data tab and ask.");
}

// ---- reading the sheet (only sent when changed) ----------------------------

async function getSnapshot() {
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.load("name");
    const used = sheet.getUsedRangeOrNullObject(true); // valuesOnly
    used.load(["address", "values"]);
    await context.sync();
    if (used.isNullObject) return { name: sheet.name, address: null, values: [], truncated: false };
    let values = used.values;
    let truncated = false;
    if (values.length > MAX_ROWS) { values = values.slice(0, MAX_ROWS); truncated = true; }
    return { name: sheet.name, address: used.address, values, truncated };
  });
}

function dataNote(s) {
  if (!s.address) return `[Active sheet "${s.name}" is empty.]`;
  return `[Active sheet "${s.name}", range ${s.address}${s.truncated ? ` (first ${MAX_ROWS} rows)` : ""}. Current data:]\n${JSON.stringify(s.values)}`;
}

function signature(s) {
  return `${s.name}|${s.address}|${s.values.length}|${hash(JSON.stringify(s.values))}`;
}
function hash(str) {
  let h = 5381;
  for (let i = 0; i < str.length; i++) h = ((h << 5) + h + str.charCodeAt(i)) | 0;
  return h;
}

// ---- parsing the agent's reply ---------------------------------------------

function splitReply(raw) {
  let actions = [];
  let prose = raw;
  const fenced = raw.match(/```json\s*([\s\S]*?)```/i);
  const target = fenced ? fenced[1] : (raw.match(/\{[\s\S]*"actions"[\s\S]*\}/) || [])[0];
  if (target) {
    try {
      const obj = JSON.parse(target);
      actions = obj.actions || (obj.editPlan ? obj.editPlan.map((e) => ({ type: "setCell", ...e })) : []);
    } catch (_) { /* leave actions empty */ }
    prose = raw.replace(fenced ? fenced[0] : target, "").trim();
  }
  return { prose: prose || "(proposed changes below)", actions: Array.isArray(actions) ? actions : [] };
}

// ---- preview + apply -------------------------------------------------------

function describe(a) {
  switch (a.type) {
    case "setCell": return `Set ${a.cell}:  "${a.old ?? ""}" → "${a.new}"`;
    case "setCells": return `Fill ${a.range} (${(a.values || []).length} rows)`;
    case "format": return `Format ${a.range}${a.numberFormat ? ` as ${a.numberFormat}` : ""}${a.bold ? " (bold)" : ""}`;
    case "createTable": return `Create table "${a.name || "Table"}" over ${a.range}`;
    case "createChart": return `Create ${a.chartType || "Column"} chart from ${a.dataRange}${a.title ? ` — "${a.title}"` : ""}`;
    case "newSheet": return `New sheet "${a.name}"`;
    case "renameSheet": return `Rename active tab → "${a.to || a.name}"`;
    default: return JSON.stringify(a);
  }
}

function renderActions(actions) {
  const box = document.getElementById("preview");
  box.innerHTML = "";
  if (!actions.length) { document.getElementById("apply").style.display = "none"; return; }
  actions.forEach((a) => {
    const div = document.createElement("div");
    div.className = "act";
    div.textContent = "• " + describe(a);
    box.appendChild(div);
  });
  document.getElementById("apply").style.display = "block";
}

async function apply() {
  if (!pendingActions.length || busy) return;
  setBusy(true, "Applying…");
  try {
    await Excel.run(async (context) => {
      const wb = context.workbook;
      let sheet = wb.worksheets.getActiveWorksheet();
      for (const a of pendingActions) {
        switch (a.type) {
          case "newSheet": {
            sheet = wb.worksheets.add(a.name || "Sheet");
            sheet.activate();
            break;
          }
          case "renameSheet": {
            sheet.name = a.to || a.name;
            break;
          }
          case "setCell": {
            const r = resolveRange(wb, sheet, a.cell);
            r.values = [[a.new]];
            r.format.fill.color = "#C6EFCE";
            break;
          }
          case "setCells": {
            resolveRange(wb, sheet, a.range).values = a.values;
            break;
          }
          case "format": {
            const r = resolveRange(wb, sheet, a.range);
            if (a.bold !== undefined) r.format.font.bold = !!a.bold;
            if (a.fill) r.format.fill.color = a.fill;
            if (a.numberFormat) {
              r.load("rowCount, columnCount");
              await context.sync();
              const fmt = Array.from({ length: r.rowCount }, () => Array.from({ length: r.columnCount }, () => a.numberFormat));
              r.numberFormat = fmt;
            }
            break;
          }
          case "createTable": {
            const t = wb.tables.add(resolveRange(wb, sheet, a.range), a.hasHeaders !== false);
            if (a.name) t.name = tableName(a.name);
            break;
          }
          case "createChart": {
            const ch = sheet.charts.add(chartType(a.chartType), resolveRange(wb, sheet, a.dataRange), Excel.ChartSeriesBy.auto);
            if (a.title) ch.title.text = a.title;
            break;
          }
        }
      }
      await context.sync();
    });
    addBubble("bot", `Applied ${pendingActions.length} action(s).`);
    clearPending();
    lastSig = null; // workbook changed — re-send fresh data on the next turn
    setStatus("Ready.");
  } catch (e) {
    addBubble("bot", "⚠ " + e.message);
    setStatus("Error.");
  } finally {
    setBusy(false);
  }
}

// ---- helpers ---------------------------------------------------------------

function resolveRange(wb, fallbackSheet, addr) {
  addr = String(addr || "").trim();
  if (addr.includes("!")) {
    const i = addr.lastIndexOf("!");
    const sn = addr.slice(0, i).replace(/^'|'$/g, "").replace(/''/g, "'");
    return wb.worksheets.getItem(sn).getRange(addr.slice(i + 1));
  }
  return fallbackSheet.getRange(addr);
}

function chartType(t) {
  const m = {
    columnclustered: "ColumnClustered", column: "ColumnClustered", columns: "ColumnClustered",
    bar: "BarClustered", barclustered: "BarClustered",
    line: "Line", pie: "Pie", doughnut: "Doughnut", area: "Area",
    scatter: "XYScatter", xyscatter: "XYScatter",
  };
  return m[String(t || "").toLowerCase().replace(/[^a-z]/g, "")] || "ColumnClustered";
}

function tableName(n) {
  return String(n).replace(/[^A-Za-z0-9_]/g, "_").replace(/^[^A-Za-z_]/, "_");
}

function clearPending() {
  pendingActions = [];
  document.getElementById("preview").innerHTML = "";
  document.getElementById("apply").style.display = "none";
}

function setBusy(b, msg) {
  busy = b;
  document.getElementById("ask").disabled = b;
  if (msg) setStatus(msg);
}

function addBubble(who, text) {
  const log = document.getElementById("log");
  const div = document.createElement("div");
  div.className = "msg " + who;
  div.textContent = text;
  log.appendChild(div);
  log.scrollTop = log.scrollHeight;
}

function setStatus(s) { document.getElementById("status").textContent = s; }
