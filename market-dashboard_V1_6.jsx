import { useState, useCallback, useMemo, useEffect } from "react";
import * as XLSX from "xlsx";

// ═══════════════════════════════════════════════════════════════════════════════
// CONFIG
// ═══════════════════════════════════════════════════════════════════════════════
const ANTHROPIC_URL = "https://api.anthropic.com/v1/messages";
const MODEL_NEWS = "claude-haiku-4-5-20251001";
const MODEL_AI = "claude-sonnet-4-20250514";
const TTL_DAY = 24 * 60 * 60 * 1000;
const HISTORY_KEY = "marketDashboard:priceHistory:v1";
const HISTORY_FILENAME_PREFIX = "historique-marche-";

// ═══════════════════════════════════════════════════════════════════════════════
// DONNÉES DE MARCHÉ — 100% via Anthropic + web_search
// Source unique : api.anthropic.com (CSP-safe, aucune autre origine appelée).
// ═══════════════════════════════════════════════════════════════════════════════

// Récupère les prix de tous les tickers stocks en UN SEUL appel API
async function fetchStocks(tickers) {
  if (!tickers || tickers.length === 0) {
    return {};
  }
  const list = tickers.join(",");
  const prompt =
    "Récupère le prix actuel (ou dernière clôture si marché fermé) de chacun de ces tickers US : " +
    list +
    ". Réponds STRICTEMENT avec un objet JSON sans markdown ni texte autour. " +
    "Pour chaque ticker, indique price (nombre USD) et changePct (variation % sur la journée, nombre). " +
    'Format exact : {"NVDA":{"price":478.32,"changePct":1.16},"SPY":{"price":597.5,"changePct":0.31}}. ' +
    "Si un ticker est introuvable, mets price:0 et changePct:0. Ne renvoie QUE le JSON.";

  const json = await callAPI(
    MODEL_NEWS,
    [{ role: "user", content: prompt }],
    {
      max_tokens: Math.min(4000, 80 + tickers.length * 30),
      tools: [{ type: "web_search_20250305", name: "web_search" }],
    }
  );
  const blocks = (json.content || []).filter((b) => b.type === "text");
  const raw = blocks.map((b) => b.text).join("");
  if (!raw) {
    throw new Error("API Anthropic : réponse vide");
  }
  const parsed = extractJSON(raw, false);
  if (!parsed || typeof parsed !== "object") {
    throw new Error("Réponse non parsable");
  }
  const result = {};
  for (const t of tickers) {
    const q = parsed[t];
    if (!q || typeof q.price !== "number" || q.price <= 0) {
      continue;
    }
    const pct = typeof q.changePct === "number" ? q.changePct : 0;
    const price = q.price;
    result[t] = {
      price,
      change: (price * pct) / 100,
      changePct: pct,
    };
  }
  if (Object.keys(result).length === 0) {
    throw new Error("Aucune donnée exploitable (réponse : " + raw.substring(0, 120) + ")");
  }
  return result;
}

// Récupère les données crypto en UN SEUL appel API (price, change24h, high, low, vol)
async function fetchCryptos(symbols) {
  if (!symbols || symbols.length === 0) {
    return {};
  }
  // Mapping symbol Binance-like → id court pour le prompt
  const ids = symbols.map((s) => s.replace(/USDT$/i, "").replace(/USD$/i, ""));
  const list = ids.join(",");
  const prompt =
    "Récupère les données crypto actuelles en USD pour ces cryptomonnaies : " +
    list +
    ". Réponds STRICTEMENT avec un objet JSON sans markdown ni texte autour. " +
    "Pour chaque crypto, indique : price (prix USD, nombre), change24h (variation % sur 24h, nombre), " +
    "high24h (plus haut 24h USD, nombre), low24h (plus bas 24h USD, nombre), quoteVol (volume 24h en USD, nombre). " +
    'Format exact : {"BTC":{"price":67500,"change24h":1.2,"high24h":68200,"low24h":66900,"quoteVol":35000000000},"ETH":{"price":3450,"change24h":-0.8,"high24h":3500,"low24h":3420,"quoteVol":15000000000}}. ' +
    "Si introuvable, mets toutes les valeurs à 0. Ne renvoie QUE le JSON.";

  const json = await callAPI(
    MODEL_NEWS,
    [{ role: "user", content: prompt }],
    {
      max_tokens: Math.min(4000, 120 + symbols.length * 90),
      tools: [{ type: "web_search_20250305", name: "web_search" }],
    }
  );
  const blocks = (json.content || []).filter((b) => b.type === "text");
  const raw = blocks.map((b) => b.text).join("");
  if (!raw) {
    throw new Error("API Anthropic crypto : réponse vide");
  }
  const parsed = extractJSON(raw, false);
  if (!parsed || typeof parsed !== "object") {
    throw new Error("Crypto : réponse non parsable");
  }
  const result = {};
  for (let i = 0; i < symbols.length; i++) {
    const sym = symbols[i];
    const id = ids[i];
    // Tolérance : le LLM peut renvoyer la clé sous plusieurs formes
    const q = parsed[id] || parsed[sym] || parsed[id.toUpperCase()] || parsed[id.toLowerCase()];
    if (!q || typeof q.price !== "number" || q.price <= 0) {
      continue;
    }
    result[sym] = {
      price: q.price,
      change24h: typeof q.change24h === "number" ? q.change24h : 0,
      high24h: typeof q.high24h === "number" && q.high24h > 0 ? q.high24h : 0,
      low24h: typeof q.low24h === "number" && q.low24h > 0 ? q.low24h : 0,
      quoteVol: typeof q.quoteVol === "number" && q.quoteVol > 0 ? q.quoteVol : 0,
    };
  }
  if (Object.keys(result).length === 0) {
    throw new Error("Aucune crypto exploitable (réponse : " + raw.substring(0, 120) + ")");
  }
  return result;
}

// ═══════════════════════════════════════════════════════════════════════════════
// ANTHROPIC API — appel commun
// ═══════════════════════════════════════════════════════════════════════════════
async function callAPI(model, messages, opts) {
  const body = Object.assign({ model, max_tokens: 1800, messages }, opts || {});
  const maxRetries = 3;
  let lastErr;
  for (let attempt = 0; attempt <= maxRetries; attempt++) {
    try {
      const res = await fetch(ANTHROPIC_URL, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(body),
      });
      if ((res.status === 429 || res.status >= 500) && attempt < maxRetries) {
        const wait = Math.min(8000, 1500 * Math.pow(2, attempt));
        console.warn("Anthropic HTTP " + res.status + " — retry dans " + Math.round(wait / 1000) + "s (" + (attempt + 1) + "/" + maxRetries + ")");
        await new Promise((r) => setTimeout(r, wait));
        continue;
      }
      if (!res.ok) {
        if (res.status === 429) {
          throw new Error("HTTP 429 (limite atteinte, attendez ~1 min)");
        }
        throw new Error("HTTP " + res.status);
      }
      return res.json();
    } catch (e) {
      lastErr = e;
      if (attempt >= maxRetries || (e.message && e.message.indexOf("HTTP 4") === 0 && e.message.indexOf("HTTP 429") !== 0)) {
        throw e;
      }
    }
  }
  throw lastErr || new Error("API échec après retries");
}

function repairJSON(raw) {
  let s = raw.trim();
  s = s.replace(/^```(?:json)?\s*/i, "").replace(/```\s*$/i, "").trim();
  s = s.replace(/,(\s*[}\]])/g, "$1");

  const out = [];
  let inString = false;
  let escaped = false;
  for (let i = 0; i < s.length; i++) {
    const ch = s[i];
    if (!inString) {
      out.push(ch);
      if (ch === '"') {
        inString = true;
        escaped = false;
      }
      continue;
    }
    if (escaped) {
      out.push(ch);
      escaped = false;
      continue;
    }
    if (ch === "\\") {
      out.push(ch);
      escaped = true;
      continue;
    }
    if (ch === '"') {
      let j = i + 1;
      while (j < s.length && /\s/.test(s[j])) {
        j++;
      }
      const next = s[j];
      if (next === "," || next === "}" || next === "]" || next === ":" || j >= s.length) {
        out.push(ch);
        inString = false;
      } else {
        out.push("'");
      }
      continue;
    }
    if (ch === "\n" || ch === "\r") {
      out.push(" ");
      continue;
    }
    out.push(ch);
  }
  let result = out.join("");
  result = closeTruncatedJSON(result);
  result = result.replace(/,(\s*[}\]])/g, "$1");
  return result;
}

function closeTruncatedJSON(s) {
  let cleaned = s;
  for (let pass = 0; pass < 10; pass++) {
    const before = cleaned;
    if (isInUnclosedString(cleaned)) {
      cleaned = cleaned + '"';
    }
    cleaned = cleaned.replace(/,\s*"[^"]*"\s*:\s*$/g, "");
    cleaned = cleaned.replace(/,\s*"[^"]*"\s*:\s*[^,}\]"]*$/g, "");
    cleaned = cleaned.replace(/,\s*"[^"]*"\s*$/g, "");
    cleaned = cleaned.replace(/,\s*[^,}\]"]*$/g, "");
    cleaned = cleaned.replace(/([\{\[])\s*"[^"]*"\s*:\s*$/g, "$1");
    cleaned = cleaned.replace(/([\{\[])\s*"[^"]*"\s*:\s*[^,}\]"]*$/g, "$1");
    cleaned = cleaned.replace(/([\{\[])\s*"[^"]*"\s*$/g, "$1");
    cleaned = cleaned.replace(/,\s*$/g, "");
    cleaned = cleaned.replace(/\s+$/g, "");
    if (cleaned === before) break;
  }

  const stack = [];
  let inString = false;
  let escaped = false;
  for (let i = 0; i < cleaned.length; i++) {
    const ch = cleaned[i];
    if (inString) {
      if (escaped) { escaped = false; continue; }
      if (ch === "\\") { escaped = true; continue; }
      if (ch === '"') inString = false;
      continue;
    }
    if (ch === '"') { inString = true; continue; }
    if (ch === "{") stack.push("}");
    else if (ch === "[") stack.push("]");
    else if (ch === "}" || ch === "]") stack.pop();
  }

  if (inString) cleaned += '"';
  while (stack.length > 0) {
    cleaned += stack.pop();
  }
  return cleaned;
}

function isInUnclosedString(s) {
  let inString = false;
  let escaped = false;
  for (let i = 0; i < s.length; i++) {
    const ch = s[i];
    if (inString) {
      if (escaped) { escaped = false; continue; }
      if (ch === "\\") { escaped = true; continue; }
      if (ch === '"') inString = false;
      continue;
    }
    if (ch === '"') inString = true;
  }
  return inString;
}

function extractJSON(text, arrayMode) {
  if (!text) {
    throw new Error("Réponse vide");
  }
  const cleaned = text.replace(/```json|```/g, "").trim();
  const re = arrayMode ? /\[[\s\S]*\]/ : /\{[\s\S]*\}/;
  const m = cleaned.match(re);
  if (m) {
    try {
      return JSON.parse(m[0]);
    } catch (e1) {
      try {
        const repaired = repairJSON(m[0]);
        return JSON.parse(repaired);
      } catch (e2) {
        try {
          const repaired = repairJSON(cleaned);
          return JSON.parse(repaired);
        } catch (e3) {
          // continue vers fallback
        }
      }
    }
  }
  if (arrayMode) {
    const objMatch = cleaned.match(/\{[\s\S]*\}/);
    if (objMatch) {
      try {
        const obj = JSON.parse(repairJSON(objMatch[0]));
        const arr = Object.values(obj).find((v) => Array.isArray(v));
        if (arr) {
          return arr;
        }
      } catch (e) {
        // ignore
      }
    }
  }
  console.error("extractJSON: irréparable. Réponse brute (1000 premiers chars):");
  console.error(text.substring(0, 1000));
  throw new Error("JSON irréparable (voir console pour la réponse brute)");
}

async function fetchNews() {
  const d = new Date();
  const today = d.toLocaleDateString("fr-FR", { day: "numeric", month: "long", year: "numeric" });
  const prompt =
    "Date: " + today + ". Trouve 10 actualités majeures macro/géopolitiques des 30 derniers jours impactant les marchés US (Fed, géopolitique, crypto, économie, énergie, tech). " +
    "Réponds STRICTEMENT avec un tableau JSON valide, du plus récent au plus ancien, sans markdown ni texte autour. " +
    "RÈGLES IMPORTANTES : (1) JAMAIS de guillemets doubles (\") à l'intérieur des valeurs string — remplace-les par des apostrophes simples ('). " +
    "(2) Aucune virgule en trop. (3) Pas de retour à la ligne dans les valeurs. (4) Exactement 10 entrées. " +
    "Format strict : [{\"date\":\"12 mai\",\"impact\":\"POSITIVE|NEUTRAL|NEGATIVE\",\"sev\":7,\"title\":\"titre court sans guillemets\",\"summary\":\"une phrase sans guillemets\",\"tags\":[\"Or up\"],\"src\":\"source\"}]";
  const json = await callAPI(
    MODEL_NEWS,
    [{ role: "user", content: prompt }],
    { max_tokens: 1800, tools: [{ type: "web_search_20250305", name: "web_search" }] }
  );
  const blocks = (json.content || []).filter((b) => b.type === "text");
  const raw = blocks.map((b) => b.text).join("");
  if (!raw) {
    throw new Error("Aucun texte retourné par l'API");
  }
  const arr = extractJSON(raw, true);
  if (!Array.isArray(arr) || arr.length === 0) {
    throw new Error("Tableau de news vide");
  }
  return arr;
}

const KEY_STOCKS = ["SPY", "QQQ", "NVDA", "AAPL", "MSFT", "GLD", "SLV", "TLT", "XLE", "JPM"];

async function fetchAIAnalysis(stocks, crypto, news) {
  const sl = KEY_STOCKS.filter((t) => stocks[t])
    .map((t) => {
      const d = stocks[t];
      const pct = d.changePct || 0;
      return t + ":$" + (d.price || 0).toFixed(0) + "(" + (pct >= 0 ? "+" : "") + pct.toFixed(1) + "%)";
    })
    .join(" ");
  const cl = Object.entries(crypto)
    .map(([s, d]) => {
      const pct = d.change24h || 0;
      return s.replace("USDT", "") + ":$" + (d.price || 0).toFixed(0) + "(" + (pct >= 0 ? "+" : "") + pct.toFixed(1) + "%)";
    })
    .join(" ");
  const nl = (news || [])
    .slice(0, 10)
    .map((n) => "[" + n.impact + "] " + n.title)
    .join(" | ");
  const schema =
    "{\"overallSentiment\":\"HAUSSIER|NEUTRE|BAISSIER\"," +
    "\"globalSummary\":\"3-4 phrases\"," +
    "\"topOpportunities\":[{\"asset\":\"\",\"reason\":\"\",\"action\":\"BUY\"}]," +
    "\"topRisks\":[{\"asset\":\"\",\"reason\":\"\",\"action\":\"SELL\"}]," +
    "\"indicators\":{" +
      "\"vix\":{\"v\":0,\"status\":\"\",\"desc\":\"\",\"tone\":\"up|warn|down|neutral\"}," +
      "\"fearGreed\":{\"v\":0,\"status\":\"\",\"desc\":\"\",\"tone\":\"up|warn|down|neutral\"}," +
      "\"dxy\":{\"v\":0,\"status\":\"\",\"desc\":\"\",\"tone\":\"up|warn|down|neutral\"}," +
      "\"yield10y\":{\"v\":0,\"status\":\"\",\"desc\":\"\",\"tone\":\"up|warn|down|neutral\"}," +
      "\"btcDom\":{\"v\":0,\"status\":\"\",\"desc\":\"\",\"tone\":\"up|warn|down|neutral\"}," +
      "\"aiSent\":{\"v\":0,\"status\":\"\",\"desc\":\"\",\"tone\":\"up|warn|down|neutral\"}" +
    "}," +
    "\"sectors\":{" +
      "\"indices\":{\"action\":\"HOLD\",\"confidence\":7,\"signal\":\"NEUTRE\",\"reason\":\"...\",\"risk\":\"...\"}," +
      "\"mag7\":{\"action\":\"BUY\",\"confidence\":8,\"signal\":\"ACHAT\",\"reason\":\"...\",\"risk\":\"...\"}," +
      "\"semis\":{\"action\":\"BUY\",\"confidence\":8,\"signal\":\"ACHAT\",\"reason\":\"...\",\"risk\":\"...\"}," +
      "\"metals\":{\"action\":\"BUY\",\"confidence\":9,\"signal\":\"FORT ACHAT\",\"reason\":\"...\",\"risk\":\"...\"}," +
      "\"energy\":{\"action\":\"TRANSFER\",\"confidence\":5,\"signal\":\"PIVOTER\",\"reason\":\"...\",\"risk\":\"...\"}," +
      "\"health\":{\"action\":\"HOLD\",\"confidence\":5,\"signal\":\"NEUTRE\",\"reason\":\"...\",\"risk\":\"...\"}," +
      "\"finance\":{\"action\":\"HOLD\",\"confidence\":6,\"signal\":\"NEUTRE\",\"reason\":\"...\",\"risk\":\"...\"}," +
      "\"bonds\":{\"action\":\"SELL\",\"confidence\":7,\"signal\":\"VENTE\",\"reason\":\"...\",\"risk\":\"...\"}," +
      "\"realestate\":{\"action\":\"SELL\",\"confidence\":7,\"signal\":\"VENTE\",\"reason\":\"...\",\"risk\":\"...\"}," +
      "\"emerging\":{\"action\":\"HOLD\",\"confidence\":5,\"signal\":\"NEUTRE\",\"reason\":\"...\",\"risk\":\"...\"}" +
    "}}";
  const prompt =
    "Gérant macro senior. Analyse du jour à partir UNIQUEMENT des données fournies (aucun a priori).\n" +
    "STOCKS: " + (sl || "N/A") + "\n" +
    "CRYPTO: " + (cl || "N/A") + "\n" +
    "NEWS: " + (nl || "N/A") + "\n" +
    "Renvoie uniquement le JSON ci-dessous (action=BUY|HOLD|SELL|TRANSFER, tone=up|warn|down|neutral) : " + schema;
  const json = await callAPI(MODEL_AI, [{ role: "user", content: prompt }], { max_tokens: 4500 });
  const blocks = (json.content || []).filter((b) => b.type === "text");
  const text = blocks.map((b) => b.text).join("");
  return extractJSON(text, false);
}

// ═══════════════════════════════════════════════════════════════════════════════
// CACHE EN MÉMOIRE (vierge au démarrage)
// ═══════════════════════════════════════════════════════════════════════════════
const MEM = {
  stocks: null,
  stocksAt: null,
  crypto: null,
  cryptoAt: null,
  news: null,
  newsAt: null,
  ai: null,
  aiAt: null,
  history: null,
  historyAt: null,
};

function ageString(at) {
  if (!at) {
    return null;
  }
  const d = Date.now() - at;
  if (d < 60000) {
    return "à l'instant";
  }
  if (d < 3600000) {
    return "il y a " + Math.floor(d / 60000) + " min";
  }
  return "il y a " + Math.floor(d / 3600000) + "h";
}

function nextRefreshString(at) {
  if (!at) {
    return null;
  }
  const ms = TTL_DAY - (Date.now() - at);
  if (ms <= 0) {
    return "maintenant";
  }
  const h = Math.floor(ms / 3600000);
  const m = Math.floor((ms % 3600000) / 60000);
  if (h > 0) {
    return "dans " + h + "h" + (m > 0 ? m + "min" : "");
  }
  return "dans " + m + "min";
}

// ═══════════════════════════════════════════════════════════════════════════════
// HISTORIQUE PRIX — Persistance + Export Excel
// ═══════════════════════════════════════════════════════════════════════════════
function todayKey() {
  const d = new Date();
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return y + "-" + m + "-" + day;
}

async function storageGet(key) {
  if (typeof window === "undefined") {
    return null;
  }
  try {
    if (window.storage && typeof window.storage.get === "function") {
      const r = await window.storage.get(key);
      return r && typeof r.value !== "undefined" ? r.value : null;
    }
  } catch (e) {
    // fallthrough
  }
  try {
    if (window.localStorage) {
      return window.localStorage.getItem(key);
    }
  } catch (e) {
    // ignore
  }
  return null;
}

async function storageSet(key, value) {
  if (typeof window === "undefined") {
    return false;
  }
  try {
    if (window.storage && typeof window.storage.set === "function") {
      await window.storage.set(key, value);
      return true;
    }
  } catch (e) {
    // fallthrough
  }
  try {
    if (window.localStorage) {
      window.localStorage.setItem(key, value);
      return true;
    }
  } catch (e) {
    // ignore
  }
  return false;
}

function emptyHistory() {
  return { version: 1, lastDate: null, tickers: [], rows: {} };
}

async function loadHistory() {
  const raw = await storageGet(HISTORY_KEY);
  if (!raw) {
    return emptyHistory();
  }
  try {
    const parsed = typeof raw === "string" ? JSON.parse(raw) : raw;
    if (
      parsed &&
      typeof parsed === "object" &&
      parsed.rows &&
      typeof parsed.rows === "object" &&
      Array.isArray(parsed.tickers)
    ) {
      return {
        version: parsed.version || 1,
        lastDate: parsed.lastDate || null,
        tickers: parsed.tickers.slice(),
        rows: Object.assign({}, parsed.rows),
      };
    }
  } catch (e) {
    console.warn("History parse failed:", e);
  }
  return emptyHistory();
}

function appendTodayPrices(history, stocks) {
  const key = todayKey();
  const base = history && history.rows ? history : emptyHistory();
  const tickerSet = new Set(base.tickers || []);
  const row = {};
  for (const t in stocks) {
    if (!Object.prototype.hasOwnProperty.call(stocks, t)) {
      continue;
    }
    const v = stocks[t];
    if (v && typeof v.price === "number" && v.price > 0 && isFinite(v.price)) {
      row[t] = Math.round(v.price * 10000) / 10000;
      tickerSet.add(t);
    }
  }
  if (Object.keys(row).length === 0) {
    return base;
  }
  const nextRows = Object.assign({}, base.rows);
  nextRows[key] = row;
  return {
    version: 1,
    lastDate: key,
    tickers: Array.from(tickerSet).sort(),
    rows: nextRows,
  };
}

function buildHistoryWorkbook(history, preferredTickerOrder) {
  const h = history && history.rows ? history : emptyHistory();
  const dates = Object.keys(h.rows).sort();
  const allTickersSet = new Set(h.tickers || []);
  for (const d of dates) {
    const r = h.rows[d] || {};
    for (const t in r) {
      if (Object.prototype.hasOwnProperty.call(r, t)) {
        allTickersSet.add(t);
      }
    }
  }
  const ordered = [];
  const seen = new Set();
  if (Array.isArray(preferredTickerOrder)) {
    for (const t of preferredTickerOrder) {
      if (allTickersSet.has(t) && !seen.has(t)) {
        ordered.push(t);
        seen.add(t);
      }
    }
  }
  const remaining = Array.from(allTickersSet).filter((t) => !seen.has(t)).sort();
  const tickers = ordered.concat(remaining);

  const aoa = [["Date"].concat(tickers)];
  for (const d of dates) {
    const r = h.rows[d] || {};
    const line = [d];
    for (const t of tickers) {
      const v = r[t];
      line.push(typeof v === "number" ? v : "");
    }
    aoa.push(line);
  }

  const ws = XLSX.utils.aoa_to_sheet(aoa);
  const cols = [{ wch: 12 }];
  for (let i = 0; i < tickers.length; i++) {
    cols.push({ wch: Math.max(8, tickers[i].length + 2) });
  }
  ws["!cols"] = cols;
  ws["!freeze"] = { xSplit: 1, ySplit: 1 };

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Historique");
  return wb;
}

function downloadHistoryExcel(history, preferredTickerOrder) {
  const wb = buildHistoryWorkbook(history, preferredTickerOrder);
  const fname = HISTORY_FILENAME_PREFIX + todayKey() + ".xlsx";
  try {
    XLSX.writeFile(wb, fname);
    return true;
  } catch (e) {
    console.error("Excel download failed:", e);
    try {
      const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array" });
      const blob = new Blob([wbout], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = fname;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      setTimeout(function () {
        URL.revokeObjectURL(url);
      }, 1000);
      return true;
    } catch (e2) {
      console.error("Excel fallback download failed:", e2);
      return false;
    }
  }
}

function historyStats(history) {
  if (!history || !history.rows) {
    return { days: 0, tickers: 0, first: null, last: null };
  }
  const dates = Object.keys(history.rows).sort();
  return {
    days: dates.length,
    tickers: (history.tickers || []).length,
    first: dates[0] || null,
    last: dates[dates.length - 1] || null,
  };
}

// ═══════════════════════════════════════════════════════════════════════════════
// DESIGN TOKENS
// ═══════════════════════════════════════════════════════════════════════════════
const C = {
  bg: "#06090F",
  surface: "#0B1018",
  card: "#0D1421",
  border: "#1A2540",
  bright: "#243560",
  buy: "#00C07A",
  buyBg: "#00C07A14",
  buyBd: "#00C07A30",
  hold: "#3B82F6",
  holdBg: "#3B82F614",
  holdBd: "#3B82F630",
  xfer: "#F5A623",
  xferBg: "#F5A62314",
  xferBd: "#F5A62330",
  sell: "#F04040",
  sellBg: "#F0404014",
  sellBd: "#F0404030",
  up: "#00C07A",
  dn: "#F04040",
  txt: "#EFF4FC",
  txt2: "#94A3B8",
  muted: "#475569",
  accent: "#4F8EF7",
  yf: "#720CE8",
  ai: "#A78BFA",
};

const ACTIONS = [
  { key: "BUY", fr: "ACHETER", icon: "↑", color: C.buy, bg: C.buyBg, bd: C.buyBd },
  { key: "HOLD", fr: "GARDER", icon: "◆", color: C.hold, bg: C.holdBg, bd: C.holdBd },
  { key: "TRANSFER", fr: "TRANSFÉRER", icon: "⇄", color: C.xfer, bg: C.xferBg, bd: C.xferBd },
  { key: "SELL", fr: "VENDRE", icon: "↓", color: C.sell, bg: C.sellBg, bd: C.sellBd },
];

function actColor(key) {
  return ACTIONS.find((a) => a.key === key) || ACTIONS[1];
}

function actionFromChange(pct) {
  if (pct >= 2) {
    return "BUY";
  }
  if (pct >= 0.3) {
    return "HOLD";
  }
  if (pct <= -2) {
    return "SELL";
  }
  if (pct <= -0.5) {
    return "TRANSFER";
  }
  return "HOLD";
}

function toneColor(tone) {
  if (tone === "up" || tone === "buy") return C.up;
  if (tone === "warn") return C.xfer;
  if (tone === "down" || tone === "sell") return C.sell;
  return C.hold;
}

// ═══════════════════════════════════════════════════════════════════════════════
// CATALOGUES
// ═══════════════════════════════════════════════════════════════════════════════
const SECTOR_META = {
  indices: { name: "Indices US", icon: "🇺🇸", accent: "#4F8EF7" },
  mag7: { name: "Mega Cap Tech", icon: "🤖", accent: "#22C55E", sub: "Magnificent 7" },
  semis: { name: "Semi-conducteurs", icon: "💾", accent: "#A78BFA" },
  metals: { name: "Métaux Précieux", icon: "🥇", accent: "#F59E0B" },
  energy: { name: "Énergie", icon: "⚡", accent: "#F97316" },
  health: { name: "Santé / Pharma", icon: "🏥", accent: "#06B6D4" },
  finance: { name: "Finance & Bancaire", icon: "🏦", accent: "#64748B" },
  bonds: { name: "Obligations / Taux", icon: "💵", accent: "#8B5CF6" },
  realestate: { name: "Immobilier / REITs", icon: "🏠", accent: "#EC4899" },
  emerging: { name: "Marchés Émergents", icon: "🌍", accent: "#10B981" },
};

const STOCKS = [
  { t: "SPY", n: "S&P 500", s: "indices" },
  { t: "QQQ", n: "NASDAQ 100", s: "indices" },
  { t: "DIA", n: "Dow Jones", s: "indices" },
  { t: "IWM", n: "Russell 2000", s: "indices" },
  { t: "NVDA", n: "NVIDIA", s: "mag7" },
  { t: "META", n: "Meta", s: "mag7" },
  { t: "MSFT", n: "Microsoft", s: "mag7" },
  { t: "AMZN", n: "Amazon", s: "mag7" },
  { t: "GOOGL", n: "Alphabet", s: "mag7" },
  { t: "AAPL", n: "Apple", s: "mag7" },
  { t: "TSLA", n: "Tesla", s: "mag7" },
  { t: "TSM", n: "TSMC", s: "semis" },
  { t: "AVGO", n: "Broadcom", s: "semis" },
  { t: "AMD", n: "AMD", s: "semis" },
  { t: "INTC", n: "Intel", s: "semis" },
  { t: "GLD", n: "Or (GLD)", s: "metals" },
  { t: "SLV", n: "Argent (SLV)", s: "metals" },
  { t: "GDX", n: "Gold Miners", s: "metals" },
  { t: "USO", n: "WTI Crude", s: "energy" },
  { t: "XLE", n: "Energy SPDR", s: "energy" },
  { t: "XOM", n: "ExxonMobil", s: "energy" },
  { t: "LLY", n: "Eli Lilly", s: "health" },
  { t: "JNJ", n: "J&J", s: "health" },
  { t: "XLV", n: "Health ETF", s: "health" },
  { t: "JPM", n: "JPMorgan", s: "finance" },
  { t: "XLF", n: "Financial ETF", s: "finance" },
  { t: "TLT", n: "T-Bonds 30Y", s: "bonds" },
  { t: "BIL", n: "T-Bills 1-3M", s: "bonds" },
  { t: "VNQ", n: "REITs ETF", s: "realestate" },
  { t: "FXI", n: "Chine", s: "emerging" },
  { t: "INDA", n: "Inde", s: "emerging" },
  { t: "EEM", n: "EM Global", s: "emerging" },
];

const ETF_LIST = [
  { t: "VOO", n: "Vanguard S&P 500", cat: "Large Cap", fee: 0.03, aum: "827B", risk: 3, desc: "Cœur de portefeuille. Frais ultra-bas." },
  { t: "QQQ", n: "Invesco NASDAQ-100", cat: "Tech", fee: 0.20, aum: "300B", risk: 6, desc: "Mag7 + IA. ~18%/an sur 10 ans." },
  { t: "VTI", n: "Vanguard Total Market", cat: "Total US", fee: 0.03, aum: "450B", risk: 3, desc: "3500+ entreprises US." },
  { t: "SCHD", n: "Schwab Dividend", cat: "Dividend", fee: 0.06, aum: "65B", risk: 3, desc: "Revenu passif qualité." },
  { t: "VUG", n: "Vanguard Growth", cat: "Growth", fee: 0.03, aum: "180B", risk: 6, desc: "150+ actions growth." },
  { t: "VGT", n: "Vanguard Tech", cat: "Tech", fee: 0.09, aum: "75B", risk: 7, desc: "Tech US pure." },
  { t: "SMH", n: "VanEck Semiconductors", cat: "Semis", fee: 0.35, aum: "30B", risk: 8, desc: "NVDA, TSM, AVGO. Leader IA." },
  { t: "XLE", n: "Energy Select", cat: "Énergie", fee: 0.09, aum: "40B", risk: 7, desc: "Exxon, Chevron. Tactique." },
  { t: "GLD", n: "SPDR Gold", cat: "Or", fee: 0.40, aum: "75B", risk: 4, desc: "Référence or physique." },
  { t: "IAU", n: "iShares Gold", cat: "Or", fee: 0.25, aum: "30B", risk: 4, desc: "Alternative moins chère à GLD." },
  { t: "SLV", n: "iShares Silver", cat: "Argent", fee: 0.50, aum: "15B", risk: 6, desc: "Argent physique." },
  { t: "GDX", n: "VanEck Gold Miners", cat: "Miners", fee: 0.51, aum: "30B", risk: 7, desc: "Levier 2-3x sur l'or." },
  { t: "XLF", n: "Financial SPDR", cat: "Finance", fee: 0.09, aum: "45B", risk: 5, desc: "JPM, Berkshire, Goldman." },
  { t: "XLV", n: "Health Care", cat: "Santé", fee: 0.09, aum: "35B", risk: 4, desc: "Défensif. Lilly, JNJ, UNH." },
  { t: "IBIT", n: "iShares Bitcoin ETF", cat: "Crypto", fee: 0.25, aum: "65B", risk: 9, desc: "BTC spot ETF." },
  { t: "VXUS", n: "Vanguard Intl", cat: "Intl", fee: 0.07, aum: "75B", risk: 5, desc: "8000+ actions hors-US." },
  { t: "AGG", n: "iShares Bonds", cat: "Bonds", fee: 0.03, aum: "105B", risk: 3, desc: "Bonds US." },
  { t: "BIL", n: "T-Bills 1-3M", cat: "Cash", fee: 0.14, aum: "35B", risk: 1, desc: "~4% sans risque." },
  { t: "VWO", n: "Vanguard Emerging", cat: "EM", fee: 0.08, aum: "82B", risk: 6, desc: "Chine, Inde, Brésil." },
  { t: "TLT", n: "T-Bonds 30Y", cat: "Bonds LT", fee: 0.15, aum: "50B", risk: 5, desc: "Bonds long terme." },
];

const CRYPTOS = [
  { sym: "BTCUSDT", id: "BTC", n: "Bitcoin", risk: 7, cats: ["ETF spot $87B inflows", "Réserve stratégique US", "Halving 2024 effet"], risks: ["Volatilité géopolitique", "Réglementation"] },
  { sym: "ETHUSDT", id: "ETH", n: "Ethereum", risk: 8, cats: ["ETF spot approuvé", "Layer-2 scaling", "DeFi anchor"], risks: ["Concurrence Solana", "Fusaka upgrade"] },
  { sym: "SOLUSDT", id: "SOL", n: "Solana", risk: 9, cats: ["Alpenglow upgrade", "ETF SOL en attente", "Revenue $2.85B"], risks: ["Pannes réseau", "Très volatil"] },
  { sym: "XRPUSDT", id: "XRP", n: "XRP", risk: 8, cats: ["ETF XRP spot approuvé", "SEC case résolu", "Adoption bancaire"], risks: ["Concentration Ripple", "Concurrence stablecoins"] },
  { sym: "ZECUSDT", id: "ZEC", n: "Zcash", risk: 10, cats: ["Privacy theme Grayscale", "Hedge anti-débasement", "Petite cap"], risks: ["Banni dans certaines juridictions", "Très spéculatif"] },
];

const PROFILES = {
  conservative: {
    label: "Prudent",
    icon: "🛡",
    desc: "Préservation du capital",
    alloc: [
      { e: "GLD", pct: 25, c: C.xfer, why: "Hedge inflation & géopo" },
      { e: "VOO", pct: 20, c: C.hold, why: "Cœur diversifié" },
      { e: "BIL", pct: 25, c: C.buy, why: "Sécurité ~4%" },
      { e: "SCHD", pct: 15, c: C.buy, why: "Revenu passif" },
      { e: "SLV", pct: 5, c: "#94A3B8", why: "Diversification métaux" },
      { e: "XLV", pct: 10, c: "#06B6D4", why: "Défensif" },
    ],
  },
  balanced: {
    label: "Équilibré",
    icon: "⚖",
    desc: "Croissance modérée + protection",
    alloc: [
      { e: "VOO", pct: 25, c: C.hold, why: "Cœur de portefeuille" },
      { e: "GLD", pct: 15, c: C.xfer, why: "Hedge structurel" },
      { e: "QQQ", pct: 15, c: C.buy, why: "Croissance IA" },
      { e: "SMH", pct: 10, c: C.buy, why: "TSMC + NVDA" },
      { e: "SLV", pct: 8, c: "#94A3B8", why: "Déficit structurel" },
      { e: "IBIT", pct: 7, c: "#F7931A", why: "Or numérique" },
      { e: "VXUS", pct: 10, c: "#10B981", why: "Diversification géo" },
      { e: "BIL", pct: 10, c: C.buy, why: "Dry powder" },
    ],
  },
  growth: {
    label: "Croissance",
    icon: "🚀",
    desc: "Maximiser la performance",
    alloc: [
      { e: "QQQ", pct: 25, c: C.buy, why: "Mag7 + IA" },
      { e: "SMH", pct: 15, c: C.buy, why: "TSMC, NVDA dominants" },
      { e: "GLD", pct: 10, c: C.xfer, why: "Hedge minimal" },
      { e: "IBIT", pct: 10, c: "#F7931A", why: "Crypto inst." },
      { e: "ETH", pct: 5, c: "#627EEA", why: "Layer 1 dominant" },
      { e: "SOL", pct: 5, c: "#9945FF", why: "Momentum" },
      { e: "VUG", pct: 15, c: C.buy, why: "Actions growth" },
      { e: "GDX", pct: 5, c: C.xfer, why: "Levier or" },
      { e: "VWO", pct: 5, c: "#10B981", why: "Inde/Chine" },
      { e: "BIL", pct: 5, c: C.buy, why: "Liquidités" },
    ],
  },
  aggressive: {
    label: "Agressif",
    icon: "⚡",
    desc: "Conviction maximale",
    alloc: [
      { e: "SMH", pct: 25, c: C.buy, why: "Leader IA" },
      { e: "IBIT", pct: 15, c: "#F7931A", why: "Adoption inst." },
      { e: "SOL", pct: 10, c: "#9945FF", why: "Momentum" },
      { e: "XRP", pct: 8, c: "#23292F", why: "Post-SEC +400%" },
      { e: "ETH", pct: 7, c: "#627EEA", why: "DeFi anchor" },
      { e: "QQQ", pct: 15, c: C.buy, why: "Tech beta" },
      { e: "GDX", pct: 10, c: C.xfer, why: "Levier or" },
      { e: "ZEC", pct: 3, c: "#F4B728", why: "Privacy theme" },
      { e: "SLV", pct: 5, c: "#94A3B8", why: "Déficit" },
      { e: "XLE", pct: 2, c: C.sell, why: "Tactique géopo" },
    ],
  },
};

// ═══════════════════════════════════════════════════════════════════════════════
// UI HELPERS
// ═══════════════════════════════════════════════════════════════════════════════
function Spinner(props) {
  const size = props.size || 14;
  const color = props.color || "#fff";
  const style = {
    display: "inline-block",
    width: size,
    height: size,
    border: "2px solid " + color + "40",
    borderTopColor: color,
    borderRadius: "50%",
    animation: "spin 0.8s linear infinite",
  };
  return <span style={style} />;
}

function Bar(props) {
  const height = props.height || 4;
  const outer = { height, background: C.border, borderRadius: 2, overflow: "hidden" };
  const inner = {
    height: "100%",
    width: Math.min(100, Math.abs(props.pct)) + "%",
    background: props.color,
    borderRadius: 2,
  };
  return (
    <div style={outer}>
      <div style={inner} />
    </div>
  );
}

function Chip(props) {
  const style = {
    fontSize: 11,
    padding: "2px 8px",
    borderRadius: 4,
    background: props.color + "22",
    color: props.color,
    border: "1px solid " + props.color + "44",
    fontFamily: "monospace",
    fontWeight: 700,
  };
  return <span style={style}>{props.children}</span>;
}

function ActionBadge(props) {
  const a = actColor(props.actionKey);
  const style = {
    padding: "3px 10px",
    borderRadius: 6,
    fontSize: 11,
    fontWeight: 700,
    fontFamily: "monospace",
    background: a.bg,
    border: "1px solid " + a.bd,
    color: a.color,
  };
  return <span style={style}>{props.signal || a.fr}</span>;
}

function StatusPill(props) {
  const color = props.err
    ? C.sell
    : props.loading
    ? C.xfer
    : props.at
    ? props.color
    : C.muted;
  const style = {
    padding: "3px 9px",
    borderRadius: 20,
    fontSize: 10,
    fontWeight: 600,
    fontFamily: "monospace",
    color,
    background: color + "18",
    border: "1px solid " + color + "44",
    display: "inline-flex",
    alignItems: "center",
    gap: 5,
  };
  let icon;
  if (props.loading) {
    icon = <Spinner size={10} color={C.xfer} />;
  } else if (props.err) {
    icon = "✕";
  } else if (props.at) {
    icon = "●";
  } else {
    icon = "○";
  }
  return (
    <span style={style}>
      {icon} {props.icon} {props.label}
      {props.at && !props.loading && !props.err ? " · " + ageString(props.at) : ""}
      {props.err ? " · ERR" : ""}
    </span>
  );
}

function ActionBtn(props) {
  const isDisabled = props.loading || props.disabled;
  const color = props.color;
  const style = {
    padding: "8px 14px",
    borderRadius: 8,
    border: "1.5px solid " + (props.active ? color : color + "55"),
    cursor: isDisabled ? "not-allowed" : "pointer",
    background: props.active ? color + "25" : "transparent",
    color: props.disabled ? C.muted : color,
    fontSize: 12,
    fontWeight: 700,
    display: "flex",
    alignItems: "center",
    gap: 7,
    fontFamily: "'Segoe UI',sans-serif",
    opacity: isDisabled ? 0.7 : 1,
    transition: "all 0.15s",
    boxShadow: props.active ? "0 0 12px " + color + "40" : "none",
    whiteSpace: "nowrap",
  };
  const iconNode = props.loading ? (
    <Spinner size={12} color={color} />
  ) : (
    <span style={{ fontSize: 14 }}>{props.icon}</span>
  );
  return (
    <button onClick={props.onClick} disabled={isDisabled} style={style}>
      {iconNode}
      {props.label}
    </button>
  );
}

// ═══════════════════════════════════════════════════════════════════════════════
// SECTORS TAB
// ═══════════════════════════════════════════════════════════════════════════════
function SectorsTab(props) {
  const { stocks, aiData, filter, setFilter } = props;
  const [expanded, setExpanded] = useState(null);

  const groups = useMemo(() => {
    return Object.keys(SECTOR_META).map((sid) => {
      const meta = SECTOR_META[sid];
      const items = STOCKS.filter((s) => s.s === sid).map((s) => {
        return Object.assign({}, s, { d: stocks[s.t] });
      });
      const valid = items.filter((i) => i.d && i.d.price);
      const avg = valid.length
        ? valid.reduce((a, i) => a + i.d.changePct, 0) / valid.length
        : null;
      const ai = aiData && aiData.sectors && aiData.sectors[sid];
      let action = "HOLD";
      let confidence = 5;
      let signal = "N/D";
      let reason = "Chargement…";
      let risk = "";
      let fromAI = false;
      if (ai) {
        action = ai.action;
        confidence = ai.confidence;
        signal = ai.signal;
        reason = ai.reason;
        risk = ai.risk;
        fromAI = true;
      } else if (avg !== null) {
        action = actionFromChange(avg);
        confidence = Math.min(9, 5 + Math.round(Math.abs(avg)));
        if (avg >= 2) {
          signal = "FORT ACHAT";
        } else if (avg >= 0.5) {
          signal = "ACHAT";
        } else if (avg <= -2) {
          signal = "FORT VENTE";
        } else if (avg <= -0.5) {
          signal = "PIVOTER";
        } else {
          signal = "NEUTRE";
        }
        reason = "Moyenne sectorielle " + (avg >= 0 ? "+" : "") + avg.toFixed(2) + "% aujourd'hui.";
      }
      return { sid, meta, items, avg, action, confidence, signal, reason, risk, fromAI };
    });
  }, [stocks, aiData]);

  const filtered = filter === "ALL" ? groups : groups.filter((g) => g.action === filter);

  const counts = useMemo(() => {
    const o = {};
    ACTIONS.forEach((a) => {
      o[a.key] = groups.filter((g) => g.action === a.key).length;
    });
    return o;
  }, [groups]);

  const filterRow = (
    <div style={{ display: "flex", gap: 8, flexWrap: "wrap", padding: "4px 0 16px", alignItems: "center" }}>
      <span style={{ fontSize: 10, color: C.muted, fontFamily: "monospace", letterSpacing: "0.1em" }}>FILTRER :</span>
      {[{ key: "ALL", fr: "Tous", icon: "◉", color: C.txt2 }].concat(ACTIONS).map((a) => {
        const count = a.key === "ALL" ? groups.length : counts[a.key] || 0;
        const on = filter === a.key;
        const bcol = a.color || C.txt2;
        const btnStyle = {
          padding: "5px 14px",
          borderRadius: 6,
          border: "1px solid " + (on ? bcol : C.border),
          background: on ? bcol + "18" : "transparent",
          color: on ? bcol : C.muted,
          fontSize: 12,
          fontWeight: on ? 700 : 400,
          cursor: "pointer",
          fontFamily: "monospace",
          display: "flex",
          alignItems: "center",
          gap: 6,
        };
        const countStyle = {
          background: on ? bcol + "33" : C.border,
          borderRadius: 10,
          padding: "1px 7px",
          fontSize: 10,
          fontWeight: 700,
        };
        return (
          <button key={a.key} onClick={() => setFilter(a.key)} style={btnStyle}>
            <span style={{ fontSize: 10 }}>{a.icon}</span>
            {a.fr}
            <span style={countStyle}>{count}</span>
          </button>
        );
      })}
    </div>
  );

  return (
    <div>
      {filterRow}
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill,minmax(420px,1fr))", gap: 14 }}>
        {filtered.map((g) => {
          const a = actColor(g.action);
          const exp = expanded === g.sid;
          const cardStyle = {
            background: C.card,
            border: "1px solid " + (exp ? a.color + "55" : C.border),
            borderRadius: 12,
            overflow: "hidden",
            boxShadow: exp ? "0 0 24px " + a.color + "18" : "none",
          };
          return (
            <div key={g.sid} style={cardStyle}>
              <div style={{ padding: "13px 16px 10px", borderBottom: "1px solid " + C.border, display: "flex", alignItems: "flex-start", justifyContent: "space-between" }}>
                <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                  <span style={{ fontSize: 24 }}>{g.meta.icon}</span>
                  <div>
                    <div style={{ fontSize: 14, fontWeight: 700 }}>{g.meta.name}</div>
                    {g.meta.sub ? <div style={{ fontSize: 10, color: C.muted, fontFamily: "monospace" }}>{g.meta.sub}</div> : null}
                  </div>
                </div>
                <div style={{ display: "flex", alignItems: "center", gap: 7 }}>
                  {g.fromAI ? <Chip color={C.ai}>✦ IA</Chip> : null}
                  {g.avg !== null ? (
                    <span style={{ padding: "2px 8px", borderRadius: 5, fontSize: 11, fontFamily: "monospace", fontWeight: 700, background: g.avg >= 0 ? C.up + "22" : C.dn + "22", color: g.avg >= 0 ? C.up : C.dn }}>
                      {g.avg >= 0 ? "+" : ""}
                      {g.avg.toFixed(2)}%
                    </span>
                  ) : null}
                  <ActionBadge actionKey={g.action} signal={g.signal} />
                  <button onClick={() => setExpanded(exp ? null : g.sid)} style={{ background: C.border, border: "none", borderRadius: 6, color: C.txt2, width: 28, height: 28, cursor: "pointer", fontSize: 12 }}>
                    {exp ? "▲" : "▼"}
                  </button>
                </div>
              </div>
              <div style={{ height: 3 }}>
                <Bar pct={g.confidence * 10} color={g.meta.accent} />
              </div>
              <div style={{ padding: "11px 16px 8px", display: "flex", flexWrap: "wrap", gap: 6 }}>
                {g.items.map((item, i) => {
                  const d = item.d;
                  if (!d || !d.price) {
                    return (
                      <span key={i} style={{ fontSize: 10, padding: "3px 8px", borderRadius: 6, background: C.border + "44", color: C.muted, fontFamily: "monospace" }}>
                        {item.t} —
                      </span>
                    );
                  }
                  const pos = d.changePct >= 0;
                  return (
                    <div key={i} title={item.n + ": $" + d.price.toFixed(2)} style={{ display: "flex", alignItems: "center", gap: 5, padding: "3px 9px", fontSize: 11, borderRadius: 6, background: pos ? C.up + "10" : C.dn + "10", border: "1px solid " + (pos ? C.up + "25" : C.dn + "25") }}>
                      <span style={{ color: C.muted, fontSize: 10, fontFamily: "monospace" }}>{item.t}</span>
                      <span style={{ color: C.txt2, fontWeight: 600 }}>${d.price.toFixed(2)}</span>
                      <span style={{ color: pos ? C.up : C.dn, fontWeight: 700, fontFamily: "monospace" }}>
                        {pos ? "+" : ""}
                        {d.changePct.toFixed(2)}%
                      </span>
                    </div>
                  );
                })}
              </div>
              {exp ? (
                <div style={{ padding: "0 16px 12px", borderTop: "1px solid " + C.border, marginTop: 4 }}>
                  <div style={{ padding: "11px 14px", background: a.color + "0C", borderLeft: "3px solid " + a.color, borderRadius: "0 6px 6px 0", marginTop: 10 }}>
                    <div style={{ fontSize: 10, color: a.color, fontFamily: "monospace", marginBottom: 5, letterSpacing: "0.1em" }}>
                      {g.fromAI ? "✦ ANALYSE IA" : "◆ RÈGLE AUTOMATIQUE"}
                    </div>
                    <div style={{ fontSize: 13, lineHeight: 1.65, color: C.txt }}>{g.reason}</div>
                  </div>
                  {g.risk ? (
                    <div style={{ padding: "9px 14px", background: "#2A1010", border: "1px solid " + C.sell + "22", borderRadius: 6, marginTop: 8 }}>
                      <div style={{ fontSize: 10, color: C.sell, fontFamily: "monospace", marginBottom: 3, letterSpacing: "0.1em" }}>⚠ RISQUE</div>
                      <div style={{ fontSize: 12, color: "#FCA5A5" }}>{g.risk}</div>
                    </div>
                  ) : null}
                </div>
              ) : null}
              <div style={{ padding: "9px 16px 13px", borderTop: "1px solid " + C.border }}>
                <div style={{ fontSize: 9, color: C.muted, fontFamily: "monospace", marginBottom: 7, letterSpacing: "0.1em" }}>DÉCISION ↓</div>
                <div style={{ display: "grid", gridTemplateColumns: "repeat(4,1fr)", gap: 6 }}>
                  {ACTIONS.map((act) => {
                    const on = act.key === g.action;
                    const cellStyle = {
                      padding: "8px 4px",
                      borderRadius: 8,
                      textAlign: "center",
                      border: "1.5px solid " + (on ? act.color : C.border),
                      background: on ? act.bg : "transparent",
                      opacity: on ? 1 : 0.4,
                      boxShadow: on ? "0 2px 12px " + act.color + "25" : "none",
                    };
                    return (
                      <div key={act.key} style={cellStyle}>
                        <div style={{ fontSize: 16, color: act.color, fontWeight: 700 }}>{act.icon}</div>
                        <div style={{ fontSize: 10, color: act.color, fontFamily: "monospace", fontWeight: on ? 700 : 400 }}>{act.fr}</div>
                      </div>
                    );
                  })}
                </div>
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════════════════
// ETFs TAB
// ═══════════════════════════════════════════════════════════════════════════════
function ETFsTab(props) {
  const { stocks } = props;
  const [sort, setSort] = useState("change");

  const rows = useMemo(() => {
    const arr = ETF_LIST.map((e) => {
      const d = stocks[e.t];
      const pct = d ? d.changePct : null;
      const action = pct != null ? actionFromChange(pct) : "HOLD";
      return Object.assign({}, e, { d, action, pct });
    });
    arr.sort((a, b) => {
      if (sort === "change") {
        return (b.pct == null ? -999 : b.pct) - (a.pct == null ? -999 : a.pct);
      }
      if (sort === "fee") {
        return a.fee - b.fee;
      }
      return a.risk - b.risk;
    });
    return arr;
  }, [stocks, sort]);

  const headers = ["ETF", "Catégorie", "Prix", "Chg%", "Frais", "AUM", "Risque", "Décision", "Note"];

  return (
    <div>
      <div style={{ background: C.card, border: "1px solid " + C.border, borderRadius: 12, padding: "14px 18px", marginBottom: 14, display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 10 }}>
        <div>
          <div style={{ fontSize: 16, fontWeight: 700 }}>📦 ETFs — {ETF_LIST.length} fonds</div>
          <div style={{ fontSize: 11, color: C.muted, marginTop: 4 }}>Prix via Anthropic · cache 24h</div>
        </div>
        <div style={{ display: "flex", gap: 6 }}>
          {[{ k: "change", l: "Performance" }, { k: "fee", l: "Frais" }, { k: "risk", l: "Risque" }].map((o) => {
            const on = sort === o.k;
            const btnStyle = {
              padding: "5px 12px",
              borderRadius: 6,
              fontSize: 11,
              cursor: "pointer",
              fontFamily: "monospace",
              border: "1px solid " + (on ? C.accent : C.border),
              background: on ? C.accent + "22" : "transparent",
              color: on ? C.accent : C.muted,
            };
            return (
              <button key={o.k} onClick={() => setSort(o.k)} style={btnStyle}>
                {o.l}
              </button>
            );
          })}
        </div>
      </div>
      <div style={{ background: C.card, borderRadius: 12, border: "1px solid " + C.border, overflow: "hidden" }}>
        <div style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse" }}>
            <thead>
              <tr style={{ background: C.surface }}>
                {headers.map((h, i) => {
                  const thStyle = {
                    padding: "9px 12px",
                    textAlign: "left",
                    fontSize: 10,
                    color: C.muted,
                    fontFamily: "monospace",
                    letterSpacing: "0.08em",
                    whiteSpace: "nowrap",
                  };
                  return (
                    <th key={i} style={thStyle}>
                      {h}
                    </th>
                  );
                })}
              </tr>
            </thead>
            <tbody>
              {rows.map((e, i) => {
                const a = actColor(e.action);
                const rowStyle = {
                  borderTop: "1px solid " + C.border,
                  background: i % 2 === 0 ? "transparent" : C.surface + "88",
                };
                return (
                  <tr key={e.t} style={rowStyle}>
                    <td style={{ padding: "9px 12px" }}>
                      <div style={{ fontWeight: 700, fontSize: 13 }}>{e.t}</div>
                      <div style={{ fontSize: 10, color: C.muted }}>{e.n}</div>
                    </td>
                    <td style={{ padding: "9px 12px" }}>
                      <span style={{ fontSize: 10, padding: "2px 7px", borderRadius: 4, background: C.border, color: C.txt2, fontFamily: "monospace" }}>{e.cat}</span>
                    </td>
                    <td style={{ padding: "9px 12px", fontFamily: "monospace", fontSize: 12 }}>
                      {e.d && e.d.price ? "$" + e.d.price.toFixed(2) : <span style={{ color: C.muted }}>—</span>}
                    </td>
                    <td style={{ padding: "9px 12px" }}>
                      {e.pct != null ? (
                        <span style={{ fontFamily: "monospace", fontWeight: 700, fontSize: 12, color: e.pct >= 0 ? C.up : C.dn }}>
                          {e.pct >= 0 ? "+" : ""}
                          {e.pct.toFixed(2)}%
                        </span>
                      ) : (
                        <span style={{ color: C.muted }}>—</span>
                      )}
                    </td>
                    <td style={{ padding: "9px 12px", fontFamily: "monospace", fontSize: 11, color: C.txt2 }}>{e.fee.toFixed(2)}%</td>
                    <td style={{ padding: "9px 12px", fontFamily: "monospace", fontSize: 11, color: C.txt2 }}>${e.aum}</td>
                    <td style={{ padding: "9px 12px" }}>
                      <div style={{ display: "flex", gap: 2 }}>
                        {Array.from({ length: 10 }).map((_, j) => {
                          const filled = j < e.risk;
                          const fillColor = e.risk >= 8 ? C.sell : e.risk >= 6 ? C.xfer : C.up;
                          return <div key={j} style={{ width: 5, height: 12, borderRadius: 1, background: filled ? fillColor : C.border }} />;
                        })}
                      </div>
                    </td>
                    <td style={{ padding: "9px 12px" }}>
                      <div style={{ display: "flex", alignItems: "center", gap: 5, padding: "3px 9px", borderRadius: 6, background: a.bg, border: "1px solid " + a.bd, width: "fit-content" }}>
                        <span style={{ color: a.color, fontWeight: 700 }}>{a.icon}</span>
                        <span style={{ fontSize: 11, color: a.color, fontWeight: 600, fontFamily: "monospace" }}>{a.fr}</span>
                      </div>
                    </td>
                    <td style={{ padding: "9px 12px", fontSize: 11, color: C.txt2, maxWidth: 240, lineHeight: 1.4 }}>{e.desc}</td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════════════════
// CRYPTO TAB
// ═══════════════════════════════════════════════════════════════════════════════
function CryptoTab(props) {
  const { crypto } = props;
  const [sel, setSel] = useState(0);
  const c = CRYPTOS[sel];
  const d = crypto[c.sym];
  const action = d ? actionFromChange(d.change24h) : "HOLD";
  const a = actColor(action);

  return (
    <div>
      <div style={{ background: C.card, border: "1px solid " + C.border, borderRadius: 12, padding: "14px 18px", marginBottom: 14 }}>
        <div style={{ fontSize: 16, fontWeight: 700 }}>₿ Cryptomonnaies — LIVE</div>
        <div style={{ fontSize: 11, color: C.muted, marginTop: 4 }}>Source : Anthropic · cache 24h</div>
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(200px,1fr))", gap: 12, marginBottom: 18 }}>
        {CRYPTOS.map((cr, i) => {
          const cd = crypto[cr.sym];
          const ak = cd ? actionFromChange(cd.change24h) : "HOLD";
          const av = actColor(ak);
          const on = i === sel;
          const cardStyle = {
            padding: "13px 15px",
            background: C.card,
            borderRadius: 12,
            cursor: "pointer",
            border: "1.5px solid " + (on ? av.color : C.border),
            boxShadow: on ? "0 0 20px " + av.color + "25" : "none",
          };
          const fmt = cd && cd.price > 10 ? 2 : 4;
          return (
            <div key={cr.sym} onClick={() => setSel(i)} style={cardStyle}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 7 }}>
                <div>
                  <div style={{ fontSize: 14, fontWeight: 700 }}>{cr.id}</div>
                  <div style={{ fontSize: 10, color: C.muted, fontFamily: "monospace" }}>{cr.n}</div>
                </div>
                <ActionBadge actionKey={ak} signal={ak} />
              </div>
              <div style={{ fontFamily: "monospace", fontSize: 18, fontWeight: 700, marginBottom: 4 }}>
                {cd ? "$" + cd.price.toLocaleString(undefined, { maximumFractionDigits: fmt }) : <span style={{ color: C.muted, fontSize: 13 }}>—</span>}
              </div>
              {cd ? (
                <div style={{ fontSize: 10, fontFamily: "monospace", color: cd.change24h >= 0 ? C.up : C.dn }}>
                  24h: {cd.change24h >= 0 ? "+" : ""}
                  {cd.change24h.toFixed(2)}%
                </div>
              ) : null}
            </div>
          );
        })}
      </div>
      <div style={{ background: C.card, border: "1px solid " + a.color + "55", borderRadius: 12, padding: 18 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 14, flexWrap: "wrap", gap: 10 }}>
          <div style={{ fontSize: 22, fontWeight: 700 }}>{c.id} — {c.n}</div>
          <div style={{ padding: "8px 14px", borderRadius: 8, background: a.bg, border: "1.5px solid " + a.color, color: a.color, fontWeight: 700, fontFamily: "monospace" }}>
            {a.icon} {a.fr}
          </div>
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(250px,1fr))", gap: 14 }}>
          <div style={{ background: C.surface, borderRadius: 10, padding: 14 }}>
            <div style={{ fontSize: 10, color: C.muted, fontFamily: "monospace", marginBottom: 10, letterSpacing: "0.1em" }}>📊 DONNÉES LIVE</div>
            {d ? (
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
                {(() => {
                  const fmt = d.price > 10 ? 2 : 4;
                  const rows = [
                    { l: "PRIX", v: "$" + d.price.toLocaleString(undefined, { maximumFractionDigits: fmt }) },
                    { l: "VAR 24H", v: (d.change24h >= 0 ? "+" : "") + d.change24h.toFixed(2) + "%", c: d.change24h >= 0 ? C.up : C.dn },
                    { l: "HAUT 24H", v: d.high24h ? "$" + d.high24h.toLocaleString(undefined, { maximumFractionDigits: 2 }) : "—" },
                    { l: "BAS 24H", v: d.low24h ? "$" + d.low24h.toLocaleString(undefined, { maximumFractionDigits: 2 }) : "—" },
                    { l: "VOL", v: d.quoteVol ? "$" + (d.quoteVol / 1e9).toFixed(2) + "B" : "—" },
                  ];
                  return rows.map((row, j) => (
                    <div key={j}>
                      <div style={{ fontSize: 9, color: C.muted }}>{row.l}</div>
                      <div style={{ fontSize: 15, fontWeight: 700, fontFamily: "monospace", color: row.c || C.txt }}>{row.v}</div>
                    </div>
                  ));
                })()}
              </div>
            ) : (
              <div style={{ color: C.muted, fontSize: 12, fontStyle: "italic" }}>Aucune donnée — cliquez sur Marché pour charger.</div>
            )}
            <div style={{ marginTop: 12, display: "flex", alignItems: "center", gap: 5 }}>
              <span style={{ fontSize: 10, color: C.muted, fontFamily: "monospace" }}>Risque :</span>
              {Array.from({ length: 10 }).map((_, j) => {
                const filled = j < c.risk;
                const fc = c.risk >= 8 ? C.sell : C.xfer;
                return <div key={j} style={{ width: 8, height: 8, borderRadius: 2, background: filled ? fc : C.border }} />;
              })}
              <span style={{ fontSize: 11, color: c.risk >= 8 ? C.sell : C.xfer, fontFamily: "monospace", fontWeight: 700 }}>{c.risk}/10</span>
            </div>
          </div>
          <div style={{ background: C.up + "08", border: "1px solid " + C.up + "22", borderRadius: 10, padding: 14 }}>
            <div style={{ fontSize: 10, color: C.up, fontFamily: "monospace", marginBottom: 9, letterSpacing: "0.1em" }}>🚀 CATALYSEURS</div>
            {c.cats.map((x, i) => (
              <div key={i} style={{ fontSize: 12, marginBottom: 5, paddingLeft: 14, position: "relative", lineHeight: 1.4 }}>
                <span style={{ position: "absolute", left: 0, color: C.up }}>✓</span>
                {x}
              </div>
            ))}
          </div>
          <div style={{ background: C.sell + "08", border: "1px solid " + C.sell + "22", borderRadius: 10, padding: 14 }}>
            <div style={{ fontSize: 10, color: C.sell, fontFamily: "monospace", marginBottom: 9, letterSpacing: "0.1em" }}>⚠ RISQUES</div>
            {c.risks.map((x, i) => (
              <div key={i} style={{ fontSize: 12, marginBottom: 5, paddingLeft: 14, position: "relative", lineHeight: 1.4 }}>
                <span style={{ position: "absolute", left: 0, color: C.sell }}>×</span>
                {x}
              </div>
            ))}
          </div>
        </div>
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════════════════
// SENTIMENT TAB — indicateurs 100% dynamiques (issus de aiData.indicators)
// ═══════════════════════════════════════════════════════════════════════════════
const IND_META = [
  { k: "vix", l: "VIX (Volatilité)", max: 50 },
  { k: "fearGreed", l: "Fear & Greed", max: 100 },
  { k: "dxy", l: "Dollar (DXY)", max: 120 },
  { k: "yield10y", l: "10Y Yield", max: 8 },
  { k: "btcDom", l: "BTC Dominance", max: 100 },
  { k: "aiSent", l: "Sentiment IA", max: 100 },
];

function SentimentTab(props) {
  const { aiData, news, newsAge } = props;
  const [fi, setFi] = useState("ALL");
  const filtered = fi === "ALL" ? news : news.filter((n) => n.impact === fi);

  const indicators = (aiData && aiData.indicators) || null;

  const IND = IND_META.map((m) => {
    const d = indicators ? indicators[m.k] : null;
    if (!d || typeof d.v !== "number") {
      return { l: m.l, max: m.max, v: null, s: "—", c: C.muted, desc: "Aucune donnée" };
    }
    return {
      l: m.l,
      max: m.max,
      v: d.v,
      s: d.status || "—",
      c: toneColor(d.tone),
      desc: d.desc || "",
    };
  });

  const filters = [
    { k: "ALL", l: "Toutes", c: C.txt2 },
    { k: "POSITIVE", l: "Positives", c: C.up },
    { k: "NEUTRAL", l: "Neutres", c: C.hold },
    { k: "NEGATIVE", l: "Négatives", c: C.sell },
  ];

  return (
    <div>
      <div style={{ background: C.card, border: "1px solid " + C.border, borderRadius: 12, padding: "14px 18px", marginBottom: 14 }}>
        <div style={{ fontSize: 16, fontWeight: 700, marginBottom: 13 }}>🌡 Indicateurs de Sentiment</div>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(210px,1fr))", gap: 11 }}>
          {IND.map((s, i) => {
            const hasData = s.v !== null;
            return (
              <div key={i} style={{ background: C.surface, borderRadius: 8, padding: "11px 13px", border: "1px solid " + C.border }}>
                <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 5 }}>
                  <span style={{ fontSize: 11, color: C.muted, fontFamily: "monospace" }}>{s.l}</span>
                  <span style={{ fontSize: 10, padding: "1px 7px", borderRadius: 10, background: s.c + "22", color: s.c, fontFamily: "monospace", fontWeight: 700 }}>{s.s}</span>
                </div>
                <div style={{ fontSize: 22, fontWeight: 700, fontFamily: "monospace", color: s.c, marginBottom: 5 }}>
                  {hasData ? s.v : "—"}
                </div>
                <Bar pct={hasData ? (s.v / s.max) * 100 : 0} color={s.c} />
                <div style={{ fontSize: 10, color: C.muted, marginTop: 5 }}>{s.desc}</div>
              </div>
            );
          })}
        </div>
      </div>

      {aiData && aiData.globalSummary ? (
        <div style={{ background: "linear-gradient(135deg,#1A0B2E,#0F1729)", border: "1px solid " + C.ai + "44", borderRadius: 12, padding: "14px 18px", marginBottom: 14 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 9, marginBottom: 10, flexWrap: "wrap" }}>
            <span style={{ fontSize: 18 }}>✦</span>
            <div style={{ flex: 1 }}>
              <div style={{ fontSize: 14, fontWeight: 700, color: C.ai }}>SYNTHÈSE IA — {aiData.overallSentiment}</div>
            </div>
          </div>
          <div style={{ fontSize: 13, lineHeight: 1.7, color: C.txt2 }}>{aiData.globalSummary}</div>
        </div>
      ) : (
        <div style={{ background: C.card, border: "1px solid " + C.xfer + "33", borderRadius: 12, padding: "14px 18px", marginBottom: 14 }}>
          <div style={{ fontSize: 13, lineHeight: 1.7, color: C.txt2, textAlign: "center", padding: "8px 0" }}>
            Lancez une analyse IA pour générer la synthèse macro et les indicateurs.
          </div>
        </div>
      )}

      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 8, marginBottom: 11 }}>
        <div style={{ display: "flex", gap: 7, flexWrap: "wrap", alignItems: "center" }}>
          <span style={{ fontSize: 10, color: C.muted, fontFamily: "monospace" }}>FILTRER :</span>
          {filters.map((o) => {
            const on = fi === o.k;
            const btnStyle = {
              padding: "4px 13px",
              borderRadius: 6,
              fontSize: 11,
              cursor: "pointer",
              fontFamily: "monospace",
              border: "1px solid " + (on ? o.c : C.border),
              background: on ? o.c + "22" : "transparent",
              color: on ? o.c : C.muted,
            };
            return (
              <button key={o.k} onClick={() => setFi(o.k)} style={btnStyle}>
                {o.l}
              </button>
            );
          })}
        </div>
        {newsAge ? (
          <span style={{ fontSize: 10, color: C.up, fontFamily: "monospace", display: "flex", alignItems: "center", gap: 5 }}>
            <span style={{ width: 6, height: 6, borderRadius: "50%", background: C.up, display: "inline-block" }} />
            Récupérées {newsAge}
          </span>
        ) : null}
      </div>

      <div style={{ background: C.card, border: "1px solid " + C.border, borderRadius: 12, overflow: "hidden" }}>
        <div style={{ padding: "12px 18px", borderBottom: "1px solid " + C.border, background: C.surface, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
          <div style={{ fontSize: 14, fontWeight: 700 }}>📰 Actualités & Géopolitique</div>
          <div style={{ fontSize: 11, color: C.muted, fontFamily: "monospace" }}>{filtered.length} événements</div>
        </div>
        {news.length === 0 ? (
          <div style={{ padding: 28, textAlign: "center", color: C.muted, fontSize: 12, lineHeight: 1.6 }}>
            <div style={{ fontSize: 22, marginBottom: 8 }}>📰</div>
            <div>Aucune actualité chargée.</div>
            <div style={{ fontSize: 11, marginTop: 4 }}>Cliquez sur le bouton <span style={{ color: C.xfer, fontWeight: 700 }}>News</span> en haut pour récupérer 10 actualités via l'API.</div>
          </div>
        ) : filtered.length === 0 ? (
          <div style={{ padding: 24, textAlign: "center", color: C.muted, fontSize: 12 }}>Aucun événement pour ce filtre</div>
        ) : (
          filtered.map((n, i) => {
            const nc = n.impact === "POSITIVE" ? C.up : n.impact === "NEGATIVE" ? C.sell : C.hold;
            const rowStyle = {
              padding: "13px 18px",
              borderTop: i > 0 ? "1px solid " + C.border : "none",
              background: i % 2 === 0 ? "transparent" : C.surface + "44",
            };
            const impactLabel = n.impact === "POSITIVE" ? "↑ POS" : n.impact === "NEGATIVE" ? "↓ NEG" : "◆ NEU";
            return (
              <div key={i} style={rowStyle}>
                <div style={{ display: "flex", gap: 11, alignItems: "flex-start" }}>
                  <div style={{ minWidth: 76 }}>
                    <div style={{ fontSize: 10, color: C.muted, fontFamily: "monospace" }}>{n.date}</div>
                    <span style={{ display: "inline-block", marginTop: 4, fontSize: 10, padding: "2px 7px", borderRadius: 4, background: nc + "22", color: nc, fontWeight: 700, fontFamily: "monospace" }}>
                      {impactLabel}
                    </span>
                  </div>
                  <div style={{ flex: 1 }}>
                    <div style={{ fontSize: 13, fontWeight: 600, marginBottom: 5, lineHeight: 1.4 }}>{n.title}</div>
                    <div style={{ fontSize: 12, color: C.txt2, lineHeight: 1.5, marginBottom: 7 }}>{n.summary}</div>
                    <div style={{ display: "flex", gap: 5, flexWrap: "wrap", alignItems: "center" }}>
                      {(n.tags || []).map((t, j) => {
                        const tagColor = t.indexOf("↑") >= 0 ? C.up : t.indexOf("↓") >= 0 ? C.sell : C.hold;
                        return (
                          <span key={j} style={{ fontSize: 10, padding: "2px 7px", borderRadius: 4, fontFamily: "monospace", background: tagColor + "15", color: tagColor }}>
                            {t}
                          </span>
                        );
                      })}
                      <span style={{ fontSize: 9, color: C.muted, marginLeft: "auto", fontFamily: "monospace" }}>{n.src}</span>
                    </div>
                  </div>
                  <div style={{ textAlign: "center", minWidth: 28 }}>
                    <div style={{ fontSize: 9, color: C.muted, fontFamily: "monospace" }}>IMPACT</div>
                    <div style={{ fontSize: 14, fontWeight: 700, color: nc, fontFamily: "monospace" }}>{n.sev}</div>
                    <div style={{ fontSize: 9, color: C.muted, fontFamily: "monospace" }}>/10</div>
                  </div>
                </div>
              </div>
            );
          })
        )}
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════════════════
// ALLOCATOR TAB
// ═══════════════════════════════════════════════════════════════════════════════
function AllocatorTab(props) {
  const { stocks, crypto } = props;
  const [amount, setAmount] = useState(1000);
  const [profile, setProfile] = useState("balanced");
  const p = PROFILES[profile];

  function getLivePrice(t) {
    if (stocks[t] && stocks[t].price) {
      return stocks[t].price;
    }
    const cryptoSym = t + "USDT";
    if (crypto[cryptoSym] && crypto[cryptoSym].price) {
      return crypto[cryptoSym].price;
    }
    return null;
  }

  const quickAmounts = [100, 500, 1000, 5000, 10000, 50000];
  const headers = ["#", "Allocation", "Ticker", "Prix LIVE", "%", "Montant €", "Parts ≈", "Raison"];

  return (
    <div>
      <div style={{ background: C.card, border: "1px solid " + C.border, borderRadius: 12, padding: 18, marginBottom: 14 }}>
        <div style={{ fontSize: 16, fontWeight: 700, marginBottom: 3 }}>🎯 Allocateur — Prix LIVE</div>
        <div style={{ fontSize: 11, color: C.muted, marginTop: 4 }}>Stocks & Crypto via Anthropic</div>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 18, marginTop: 18 }}>
          <div>
            <label style={{ fontSize: 11, color: C.muted, fontFamily: "monospace", display: "block", marginBottom: 7 }}>💰 MONTANT (€)</label>
            <input
              type="number"
              value={amount}
              onChange={(e) => setAmount(Math.max(0, +e.target.value))}
              style={{ width: "100%", padding: "9px 13px", background: C.surface, border: "1px solid " + C.bright, borderRadius: 8, color: C.txt, fontSize: 18, fontWeight: 700, fontFamily: "monospace", outline: "none" }}
            />
            <div style={{ display: "flex", gap: 5, marginTop: 7, flexWrap: "wrap" }}>
              {quickAmounts.map((v) => {
                const on = amount === v;
                const btnStyle = {
                  padding: "3px 9px",
                  borderRadius: 5,
                  fontSize: 10,
                  cursor: "pointer",
                  fontFamily: "monospace",
                  border: "1px solid " + (on ? C.accent : C.border),
                  background: on ? C.accent + "22" : "transparent",
                  color: on ? C.accent : C.muted,
                };
                return (
                  <button key={v} onClick={() => setAmount(v)} style={btnStyle}>
                    {v.toLocaleString()}€
                  </button>
                );
              })}
            </div>
          </div>
          <div>
            <label style={{ fontSize: 11, color: C.muted, fontFamily: "monospace", display: "block", marginBottom: 7 }}>📊 PROFIL</label>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(4,1fr)", gap: 6 }}>
              {Object.entries(PROFILES).map(([k, pr]) => {
                const on = profile === k;
                const btnStyle = {
                  padding: "7px 4px",
                  borderRadius: 8,
                  textAlign: "center",
                  cursor: "pointer",
                  border: "1.5px solid " + (on ? C.accent : C.border),
                  background: on ? C.accent + "22" : "transparent",
                  color: on ? C.accent : C.txt2,
                };
                return (
                  <button key={k} onClick={() => setProfile(k)} style={btnStyle}>
                    <div style={{ fontSize: 18 }}>{pr.icon}</div>
                    <div style={{ fontSize: 10, fontWeight: on ? 700 : 400, marginTop: 2 }}>{pr.label}</div>
                  </button>
                );
              })}
            </div>
            <div style={{ fontSize: 11, color: C.muted, marginTop: 7, fontStyle: "italic" }}>→ {p.desc}</div>
          </div>
        </div>
      </div>
      <div style={{ background: C.card, border: "1px solid " + C.border, borderRadius: 12, padding: "13px 18px", marginBottom: 14 }}>
        <div style={{ fontSize: 10, color: C.muted, fontFamily: "monospace", marginBottom: 9 }}>RÉPARTITION — {amount.toLocaleString()}€</div>
        <div style={{ display: "flex", height: 32, borderRadius: 7, overflow: "hidden", gap: 1 }}>
          {p.alloc.map((a, i) => {
            const segStyle = {
              width: a.pct + "%",
              background: a.c,
              display: "flex",
              alignItems: "center",
              justifyContent: "center",
              fontSize: 10,
              fontWeight: 700,
              color: "#000",
            };
            return (
              <div key={i} title={a.e + ": " + a.pct + "%"} style={segStyle}>
                {a.pct >= 8 ? a.pct + "%" : ""}
              </div>
            );
          })}
        </div>
      </div>
      <div style={{ background: C.card, border: "1px solid " + C.border, borderRadius: 12, overflow: "hidden" }}>
        <div style={{ padding: "12px 18px", borderBottom: "1px solid " + C.border, background: C.surface, display: "flex", justifyContent: "space-between" }}>
          <div style={{ fontSize: 14, fontWeight: 700 }}>📋 Plan détaillé</div>
          <div style={{ fontSize: 11, color: C.muted, fontFamily: "monospace" }}>{p.alloc.length} positions</div>
        </div>
        <div style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse" }}>
            <thead>
              <tr style={{ background: C.surface }}>
                {headers.map((h, i) => (
                  <th key={i} style={{ padding: "9px 13px", textAlign: "left", fontSize: 10, color: C.muted, fontFamily: "monospace", letterSpacing: "0.08em", whiteSpace: "nowrap" }}>
                    {h}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {p.alloc.map((a, i) => {
                const eur = (amount * a.pct) / 100;
                const livePrice = getLivePrice(a.e);
                let shares = null;
                if (livePrice) {
                  const fmt = livePrice < 1 ? 0 : livePrice < 10 ? 3 : 4;
                  shares = ((eur * 1.08) / livePrice).toFixed(fmt);
                }
                const rowStyle = {
                  borderTop: "1px solid " + C.border,
                  background: i % 2 === 0 ? "transparent" : C.surface + "88",
                };
                const priceFmt = livePrice && livePrice > 100 ? 2 : 4;
                return (
                  <tr key={i} style={rowStyle}>
                    <td style={{ padding: "11px 13px", color: C.muted, fontFamily: "monospace", fontWeight: 700 }}>{i + 1}</td>
                    <td style={{ padding: "11px 13px" }}>
                      <div style={{ display: "flex", alignItems: "center", gap: 7 }}>
                        <div style={{ width: 5, height: 22, background: a.c, borderRadius: 2 }} />
                        <span style={{ fontSize: 13, fontWeight: 600 }}>{a.e}</span>
                      </div>
                    </td>
                    <td style={{ padding: "11px 13px" }}>
                      <span style={{ fontSize: 11, padding: "2px 9px", borderRadius: 5, background: a.c + "22", color: a.c, fontFamily: "monospace", fontWeight: 700, border: "1px solid " + a.c + "44" }}>{a.e}</span>
                    </td>
                    <td style={{ padding: "11px 13px", fontFamily: "monospace", fontSize: 12, color: livePrice ? C.yf : C.muted }}>
                      {livePrice ? "$" + livePrice.toFixed(priceFmt) : "—"}
                    </td>
                    <td style={{ padding: "11px 13px", fontFamily: "monospace", fontSize: 12, fontWeight: 700, color: a.c }}>{a.pct}%</td>
                    <td style={{ padding: "11px 13px", fontFamily: "monospace", fontSize: 14, fontWeight: 700 }}>{eur.toLocaleString("fr-FR", { maximumFractionDigits: 0 })}€</td>
                    <td style={{ padding: "11px 13px", fontFamily: "monospace", fontSize: 11, color: C.txt2 }}>{shares || "—"}</td>
                    <td style={{ padding: "11px 13px", fontSize: 11, color: C.txt2, maxWidth: 200, lineHeight: 1.4 }}>{a.why}</td>
                  </tr>
                );
              })}
            </tbody>
            <tfoot>
              <tr style={{ background: C.surface, borderTop: "2px solid " + C.accent }}>
                <td colSpan={4} style={{ padding: "11px 13px", fontSize: 12, fontWeight: 700, color: C.accent }}>TOTAL</td>
                <td style={{ padding: "11px 13px", fontSize: 13, fontWeight: 700, color: C.accent, fontFamily: "monospace" }}>100%</td>
                <td style={{ padding: "11px 13px", fontSize: 16, fontWeight: 700, color: C.accent, fontFamily: "monospace" }}>{amount.toLocaleString()}€</td>
                <td colSpan={2} />
              </tr>
            </tfoot>
          </table>
        </div>
      </div>
      <div style={{ marginTop: 11, padding: "10px 14px", background: "#100A08", borderRadius: 8, border: "1px solid #2A1A0E", fontSize: 11, color: "#92400E", fontFamily: "monospace", lineHeight: 1.6 }}>
        ⚠ Conversion EUR→USD approximative (×1.08). Ceci n'est pas un conseil financier.
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════════════════
// AI SUMMARY HEADER
// ═══════════════════════════════════════════════════════════════════════════════
function AISummary(props) {
  const { aiData } = props;
  if (!aiData) {
    return null;
  }
  return (
    <div style={{ background: "linear-gradient(135deg,#1A0B2E,#0F1729)", border: "1px solid " + C.ai + "44", borderRadius: 12, padding: "14px 18px", marginBottom: 14 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 10, flexWrap: "wrap", gap: 8 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 9 }}>
          <span style={{ fontSize: 16 }}>✦</span>
          <div>
            <div style={{ fontSize: 14, fontWeight: 700, color: C.ai }}>ANALYSE IA — {aiData.overallSentiment}</div>
            <div style={{ fontSize: 10, color: C.muted, fontFamily: "monospace", marginTop: 2 }}>
              sonnet-4 · {ageString(MEM.aiAt)}
            </div>
          </div>
        </div>
      </div>
      {aiData.globalSummary ? (
        <div style={{ fontSize: 12, color: C.txt2, lineHeight: 1.6, marginBottom: 10 }}>{aiData.globalSummary}</div>
      ) : null}
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
        {aiData.topOpportunities ? (
          <div style={{ background: C.up + "08", border: "1px solid " + C.up + "22", borderRadius: 8, padding: "10px 12px" }}>
            <div style={{ fontSize: 10, color: C.up, fontFamily: "monospace", marginBottom: 7, letterSpacing: "0.1em" }}>🚀 OPPORTUNITÉS</div>
            {aiData.topOpportunities.map((o, i) => (
              <div key={i} style={{ fontSize: 11, color: C.txt2, marginBottom: 4, lineHeight: 1.4 }}>
                <span style={{ color: C.up, fontWeight: 700, fontFamily: "monospace" }}>{o.asset}</span> — {o.reason}
              </div>
            ))}
          </div>
        ) : null}
        {aiData.topRisks ? (
          <div style={{ background: C.sell + "08", border: "1px solid " + C.sell + "22", borderRadius: 8, padding: "10px 12px" }}>
            <div style={{ fontSize: 10, color: C.sell, fontFamily: "monospace", marginBottom: 7, letterSpacing: "0.1em" }}>⚠ RISQUES</div>
            {aiData.topRisks.map((o, i) => (
              <div key={i} style={{ fontSize: 11, color: C.txt2, marginBottom: 4, lineHeight: 1.4 }}>
                <span style={{ color: C.sell, fontWeight: 700, fontFamily: "monospace" }}>{o.asset}</span> — {o.reason}
              </div>
            ))}
          </div>
        ) : null}
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════════════════
// WELCOME SCREEN
// ═══════════════════════════════════════════════════════════════════════════════
function WelcomeScreen() {
  const cards = [
    { color: C.yf, title: "📈 Marché", desc: "Stocks + ETFs + Crypto via Anthropic web_search." },
    { color: C.xfer, title: "📰 News", desc: "10 actus macro via Haiku 4.5." },
    { color: C.ai, title: "✦ Analyse IA", desc: "Synthèse + indicateurs via Sonnet 4. Nécessite marché + news." },
    { color: C.up, title: "🤖 Full Auto", desc: "Refresh complet 1×/jour tant que l'onglet reste ouvert." },
  ];
  return (
    <div style={{ background: C.card, border: "1px solid " + C.accent + "33", borderRadius: 12, padding: 24, marginBottom: 16 }}>
      <div style={{ display: "flex", alignItems: "center", gap: 14, marginBottom: 14, flexWrap: "wrap" }}>
        <div style={{ fontSize: 26 }}>📡</div>
        <div>
          <div style={{ fontSize: 16, fontWeight: 700, color: C.accent }}>Prêt à charger les données</div>
          <div style={{ fontSize: 11, color: C.muted, marginTop: 3 }}>
            Aucune donnée pré-chargée. Cliquez sur les boutons pour récupérer les infos via l'API Anthropic.
          </div>
        </div>
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(220px,1fr))", gap: 10 }}>
        {cards.map((c, i) => (
          <div key={i} style={{ padding: "10px 12px", background: c.color + "0E", border: "1px solid " + c.color + "33", borderRadius: 8, fontSize: 11, color: C.txt2, lineHeight: 1.5 }}>
            <div style={{ color: c.color, fontWeight: 700, marginBottom: 3 }}>{c.title}</div>
            {c.desc}
          </div>
        ))}
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════════════════
// MAIN APP
// ═══════════════════════════════════════════════════════════════════════════════
const TABS = [
  { id: "sectors", l: "Secteurs", i: "📊" },
  { id: "etfs", l: "ETFs", i: "📦" },
  { id: "crypto", l: "Crypto", i: "₿" },
  { id: "sentiment", l: "Sentiment", i: "🌡" },
  { id: "allocator", l: "Allocateur", i: "🎯" },
];

export default function App() {
  const [stocks, setStocks] = useState(MEM.stocks || {});
  const [crypto, setCrypto] = useState(MEM.crypto || {});
  const [newsData, setNewsData] = useState(MEM.news || []);
  const [aiData, setAiData] = useState(MEM.ai || null);
  const [tab, setTab] = useState("sectors");
  const [filter, setFilter] = useState("ALL");

  const [loadMkt, setLoadMkt] = useState(false);
  const [loadNews, setLoadNews] = useState(false);
  const [loadAI, setLoadAI] = useState(false);
  const [errors, setErrors] = useState({});
  const [autoMode, setAutoMode] = useState(false);
  const [history, setHistory] = useState(MEM.history || null);

  // Charge l'historique persisté au montage (uniquement l'historique, pas de news)
  useEffect(() => {
    let cancelled = false;
    (async () => {
      try {
        if (!MEM.history) {
          const h = await loadHistory();
          if (cancelled) {
            return;
          }
          MEM.history = h;
          MEM.historyAt = Date.now();
          setHistory(h);
        }
      } catch (e) {
        console.warn("History initial load failed:", e);
      }
    })();
    return () => {
      cancelled = true;
    };
  }, []);

  // Tick pour rafraîchir les compteurs d'âge (30s)
  const setTick = useState(0)[1];
  useEffect(() => {
    const id = setInterval(() => setTick((t) => t + 1), 30000);
    return () => clearInterval(id);
  }, [setTick]);

  const ALL_TICKERS = useMemo(() => {
    const s = new Set();
    STOCKS.forEach((x) => s.add(x.t));
    ETF_LIST.forEach((x) => s.add(x.t));
    return Array.from(s);
  }, []);

  const ALL_CRYPTOS = useMemo(() => CRYPTOS.map((c) => c.sym), []);

  function clearErr(k) {
    setErrors((e) => {
      const n = Object.assign({}, e);
      delete n[k];
      return n;
    });
  }

  const refreshMarket = useCallback(async () => {
    if (loadMkt) {
      return null;
    }
    setLoadMkt(true);
    clearErr("market");
    try {
      // Parallélise stocks + crypto en deux appels distincts (chaque payload tient
      // confortablement dans max_tokens et conserve la résilience indépendante).
      const [sRes, cRes] = await Promise.allSettled([
        fetchStocks(ALL_TICKERS),
        fetchCryptos(ALL_CRYPTOS),
      ]);

      let s = MEM.stocks || {};
      let c = MEM.crypto || {};
      const errs = [];

      if (sRes.status === "fulfilled") {
        s = sRes.value;
        MEM.stocks = s;
        MEM.stocksAt = Date.now();
        setStocks(s);
        // Historique : append/MAJ de la ligne du jour
        try {
          const current = MEM.history || (await loadHistory());
          const next = appendTodayPrices(current, s);
          MEM.history = next;
          MEM.historyAt = Date.now();
          setHistory(next);
          await storageSet(HISTORY_KEY, JSON.stringify(next));
        } catch (he) {
          console.warn("History save skipped:", he);
        }
      } else {
        console.error("Stocks refresh failed:", sRes.reason);
        errs.push("stocks: " + (sRes.reason && sRes.reason.message ? sRes.reason.message : "err"));
      }

      if (cRes.status === "fulfilled") {
        c = cRes.value;
        MEM.crypto = c;
        MEM.cryptoAt = Date.now();
        setCrypto(c);
      } else {
        console.error("Crypto refresh failed:", cRes.reason);
        errs.push("crypto: " + (cRes.reason && cRes.reason.message ? cRes.reason.message : "err"));
      }

      if (errs.length > 0) {
        setErrors((x) => Object.assign({}, x, { market: errs.join(" · ") }));
      }
      if (sRes.status === "rejected" && cRes.status === "rejected") {
        return null;
      }
      return { stocks: s, crypto: c };
    } catch (e) {
      console.error("Market refresh:", e);
      setErrors((x) => Object.assign({}, x, { market: e.message || "erreur" }));
      return null;
    } finally {
      setLoadMkt(false);
    }
  }, [loadMkt, ALL_TICKERS, ALL_CRYPTOS]);

  const refreshNews = useCallback(async () => {
    if (loadNews) {
      return null;
    }
    setLoadNews(true);
    clearErr("news");
    try {
      const n = await fetchNews();
      MEM.news = n;
      MEM.newsAt = Date.now();
      setNewsData(n);
      return n;
    } catch (e) {
      console.error("News refresh:", e);
      setErrors((x) => Object.assign({}, x, { news: e.message || "erreur" }));
      return null;
    } finally {
      setLoadNews(false);
    }
  }, [loadNews]);

  const refreshAI = useCallback(
    async (stocksOverride, cryptoOverride, newsOverride) => {
      if (loadAI) {
        return null;
      }
      setLoadAI(true);
      clearErr("ai");
      try {
        const s = stocksOverride || stocks || MEM.stocks || {};
        const c = cryptoOverride || crypto || MEM.crypto || {};
        const n = newsOverride || newsData || MEM.news || [];
        const ai = await fetchAIAnalysis(s, c, n);
        MEM.ai = ai;
        MEM.aiAt = Date.now();
        setAiData(ai);
        return ai;
      } catch (e) {
        console.error("AI refresh:", e);
        setErrors((x) => Object.assign({}, x, { ai: e.message || "erreur" }));
        return null;
      } finally {
        setLoadAI(false);
      }
    },
    [loadAI, stocks, crypto, newsData]
  );

  const refreshAll = useCallback(async () => {
    const [mkt, nws] = await Promise.all([refreshMarket(), refreshNews()]);
    const s = mkt ? mkt.stocks : undefined;
    const c = mkt ? mkt.crypto : undefined;
    await refreshAI(s, c, nws);
  }, [refreshMarket, refreshNews, refreshAI]);

  // Mode auto : check toutes les 60s, refresh si 24h écoulées
  useEffect(() => {
    if (!autoMode) {
      return undefined;
    }
    function check() {
      const last = Math.max(MEM.stocksAt || 0, MEM.cryptoAt || 0, MEM.newsAt || 0, MEM.aiAt || 0);
      if (last && Date.now() - last >= TTL_DAY) {
        refreshAll();
      }
    }
    const id = setInterval(check, 60000);
    return () => clearInterval(id);
  }, [autoMode, refreshAll]);

  const dateStr = new Date().toLocaleDateString("fr-FR", { weekday: "long", day: "numeric", month: "long", year: "numeric" });
  const hasData = Object.keys(stocks).length > 0 || Object.keys(crypto).length > 0;
  const anyLoading = loadMkt || loadNews || loadAI;
  const lastRefresh = Math.max(MEM.stocksAt || 0, MEM.cryptoAt || 0, MEM.newsAt || 0, MEM.aiAt || 0);
  const countdown = autoMode && lastRefresh ? nextRefreshString(lastRefresh) : null;

  function handleAutoToggle() {
    const next = !autoMode;
    setAutoMode(next);
    if (next && !lastRefresh) {
      refreshAll();
    }
  }

  const hStats = useMemo(() => historyStats(history), [history]);

  function handleDownloadHistory() {
    if (!history || hStats.days === 0) {
      return;
    }
    downloadHistoryExcel(history, ALL_TICKERS);
  }

  return (
    <div style={{ minHeight: "100vh", background: C.bg, color: C.txt, fontFamily: "'Segoe UI',system-ui,sans-serif" }}>
      <div style={{ position: "sticky", top: 0, zIndex: 100, background: C.surface, borderBottom: "1px solid " + C.border }}>
        <div style={{ padding: "10px 22px", borderBottom: "1px solid " + C.border }}>
          <div style={{ maxWidth: 1500, margin: "0 auto", display: "flex", alignItems: "center", justifyContent: "space-between", flexWrap: "wrap", gap: 11 }}>
            <div style={{ display: "flex", alignItems: "center", gap: 13, flexWrap: "wrap" }}>
              <div>
                <div style={{ fontSize: 17, fontWeight: 700 }}>📊 Radar de Marché LIVE</div>
                <div style={{ fontSize: 10, color: C.muted, fontFamily: "monospace", marginTop: 1 }}>{dateStr}</div>
              </div>
              <div style={{ display: "flex", gap: 6, flexWrap: "wrap", alignItems: "center" }}>
                <StatusPill label="Marché" icon="📈" at={MEM.stocksAt} loading={loadMkt} err={errors.market} color={C.yf} />
                <StatusPill label="News" icon="📰" at={MEM.newsAt} loading={loadNews} err={errors.news} color={C.xfer} />
                <StatusPill label="IA" icon="✦" at={MEM.aiAt} loading={loadAI} err={errors.ai} color={C.ai} />
                {hStats.days > 0 ? (
                  <span style={{ padding: "3px 9px", borderRadius: 20, fontSize: 10, fontWeight: 600, fontFamily: "monospace", color: C.accent, background: C.accent + "18", border: "1px solid " + C.accent + "44", display: "inline-flex", alignItems: "center", gap: 5 }}>
                    📒 {hStats.days}j · {hStats.tickers} actifs
                  </span>
                ) : null}
                {autoMode ? (
                  <span style={{ padding: "3px 9px", borderRadius: 20, fontSize: 10, fontWeight: 700, fontFamily: "monospace", color: C.up, background: C.up + "18", border: "1px solid " + C.up + "55", display: "inline-flex", alignItems: "center", gap: 5 }}>
                    🤖 AUTO ON{countdown ? " · " + countdown : ""}
                  </span>
                ) : null}
              </div>
            </div>
            <div style={{ display: "flex", alignItems: "center", gap: 6, flexWrap: "wrap" }}>
              <ActionBtn onClick={refreshMarket} loading={loadMkt} icon="📈" label="Marché" color={C.yf} />
              <ActionBtn onClick={refreshNews} loading={loadNews} icon="📰" label="News" color={C.xfer} />
              <ActionBtn onClick={() => refreshAI()} loading={loadAI} disabled={!hasData} icon="✦" label="Analyse IA" color={C.ai} />
              <ActionBtn onClick={handleDownloadHistory} disabled={hStats.days === 0} icon="📒" label={"Excel (" + hStats.days + "j)"} color={C.accent} />
              <ActionBtn
                onClick={handleAutoToggle}
                loading={anyLoading && autoMode}
                icon={autoMode ? "🛑" : "🤖"}
                label={autoMode ? "Stop Auto" : "Full Auto"}
                color={autoMode ? C.sell : C.up}
                active={autoMode}
              />
            </div>
          </div>
        </div>
        <div style={{ maxWidth: 1500, margin: "0 auto", display: "flex", padding: "0 22px" }}>
          {TABS.map((t) => {
            const on = tab === t.id;
            const btnStyle = {
              padding: "11px 17px",
              background: "transparent",
              border: "none",
              cursor: "pointer",
              color: on ? C.txt : C.muted,
              fontSize: 13,
              fontWeight: on ? 600 : 400,
              borderBottom: on ? "2px solid " + C.accent : "2px solid transparent",
              display: "flex",
              alignItems: "center",
              gap: 6,
            };
            return (
              <button key={t.id} onClick={() => setTab(t.id)} style={btnStyle}>
                <span>{t.i}</span>
                {t.l}
              </button>
            );
          })}
        </div>
      </div>

      <div style={{ maxWidth: 1500, margin: "0 auto", padding: "18px 22px" }}>
        {Object.keys(errors).length > 0 ? (
          <div style={{ background: C.sell + "12", border: "1px solid " + C.sell + "55", borderRadius: 8, padding: "8px 14px", marginBottom: 12, fontSize: 11, color: C.sell, fontFamily: "monospace", display: "flex", gap: 14, flexWrap: "wrap" }}>
            ⚠ {Object.entries(errors).map(([k, v]) => k + ": " + v).join(" · ")}
          </div>
        ) : null}

        <AISummary aiData={aiData} />

        {!hasData && !anyLoading ? <WelcomeScreen /> : null}

        {tab === "sectors" ? <SectorsTab stocks={stocks} aiData={aiData} filter={filter} setFilter={setFilter} /> : null}
        {tab === "etfs" ? <ETFsTab stocks={stocks} /> : null}
        {tab === "crypto" ? <CryptoTab crypto={crypto} /> : null}
        {tab === "sentiment" ? <SentimentTab aiData={aiData} news={newsData} newsAge={ageString(MEM.newsAt)} /> : null}
        {tab === "allocator" ? <AllocatorTab stocks={stocks} crypto={crypto} /> : null}
      </div>

      <style>{`
        @keyframes spin { from { transform: rotate(0deg); } to { transform: rotate(360deg); } }
        * { box-sizing: border-box; margin: 0; padding: 0; }
        ::-webkit-scrollbar { width: 5px; height: 5px; }
        ::-webkit-scrollbar-track { background: ${C.bg}; }
        ::-webkit-scrollbar-thumb { background: ${C.border}; border-radius: 3px; }
        button:hover:not(:disabled) { opacity: 0.92; }
      `}</style>
    </div>
  );
}
