// Servidor do Render: publica as páginas de public/ (montada pelo render-build.sh) e guarda as
// alterações feitas nas páginas de escalas no Postgres do Supabase (tabela esc_est, ver
// supabase/schema.sql), como nos outros projetos. Sem DATABASE_URL (rodando na máquina), grava
// em estado-local/.
//
//   GET  /api/estado/<página>            estado atual e quantas versões dá para desfazer
//   PUT  /api/estado/<página>            {estado}: grava, guardando o anterior no histórico
//   POST /api/estado/<página>/desfazer   volta ao estado anterior
//
// Variáveis: DATABASE_URL (connection string do Supabase, Session pooler), PORT.
"use strict";
const http = require("http");
const fs = require("fs");
const path = require("path");

const PUBLIC = path.join(__dirname, "public");
const LOCAL = path.join(__dirname, "estado-local");
const DATABASE_URL = process.env.DATABASE_URL || "";
const HISTORICO = 100;            // versões guardadas para o Desfazer
const MAX_CORPO = 512 * 1024;
const NOME_OK = /^[a-z0-9_-]{1,60}$/;
const TIPOS = { ".html": "text/html; charset=utf-8", ".js": "text/javascript; charset=utf-8", ".css": "text/css; charset=utf-8",
  ".json": "application/json; charset=utf-8", ".pdf": "application/pdf", ".png": "image/png", ".svg": "image/svg+xml", ".ico": "image/x-icon" };

/* ---------- onde o estado fica: Supabase ou pasta local ---------- */
if (process.env.RENDER && !DATABASE_URL) console.error("[estado] ATENÇÃO: DATABASE_URL não configurada: as alterações vão para o disco do Render e somem no próximo deploy.");
let pool = null, tabela = null;
function banco() {
  if (!pool) {
    const { Pool } = require("pg");
    // PGSSLMODE=disable só para um banco local de testes; no Supabase a conexão é sempre com SSL
    pool = new Pool({ connectionString: DATABASE_URL, ssl: process.env.PGSSLMODE === "disable" ? false : { rejectUnauthorized: false }, max: 3 });
  }
  // a tabela também é criada aqui, para o primeiro deploy funcionar mesmo sem rodar o schema.sql
  if (!tabela) tabela = pool.query(fs.readFileSync(path.join(__dirname, "supabase", "schema.sql"), "utf-8")).catch(e => { tabela = null; throw e; });
  return tabela.then(() => pool);
}
const cache = new Map();          // nome -> doc
async function le(nome) {
  if (cache.has(nome)) return cache.get(nome);
  let doc = null;
  if (DATABASE_URL) {
    const r = await (await banco()).query("select doc from esc_est where pag = $1", [nome]);
    if (r.rows.length) doc = r.rows[0].doc;
  } else {
    try { doc = JSON.parse(fs.readFileSync(path.join(LOCAL, `${nome}.json`), "utf-8")); } catch (e) {}
  }
  doc = doc || { estado: null, quando: null, anteriores: [] };
  cache.set(nome, doc);
  return doc;
}
async function grava(nome, doc) {
  if (DATABASE_URL) {
    await (await banco()).query(
      "insert into esc_est (pag, doc, atu_em) values ($1, $2, now()) on conflict (pag) do update set doc = excluded.doc, atu_em = now()",
      [nome, JSON.stringify(doc)]);
  } else {
    fs.mkdirSync(LOCAL, { recursive: true });
    fs.writeFileSync(path.join(LOCAL, `${nome}.json`), JSON.stringify(doc, null, 1) + "\n");
  }
  cache.set(nome, doc);
}
// uma gravação por vez: duas pessoas salvando juntas não perdem o histórico
let fila = Promise.resolve();
const emFila = f => (fila = fila.then(f, f));

const resposta = doc => ({ estado: doc.estado, quando: doc.quando, desfazer: doc.anteriores.length });

/* ---------- HTTP ---------- */
function envia(res, status, corpo, tipo = "application/json; charset=utf-8") {
  res.writeHead(status, { "Content-Type": tipo, "Cache-Control": "no-cache", "X-Content-Type-Options": "nosniff" });
  res.end(typeof corpo === "string" || Buffer.isBuffer(corpo) ? corpo : JSON.stringify(corpo));
}
function leCorpo(req) {
  return new Promise((ok, erro) => {
    let n = 0; const partes = [];
    req.on("data", p => { n += p.length; if (n > MAX_CORPO) { erro(new Error("grande demais")); req.destroy(); } else partes.push(p); });
    req.on("end", () => { try { ok(JSON.parse(Buffer.concat(partes).toString("utf-8") || "{}")); } catch (e) { erro(e); } });
    req.on("error", erro);
  });
}
async function api(req, res, nome, acao) {
  if (!NOME_OK.test(nome)) return envia(res, 404, { erro: "página desconhecida" });
  if (req.method === "GET" && !acao) return envia(res, 200, resposta(await le(nome)));
  if (req.method === "PUT" && !acao) {
    const { estado } = await leCorpo(req);
    if (!estado || typeof estado !== "object") return envia(res, 400, { erro: "falta o estado" });
    const doc = await emFila(async () => {
      const d = await le(nome);
      if (JSON.stringify(d.estado) === JSON.stringify(estado)) return d;
      const novo = { estado, quando: new Date().toISOString(),
        // o primeiro anterior é null: desfazer até o fim volta a página como ela veio
        anteriores: [{ estado: d.estado, quando: d.quando }, ...d.anteriores].slice(0, HISTORICO) };
      await grava(nome, novo);
      return novo;
    });
    return envia(res, 200, resposta(doc));
  }
  if (req.method === "POST" && acao === "desfazer") {
    const doc = await emFila(async () => {
      const d = await le(nome);
      if (!d.anteriores.length) return d;
      const [volta, ...resto] = d.anteriores;
      const novo = { estado: volta.estado, quando: volta.quando, anteriores: resto };
      await grava(nome, novo);
      return novo;
    });
    return envia(res, 200, resposta(doc));
  }
  return envia(res, 405, { erro: "método não aceito" });
}
function arquivo(req, res, url) {
  let rel;
  try { rel = decodeURIComponent(url.pathname); } catch (e) { return envia(res, 400, "Endereço inválido", "text/plain; charset=utf-8"); }
  let alvo = path.normalize(path.join(PUBLIC, rel));
  if (alvo !== PUBLIC && !alvo.startsWith(PUBLIC + path.sep)) return envia(res, 404, "Não encontrado", "text/plain; charset=utf-8");
  fs.stat(alvo, (e, st) => {
    if (!e && st.isDirectory()) {
      if (!url.pathname.endsWith("/")) { res.writeHead(301, { Location: url.pathname + "/" + url.search }); return res.end(); }
      alvo = path.join(alvo, "index.html");
    }
    fs.readFile(alvo, (e2, buf) => {
      if (e2) return envia(res, 404, "Não encontrado", "text/plain; charset=utf-8");
      envia(res, 200, buf, TIPOS[path.extname(alvo).toLowerCase()] || "application/octet-stream");
    });
  });
}
http.createServer(async (req, res) => {
  const url = new URL(req.url, "http://x");
  const m = url.pathname.match(/^\/api\/estado\/([^/]+)(?:\/([a-z]+))?\/?$/);
  try {
    if (m) return await api(req, res, m[1], m[2]);
    if (req.method !== "GET" && req.method !== "HEAD") return envia(res, 405, { erro: "método não aceito" });
    arquivo(req, res, url);
  } catch (e) {
    console.error(e);
    if (!res.headersSent) envia(res, 500, { erro: "não foi possível salvar agora" });
  }
}).listen(process.env.PORT || 3000, () => console.log(`no ar na porta ${process.env.PORT || 3000}; estado em ${DATABASE_URL ? "Supabase (esc_est)" : LOCAL}`));
