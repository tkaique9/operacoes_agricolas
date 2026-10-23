// Servidor do Render: publica as páginas de public/ (montada pelo render-build.sh) e guarda as
// alterações feitas nas páginas de escalas. Cada alteração vira um commit no arquivo
// estado/<página>.json do branch GITHUB_BRANCH (padrão "dados", fora da main para não disparar
// deploy a cada clique). Sem GITHUB_TOKEN (rodando na máquina), grava em estado-local/.
//
//   GET  /api/estado/<página>            estado atual e quantas versões dá para desfazer
//   PUT  /api/estado/<página>            {estado}: grava, guardando o anterior no histórico
//   POST /api/estado/<página>/desfazer   volta ao estado anterior
//
// Variáveis: GITHUB_TOKEN (token com permissão de escrita em Contents no repositório),
// GITHUB_REPO (padrão tkaique9/operacoes_agricolas), GITHUB_BRANCH (padrão dados), PORT.
"use strict";
const http = require("http");
const fs = require("fs");
const path = require("path");

const PUBLIC = path.join(__dirname, "public");
const LOCAL = path.join(__dirname, "estado-local");
const TOKEN = process.env.GITHUB_TOKEN || "";
const REPO = process.env.GITHUB_REPO || "tkaique9/operacoes_agricolas";
const BRANCH = process.env.GITHUB_BRANCH || "dados";
const HISTORICO = 100;            // versões guardadas para o Desfazer
const MAX_CORPO = 512 * 1024;
const NOME_OK = /^[a-z0-9_-]{1,60}$/;
const TIPOS = { ".html": "text/html; charset=utf-8", ".js": "text/javascript; charset=utf-8", ".css": "text/css; charset=utf-8",
  ".json": "application/json; charset=utf-8", ".pdf": "application/pdf", ".png": "image/png", ".svg": "image/svg+xml", ".ico": "image/x-icon" };

/* ---------- onde o estado fica: GitHub ou pasta local ---------- */
const gh = async (metodo, rota, corpo) => {
  const r = await fetch(`https://api.github.com/repos/${REPO}${rota}`, {
    method: metodo,
    headers: { Authorization: `Bearer ${TOKEN}`, Accept: "application/vnd.github+json", "User-Agent": "operacoes-agricolas",
      ...(corpo ? { "Content-Type": "application/json" } : {}) },
    body: corpo ? JSON.stringify(corpo) : undefined,
  });
  const txt = await r.text();
  return { status: r.status, json: txt ? JSON.parse(txt) : null };
};
let branchPronto = null;
function garanteBranch() {
  // cria o branch de dados a partir do branch padrão na primeira gravação
  if (!branchPronto) branchPronto = (async () => {
    if ((await gh("GET", `/branches/${BRANCH}`)).status === 200) return;
    const repo = await gh("GET", "");
    const base = await gh("GET", `/git/ref/heads/${repo.json.default_branch}`);
    const r = await gh("POST", "/git/refs", { ref: `refs/heads/${BRANCH}`, sha: base.json.object.sha });
    if (r.status !== 201 && r.status !== 422) throw new Error(`não criou o branch ${BRANCH}: ${r.status}`);
  })().catch(e => { branchPronto = null; throw e; });
  return branchPronto;
}
const cache = new Map();          // nome -> {doc, sha}
async function le(nome) {
  if (cache.has(nome)) return cache.get(nome).doc;
  let doc = null, sha = null;
  if (TOKEN) {
    await garanteBranch();
    const r = await gh("GET", `/contents/estado/${nome}.json?ref=${BRANCH}`);
    if (r.status === 200) { doc = JSON.parse(Buffer.from(r.json.content, "base64").toString("utf-8")); sha = r.json.sha; }
    else if (r.status !== 404) throw new Error(`GitHub respondeu ${r.status} ao ler`);
  } else {
    try { doc = JSON.parse(fs.readFileSync(path.join(LOCAL, `${nome}.json`), "utf-8")); } catch (e) {}
  }
  doc = doc || { estado: null, quando: null, anteriores: [] };
  cache.set(nome, { doc, sha });
  return doc;
}
async function grava(nome, doc, msg) {
  const txt = JSON.stringify(doc, null, 1) + "\n";
  if (TOKEN) {
    const c = cache.get(nome) || {};
    const r = await gh("PUT", `/contents/estado/${nome}.json`, { message: msg, branch: BRANCH,
      content: Buffer.from(txt, "utf-8").toString("base64"), ...(c.sha ? { sha: c.sha } : {}) });
    if (r.status !== 200 && r.status !== 201) { cache.delete(nome); throw new Error(`GitHub respondeu ${r.status} ao gravar`); }
    cache.set(nome, { doc, sha: r.json.content.sha });
  } else {
    fs.mkdirSync(LOCAL, { recursive: true });
    fs.writeFileSync(path.join(LOCAL, `${nome}.json`), txt);
    cache.set(nome, { doc, sha: null });
  }
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
      await grava(nome, novo, `Alteração em ${nome}`);
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
      await grava(nome, novo, `Desfazer em ${nome}`);
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
}).listen(process.env.PORT || 3000, () => console.log(`no ar na porta ${process.env.PORT || 3000}; estado em ${TOKEN ? `GitHub ${REPO}@${BRANCH}` : LOCAL}`));
