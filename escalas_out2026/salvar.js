// Alterações que ficam no site: no servidor do Render (server.js) cada mudança na página é
// gravada no Git e o botão Desfazer volta à versão anterior, para todo mundo.
// Aberta do disco, no GitHub Pages ou no claude.ai não há servidor: a página funciona como
// antes e a barra de salvar não aparece.
//
// Uso: salvaSite({ nome, estado, aplica, onde }) — estado() devolve o que a pessoa escolheu,
// aplica(st) põe a página nesse estado (st nulo = como veio) e onde é o elemento que recebe a
// barra. Depois a página chama salvaSite.mudou() sempre que redesenha.
(function () {
  const servido = /^https?:$/.test(location.protocol) && !/(^|\.)github\.io$/.test(location.hostname);
  let cfg = null, ultimo = null, pronto = false, timer = 0, ocupado = false, el = {};
  const status = (txt, cls) => { if (el.st) { el.st.textContent = txt; el.st.className = "salva-st" + (cls ? " " + cls : ""); } };
  const hora = iso => { try { return new Date(iso).toLocaleString("pt-BR", { day: "2-digit", month: "2-digit", hour: "2-digit", minute: "2-digit" }); } catch (e) { return ""; } };
  function mostra(r) {
    el.desf.disabled = !r.desfazer;
    el.desf.title = r.desfazer ? `Voltar à versão anterior (${r.desfazer} no histórico)` : "Nada para desfazer";
    status(r.quando ? `Salvo ${hora(r.quando)}` : "Sem alterações salvas", "ok");
  }
  async function chama(metodo, sufixo, corpo) {
    const r = await fetch(`/api/estado/${cfg.nome}${sufixo}`, { method: metodo, headers: corpo ? { "Content-Type": "application/json" } : {},
      body: corpo ? JSON.stringify(corpo) : undefined, cache: "no-store" });
    if (!r.ok) throw new Error(String(r.status));
    return r.json();
  }
  function aplica(st) { ocupado = true; try { cfg.aplica(st); } finally { ocupado = false; } ultimo = JSON.stringify(cfg.estado()); }
  async function salva() {
    const st = cfg.estado(), txt = JSON.stringify(st);
    if (txt === ultimo) return;
    status("Salvando…");
    try { const r = await chama("PUT", "", { estado: st }); ultimo = txt; mostra(r); }
    catch (e) { status("Não salvou — tente de novo", "erro"); }
  }
  async function desfaz() {
    clearTimeout(timer); el.desf.disabled = true; status("Desfazendo…");
    try { const r = await chama("POST", "/desfazer"); aplica(r.estado); mostra(r); }
    catch (e) { status("Não desfez — tente de novo", "erro"); el.desf.disabled = false; }
  }
  window.salvaSite = async function (c) {
    if (!servido) return;
    cfg = c;
    const css = document.createElement("style");
    css.textContent = ".salva{display:inline-flex;align-items:center;gap:8px;margin-left:8px}"
      + ".salva-st{font-size:12px;color:var(--muted,#5A6072);white-space:nowrap}.salva-st.erro{color:#C0392B;font-weight:600}"
      + ".salva-desf.solto{font:inherit;font-weight:600;font-size:14px;padding:6px 10px;border:1px solid var(--line,#D6DAE2);border-radius:var(--r,4px);background:transparent;color:var(--ink,#2B2B2B);cursor:pointer;white-space:nowrap}"
      + ".salva-desf:disabled{opacity:.45;cursor:default}@media print{.salva{display:none!important}}";
    document.head.appendChild(css);
    const barra = document.createElement("span");
    barra.className = "salva";
    barra.innerHTML = '<button type="button" class="salva-desf" disabled>↶ Desfazer</button><span class="salva-st" role="status">Carregando…</span>';
    c.onde.appendChild(barra);
    el = { desf: barra.querySelector("button"), st: barra.querySelector(".salva-st") };
    if (c.solto) el.desf.classList.add("solto");
    el.desf.addEventListener("click", desfaz);
    document.addEventListener("keydown", e => {
      if ((e.ctrlKey || e.metaKey) && !e.shiftKey && e.key.toLowerCase() === "z" && !/^(INPUT|TEXTAREA|SELECT)$/.test(e.target.tagName) && !el.desf.disabled) { e.preventDefault(); desfaz(); }
    });
    try {
      const r = await chama("GET", "");
      if (r.estado) aplica(r.estado); else ultimo = JSON.stringify(cfg.estado());
      mostra(r); pronto = true;
    } catch (e) { barra.remove(); }
  };
  window.salvaSite.mudou = function () {
    if (!pronto || ocupado) return;
    clearTimeout(timer);
    if (JSON.stringify(cfg.estado()) !== ultimo) { status("Falta salvar…"); timer = setTimeout(salva, 800); }
  };
  addEventListener("beforeunload", e => { if (pronto && cfg && JSON.stringify(cfg.estado()) !== ultimo) { salva(); e.preventDefault(); e.returnValue = ""; } });
})();
