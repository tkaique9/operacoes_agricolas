// Gera a apresentação de necessidade de fertilizantes a partir de dados_consolidados.json.
//
//   node gerar_apresentacao.js
//
// Saídas em build/: base.pptx (slides, gráficos nativos, notas) e animacoes.json (o roteiro
// que animar_e_exportar.ps1 aplica no PowerPoint: grupos, animações, transições e ajustes
// finos dos gráficos que o pptxgenjs não escreve).
"use strict";
const fs = require("fs");
const path = require("path");
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const JSZip = require("jszip");
const fa = require("react-icons/fa");
const gi = require("react-icons/gi");

const D = JSON.parse(fs.readFileSync(path.join(__dirname, "dados_consolidados.json"), "utf8"));
const A = D.cenarios.A;
const B = D.cenarios.B;
const dA = A.derivados;
const dB = B.derivados;
const BUILD = path.join(__dirname, "build");
fs.mkdirSync(BUILD, { recursive: true });

// ---------------------------------------------------------------- tema
const THEME = {
	name: "Fertilizantes CRV",
	headFontFace: "Calibri",
	bodyFontFace: "Calibri",
	colors: {
		dk1: "1E2B25", lt1: "FFFFFF", dk2: "0E3B2C", lt2: "EEF3EA",
		accent1: "1F6B3A", // verde-floresta: necessidade, 16-00-20
		accent2: "8DBE3F", // verde-cana: realizado
		accent3: "D9A21B", // âmbar: atenção, excedente
		accent4: "B3402E", // vermelho: déficit, compra
		accent5: "7A5230", // marrom: vinhaça
		accent6: "66736D", // cinza-verde: texto secundário
		hlink: "1F6B3A", folHlink: "66736D",
	},
};
const HEX = THEME.colors;
const EXTRA = { cinzaClaro: "A9B4AE", vinhacaClara: "A8825A", canaClara: "C5DE9C", grade: "DCE4D8", linhaTabela: "D5DED2" };

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13,333 × 7,5 pol
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = "Planejamento de Necessidade de Fertilizantes";
pres.subject = "Necessidade de fertilizantes, estoque, compras e cenários — posição 16/10/2026";
pres.author = "Departamento Agrícola";
pres.company = "Usina CRV Industrial";
const C = pres.SchemeColor;

// ---------------------------------------------------------------- utilidades
const fmt = (v, d = 0) => Number(v).toLocaleString("pt-BR", { minimumFractionDigits: d, maximumFractionDigits: d });
const pct = (v, d = 0) => fmt(v * 100, d) + "%";
const dataBR = (iso) => iso.slice(8, 10) + "/" + iso.slice(5, 7);
const ARQ = "Necessidade Fertilizante OUT 2026.xlsx";
const ABA_A = A.aba;
const ABA_B = B.aba;

const SPEC = { transicoes: {}, grupos: [], efeitos: [], graficos: [] };
let n = 0;
const fx = (forma, efeito = "fade", gatilho = "apos", extra = {}) =>
	SPEC.efeitos.push(Object.assign({ slide: n, forma, efeito, gatilho, duracao: 0.45, atraso: 0 }, extra));
const grp = (nome, membros) => SPEC.grupos.push({ slide: n, nome, membros });
const ajusteGrafico = (forma, ops) => SPEC.graficos.push(Object.assign({ slide: n, forma }, ops));

function novoSlide(master, secao, transicao = "fade") {
	n++;
	SPEC.transicoes[n] = transicao;
	return pres.addSlide({ masterName: master, sectionTitle: secao });
}

// Espaço inseparável entre número e unidade ("2.260 t", "R$ 226 mil"), para a quebra de linha
// nunca separar o valor da unidade.
const UNIDADE = /(\d) (?=(t|ha|kg|dias|dia|mil|t\/dia|ha\/dia|kg\/ha)\b)|(R\$) (?=\d)/g;
const colar = (s) => s.replace(UNIDADE, (m, d, _u, rs) => (rs ? "R$ " : d + " "));
const colarTudo = (t) => (typeof t === "string" ? colar(t) : t.map((r) => Object.assign({}, r, { text: colar(r.text) })));

function txt(slide, text, o) {
	slide.addText(colarTudo(text), Object.assign({ isTextBox: true, margin: 0, fontSize: 14, color: C.text1, valign: "top", fit: "none" }, o));
}

function caixa(slide, o) {
	const opts = Object.assign({ rectRadius: 0.08 }, o);
	if (!opts.fill) opts.fill = { color: C.background2 };
	slide.addShape(opts.rectRadius ? pres.shapes.ROUNDED_RECTANGLE : pres.shapes.RECTANGLE, opts);
}

const sombra = () => ({ type: "outer", color: "0E3B2C", blur: 8, offset: 2, angle: 90, opacity: 0.12 });

function fonte(slide, texto) {
	txt(slide, "Fonte: " + texto, { x: 0.6, y: 6.66, w: 12.13, h: 0.28, fontSize: 10, color: C.accent6, valign: "middle", objectName: `s${n}_fonte` });
}

function titulos(slide, kicker, titulo) {
	slide.addText(kicker, { placeholder: "kicker" });
	slide.addText(colar(titulo), { placeholder: "title" });
}

// ---------------------------------------------------------------- ícones (react-icons → PNG)
const cacheIcone = new Map();
async function icone(Comp, fg, bg = null, px = 256) {
	const chave = [Comp.name, fg, bg, px].join("|");
	if (cacheIcone.has(chave)) return cacheIcone.get(chave);
	const inner = bg ? Math.round(px * 0.56) : px;
	let svg = ReactDOMServer.renderToStaticMarkup(React.createElement(Comp, { color: "#" + fg, size: String(inner) }));
	if (bg) {
		const off = Math.round((px - inner) / 2);
		svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${px}" height="${px}" viewBox="0 0 ${px} ${px}">` +
			`<circle cx="${px / 2}" cy="${px / 2}" r="${px / 2}" fill="#${bg}"/>` +
			svg.replace("<svg ", `<svg x="${off}" y="${off}" `) + "</svg>";
	}
	const png = await sharp(Buffer.from(svg)).png().toBuffer();
	const data = "image/png;base64," + png.toString("base64");
	cacheIcone.set(chave, data);
	return data;
}

// ---------------------------------------------------------------- imagem de fundo (canavial estilizado)
function prng(seed) {
	return () => {
		seed |= 0; seed = (seed + 0x6d2b79f5) | 0;
		let t = Math.imul(seed ^ (seed >>> 15), 1 | seed);
		t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
		return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
	};
}

// Folha de cana: sai do nó (x, y), sobe até a altura `subida`·L e cai em arco até a ponta,
// `queda`·L abaixo do nó. `dir` (−1 a 1) dá o lado e quanto ela abre.
function folha(x, y, dir, L, larg, subida, queda, cor, op) {
	const c1x = x + dir * L * 0.38, c1y = y - L * subida;
	const tx = x + dir * L, ty = y + L * queda;
	const c2x = c1x + dir * larg * 0.6, c2y = c1y + larg * 2.2;
	const f = (v) => v.toFixed(1);
	return `<path d="M${f(x)},${f(y)} Q${f(c1x)},${f(c1y)} ${f(tx)},${f(ty)} Q${f(c2x)},${f(c2y)} ${f(x + larg * 0.4)},${f(y + larg)} Z" fill="${cor}" opacity="${op.toFixed(3)}"/>`;
}

function canavialSVG({ w = 2667, h = 1500, seed = 11, xMin = 0.44, xMax = 1.02, escurecerEsq = true, opacidade = 1 }) {
	const r = prng(seed);
	const rr = (a, b) => a + (b - a) * r();
	const camadas = [
		{ n: 30, cor: "#1F5236", folha: "#2A6640", op: 0.42, wmin: 7, wmax: 11, topo: [260, 760], escala: 0.7 },
		{ n: 20, cor: "#33784A", folha: "#46904F", op: 0.6, wmin: 11, wmax: 16, topo: [170, 600], escala: 0.9 },
		{ n: 11, cor: "#5C9C3F", folha: "#79B548", op: 0.85, wmin: 16, wmax: 23, topo: [80, 420], escala: 1.15 },
	];
	let g = "";
	for (const c of camadas) {
		for (let i = 0; i < c.n; i++) {
			const x0 = rr(xMin, xMax) * w;
			const topo = rr(c.topo[0], c.topo[1]);
			const curva = rr(-90, 90);
			const larg = rr(c.wmin, c.wmax);
			const y0 = h + 30;
			const cx = x0 + curva * 1.4;
			const x1 = x0 + curva;
			const ponto = (t) => ({
				x: (1 - t) ** 2 * x0 + 2 * (1 - t) * t * cx + t * t * x1,
				y: (1 - t) ** 2 * y0 + 2 * (1 - t) * t * ((y0 + topo) / 2) + t * t * topo,
			});
			const op = (c.op * opacidade).toFixed(3);
			g += `<path d="M${x0.toFixed(1)},${y0} Q${cx.toFixed(1)},${((y0 + topo) / 2).toFixed(1)} ${x1.toFixed(1)},${topo.toFixed(1)}" stroke="${c.cor}" stroke-width="${larg.toFixed(1)}" fill="none" stroke-linecap="round" opacity="${op}"/>`;
			// nós do colmo
			for (let t = rr(0.05, 0.1); t < 0.97; t += rr(0.07, 0.1)) {
				const p = ponto(t);
				g += `<line x1="${(p.x - larg * 0.75).toFixed(1)}" y1="${p.y.toFixed(1)}" x2="${(p.x + larg * 0.75).toFixed(1)}" y2="${p.y.toFixed(1)}" stroke="#0A2A1E" stroke-width="${(larg * 0.28).toFixed(1)}" opacity="${(0.5 * opacidade).toFixed(3)}"/>`;
				// folhas nos nós superiores: fitas longas e estreitas que sobem e caem em arco
				if (t > 0.5 && r() < 0.5) {
					const dir = r() < 0.5 ? -1 : 1;
					g += folha(p.x, p.y, dir, rr(300, 600) * c.escala, rr(7, 12) * c.escala, rr(0.35, 0.55), rr(-0.05, 0.35), c.folha, c.op * 0.8 * opacidade);
				}
			}
			// palmito: leque de folhas no topo do colmo
			const pt = ponto(1);
			for (let k = 0; k < 6; k++) {
				const dir = k % 2 === 0 ? -1 : 1;
				const abertura = k < 2 ? rr(0.15, 0.35) : rr(0.55, 1);
				g += folha(pt.x, pt.y, dir * abertura, rr(280, 560) * c.escala, rr(7, 12) * c.escala, rr(0.55, 0.85), rr(-0.25, 0.3), c.folha, c.op * 0.85 * opacidade);
			}
		}
	}
	const escurecer = escurecerEsq
		? `<rect width="${w}" height="${h}" fill="url(#esq)"/>`
		: `<rect width="${w}" height="${h}" fill="#0B3226" opacity="0.55"/>`;
	return `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}" viewBox="0 0 ${w} ${h}">
<defs>
<linearGradient id="fundo" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#06241A"/><stop offset="0.6" stop-color="#0E3B2C"/><stop offset="1" stop-color="#174F39"/></linearGradient>
<radialGradient id="sol" cx="0.86" cy="0.18" r="0.55"><stop offset="0" stop-color="#E9B949" stop-opacity="0.33"/><stop offset="1" stop-color="#E9B949" stop-opacity="0"/></radialGradient>
<linearGradient id="esq" x1="0" y1="0" x2="1" y2="0"><stop offset="0" stop-color="#05201A" stop-opacity="0.97"/><stop offset="0.46" stop-color="#06241A" stop-opacity="0.88"/><stop offset="0.7" stop-color="#06241A" stop-opacity="0"/></linearGradient>
<linearGradient id="chao" x1="0" y1="0" x2="0" y2="1"><stop offset="0.7" stop-color="#05201A" stop-opacity="0"/><stop offset="1" stop-color="#05201A" stop-opacity="0.75"/></linearGradient>
</defs>
<rect width="${w}" height="${h}" fill="url(#fundo)"/>
<rect width="${w}" height="${h}" fill="url(#sol)"/>
${g}
<rect width="${w}" height="${h}" fill="url(#chao)"/>
${escurecer}
</svg>`;
}

// ---------------------------------------------------------------- gráficos: padrão único
function graf(extra) {
	return Object.assign({
		catAxisLabelColor: "33413A", catAxisLabelFontSize: 12, catAxisLabelFontFace: "+mn-lt",
		valAxisLabelColor: "66736D", valAxisLabelFontSize: 11, valAxisLabelFontFace: "+mn-lt",
		valGridLine: { style: "none" }, catGridLine: { style: "none" },
		catAxisLineShow: true, catAxisLineColor: "B9C4BD", valAxisLineShow: false,
		showValue: true, dataLabelFontFace: "+mn-lt", dataLabelFontSize: 12, dataLabelColor: "1E2B25", dataLabelFontBold: true,
		showLegend: false, legendFontFace: "+mn-lt", legendFontSize: 12, legendColor: "33413A",
		showTitle: false,
	}, extra);
}

function legendaChips(slide, itens, x, y, nome) {
	// itens: [{cor, texto, larg}]
	let cx = x;
	const membros = [];
	itens.forEach((it, i) => {
		slide.addShape(pres.shapes.RECTANGLE, { x: cx, y: y + 0.06, w: 0.18, h: 0.18, fill: { color: it.cor }, objectName: `${nome}_q${i}` });
		txt(slide, it.texto, { x: cx + 0.26, y, w: it.larg, h: 0.3, fontSize: 12, color: C.text1, valign: "middle", objectName: `${nome}_t${i}` });
		membros.push(`${nome}_q${i}`, `${nome}_t${i}`);
		cx += 0.26 + it.larg + 0.25;
	});
	grp(nome, membros);
}

// =====================================================================================
async function main() {
	const I = {
		cana: await icone(gi.GiSugarCane, "FFFFFF", HEX.accent1),
		canaRodape: await icone(gi.GiSugarCane, HEX.accent1),
		canaCapa: await icone(gi.GiSugarCane, HEX.accent2),
		canaFech: await icone(gi.GiSugarCane, "A9C9A0"),
		saco: await icone(gi.GiFertilizerBag, "FFFFFF", HEX.accent1),
		sacoVerm: await icone(gi.GiFertilizerBag, "FFFFFF", HEX.accent4),
		carrinho: await icone(fa.FaShoppingCart, "FFFFFF", HEX.accent4),
		armazem: await icone(fa.FaWarehouse, "1E2B25", HEX.accent3),
		armazemVerde: await icone(fa.FaWarehouse, "FFFFFF", HEX.accent1),
		calendario: await icone(fa.FaCalendarAlt, "FFFFFF", HEX.dk2),
		gota: await icone(fa.FaTint, "FFFFFF", HEX.accent5),
		frasco: await icone(fa.FaFlask, "FFFFFF", HEX.accent1),
		mapa: await icone(fa.FaMapMarkedAlt, "FFFFFF", HEX.accent6),
		alerta: await icone(fa.FaExclamationTriangle, "1E2B25", HEX.accent3),
		caminhao: await icone(fa.FaTruck, "FFFFFF", HEX.accent1),
		caminhaoVerm: await icone(fa.FaTruck, "FFFFFF", HEX.accent4),
		dinheiro: await icone(fa.FaMoneyBillWave, "FFFFFF", HEX.accent1),
		balanca: await icone(fa.FaBalanceScale, "FFFFFF", HEX.accent1),
		base: await icone(fa.FaDatabase, "FFFFFF", HEX.accent6),
		lacuna: await icone(fa.FaQuestion, "FFFFFF", HEX.accent6),
		check: await icone(fa.FaCheck, "FFFFFF", HEX.accent1),
		ampulheta: await icone(fa.FaHourglassHalf, "1E2B25", HEX.accent3),
		lupa: await icone(fa.FaSearch, "FFFFFF", HEX.accent1),
		camadas: await icone(fa.FaLayerGroup, "FFFFFF", HEX.accent1),
		contrato: await icone(fa.FaFileContract, "FFFFFF", HEX.accent1),
		grafico: await icone(fa.FaChartLine, "FFFFFF", HEX.accent1),
		trator: await icone(fa.FaTractor, "FFFFFF", HEX.accent1),
		cenario: await icone(fa.FaRoute, "FFFFFF", HEX.accent1),
		cenaClara: await icone(gi.GiSugarCane, HEX.dk2, HEX.accent2),
		sacoClaro: await icone(gi.GiFertilizerBag, HEX.dk2, HEX.accent2),
		alertaClaro: await icone(fa.FaExclamationTriangle, HEX.dk2, HEX.accent2),
		relogioClaro: await icone(fa.FaHourglassHalf, HEX.dk2, HEX.accent2),
	};

	const capaPng = path.join(BUILD, "fundo_capa.png");
	const fechPng = path.join(BUILD, "fundo_fechamento.png");
	await sharp(Buffer.from(canavialSVG({ seed: 11 }))).png().toFile(capaPng);
	await sharp(Buffer.from(canavialSVG({ seed: 23, xMin: -0.02, xMax: 1.02, escurecerEsq: false, opacidade: 0.55 }))).png().toFile(fechPng);

	// ------------------------------------------------------------ layouts
	const rodape = "Planejamento de Necessidade de Fertilizantes  ·  Usina CRV Industrial – Unidade Capinópolis-MG  ·  Posição de 16/10/2026";
	pres.defineSlideMaster({
		title: "CAPA",
		background: { path: capaPng },
		objects: [
			{ placeholder: { options: { name: "title", type: "title", x: 0.7, y: 2.2, w: 7.9, h: 1.5, fontSize: 40, bold: true, color: C.background1, align: "left", valign: "top", margin: 0 }, text: "" } },
			{ placeholder: { options: { name: "subtitle", type: "body", x: 0.7, y: 3.85, w: 7.9, h: 0.5, fontSize: 20, color: C.accent2, align: "left", valign: "top", margin: 0 }, text: "" } },
		],
	});
	pres.defineSlideMaster({
		title: "CONTEUDO",
		background: { color: "FFFFFF" },
		objects: [
			{ placeholder: { options: { name: "kicker", type: "body", x: 0.6, y: 0.36, w: 12.13, h: 0.3, fontSize: 12, bold: true, color: C.accent1, align: "left", valign: "top", margin: 0 }, text: "" } },
			{ placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.66, w: 12.13, h: 0.72, fontSize: 30, bold: true, color: C.text2, align: "left", valign: "top", margin: 0 }, text: "" } },
			{ image: { x: 0.6, y: 7.06, w: 0.2, h: 0.2, data: I.canaRodape } },
			{ text: { text: rodape, options: { x: 0.88, y: 7.01, w: 10.5, h: 0.3, fontSize: 10, color: C.accent6, valign: "middle", margin: 0 } } },
		],
		slideNumber: { x: 12.23, y: 7.01, w: 0.5, h: 0.3, fontSize: 10, color: C.accent6, align: "right", valign: "middle", margin: 0 },
	});
	pres.defineSlideMaster({
		title: "FECHAMENTO",
		background: { path: fechPng },
		objects: [
			{ placeholder: { options: { name: "kicker", type: "body", x: 0.6, y: 0.36, w: 12.13, h: 0.3, fontSize: 12, bold: true, color: C.accent2, align: "left", valign: "top", margin: 0 }, text: "" } },
			{ placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.66, w: 12.13, h: 0.72, fontSize: 30, bold: true, color: C.background1, align: "left", valign: "top", margin: 0 }, text: "" } },
			{ image: { x: 0.6, y: 7.06, w: 0.2, h: 0.2, data: I.canaFech } },
			{ text: { text: rodape, options: { x: 0.88, y: 7.01, w: 10.5, h: 0.3, fontSize: 10, color: "A9C9A0", valign: "middle", margin: 0 } } },
		],
		slideNumber: { x: 12.23, y: 7.01, w: 0.5, h: 0.3, fontSize: 10, color: "A9C9A0", align: "right", valign: "middle", margin: 0 },
	});

	const b16A = A.balanco["16-00-20"], b46A = A.balanco["46-00-00"];
	const b16B = B.balanco["16-00-20"], b46B = B.balanco["46-00-00"];
	const realizadoTotal = A.crv.realizado_ha + A.cargill.realizado_ha;

	// ================================================================== 1. CAPA
	pres.addSection({ title: "Abertura" });
	let s = novoSlide("CAPA", "Abertura");
	s.addText([{ text: "Planejamento de Necessidade", options: { breakLine: true } }, { text: "de Fertilizantes" }], { placeholder: "title" });
	s.addText("Planejamento agrícola, suprimentos e orçamento", { placeholder: "subtitle" });
	s.addImage({ data: I.canaCapa, x: 0.7, y: 1.42, w: 0.42, h: 0.42, objectName: "s1_marca" });
	txt(s, "USINA CRV INDUSTRIAL  ·  DEPARTAMENTO AGRÍCOLA", { x: 1.25, y: 1.48, w: 7, h: 0.32, fontSize: 13, bold: true, color: C.background1, charSpacing: 2, valign: "middle", objectName: "s1_usina" });
	txt(s, [
		{ text: "Unidade Capinópolis – MG", options: { bold: true, breakLine: true } },
		{ text: `Posição de ${dataBR(A.data_base)}/2026  ·  término da safra previsto para ${dataBR(A.termino_safra)}/2026` },
	], { x: 0.7, y: 4.75, w: 7.9, h: 0.75, fontSize: 15, color: "D7E5D2", paraSpaceAfter: 4, objectName: "s1_info" });
	txt(s, `Base: planilha "${ARQ}" — abas ${ABA_A} e ${ABA_B}`, { x: 0.7, y: 6.75, w: 8, h: 0.3, fontSize: 11, color: "A9C9A0", objectName: "s1_base" });
	grp("s1_cabec", ["s1_marca", "s1_usina"]);
	fx("s1_cabec", "fade", "apos", { duracao: 0.6 });
	fx("s1_info", "fade", "apos", { duracao: 0.6 });
	s.addNotes(
		"Abertura. Apresentação da necessidade de fertilizantes para concluir a adubação da safra, na posição de 16/10/2026. " +
		"Os números vêm da planilha 'Necessidade Fertilizante OUT 2026', que traz dois cenários em abas separadas: " +
		"LOC. 1.000ha (cenário A, 20 dias de vinhaça localizada com potássio adequado) e LOC. 450ha (cenário B, sem novos dias de localizada na CRV). " +
		"Todos os valores foram recalculados e conferidos com a planilha (87 conferências sem divergência)."
	);

	// ================================================================== 2. RESUMO EXECUTIVO
	s = novoSlide("CONTEUDO", "Abertura");
	titulos(s, "RESUMO EXECUTIVO", `Faltam ${fmt(A.compra_16_t)} t de 16-00-20 para fechar a safra`);
	const kpis = [
		{ ic: I.cana, num: fmt(A.total_a_tratar_ha), un: " ha", rot: "área ainda a adubar", sub: "igual nos dois cenários", cor: C.text2 },
		{ ic: I.saco, num: fmt(dA.necessidade_total_t), un: " t", rot: "necessidade total", sub: `cenário A · ${fmt(dB.necessidade_total_t)} t no B`, cor: C.text2 },
		{ ic: I.carrinho, num: fmt(A.compra_16_t), un: " t", rot: "16-00-20 a comprar", sub: `cenário A · ${fmt(B.compra_16_t)} t no B`, cor: C.accent4 },
		{ ic: I.armazem, num: "+" + fmt(b46A.saldo_t), un: " t", rot: "excedente de 46-00-00", sub: `cenário A · +${fmt(b46B.saldo_t)} t no B`, cor: C.text2 },
		{ ic: I.calendario, num: String(A.dias_safra), un: " dias", rot: "até o término da safra", sub: `${dataBR(A.termino_safra)}/2026`, cor: C.text2 },
	];
	const kw = 2.29, kg = 0.17;
	kpis.forEach((k, i) => {
		const x = 0.6 + i * (kw + kg), y = 1.62, nm = `s2_kpi${i + 1}`;
		caixa(s, { x, y, w: kw, h: 2.05, objectName: `${nm}_bg` });
		s.addImage({ data: k.ic, x: x + 0.2, y: y + 0.2, w: 0.5, h: 0.5, objectName: `${nm}_ic` });
		txt(s, [{ text: k.num, options: { fontSize: 32, bold: true, color: k.cor } }, { text: k.un, options: { fontSize: 16, bold: true, color: k.cor } }],
			{ x: x + 0.2, y: y + 0.82, w: kw - 0.3, h: 0.55, valign: "bottom", objectName: `${nm}_num` });
		txt(s, k.rot, { x: x + 0.2, y: y + 1.42, w: kw - 0.3, h: 0.28, fontSize: 13, bold: true, objectName: `${nm}_rot` });
		txt(s, k.sub, { x: x + 0.2, y: y + 1.7, w: kw - 0.3, h: 0.26, fontSize: 11, color: C.accent6, objectName: `${nm}_sub` });
		grp(nm, [`${nm}_bg`, `${nm}_ic`, `${nm}_num`, `${nm}_rot`, `${nm}_sub`]);
		fx(nm, "fade", "apos", { duracao: 0.4, atraso: i === 0 ? 0.2 : 0 });
	});
	txt(s, "O que os números mostram", { x: 0.6, y: 3.97, w: 6, h: 0.35, fontSize: 16, bold: true, color: C.text2, objectName: "s2_leit_t" });
	const leituras = [
		{ ic: I.saco, t: `${pct(dA.participacao_16)} do volume restante é 16-00-20; estoque e compras já feitas cobrem ${pct(dA.cobertura_16_disponivel)} dele.` },
		{ ic: I.armazemVerde, t: `46-00-00 sobra: há ${fmt(b46A.disponivel_t)} t disponíveis para ${fmt(b46A.necessidade_t)} t de uso previsto no cenário A.` },
		{ ic: I.alerta, t: "A planilha não traz preços nem datas de entrega: orçamento e risco de ruptura ainda estão em aberto." },
	];
	const mLeit = ["s2_leit_t"];
	leituras.forEach((l, i) => {
		const y = 4.45 + i * 0.7;
		s.addImage({ data: l.ic, x: 0.6, y: y + 0.02, w: 0.4, h: 0.4, objectName: `s2_leit${i}_ic` });
		txt(s, l.t, { x: 1.15, y, w: 5.35, h: 0.62, fontSize: 14, objectName: `s2_leit${i}_tx` });
		mLeit.push(`s2_leit${i}_ic`, `s2_leit${i}_tx`);
	});
	grp("s2_leitura", mLeit);
	fx("s2_leitura", "fade", "apos", { duracao: 0.5 });
	caixa(s, { x: 6.85, y: 3.97, w: 5.88, h: 2.55, fill: { color: C.text2 }, objectName: "s2_dec_bg" });
	txt(s, "Decisões para a diretoria", { x: 7.15, y: 4.15, w: 5.3, h: 0.35, fontSize: 16, bold: true, color: C.accent2, objectName: "s2_dec_t" });
	const decs = [
		`Aprovar a compra de 16-00-20: ${fmt(A.compra_16_t)} t (A) ou ${fmt(B.compra_16_t)} t (B)`,
		"Definir o destino do excedente de 46-00-00",
		`Exigir o cronograma de entrega das ${fmt(dA.a_receber_total_t)} t já compradas`,
	];
	const mDec = ["s2_dec_bg", "s2_dec_t"];
	decs.forEach((d, i) => {
		const y = 4.65 + i * 0.6;
		s.addShape(pres.shapes.OVAL, { x: 7.15, y: y + 0.02, w: 0.36, h: 0.36, fill: { color: C.accent3 }, objectName: `s2_dec${i}_n` });
		txt(s, String(i + 1), { x: 7.15, y: y + 0.02, w: 0.36, h: 0.36, fontSize: 14, bold: true, color: C.text2, align: "center", valign: "middle", objectName: `s2_dec${i}_nt` });
		txt(s, d, { x: 7.68, y, w: 4.85, h: 0.55, fontSize: 14, color: C.background1, objectName: `s2_dec${i}_tx` });
		mDec.push(`s2_dec${i}_n`, `s2_dec${i}_nt`, `s2_dec${i}_tx`);
	});
	grp("s2_decisoes", mDec);
	fx("s2_decisoes", "wipe", "clique", { direcao: "esquerda", duracao: 0.6 });
	fonte(s, `${ARQ} — ${ABA_A} (cenário A) e ${ABA_B} (cenário B), quadros 3 (áreas remanescentes) e 4 (resumo).`);
	s.addNotes(
		`Mensagem principal: para concluir a adubação da safra faltam ${fmt(A.compra_16_t, 1)} t de 16-00-20 no cenário A (${fmt(B.compra_16_t, 1)} t no B). ` +
		`Restam ${fmt(A.total_a_tratar_ha, 2)} ha a adubar (H40), iguais nos dois cenários; muda só o produto. ` +
		`Necessidade total: ${fmt(dA.necessidade_total_t, 1)} t no A (K38+K39) e ${fmt(dB.necessidade_total_t, 1)} t no B. ` +
		`O 46-00-00 tem ${fmt(b46A.disponivel_t, 2)} t disponíveis (estoque ${fmt(b46A.estoque_t, 2)} + a receber ${fmt(b46A.a_receber_t)}) para uso de ${fmt(b46A.necessidade_t, 1)} t. ` +
		"O quadro de decisões entra no clique."
	);

	// ================================================================== 3. VISÃO GERAL
	pres.addSection({ title: "Demanda" });
	s = novoSlide("CONTEUDO", "Demanda");
	titulos(s, "VISÃO GERAL DA DEMANDA", `${fmt(A.total_a_tratar_ha)} ha ainda precisam de adubação nesta safra`);
	const tratos = dA.area.tratos_total_ha;
	const bisada = A.crv.bisada_ha;
	txt(s, "Da área de tratos culturais à área a adubar (ha)", { x: 0.6, y: 1.62, w: 7.6, h: 0.32, fontSize: 14, bold: true, color: C.text2, objectName: "s3_cap" });
	s.addChart(pres.charts.BAR, [
		{ name: "apoio da cascata (oculto)", labels: ["Área de tratos culturais", "Já adubado", "Cana bisada", "A adubar"], values: [0, tratos - realizadoTotal, tratos - realizadoTotal - bisada, 0] },
		{ name: "Área (ha)", labels: ["Área de tratos culturais", "Já adubado", "Cana bisada", "A adubar"], values: [tratos, realizadoTotal, bisada, A.total_a_tratar_ha] },
	], graf({
		x: 0.6, y: 1.98, w: 7.6, h: 4.55, barDir: "col", barGrouping: "stacked", barGapWidthPct: 45,
		chartColors: ["FFFFFF", HEX.accent1], valAxisHidden: true, dataLabelPosition: "ctr", valAxisMinVal: 0, valAxisMaxVal: 31000,
		dataLabelFormatCode: "#,##0", dataLabelColor: "FFFFFF", dataLabelFontSize: 14, catAxisLabelFontSize: 13,
		objectName: "s3_graf",
	}));
	ajusteGrafico("s3_graf", {
		ocultarSerie: 1,
		corPontos: { serie: 2, cores: [HEX.dk2, HEX.accent2, EXTRA.cinzaClaro, HEX.accent3] },
		corRotulos: { serie: 2, cores: ["FFFFFF", HEX.dk2, HEX.dk2, HEX.dk2] },
	});
	fx("s3_cap", "fade", "apos", { duracao: 0.3 });
	fx("s3_graf", "wipe", "apos", { grafico: "categoria", etapas: 4, direcao: "baixo", duracao: 0.5 });
	const stats3 = [
		{ v: pct(dA.area.pct_realizado_total), cor: C.accent1, t: `da área de tratos já foi adubada: ${fmt(realizadoTotal)} ha (CRV ${fmt(A.crv.realizado_ha)} + Cargill ${fmt(A.cargill.realizado_ha)})` },
		{ v: pct(bisada / tratos), cor: C.accent6, t: `fica como cana bisada: ${fmt(bisada)} ha, ~${fmt(A.crv.bisada_t / 1000)} mil t para a próxima safra` },
		{ v: pct(A.total_a_tratar_ha / tratos), cor: C.text2, t: `ainda a adubar: ${fmt(A.total_a_tratar_ha)} ha em ${A.dias_safra} dias de safra` },
	];
	stats3.forEach((st, i) => {
		const y = 1.62 + i * 1.2;
		txt(s, st.v, { x: 8.6, y, w: 1.55, h: 0.75, fontSize: 36, bold: true, color: st.cor, valign: "middle", objectName: `s3_st${i}_v` });
		txt(s, st.t, { x: 10.2, y: y + 0.02, w: 2.53, h: 1.05, fontSize: 13, valign: "top", objectName: `s3_st${i}_t` });
		grp(`s3_st${i}`, [`s3_st${i}_v`, `s3_st${i}_t`]);
		fx(`s3_st${i}`, "fade", "apos", { duracao: 0.4 });
	});
	caixa(s, { x: 8.6, y: 5.3, w: 4.13, h: 1.23, fill: { color: C.accent3, transparency: 82 }, objectName: "s3_nota_bg" });
	s.addImage({ data: I.alerta, x: 8.78, y: 5.45, w: 0.4, h: 0.4, objectName: "s3_nota_ic" });
	txt(s, "A área a adubar é a mesma nos dois cenários. O cenário muda só o produto: vinhaça + 46-00-00 ou 16-00-20.",
		{ x: 9.32, y: 5.42, w: 3.28, h: 1.0, fontSize: 13, objectName: "s3_nota_tx" });
	grp("s3_nota", ["s3_nota_bg", "s3_nota_ic", "s3_nota_tx"]);
	fx("s3_nota", "fade", "clique", { duracao: 0.5 });
	fonte(s, `${ARQ} — ${ABA_A}, quadro 3: E40 (área de tratos), F40 e F36 (adubado), G39/A45 (cana bisada), H40 (a adubar).`);
	s.addNotes(
		`A área de tratos culturais soma ${fmt(tratos)} ha (CRV ${fmt(A.crv.tratos_ha)} = área à colher ${fmt(A.crv.a_colher_ha)} menos ${fmt(A.crv.plantio_ha)} ha de renovação; Cargill ${fmt(A.cargill.tratos_ha)} ha). ` +
		`Já foram adubados ${fmt(realizadoTotal, 2)} ha. A cana bisada (${fmt(bisada, 2)} ha × ${A.crv.tch_bisada} t/ha ≈ ${fmt(A.crv.bisada_t)} t) não recebe adubação nesta safra. ` +
		`Restam ${fmt(A.total_a_tratar_ha, 2)} ha. Esse total não muda entre os cenários: o que muda é quanto vai para vinhaça localizada (com 46-00-00) e quanto vai para fertilizante mineral (16-00-20).`
	);

	// ================================================================== 4. POR PRODUTO
	s = novoSlide("CONTEUDO", "Demanda");
	titulos(s, "NECESSIDADE POR PRODUTO", `16-00-20 responde por ${pct(dA.participacao_16)} da necessidade restante`);
	txt(s, "Necessidade restante por fórmula (t)", { x: 0.6, y: 1.62, w: 7.3, h: 0.32, fontSize: 14, bold: true, color: C.text2, objectName: "s4_cap" });
	s.addChart(pres.charts.BAR, [
		{ name: "Cenário A · localizada 1.000 ha", labels: ["16-00-20", "46-00-00"], values: [b16A.necessidade_t, b46A.necessidade_t] },
		{ name: "Cenário B · localizada 450 ha", labels: ["16-00-20", "46-00-00"], values: [b16B.necessidade_t, b46B.necessidade_t] },
	], graf({
		x: 0.6, y: 1.98, w: 7.3, h: 3.25, barDir: "bar", barGrouping: "clustered", barGapWidthPct: 55,
		chartColors: [HEX.accent1, HEX.accent2], valAxisHidden: true, catAxisOrientation: "maxMin",
		dataLabelPosition: "outEnd", dataLabelFormatCode: '#,##0" t"', catAxisLabelFontSize: 15,
		showLegend: true, legendPos: "t", valAxisMaxVal: 7000, valAxisMinVal: 0,
		objectName: "s4_graf",
	}));
	fx("s4_cap", "fade", "apos", { duracao: 0.3 });
	fx("s4_graf", "wipe", "apos", { grafico: "serie", etapas: 2, direcao: "esquerda", duracao: 0.6 });
	const part = [
		{ v: pct(dA.participacao_16, 1), t: "do volume é 16-00-20 no cenário A" },
		{ v: pct(dB.participacao_16), t: "no cenário B, sem localizada na CRV" },
	];
	part.forEach((p, i) => {
		const x = 0.6 + i * 3.7;
		txt(s, p.v, { x, y: 5.45, w: 1.75, h: 0.7, fontSize: 32, bold: true, color: C.accent1, valign: "middle", objectName: `s4_p${i}_v` });
		txt(s, p.t, { x: x + 1.8, y: 5.5, w: 1.8, h: 0.75, fontSize: 13, valign: "middle", objectName: `s4_p${i}_t` });
		grp(`s4_p${i}`, [`s4_p${i}_v`, `s4_p${i}_t`]);
		fx(`s4_p${i}`, "fade", "apos", { duracao: 0.4 });
	});
	const prods = [
		{ ic: I.saco, nome: "16-00-20", desc: "Fórmula NK para as áreas sem vinhaça localizada", dose: `Dose média ${fmt(A.doses_kg_ha["16-00-20"])} kg/ha · ${fmt(dA.nutrientes_por_ha["16-00-20"].N)} kg de N + ${fmt(dA.nutrientes_por_ha["16-00-20"].K2O)} kg de K₂O`, vol: `A: ${fmt(b16A.necessidade_t)} t  ·  B: ${fmt(b16B.necessidade_t)} t` },
		{ ic: I.gota, nome: "46-00-00", desc: "Fonte de N nas áreas com vinhaça localizada — o K vem da vinhaça", dose: `Dose média ${fmt(A.doses_kg_ha["46-00-00"])} kg/ha · ${fmt(dA.nutrientes_por_ha["46-00-00"].N)} kg de N`, vol: `A: ${fmt(b46A.necessidade_t)} t  ·  B: ${fmt(b46B.necessidade_t)} t` },
	];
	prods.forEach((p, i) => {
		const y = 1.62 + i * 2.48, nm = `s4_prod${i}`;
		caixa(s, { x: 8.3, y, w: 4.43, h: 2.3, objectName: `${nm}_bg` });
		s.addImage({ data: p.ic, x: 8.5, y: y + 0.22, w: 0.55, h: 0.55, objectName: `${nm}_ic` });
		txt(s, p.nome, { x: 9.22, y: y + 0.22, w: 3.3, h: 0.55, fontSize: 22, bold: true, color: C.text2, valign: "middle", objectName: `${nm}_n` });
		txt(s, p.desc, { x: 8.5, y: y + 0.85, w: 4.05, h: 0.55, fontSize: 13, objectName: `${nm}_d` });
		txt(s, p.dose, { x: 8.5, y: y + 1.47, w: 4.05, h: 0.3, fontSize: 12, color: C.accent6, objectName: `${nm}_dose` });
		txt(s, p.vol, { x: 8.5, y: y + 1.85, w: 4.05, h: 0.32, fontSize: 15, bold: true, color: C.accent1, objectName: `${nm}_v` });
		grp(nm, [`${nm}_bg`, `${nm}_ic`, `${nm}_n`, `${nm}_d`, `${nm}_dose`, `${nm}_v`]);
		fx(nm, "fade", "apos", { duracao: 0.45 });
	});
	fonte(s, `${ARQ} — ${ABA_A} e ${ABA_B}: K38:K39 e B50:B51. Demais fórmulas em estoque não têm necessidade calculada na planilha.`);
	s.addNotes(
		`Só duas fórmulas têm necessidade calculada: 16-00-20 (${fmt(b16A.necessidade_t, 1)} t no A; ${fmt(b16B.necessidade_t, 1)} t no B) e 46-00-00 (${fmt(b46A.necessidade_t, 1)} t no A; zero no B). ` +
		`O 16-00-20 vai para as áreas de fertilizante mineral a 500 kg/ha; o 46-00-00 complementa o N nas áreas de vinhaça localizada a 200 kg/ha, porque o potássio vem da própria vinhaça. ` +
		"Nutrientes calculados pela garantia da fórmula (N-P₂O₅-K₂O em %)."
	);

	// ================================================================== 5. POR OPERAÇÃO E UNIDADE
	s = novoSlide("CONTEUDO", "Demanda");
	titulos(s, "NECESSIDADE POR OPERAÇÃO E UNIDADE", `Adubação mineral cobre ${pct(A.crv.a_tratar_mineral_ha / A.total_a_tratar_ha)} a ${pct(B.crv.a_tratar_mineral_ha / B.total_a_tratar_ha)} da área restante`);
	txt(s, "Área a adubar por operação (ha)", { x: 0.6, y: 1.62, w: 6, h: 0.32, fontSize: 14, bold: true, color: C.text2, objectName: "s5_cap" });
	const cats5 = ["Cenário A · localizada 1.000 ha", "Cenário B · localizada 450 ha"];
	s.addChart(pres.charts.BAR, [
		{ name: "Vinhaça localizada Cargill", labels: cats5, values: [A.cargill.a_tratar_ha, B.cargill.a_tratar_ha] },
		{ name: "Vinhaça localizada CRV (46-00-00)", labels: cats5, values: [A.crv.a_tratar_localizada_ha, B.crv.a_tratar_localizada_ha] },
		{ name: "Fertilizante mineral CRV (16-00-20)", labels: cats5, values: [A.crv.a_tratar_mineral_ha, B.crv.a_tratar_mineral_ha] },
	], graf({
		x: 0.6, y: 1.95, w: 12.13, h: 2.05, barDir: "bar", barGrouping: "stacked", barGapWidthPct: 40,
		chartColors: [EXTRA.vinhacaClara, HEX.accent5, HEX.accent1], valAxisHidden: true, catAxisOrientation: "maxMin",
		dataLabelPosition: "ctr", dataLabelFormatCode: "#,##0;;;", dataLabelColor: "FFFFFF", catAxisLabelFontSize: 13,
		showLegend: true, legendPos: "t", objectName: "s5_graf",
	}));
	fx("s5_cap", "fade", "apos", { duracao: 0.3 });
	fx("s5_graf", "wipe", "apos", { grafico: "categoria", etapas: 2, direcao: "esquerda", duracao: 0.6 });
	const th = (t, al = "left") => ({ text: t, options: { bold: true, color: "FFFFFF", fill: { color: HEX.dk2 }, align: al } });
	const td = (t, al = "left", o = {}) => ({ text: t, options: Object.assign({ align: al }, o) });
	const tot = { bold: true, fill: { color: HEX.lt2 } };
	s.addTable([
		[th("Operação"), th("Unidade"), th("Fertilizante"), th("Dose", "right"), th("Área A (ha)", "right"), th("Volume A (t)", "right"), th("Área B (ha)", "right"), th("Volume B (t)", "right")],
		[td("Vinhaça localizada"), td("Cargill"), td("sem fertilizante na planilha"), td("—", "right"), td(fmt(A.cargill.a_tratar_ha, 2), "right"), td("—", "right"), td(fmt(B.cargill.a_tratar_ha, 2), "right"), td("—", "right")],
		[td("Vinhaça localizada"), td("CRV"), td("46-00-00"), td("200 kg/ha", "right"), td(fmt(A.crv.a_tratar_localizada_ha, 2), "right"), td(fmt(b46A.necessidade_t, 1), "right"), td(fmt(B.crv.a_tratar_localizada_ha, 2), "right"), td(fmt(b46B.necessidade_t, 1), "right")],
		[td("Fertilizante mineral"), td("CRV"), td("16-00-20"), td("500 kg/ha", "right"), td(fmt(A.crv.a_tratar_mineral_ha, 2), "right"), td(fmt(b16A.necessidade_t, 1), "right"), td(fmt(B.crv.a_tratar_mineral_ha, 2), "right"), td(fmt(b16B.necessidade_t, 1), "right")],
		[td("Total", "left", tot), td("", "left", tot), td("", "left", tot), td("", "right", tot), td(fmt(A.total_a_tratar_ha, 2), "right", tot), td(fmt(dA.necessidade_total_t, 1), "right", tot), td(fmt(B.total_a_tratar_ha, 2), "right", tot), td(fmt(dB.necessidade_total_t, 1), "right", tot)],
	], {
		x: 0.6, y: 4.2, w: 12.13, colW: [2.45, 1.05, 2.63, 1.15, 1.25, 1.25, 1.2, 1.15], rowH: 0.36,
		fontSize: 13, color: HEX.dk1, valign: "middle", border: { type: "solid", pt: 0.75, color: EXTRA.linhaTabela },
		margin: [0, 0.08, 0, 0.08], objectName: "s5_tab",
	});
	fx("s5_tab", "fade", "apos", { duracao: 0.5 });
	s.addImage({ data: I.mapa, x: 0.6, y: 6.12, w: 0.36, h: 0.36, objectName: "s5_nota_ic" });
	txt(s, `Sem detalhe por fazenda ou bloco: a área à colher da CRV soma ${A.crv.a_colher_partes.map((v) => fmt(v)).join(" + ")} ha sem identificação. A logística por frente depende desse detalhe.`,
		{ x: 1.1, y: 6.08, w: 11.6, h: 0.48, fontSize: 13, color: C.accent6, valign: "middle", objectName: "s5_nota_tx" });
	grp("s5_nota", ["s5_nota_ic", "s5_nota_tx"]);
	fx("s5_nota", "fade", "apos", { duracao: 0.4 });
	fonte(s, `${ARQ} — ${ABA_A} e ${ABA_B}, quadro 3: H36, H38:H40, I38:K39.`);
	s.addNotes(
		"A área restante se divide em três operações. A vinhaça localizada da Cargill (446,63 ha) não tem fertilizante associado na planilha. " +
		`No cenário A, ${fmt(A.crv.a_tratar_localizada_ha, 2)} ha da CRV recebem vinhaça localizada com 46-00-00; no B, essa área passa para fertilizante mineral. ` +
		"A planilha não detalha fazendas ou blocos, então não dá para separar a necessidade por frente ou por local de entrega."
	);

	// ================================================================== 6. ÁREA × CONSUMO
	s = novoSlide("CONTEUDO", "Demanda");
	titulos(s, "RELAÇÃO ENTRE ÁREA E CONSUMO", "Cada hectare com vinhaça poupa 500 kg de 16-00-20");
	txt(s, "Nutrientes entregues por hectare (kg/ha)", { x: 0.6, y: 1.62, w: 6.3, h: 0.32, fontSize: 14, bold: true, color: C.text2, objectName: "s6_cap" });
	s.addChart(pres.charts.BAR, [
		{ name: "N", labels: ["16-00-20 · 500 kg/ha", "46-00-00 · 200 kg/ha"], values: [dA.nutrientes_por_ha["16-00-20"].N, dA.nutrientes_por_ha["46-00-00"].N] },
		{ name: "K₂O", labels: ["16-00-20 · 500 kg/ha", "46-00-00 · 200 kg/ha"], values: [dA.nutrientes_por_ha["16-00-20"].K2O, 0] },
	], graf({
		x: 0.6, y: 1.98, w: 6.3, h: 3.9, barDir: "col", barGrouping: "clustered", barGapWidthPct: 70,
		chartColors: [HEX.accent1, HEX.accent3], valAxisHidden: true, valAxisMinVal: 0, valAxisMaxVal: 120,
		dataLabelPosition: "outEnd", dataLabelFormatCode: '#,##0" kg";;;', catAxisLabelFontSize: 14,
		showLegend: true, legendPos: "t", objectName: "s6_graf",
	}));
	fx("s6_cap", "fade", "apos", { duracao: 0.3 });
	fx("s6_graf", "wipe", "apos", { grafico: "categoria", etapas: 2, direcao: "baixo", duracao: 0.5 });
	txt(s, "No 46-00-00, o K₂O vem da vinhaça localizada — o teor de K da vinhaça não consta na planilha.",
		{ x: 0.6, y: 5.95, w: 6.3, h: 0.55, fontSize: 12, italic: true, color: C.accent6, objectName: "s6_obs" });
	fx("s6_obs", "fade", "com", { duracao: 0.4 });
	txt(s, "Troca por hectare levado à vinhaça localizada", { x: 7.35, y: 1.62, w: 5.38, h: 0.32, fontSize: 16, bold: true, color: C.text2, objectName: "s6_troca_t" });
	const troca = [
		{ v: "−500 kg", cor: C.accent4, t: "de 16-00-20, o produto em falta" },
		{ v: "+200 kg", cor: C.accent5, t: "de 46-00-00, que já sobra no estoque" },
		{ v: "−300 kg", cor: C.accent1, t: "de produto comercial por hectare" },
	];
	const mTroca = ["s6_troca_t"];
	troca.forEach((t, i) => {
		const y = 2.08 + i * 0.66;
		txt(s, t.v, { x: 7.35, y, w: 1.75, h: 0.55, fontSize: 24, bold: true, color: t.cor, valign: "middle", objectName: `s6_tr${i}_v` });
		txt(s, t.t, { x: 9.15, y, w: 3.58, h: 0.55, fontSize: 14, valign: "middle", objectName: `s6_tr${i}_t` });
		mTroca.push(`s6_tr${i}_v`, `s6_tr${i}_t`);
	});
	grp("s6_troca", mTroca);
	fx("s6_troca", "fade", "apos", { duracao: 0.5 });
	txt(s, "Nutrientes da necessidade restante (cenário A)", { x: 7.35, y: 4.2, w: 5.38, h: 0.32, fontSize: 16, bold: true, color: C.text2, objectName: "s6_nut_t" });
	const nut = [
		{ v: fmt(dA.nutrientes_t.N), u: " t de N", s: `${fmt(dB.nutrientes_t.N)} t no cenário B` },
		{ v: fmt(dA.nutrientes_t.K2O), u: " t de K₂O", s: `${fmt(dB.nutrientes_t.K2O)} t no cenário B` },
	];
	const mNut = ["s6_nut_t"];
	nut.forEach((u, i) => {
		const x = 7.35 + i * 2.78, nm = `s6_nut${i}`;
		caixa(s, { x, y: 4.62, w: 2.6, h: 1.15, objectName: `${nm}_bg` });
		txt(s, [{ text: u.v, options: { fontSize: 26, bold: true, color: C.text2 } }, { text: u.u, options: { fontSize: 14, bold: true, color: C.text2 } }],
			{ x: x + 0.18, y: 4.72, w: 2.3, h: 0.5, valign: "middle", objectName: `${nm}_v` });
		txt(s, u.s, { x: x + 0.18, y: 5.25, w: 2.3, h: 0.35, fontSize: 12, color: C.accent6, objectName: `${nm}_s` });
		mNut.push(`${nm}_bg`, `${nm}_v`, `${nm}_s`);
	});
	grp("s6_nut", mNut);
	fx("s6_nut", "fade", "apos", { duracao: 0.5 });
	const cr = A.consumo_realizado;
	txt(s, `Já aplicado com vinhaça localizada: ${fmt(cr[0].volume_t)} t de ${cr[0].produto} em ${fmt(cr[0].area_ha)} ha (${cr[0].fazendas}) e ${fmt(cr[1].volume_t)} t de ${cr[1].produto} em ${fmt(cr[1].area_ha)} ha (${cr[1].fazendas}).`,
		{ x: 7.35, y: 5.92, w: 5.38, h: 0.6, fontSize: 12, color: C.accent6, objectName: "s6_real" });
	fx("s6_real", "fade", "apos", { duracao: 0.4 });
	fonte(s, `${ARQ} — doses J38:J39; consumo realizado nas linhas 27–28 (E27:E28). Teores pela garantia da fórmula (N-P₂O₅-K₂O, %).`);
	s.addNotes(
		"Comparação por hectare: 500 kg de 16-00-20 entregam 80 kg de N e 100 kg de K₂O; 200 kg de 46-00-00 entregam 92 kg de N e o potássio vem da vinhaça. " +
		"Por isso, cada hectare que vai para a vinhaça localizada deixa de consumir 500 kg de 16-00-20 (o produto em falta) e consome 200 kg de 46-00-00 (o produto que sobra). " +
		"A conversão em nutrientes usa apenas a garantia expressa no nome da fórmula; o teor de K da vinhaça não está na planilha e não foi estimado."
	);

	// ================================================================== 7. JANELAS NO TEMPO
	s = novoSlide("CONTEUDO", "Demanda");
	titulos(s, "DISTRIBUIÇÃO NO TEMPO", `Toda a demanda restante cai em ${A.dias_safra} dias de safra`);
	const x0 = 1.0, x1 = 12.3, dias = A.dias_safra;
	const xd = (d) => x0 + (x1 - x0) * (d / dias);
	const fimK = A.dias_k_vinhaca;
	const isoMais = (iso, d) => { const t = new Date(iso + "T12:00:00Z"); t.setUTCDate(t.getUTCDate() + d); return t.toISOString().slice(0, 10); };
	const marcos = [
		{ d: 0, data: dataBR(A.data_base), t: "Posição de estoque e áreas", al: "left" },
		{ d: fimK, data: "~" + dataBR(isoMais(A.data_base, fimK)), t: `Fim dos ${fimK} dias de vinhaça com K adequado (cenário A)`, al: "center" },
		{ d: A.dias_limite_localizada, data: dataBR(A.data_limite_localizada), t: "Data limite da vinhaça localizada", al: "center" },
		{ d: dias, data: dataBR(A.termino_safra), t: "Término da safra", al: "right" },
	];
	s.addShape(pres.shapes.LINE, { x: x0, y: 2.3, w: x1 - x0, h: 0, line: { color: C.text2, width: 2.5 }, objectName: "s7_eixo" });
	fx("s7_eixo", "wipe", "apos", { direcao: "esquerda", duracao: 0.7 });
	marcos.forEach((m, i) => {
		const cx = xd(m.d), nm = `s7_m${i}`;
		const larg = i === 3 ? 1.7 : 2.5;
		const bx = m.al === "left" ? cx - 0.14 : m.al === "right" ? cx + 0.14 - larg : cx - larg / 2;
		txt(s, m.data, { x: bx, y: 1.6, w: larg, h: 0.4, fontSize: 18, bold: true, color: C.text2, align: m.al, valign: "bottom", objectName: `${nm}_d` });
		s.addShape(pres.shapes.OVAL, { x: cx - 0.14, y: 2.16, w: 0.28, h: 0.28, fill: { color: i === 3 ? C.accent4 : C.accent1 }, line: { color: "FFFFFF", width: 2 }, objectName: `${nm}_p` });
		txt(s, m.t, { x: bx, y: 2.55, w: larg, h: 0.62, fontSize: 13, color: C.text1, align: m.al, objectName: `${nm}_t` });
		grp(nm, [`${nm}_d`, `${nm}_p`, `${nm}_t`]);
		fx(nm, "fade", "apos", { duracao: 0.35 });
	});
	const faixas = [
		{ d0: 0, d1: fimK, y: 3.45, cor: C.accent5, txtCor: C.background1, t: `Localizada (A): ${fmt(A.area_localizada_estimada_ha)} ha = ${fmt(A.ritmo_localizada_ha_dia, 2)} ha/dia × ${fimK} dias` },
		{ d0: 0, d1: A.dias_limite_localizada, y: 3.93, cor: null, txtCor: C.accent5, t: `Prazo da vinhaça localizada: até ${dataBR(A.data_limite_localizada)} (${A.dias_limite_localizada} dias)` },
		{ d0: 0, d1: dias, y: 4.41, cor: C.accent1, txtCor: C.background1, t: `Adubação mineral com 16-00-20: ${fmt(A.crv.a_tratar_mineral_ha)} ha (A) ou ${fmt(B.crv.a_tratar_mineral_ha)} ha (B) — ritmo não informado na planilha` },
	];
	faixas.forEach((f, i) => {
		const nm = `s7_f${i}`;
		const opts = { x: xd(f.d0), y: f.y, w: xd(f.d1) - xd(f.d0), h: 0.38, rectRadius: 0.05, objectName: `${nm}_bg` };
		opts.fill = f.cor ? { color: f.cor } : { color: C.accent5, transparency: 85 };
		if (!f.cor) opts.line = { color: C.accent5, width: 1, dashType: "dash" };
		caixa(s, opts);
		txt(s, f.t, { x: xd(f.d0) + 0.12, y: f.y, w: xd(f.d1) - xd(f.d0) - 0.2, h: 0.38, fontSize: 12, bold: true, color: f.txtCor, valign: "middle", objectName: `${nm}_t` });
		grp(nm, [`${nm}_bg`, `${nm}_t`]);
		fx(nm, "wipe", "apos", { direcao: "esquerda", duracao: 0.5 });
	});
	const cards7 = [
		{ v: `${fmt(dA.area.ritmo_para_fechar_na_safra_ha_dia)} ha/dia`, t: `ritmo médio para adubar os ${fmt(A.total_a_tratar_ha)} ha até ${dataBR(A.termino_safra)} (referência, não meta da planilha)` },
		{ v: `${fmt(A.ritmo_localizada_ha_dia, 2)} ha/dia`, t: "ritmo da vinhaça localizada, média atingida em setembro" },
		{ v: "Sem dado mensal", t: "a planilha não distribui a demanda por mês; é preciso a programação de colheita" },
	];
	cards7.forEach((c, i) => {
		const x = 0.6 + i * 4.13, nm = `s7_c${i}`;
		caixa(s, { x, y: 5.15, w: 3.87, h: 1.38, objectName: `${nm}_bg` });
		txt(s, c.v, { x: x + 0.22, y: 5.27, w: 3.45, h: 0.5, fontSize: 24, bold: true, color: i === 2 ? C.accent6 : C.text2, valign: "middle", objectName: `${nm}_v` });
		txt(s, c.t, { x: x + 0.22, y: 5.8, w: 3.45, h: 0.66, fontSize: 12, objectName: `${nm}_t` });
		grp(nm, [`${nm}_bg`, `${nm}_v`, `${nm}_t`]);
		fx(nm, "fade", "apos", { duracao: 0.4 });
	});
	fonte(s, `${ARQ} — ${ABA_A}, quadro 1: A10:C14 (datas, dias, ritmo de ${fmt(A.ritmo_localizada_ha_dia, 2)} ha/dia e ${fimK} dias de K adequado).`);
	s.addNotes(
		`Calendário da planilha: posição em 16/10; data limite da vinhaça localizada em 30/11 (45 dias); término da safra em 15/12 (60 dias). ` +
		`No cenário A, a vinhaça localizada atende o nível de potássio por ${fimK} dias, a ${fmt(A.ritmo_localizada_ha_dia, 2)} ha/dia (média de setembro), o que dá ${fmt(A.area_localizada_estimada_ha, 1)} ha — dos quais 446,63 ha são da Cargill. ` +
		`A planilha não informa o ritmo da adubação mineral nem a distribuição mensal da demanda. O ritmo de ${fmt(dA.area.ritmo_para_fechar_na_safra_ha_dia)} ha/dia é apenas a conta de ${fmt(A.total_a_tratar_ha, 2)} ha ÷ 60 dias, como referência.`
	);

	// ================================================================== 8. CURVA DE ABASTECIMENTO
	pres.addSection({ title: "Abastecimento" });
	s = novoSlide("CONTEUDO", "Abastecimento");
	const cv = dA.curva_16;
	titulos(s, "CURVA DE ABASTECIMENTO · 16-00-20 · CENÁRIO A", `16-00-20 disponível se esgota por volta de ${dataBR(cv.data_esgota_disponivel)}`);
	txt(s, "Necessidade acumulada × disponível de 16-00-20 (t)", { x: 0.6, y: 1.62, w: 8.2, h: 0.32, fontSize: 14, bold: true, color: C.text2, objectName: "s8_cap" });
	const labs8 = cv.pontos.map((p) => dataBR(p.data));
	s.addChart(pres.charts.LINE, [
		{ name: "Necessidade acumulada (premissa)", labels: labs8, values: cv.pontos.map((p) => p.necessidade_acum_t) },
		{ name: `Estoque em 16/10 (${fmt(b16A.estoque_t)} t)`, labels: labs8, values: cv.pontos.map(() => b16A.estoque_t) },
		{ name: `Estoque + comprado a receber (${fmt(b16A.disponivel_t)} t)`, labels: labs8, values: cv.pontos.map(() => b16A.disponivel_t) },
	], graf({
		x: 0.6, y: 1.98, w: 8.2, h: 4.0, chartColors: [HEX.accent1, HEX.accent3, HEX.accent2], lineSize: 3, lineDataSymbol: "none",
		showValue: false, valAxisHidden: false, valAxisLabelFormatCode: "#,##0", valAxisMinVal: 0, valAxisMaxVal: 6000, valAxisMajorUnit: 1000,
		valGridLine: { color: EXTRA.grade, size: 0.75 }, showLegend: true, legendPos: "b", catAxisLabelFontSize: 12,
		objectName: "s8_graf",
	}));
	ajusteGrafico("s8_graf", { tracejado: [2, 3], espessura: { 1: 3.5, 2: 2.25, 3: 2.25 }, rotuloUltimo: [1] });
	fx("s8_cap", "fade", "apos", { duracao: 0.3 });
	fx("s8_graf", "wipe", "apos", { direcao: "esquerda", duracao: 1.2 });
	const call8 = [
		{ v: "~" + dataBR(cv.data_esgota_estoque), cor: C.text2, t: `acaba o estoque atual de ${fmt(b16A.estoque_t)} t (cerca de ${fmt(cv.dias_cobertura_estoque, 1)} dias de consumo)` },
		{ v: "~" + dataBR(cv.data_esgota_disponivel), cor: C.text2, t: `acaba o disponível (${fmt(b16A.disponivel_t)} t), se as ${fmt(b16A.a_receber_t)} t a receber chegarem até lá` },
		{ v: `${fmt(A.compra_16_t)} t`, cor: C.accent4, t: `precisam chegar antes de ~${dataBR(cv.data_esgota_disponivel)}, descontado o prazo do fornecedor (não informado)` },
	];
	call8.forEach((c, i) => {
		const y = 1.62 + i * 1.47, nm = `s8_c${i}`;
		caixa(s, { x: 9.15, y, w: 3.58, h: 1.32, fill: i === 2 ? { color: C.accent4, transparency: 88 } : { color: C.background2 }, objectName: `${nm}_bg` });
		txt(s, c.v, { x: 9.35, y: y + 0.12, w: 3.2, h: 0.5, fontSize: 26, bold: true, color: c.cor, valign: "middle", objectName: `${nm}_v` });
		txt(s, c.t, { x: 9.35, y: y + 0.64, w: 3.25, h: 0.62, fontSize: 12, objectName: `${nm}_t` });
		grp(nm, [`${nm}_bg`, `${nm}_v`, `${nm}_t`]);
		fx(nm, "fade", i === 2 ? "clique" : "apos", { duracao: 0.45 });
	});
	txt(s, `Premissa ilustrativa: ${fmt(b16A.necessidade_t, 1)} t aplicadas por igual em ${dias} dias (${fmt(cv.consumo_diario_t, 1)} t/dia). A curva real depende da programação de colheita e do ritmo de aplicação, que não estão na planilha. No cenário B, as datas antecipam para ~${dataBR(dB.curva_16.data_esgota_estoque)} e ~${dataBR(dB.curva_16.data_esgota_disponivel)}.`,
		{ x: 0.6, y: 6.05, w: 12.13, h: 0.52, fontSize: 12, italic: true, color: C.accent6, objectName: "s8_prem" });
	fx("s8_prem", "fade", "com", { duracao: 0.4 });
	fonte(s, `${ARQ} — ${ABA_A}: K39 (necessidade), C51 (estoque), D51 (comprado a receber), C11 (dias até o término da safra).`);
	s.addNotes(
		"Esta curva não é previsão da usina: a planilha não traz a demanda mês a mês nem as datas de entrega. Para ter uma ordem de grandeza, a necessidade de 16-00-20 do cenário A foi distribuída por igual nos 60 dias até o término da safra. " +
		`Nesse ritmo (${fmt(cv.consumo_diario_t, 1)} t/dia), o estoque de 1.132 t dura cerca de ${fmt(cv.dias_cobertura_estoque, 1)} dias (até ~${dataBR(cv.data_esgota_estoque)}); somando as 2.030 t compradas a receber, a cobertura vai até ~${dataBR(cv.data_esgota_disponivel)}. ` +
		`A compra adicional de ${fmt(A.compra_16_t)} t precisa chegar antes disso; a antecedência do pedido depende do prazo do fornecedor, que é variável a definir. O terceiro cartão entra no clique.`
	);

	// ================================================================== 9. ESTOQUE POR FÓRMULA
	s = novoSlide("CONTEUDO", "Abastecimento");
	const est = [...A.estoque].sort((a, b) => b.estoque_t - a.estoque_t);
	const corGrupo = { Socaria: HEX.accent1, Plantio: HEX.accent2, "Não classificado": EXTRA.cinzaClaro };
	titulos(s, "ESTOQUE EM 16/10/2026", `Estoque de ${fmt(dA.estoque_total_t)} t: ${pct(dA.estoque_por_grupo_t.Socaria / dA.estoque_total_t)} em fórmulas de socaria`);
	txt(s, "Estoque por fórmula (t)", { x: 0.6, y: 1.62, w: 7.5, h: 0.32, fontSize: 14, bold: true, color: C.text2, objectName: "s9_cap" });
	s.addChart(pres.charts.BAR, [{ name: "Estoque (t)", labels: est.map((e) => e.produto), values: est.map((e) => e.estoque_t) }], graf({
		x: 0.6, y: 1.95, w: 7.5, h: 4.15, barDir: "bar", barGapWidthPct: 35, chartColors: est.map((e) => corGrupo[e.grupo]),
		valAxisHidden: true, catAxisOrientation: "maxMin", dataLabelPosition: "outEnd", dataLabelFormatCode: '#,##0.0" t"',
		valAxisMinVal: 0, valAxisMaxVal: 1400, catAxisLabelFontSize: 13, objectName: "s9_graf",
	}));
	fx("s9_cap", "fade", "apos", { duracao: 0.3 });
	fx("s9_graf", "wipe", "apos", { grafico: "categoria", etapas: est.length, direcao: "esquerda", duracao: 0.25 });
	legendaChips(s, [
		{ cor: HEX.accent1, texto: "Socaria", larg: 0.75 },
		{ cor: HEX.accent2, texto: "Plantio", larg: 0.7 },
		{ cor: EXTRA.cinzaClaro, texto: "Sem grupo na planilha", larg: 1.95 },
	], 0.6, 6.2, "s9_leg");
	fx("s9_leg", "fade", "com", { duracao: 0.4 });
	txt(s, "Área equivalente (estoque ÷ dose implícita)", { x: 8.5, y: 1.62, w: 4.23, h: 0.32, fontSize: 16, bold: true, color: C.text2, objectName: "s9_eq_t" });
	const gr = [
		{ g: "Socaria", rot: "Socaria", ic: I.cana, t: `5 fórmulas · ${fmt(dA.estoque_por_grupo_t.Socaria)} t` },
		{ g: "Plantio", rot: "Plantio", ic: I.cenaClara, t: `4 fórmulas · ${fmt(dA.estoque_por_grupo_t.Plantio)} t` },
		{ g: "Não classificado", rot: "Sem grupo", ic: I.lacuna, t: `Oxyfertil · ${fmt(dA.estoque_por_grupo_t["Não classificado"])} t` },
	];
	const mEq = ["s9_eq_t"];
	gr.forEach((g, i) => {
		const y = 2.08 + i * 0.78;
		s.addImage({ data: g.ic, x: 8.5, y: y + 0.06, w: 0.48, h: 0.48, objectName: `s9_eq${i}_ic` });
		txt(s, [{ text: fmt(dA.area_eq_por_grupo_ha[g.g]), options: { fontSize: 24, bold: true, color: C.text2 } }, { text: " ha", options: { fontSize: 14, bold: true, color: C.text2 } }],
			{ x: 9.12, y, w: 1.75, h: 0.6, valign: "middle", objectName: `s9_eq${i}_v` });
		txt(s, [{ text: g.rot, options: { bold: true, breakLine: true } }, { text: g.t }], { x: 10.85, y, w: 1.88, h: 0.6, fontSize: 12, valign: "middle", objectName: `s9_eq${i}_t` });
		mEq.push(`s9_eq${i}_ic`, `s9_eq${i}_v`, `s9_eq${i}_t`);
	});
	grp("s9_eq", mEq);
	fx("s9_eq", "fade", "apos", { duracao: 0.5 });
	const eq46 = A.estoque.find((e) => e.produto === "46-00-00");
	caixa(s, { x: 8.5, y: 4.85, w: 4.23, h: 1.5, fill: { color: C.accent3, transparency: 82 }, objectName: "s9_at_bg" });
	s.addImage({ data: I.alerta, x: 8.68, y: 5.0, w: 0.4, h: 0.4, objectName: "s9_at_ic" });
	txt(s, [
		{ text: "Áreas não se somam: ", options: { bold: true } },
		{ text: `os ${fmt(eq46.area_eq_ha)} ha do 46-00-00 só valem onde há vinhaça localizada.`, options: { breakLine: true } },
		{ text: "Plantio: ", options: { bold: true } },
		{ text: `${fmt(D.plantio_projetado_ha.fundacao + D.plantio_projetado_ha.renovacao)} ha projetados × ${fmt(dA.area_eq_por_grupo_ha.Plantio)} ha cobertos; a necessidade do plantio não está calculada.` },
	], { x: 9.2, y: 4.98, w: 3.4, h: 1.3, fontSize: 12, paraSpaceAfter: 6, objectName: "s9_at_tx" });
	grp("s9_at", ["s9_at_bg", "s9_at_ic", "s9_at_tx"]);
	fx("s9_at", "fade", "clique", { duracao: 0.5 });
	fonte(s, `${ARQ} — ${ABA_A}, quadro "Estoque de fertilizantes 16/10/2026" (I47:K61) e nota A42 (plantio projetado).`);
	s.addNotes(
		`Estoque total de ${fmt(dA.estoque_total_t, 2)} t em 10 fórmulas. Socaria (16-00-20, 46-00-00, 21-00-00, 25-00-00, 13-00-34): ${fmt(dA.estoque_por_grupo_t.Socaria, 2)} t; plantio (06-26-25, 06-30-25, 06-27-24, 05-25-25): ${fmt(dA.estoque_por_grupo_t.Plantio, 2)} t; Oxyfertil ${fmt(dA.estoque_por_grupo_t["Não classificado"], 2)} t fica fora dos subtotais da planilha. ` +
		"A área equivalente é o estoque dividido pela dose implícita da própria planilha (0,5 t/ha no 16-00-20, 0,2 no 46-00-00, 0,4 no 21-00-00 e 25-00-00, 0,6 nas fórmulas de plantio). " +
		"Essas áreas não podem ser somadas como se os produtos fossem equivalentes: o 46-00-00 depende da vinhaça para o potássio. O plantio projetado (1.400 ha de fundação + 1.000 ha de renovação) não tem necessidade calculada na planilha."
	);

	// ================================================================== 10. BALANÇO
	s = novoSlide("CONTEUDO", "Abastecimento");
	titulos(s, "PLANEJAMENTO DE COMPRAS E ESTOQUE", `Estoque e compras cobrem ${pct(dA.cobertura_16_disponivel)} do 16-00-20 necessário`);
	txt(s, "Da necessidade bruta à compra líquida · 16-00-20 · cenário A (t)", { x: 0.6, y: 1.62, w: 7.3, h: 0.32, fontSize: 14, bold: true, color: C.text2, objectName: "s10_cap" });
	const cats10 = ["Necessidade", "Estoque em 16/10", "Comprado a receber", "A comprar"];
	s.addChart(pres.charts.BAR, [
		{ name: "apoio da cascata (oculto)", labels: cats10, values: [0, b16A.necessidade_t - b16A.estoque_t, A.compra_16_t, 0] },
		{ name: "16-00-20 (t)", labels: cats10, values: [b16A.necessidade_t, b16A.estoque_t, b16A.a_receber_t, A.compra_16_t] },
	], graf({
		x: 0.6, y: 1.98, w: 7.3, h: 4.55, barDir: "col", barGrouping: "stacked", barGapWidthPct: 45,
		chartColors: ["FFFFFF", HEX.accent1], valAxisHidden: true, dataLabelPosition: "ctr", valAxisMinVal: 0, valAxisMaxVal: 5800,
		dataLabelFormatCode: "#,##0.0", dataLabelColor: "FFFFFF", dataLabelFontSize: 14, catAxisLabelFontSize: 13,
		objectName: "s10_graf",
	}));
	ajusteGrafico("s10_graf", {
		ocultarSerie: 1,
		corPontos: { serie: 2, cores: [HEX.dk2, HEX.accent2, EXTRA.canaClara, HEX.accent4] },
		corRotulos: { serie: 2, cores: ["FFFFFF", HEX.dk2, HEX.dk2, "FFFFFF"] },
	});
	fx("s10_cap", "fade", "apos", { duracao: 0.3 });
	fx("s10_graf", "wipe", "apos", { grafico: "categoria", etapas: 4, direcao: "baixo", duracao: 0.5 });
	// cartão que viaja (Morph) para o slide de impacto financeiro
	const cartaoCompra = (sl, x, y, w, h, tam) => {
		caixa(sl, { x, y, w, h, fill: { color: C.accent4, transparency: 88 }, objectName: "cmp_bg" });
		txt(sl, [{ text: fmt(A.compra_16_t), options: { fontSize: tam, bold: true, color: C.accent4 } }, { text: " t", options: { fontSize: tam * 0.55, bold: true, color: C.accent4 } }],
			{ x: x + 0.22, y: y + 0.12, w: w - 0.4, h: h * 0.45, valign: "middle", objectName: "cmp_v" });
		txt(sl, "de 16-00-20 a comprar no cenário A", { x: x + 0.22, y: y + h * 0.52, w: w - 0.4, h: 0.3, fontSize: 13, bold: true, objectName: "cmp_t" });
		txt(sl, `${fmt(B.compra_16_t)} t no cenário B`, { x: x + 0.22, y: y + h * 0.52 + 0.3, w: w - 0.4, h: 0.28, fontSize: 12, color: C.accent6, objectName: "cmp_s" });
		grp("!!compra16", ["cmp_bg", "cmp_v", "cmp_t", "cmp_s"]);
	};
	cartaoCompra(s, 8.3, 1.62, 4.43, 1.55, 36);
	fx("!!compra16", "fade", "apos", { duracao: 0.5 });
	caixa(s, { x: 8.3, y: 3.37, w: 4.43, h: 2.1, objectName: "s10_46_bg" });
	s.addImage({ data: I.gota, x: 8.5, y: 3.52, w: 0.45, h: 0.45, objectName: "s10_46_ic" });
	txt(s, "46-00-00 · sobra no cenário A", { x: 9.08, y: 3.52, w: 3.5, h: 0.45, fontSize: 16, bold: true, color: C.text2, valign: "middle", objectName: "s10_46_t" });
	const linhas46 = [["Necessidade", fmt(b46A.necessidade_t, 1) + " t"], ["Estoque", fmt(b46A.estoque_t, 1) + " t"], ["Comprado a receber", fmt(b46A.a_receber_t, 1) + " t"]];
	txt(s, linhas46.map((l, i) => ({ text: l[0], options: { breakLine: i < linhas46.length - 1 } })), { x: 8.5, y: 4.07, w: 2.3, h: 0.85, fontSize: 13, objectName: "s10_46_l" });
	txt(s, linhas46.map((l, i) => ({ text: l[1], options: { breakLine: i < linhas46.length - 1 } })), { x: 10.8, y: 4.07, w: 1.75, h: 0.85, fontSize: 13, align: "right", objectName: "s10_46_r" });
	txt(s, [{ text: "Saldo  ", options: { fontSize: 14, bold: true, color: C.text1 } }, { text: "+" + fmt(b46A.saldo_t, 1) + " t", options: { fontSize: 20, bold: true, color: C.accent5 } }],
		{ x: 8.5, y: 4.95, w: 4.05, h: 0.42, align: "right", valign: "middle", objectName: "s10_46_s" });
	grp("s10_46", ["s10_46_bg", "s10_46_ic", "s10_46_t", "s10_46_l", "s10_46_r", "s10_46_s"]);
	fx("s10_46", "fade", "apos", { duracao: 0.5 });
	txt(s, `Compra líquida = necessidade − estoque − comprado a receber. No cenário B: ${fmt(b16B.necessidade_t, 1)} t de necessidade e ${fmt(B.compra_16_t, 1)} t a comprar; o 46-00-00 fica sem uso previsto (${fmt(b46B.disponivel_t, 1)} t).`,
		{ x: 8.3, y: 5.62, w: 4.43, h: 0.92, fontSize: 12, color: C.accent6, objectName: "s10_obs" });
	fx("s10_obs", "fade", "apos", { duracao: 0.4 });
	fonte(s, `${ARQ} — ${ABA_A} e ${ABA_B}, quadro 4 (Resumo): B50:F51 e B56 ("Necessidade de aquisição").`);
	s.addNotes(
		`Necessidade líquida de compra, sem confundir com a bruta: necessidade de 16-00-20 ${fmt(b16A.necessidade_t, 1)} t, menos estoque ${fmt(b16A.estoque_t)} t, menos ${fmt(b16A.a_receber_t)} t já compradas a receber = ${fmt(A.compra_16_t, 1)} t a comprar (cenário A). ` +
		`Cobertura atual: ${pct(dA.cobertura_16_estoque, 1)} só com estoque; ${pct(dA.cobertura_16_disponivel, 1)} com o que está a receber. ` +
		`No 46-00-00 o quadro se inverte: ${fmt(b46A.disponivel_t, 2)} t disponíveis para ${fmt(b46A.necessidade_t, 1)} t de necessidade — sobra de ${fmt(b46A.saldo_t, 1)} t no A e ${fmt(b46B.saldo_t, 1)} t no B. ` +
		"Fórmulas sem necessidade calculada (21-00-00, 13-00-34, 25-00-00, Oxyfertil) não entram no balanço."
	);

	// ================================================================== 11. IMPACTO FINANCEIRO
	pres.addSection({ title: "Finanças e cenários" });
	s = novoSlide("CONTEUDO", "Finanças e cenários", "morph");
	const sp = dA.sens_preco_por_100;
	titulos(s, "IMPACTO FINANCEIRO", `Cada R$ 100/t no 16-00-20 vale R$ ${fmt(sp["compra_16_R$"] / 1000)} mil`);
	cartaoCompra(s, 0.6, 1.62, 3.55, 1.55, 32);
	txt(s, "×", { x: 4.15, y: 1.62, w: 0.55, h: 1.55, fontSize: 36, bold: true, color: C.accent6, align: "center", valign: "middle", objectName: "s11_x" });
	caixa(s, { x: 4.7, y: 1.62, w: 3.55, h: 1.55, fill: { color: "FFFFFF" }, line: { color: C.accent6, width: 1.25, dashType: "dash" }, objectName: "s11_preco_bg" });
	txt(s, "Preço (R$/t)", { x: 4.92, y: 1.76, w: 3.2, h: 0.55, fontSize: 24, bold: true, color: C.accent6, valign: "middle", objectName: "s11_preco_v" });
	txt(s, "a informar — a planilha não traz preços", { x: 4.92, y: 2.4, w: 3.2, h: 0.6, fontSize: 13, objectName: "s11_preco_t" });
	txt(s, "=", { x: 8.25, y: 1.62, w: 0.55, h: 1.55, fontSize: 36, bold: true, color: C.accent6, align: "center", valign: "middle", objectName: "s11_ig" });
	caixa(s, { x: 8.8, y: 1.62, w: 3.93, h: 1.55, fill: { color: C.text2 }, objectName: "s11_orc_bg" });
	txt(s, "Orçamento da compra", { x: 9.02, y: 1.76, w: 3.6, h: 0.55, fontSize: 22, bold: true, color: C.background1, valign: "middle", objectName: "s11_orc_v" });
	txt(s, "calculável assim que o preço for informado", { x: 9.02, y: 2.4, w: 3.6, h: 0.6, fontSize: 13, color: "D7E5D2", objectName: "s11_orc_t" });
	grp("s11_formula", ["s11_x", "s11_preco_bg", "s11_preco_v", "s11_preco_t", "s11_ig", "s11_orc_bg", "s11_orc_v", "s11_orc_t"]);
	fx("s11_formula", "fade", "apos", { duracao: 0.5, atraso: 0.2 });
	txt(s, "Sensibilidade a cada R$ 100/t de preço", { x: 0.6, y: 3.47, w: 8, h: 0.35, fontSize: 16, bold: true, color: C.text2, objectName: "s11_sens_t" });
	const sens = [
		{ v: `R$ ${fmt(sp["compra_16_R$"] / 1000)} mil`, t: `na compra de 16-00-20 do cenário A · R$ ${fmt(dB.sens_preco_por_100["compra_16_R$"] / 1000)} mil no B` },
		{ v: `R$ ${fmt(sp["custo_ha_16_R$"])}/ha`, t: "no custo do 16-00-20 por hectare, a 500 kg/ha" },
		{ v: `R$ ${fmt(A.doses_kg_ha["46-00-00"] / 1000 * 100)}/ha`, t: "no custo do 46-00-00 por hectare, a 200 kg/ha" },
		{ v: `R$ ${fmt(sp["excedente_46_R$"] / 1000)} mil`, t: `no valor do excedente de 46-00-00 (A) · R$ ${fmt(dB.sens_preco_por_100["excedente_46_R$"] / 1000)} mil no B` },
	];
	const mSens = ["s11_sens_t"];
	sens.forEach((c, i) => {
		const x = 0.6 + i * 3.08, nm = `s11_s${i}`;
		caixa(s, { x, y: 3.9, w: 2.9, h: 1.42, objectName: `${nm}_bg` });
		txt(s, c.v, { x: x + 0.2, y: 4.0, w: 2.55, h: 0.5, fontSize: 24, bold: true, color: C.text2, valign: "middle", objectName: `${nm}_v` });
		txt(s, c.t, { x: x + 0.2, y: 4.55, w: 2.55, h: 0.7, fontSize: 12, objectName: `${nm}_t` });
		mSens.push(`${nm}_bg`, `${nm}_v`, `${nm}_t`);
	});
	grp("s11_sens", mSens);
	fx("s11_sens", "fade", "apos", { duracao: 0.5 });
	txt(s, "Para fechar o orçamento faltam", { x: 0.6, y: 5.5, w: 4, h: 0.32, fontSize: 14, bold: true, color: C.text2, objectName: "s11_falta_t" });
	const faltas = ["Preço por fórmula e condição (CIF/FOB, prazo)", "Preço das 3.180 t já compradas", "Frete e armazenagem", "Custo de aplicação por hectare"];
	const mFalta = ["s11_falta_t"];
	faltas.forEach((f, i) => {
		const x = 0.6 + i * 3.08, nm = `s11_f${i}`;
		caixa(s, { x, y: 5.88, w: 2.9, h: 0.62, fill: { color: C.accent3, transparency: 80 }, rectRadius: 0.06, objectName: `${nm}_bg` });
		s.addImage({ data: I.ampulheta, x: x + 0.12, y: 6.0, w: 0.38, h: 0.38, objectName: `${nm}_ic` });
		txt(s, f, { x: x + 0.6, y: 5.88, w: 2.2, h: 0.62, fontSize: 12, valign: "middle", objectName: `${nm}_t` });
		mFalta.push(`${nm}_bg`, `${nm}_ic`, `${nm}_t`);
	});
	grp("s11_falta", mFalta);
	fx("s11_falta", "fade", "apos", { duracao: 0.5 });
	fonte(s, `${ARQ} — volumes de B56 e F50 (${ABA_A} e ${ABA_B}); doses J38:J39. Nenhum preço foi assumido.`);
	s.addNotes(
		"A planilha não traz preços, então o orçamento não foi calculado e nenhum preço foi suposto. O slide mostra a estrutura da conta e a sensibilidade: " +
		`cada R$ 100/t no preço do 16-00-20 muda a compra em R$ ${fmt(sp["compra_16_R$"])} no cenário A (${fmt(A.compra_16_t, 1)} t) e em R$ ${fmt(dB.sens_preco_por_100["compra_16_R$"])} no B. ` +
		"Por hectare, R$ 100/t equivalem a R$ 50/ha no 16-00-20 (500 kg/ha) e R$ 20/ha no 46-00-00 (200 kg/ha). " +
		"Assim que os preços forem informados, a planilha auxiliar (aba Sensibilidade) calcula o orçamento. O cartão de 2.260 t vem do slide anterior com transição Morph."
	);

	// ================================================================== 12. CENÁRIOS
	s = novoSlide("CONTEUDO", "Finanças e cenários");
	const porDia16 = A.ritmo_localizada_ha_dia * A.doses_kg_ha["16-00-20"] / 1000;
	const porDia46 = A.ritmo_localizada_ha_dia * A.doses_kg_ha["46-00-00"] / 1000;
	titulos(s, "CENÁRIOS DE PLANEJAMENTO", `Cada dia de vinhaça com K poupa ${fmt(porDia16)} t de 16-00-20`);
	const dif = (a, b, d = 1) => (b - a > 0 ? "+" : b - a < 0 ? "−" : "") + fmt(Math.abs(b - a), d);
	const lin = [
		["Dias de vinhaça com K adequado (C13)", fmt(A.dias_k_vinhaca), fmt(B.dias_k_vinhaca), dif(A.dias_k_vinhaca, B.dias_k_vinhaca, 0)],
		["Área de localizada estimada, ha (C14)", fmt(A.area_localizada_estimada_ha, 1), fmt(B.area_localizada_estimada_ha, 1), dif(A.area_localizada_estimada_ha, B.area_localizada_estimada_ha)],
		["Localizada CRV a adubar, ha (H38)", fmt(A.crv.a_tratar_localizada_ha, 1), fmt(B.crv.a_tratar_localizada_ha, 1), dif(A.crv.a_tratar_localizada_ha, B.crv.a_tratar_localizada_ha)],
		["Mineral CRV a adubar, ha (H39)", fmt(A.crv.a_tratar_mineral_ha, 1), fmt(B.crv.a_tratar_mineral_ha, 1), dif(A.crv.a_tratar_mineral_ha, B.crv.a_tratar_mineral_ha)],
		["Necessidade de 16-00-20, t", fmt(b16A.necessidade_t, 1), fmt(b16B.necessidade_t, 1), dif(b16A.necessidade_t, b16B.necessidade_t)],
		["Necessidade de 46-00-00, t", fmt(b46A.necessidade_t, 1), fmt(b46B.necessidade_t, 1), dif(b46A.necessidade_t, b46B.necessidade_t)],
		["Compra de 16-00-20, t", fmt(A.compra_16_t, 1), fmt(B.compra_16_t, 1), dif(A.compra_16_t, B.compra_16_t)],
		["Excedente de 46-00-00, t", fmt(b46A.saldo_t, 1), fmt(b46B.saldo_t, 1), dif(b46A.saldo_t, b46B.saldo_t)],
	];
	const destaque = { bold: true, fill: { color: "F6E3DF" } };
	s.addTable([
		[th("Premissa / resultado"), th("A · 1.000 ha", "right"), th("B · 450 ha", "right"), th("B − A", "right")],
		...lin.map((l, i) => l.map((c, j) => td(c, j === 0 ? "left" : "right", i === 6 ? destaque : {}))),
	], {
		x: 0.6, y: 1.62, w: 6.75, colW: [3.45, 1.15, 1.15, 1.0], rowH: 0.4, fontSize: 13, color: HEX.dk1, valign: "middle",
		border: { type: "solid", pt: 0.75, color: EXTRA.linhaTabela }, margin: [0, 0.08, 0, 0.08], objectName: "s12_tab",
	});
	fx("s12_tab", "fade", "apos", { duracao: 0.5 });
	txt(s, "Compra de 16-00-20 conforme os dias de vinhaça com K (t)", { x: 7.65, y: 1.62, w: 5.08, h: 0.32, fontSize: 14, bold: true, color: C.text2, objectName: "s12_cap" });
	const sensD = D.sensibilidade_dias;
	s.addChart(pres.charts.BAR, [{ name: "Compra de 16-00-20 (t)", labels: sensD.map((x) => x.dias === 20 ? "20 dias (A)" : x.dias === 0 ? "0 dias (B)" : `${x.dias} dias`), values: sensD.map((x) => x.compra_16_t) }], graf({
		x: 7.65, y: 1.98, w: 5.08, h: 3.25, barDir: "col", barGapWidthPct: 45,
		chartColors: sensD.map((x) => x.dias === 20 ? HEX.accent1 : x.dias === 0 ? HEX.accent4 : EXTRA.cinzaClaro),
		valAxisHidden: true, valAxisMinVal: 0, valAxisMaxVal: 3000, dataLabelPosition: "outEnd", dataLabelFormatCode: "#,##0", catAxisLabelFontSize: 12,
		objectName: "s12_graf",
	}));
	fx("s12_cap", "fade", "apos", { duracao: 0.3 });
	fx("s12_graf", "wipe", "apos", { grafico: "categoria", etapas: sensD.length, direcao: "baixo", duracao: 0.4 });
	txt(s, `Até ~${fmt(A.cargill.a_tratar_ha / A.ritmo_localizada_ha_dia)} dias, a localizada atende primeiro os ${fmt(A.cargill.a_tratar_ha, 2)} ha da Cargill; depois, cada dia (${fmt(A.ritmo_localizada_ha_dia, 2)} ha) troca ${fmt(porDia16, 1)} t de 16-00-20 por ${fmt(porDia46, 1)} t de 46-00-00. Acima de 20 dias, só com vinhaça de K adequado e até 30/11 (45 dias).`,
		{ x: 7.65, y: 5.3, w: 5.08, h: 1.25, fontSize: 12, color: C.accent6, objectName: "s12_obs" });
	fx("s12_obs", "fade", "com", { duracao: 0.4 });
	caixa(s, { x: 0.6, y: 5.42, w: 6.75, h: 1.13, fill: { color: C.accent3, transparency: 82 }, objectName: "s12_alt_bg" });
	s.addImage({ data: I.alerta, x: 0.78, y: 5.57, w: 0.4, h: 0.4, objectName: "s12_alt_ic" });
	txt(s, [
		{ text: "Alternativa a avaliar com a agronomia: ", options: { bold: true } },
		{ text: `o excedente de 46-00-00 (A) contém ${fmt(dA.excedente_46_N_t)} t de N, acima das ${fmt(dA.deficit_nutrientes_t.N)} t de N do déficit de 16-00-20; faltaria uma fonte de K para ${fmt(dA.deficit_nutrientes_t.K2O)} t de K₂O.` },
	], { x: 1.32, y: 5.52, w: 5.9, h: 0.98, fontSize: 12, valign: "middle", objectName: "s12_alt_tx" });
	grp("s12_alt", ["s12_alt_bg", "s12_alt_ic", "s12_alt_tx"]);
	fx("s12_alt", "fade", "clique", { duracao: 0.5 });
	fonte(s, `${ARQ} — cenários = abas ${ABA_A} (A) e ${ABA_B} (B). Barras de 10, 30 e 45 dias: mesma lógica da aba A, só variando C13 (sensibilidade, não cenário oficial).`);
	s.addNotes(
		"Os dois cenários oficiais são as duas abas da planilha. A única premissa que muda é o número de dias em que a vinhaça localizada atende o nível de potássio (C13): 20 dias no A, zero no B. " +
		`A diferença final é de ${fmt(B.compra_16_t - A.compra_16_t, 1)} t de 16-00-20 a mais no B. ` +
		`As barras de 10, 30 e 45 dias aplicam a mesma lógica da aba A só mudando C13; servem para mostrar o valor de cada dia de vinhaça com K (${fmt(porDia16, 1)} t de 16-00-20 por dia) e não são previsões da usina. ` +
		"O cenário de preço fica para quando houver cotação (estrutura pronta na planilha auxiliar). " +
		`A alternativa técnica (clique) é hipótese a validar: o N contido no excedente de 46-00-00 cobre o N do déficit, mas o K precisaria de outra fonte, a ser definida e cotada.`
	);

	// ================================================================== 13. RISCOS
	pres.addSection({ title: "Riscos e ação" });
	s = novoSlide("CONTEUDO", "Riscos e ação");
	titulos(s, "RISCOS E PONTOS DE ATENÇÃO", "Ruptura de 16-00-20 é o principal risco da safra");
	const ql = A.quadro_localizada;
	const riscos = [
		{ ic: I.carrinho, sev: "ALTO", t: "Déficit de 16-00-20", v: `${fmt(A.compra_16_t)}–${fmt(B.compra_16_t)} t`, d: `sem cobertura: ${pct(1 - dA.cobertura_16_disponivel)} a ${pct(1 - dB.cobertura_16_disponivel)} da necessidade do produto` },
		{ ic: I.caminhaoVerm, sev: "ALTO", t: "Entregas sem data", v: `${fmt(dA.a_receber_total_t)} t`, d: `compradas sem cronograma; sem elas, o estoque cobre ~${fmt(dA.curva_16.dias_cobertura_estoque, 1)} dias (premissa linear)` },
		{ ic: I.lacuna, sev: "ALTO", t: "Lacunas de informação", v: "4 dados", d: "faltam preços, prazos de fornecimento, demanda mensal e necessidade do plantio (2.400 ha)" },
		{ ic: I.armazem, sev: "MÉDIO", t: "Excedente de 46-00-00", v: `${fmt(b46A.saldo_t)}–${fmt(b46B.saldo_t)} t`, d: "acima do uso previsto: capital parado e espaço de armazém" },
		{ ic: I.gota, sev: "MÉDIO", t: "Dependência da vinhaça com K", v: `+${fmt(porDia16, 1)} t/dia`, d: `de 16-00-20 a cada dia a menos de localizada; até ${fmt(B.compra_16_t - A.compra_16_t)} t entre A e B` },
		{ ic: I.base, sev: "MÉDIO", t: "Base com inconsistências", v: `${fmt(ql.total_ha - ql.partes_ha.reduce((a, b) => a + b, 0))} ha`, d: "de diferença no quadro de vinhaça localizada (B26 × B27:B29)" },
	];
	riscos.forEach((r, i) => {
		const col = i % 3, row = Math.floor(i / 3);
		const x = 0.6 + col * 4.12, y = 1.62 + row * 2.47, nm = `s13_r${i}`;
		caixa(s, { x, y, w: 3.89, h: 2.3, objectName: `${nm}_bg` });
		s.addImage({ data: r.ic, x: x + 0.2, y: y + 0.2, w: 0.5, h: 0.5, objectName: `${nm}_ic` });
		const alto = r.sev === "ALTO";
		caixa(s, { x: x + 2.84, y: y + 0.27, w: 0.85, h: 0.32, rectRadius: 0.16, fill: { color: alto ? C.accent4 : C.accent3 }, objectName: `${nm}_sv` });
		txt(s, r.sev, { x: x + 2.84, y: y + 0.27, w: 0.85, h: 0.32, fontSize: 11, bold: true, color: alto ? C.background1 : C.text1, align: "center", valign: "middle", objectName: `${nm}_svt` });
		txt(s, r.t, { x: x + 0.2, y: y + 0.82, w: 3.5, h: 0.32, fontSize: 15, bold: true, color: C.text2, objectName: `${nm}_t` });
		txt(s, r.v, { x: x + 0.2, y: y + 1.14, w: 3.5, h: 0.45, fontSize: 22, bold: true, color: alto ? C.accent4 : C.text1, valign: "middle", objectName: `${nm}_v` });
		txt(s, r.d, { x: x + 0.2, y: y + 1.62, w: 3.5, h: 0.6, fontSize: 12, objectName: `${nm}_d` });
		grp(nm, [`${nm}_bg`, `${nm}_ic`, `${nm}_sv`, `${nm}_svt`, `${nm}_t`, `${nm}_v`, `${nm}_d`]);
		fx(nm, "fade", "apos", { duracao: 0.35 });
	});
	fonte(s, "classificação por impacto — ALTO: mais de 2.000 t ou impede a decisão; MÉDIO: demais. A planilha não permite estimar probabilidade, por isso não há matriz.");
	s.addNotes(
		"Riscos sustentados pelos dados da planilha, ordenados por impacto em volume. " +
		`1) Déficit de 16-00-20 de ${fmt(A.compra_16_t, 1)} t (A) a ${fmt(B.compra_16_t, 1)} t (B). ` +
		`2) As ${fmt(dA.a_receber_total_t)} t compradas a receber (2.030 t de 16-00-20 e 1.150 t de 46-00-00) não têm data de entrega na planilha. ` +
		"3) Faltam dados para orçamento e logística. 4) O 46-00-00 sobra. 5) Cada dia a menos de vinhaça com potássio aumenta a compra de 16-00-20. " +
		`6) O quadro de vinhaça localizada (linhas 26 a 29) não fecha: total ${fmt(ql.total_ha, 2)} ha contra ${fmt(ql.partes_ha.reduce((a, b) => a + b, 0), 2)} ha nas linhas. ` +
		"Não foi feita matriz de probabilidade × impacto porque a planilha não tem base para probabilidade."
	);

	// ================================================================== 14. PLANO DE AÇÃO
	s = novoSlide("CONTEUDO", "Riscos e ação");
	titulos(s, "PLANO DE AÇÃO", "Seis frentes para fechar o abastecimento da safra");
	txt(s, "Frente", { x: 1.3, y: 1.6, w: 3, h: 0.3, fontSize: 12, bold: true, color: C.accent6, objectName: "s14_h1" });
	txt(s, "O que fazer", { x: 4.65, y: 1.6, w: 5.4, h: 0.3, fontSize: 12, bold: true, color: C.accent6, objectName: "s14_h2" });
	txt(s, "Área sugerida  ·  prazo", { x: 10.05, y: 1.6, w: 2.68, h: 0.3, fontSize: 12, bold: true, color: C.accent6, objectName: "s14_h3" });
	grp("s14_cab", ["s14_h1", "s14_h2", "s14_h3"]);
	fx("s14_cab", "fade", "apos", { duracao: 0.3 });
	const acoes = [
		{ ic: I.lupa, t: "Validar a demanda agrícola", d: `Confirmar ${fmt(A.total_a_tratar_ha)} ha a adubar, ${fmt(bisada)} ha de cana bisada e os dias de vinhaça com K; corrigir o quadro de localizada.`, a: "Planejamento Agrícola" },
		{ ic: I.camadas, t: "Consolidar os volumes", d: `Escolher o cenário (A ou B) e fixar o 16-00-20 a comprar; decidir se o plantio de 2.400 ha entra no balanço.`, a: "Planejamento Agrícola" },
		{ ic: I.contrato, t: "Rever estoque e contratos", d: `Conferir ${fmt(b16A.estoque_t)} t de 16-00-20 e ${fmt(b46A.estoque_t)} t de 46-00-00; obter datas das ${fmt(dA.a_receber_total_t)} t a receber.`, a: "Suprimentos e Almoxarifado" },
		{ ic: I.carrinho, t: "Planejar as compras", d: `Cotar ${fmt(A.compra_16_t)}–${fmt(B.compra_16_t)} t de 16-00-20; negociar o 46-00-00 a receber; avaliar N do excedente + fonte de K.`, a: "Suprimentos e Agronomia" },
		{ ic: I.caminhao, t: "Programar entrega e armazenagem", d: `Escalonar a chegada antes de ~${dataBR(cv.data_esgota_disponivel)} (premissa linear) e checar a capacidade de armazém.`, a: "Logística e Suprimentos" },
		{ ic: I.grafico, t: "Acompanhar a execução", d: "Painel semanal: comprado, entregue, aplicado, saldo a comprar e custo previsto × realizado.", a: "Planejamento Agrícola" },
	];
	acoes.forEach((a, i) => {
		const y = 1.98 + i * 0.76, nm = `s14_a${i}`;
		caixa(s, { x: 0.6, y, w: 12.13, h: 0.68, fill: i % 2 === 0 ? { color: C.background2 } : { color: "FFFFFF" }, line: { color: EXTRA.grade, width: 0.75 }, rectRadius: 0.06, objectName: `${nm}_bg` });
		s.addImage({ data: a.ic, x: 0.75, y: y + 0.12, w: 0.44, h: 0.44, objectName: `${nm}_ic` });
		txt(s, `${i + 1}. ${a.t}`, { x: 1.3, y, w: 3.25, h: 0.68, fontSize: 14, bold: true, color: C.text2, valign: "middle", objectName: `${nm}_t` });
		txt(s, a.d, { x: 4.65, y, w: 5.25, h: 0.68, fontSize: 12, valign: "middle", objectName: `${nm}_d` });
		txt(s, [{ text: a.a, options: { bold: true, breakLine: true } }, { text: "prazo a definir", options: { color: C.accent6 } }], { x: 10.05, y, w: 2.6, h: 0.68, fontSize: 12, valign: "middle", objectName: `${nm}_a` });
		grp(nm, [`${nm}_bg`, `${nm}_ic`, `${nm}_t`, `${nm}_d`, `${nm}_a`]);
		fx(nm, "wipe", "apos", { direcao: "esquerda", duracao: 0.4 });
	});
	fonte(s, "áreas sugeridas pela natureza de cada frente; responsáveis nominais e datas não constam da planilha e ficam a definir pela diretoria.");
	s.addNotes(
		"Seis frentes, na ordem em que dependem umas das outras: validar a demanda, consolidar volumes, rever estoque e contratos, comprar, programar entrega e acompanhar. " +
		"As áreas são sugestões pela natureza da tarefa; nomes e prazos não estão na planilha e devem ser definidos pela diretoria. " +
		`A referência de ~${dataBR(cv.data_esgota_disponivel)} para a chegada do produto vem da premissa linear do slide de curva de abastecimento.`
	);

	// ================================================================== 15. PAINEL
	s = novoSlide("CONTEUDO", "Riscos e ação");
	const paineis = [
		{ r: "Necessidade total prevista", v: `${fmt(dA.necessidade_total_t)} t`, s: `cenário A · ${fmt(dB.necessidade_total_t)} t no B`, ok: true },
		{ r: "Volume comprado", v: `${fmt(dA.a_receber_total_t)} t`, s: `a receber (2.030 t de 16-00-20 + 1.150 t de 46-00-00)`, ok: true },
		{ r: "Volume entregue", v: "—", s: "depende do registro de recebimentos", ok: false },
		{ r: "Área já adubada", v: `${fmt(realizadoTotal)} ha`, s: `${pct(dA.area.pct_realizado_total)} da área de tratos; t aplicadas só na localizada`, ok: true },
		{ r: "Saldo a comprar", v: `${fmt(A.compra_16_t)} t`, s: `16-00-20, cenário A · ${fmt(B.compra_16_t)} t no B`, ok: true },
		{ r: "Aderência ao cronograma", v: "—", s: "depende do cronograma de entrega e de aplicação", ok: false },
		{ r: "Custo previsto × realizado", v: "—", s: "depende de preços e notas fiscais", ok: false },
		{ r: "Estoque projetado no fim da safra", v: `+${fmt(b46A.saldo_t)} t`, s: "de 46-00-00 (A); 16-00-20 zera após a compra", ok: true },
	];
	const nOk = paineis.filter((p) => p.ok).length;
	titulos(s, "PAINEL DE CONTROLE", `${nOk} de ${paineis.length} indicadores já podem ser medidos`);
	paineis.forEach((p, i) => {
		const col = i % 4, row = Math.floor(i / 4);
		const x = 0.6 + col * 3.08, y = 1.62 + row * 2.47, nm = `s15_k${i}`;
		const fundo = { x, y, w: 2.9, h: 2.3, fill: { color: p.ok ? C.background2 : "FFFFFF" }, objectName: `${nm}_bg` };
		if (!p.ok) fundo.line = { color: C.accent3, width: 1.25, dashType: "dash" };
		caixa(s, fundo);
		txt(s, p.r, { x: x + 0.2, y: y + 0.2, w: 2.5, h: 0.55, fontSize: 13, bold: true, color: C.accent6, objectName: `${nm}_r` });
		txt(s, p.v, { x: x + 0.2, y: y + 0.78, w: 2.5, h: 0.55, fontSize: 28, bold: true, color: p.ok ? C.text2 : C.accent6, valign: "middle", objectName: `${nm}_v` });
		txt(s, p.s, { x: x + 0.2, y: y + 1.35, w: 2.5, h: 0.5, fontSize: 11, objectName: `${nm}_s` });
		s.addImage({ data: p.ok ? I.check : I.ampulheta, x: x + 0.2, y: y + 1.88, w: 0.26, h: 0.26, objectName: `${nm}_ic` });
		txt(s, p.ok ? "mensurável hoje" : "depende de dado", { x: x + 0.54, y: y + 1.86, w: 2.1, h: 0.3, fontSize: 11, bold: true, color: p.ok ? C.accent1 : C.text1, valign: "middle", objectName: `${nm}_st` });
		grp(nm, [`${nm}_bg`, `${nm}_r`, `${nm}_v`, `${nm}_s`, `${nm}_ic`, `${nm}_st`]);
		fx(nm, "fade", "apos", { duracao: 0.3 });
	});
	fonte(s, `${ARQ} — ${ABA_A} e ${ABA_B}. Indicadores sem dado ficam com "—" até que a informação seja incorporada à planilha.`);
	s.addNotes(
		`Painel de acompanhamento proposto. ${nOk} indicadores já são calculáveis com a planilha; os outros 3 (volume entregue, aderência ao cronograma, custo previsto × realizado) dependem de dados que precisam ser incorporados: registro de recebimentos, cronograma de entregas e de aplicação, preços e notas. ` +
		"A área já adubada é mostrada em hectares porque a planilha só traz toneladas aplicadas nas áreas de vinhaça localizada (linhas 27-28)."
	);

	// ================================================================== 16. CONCLUSÕES
	pres.addSection({ title: "Decisão" });
	s = novoSlide("FECHAMENTO", "Decisão");
	titulos(s, "CONCLUSÕES E DECISÕES", "Três decisões para fechar a adubação da safra");
	txt(s, "Conclusões", { x: 0.6, y: 1.62, w: 6, h: 0.35, fontSize: 16, bold: true, color: C.accent2, objectName: "s16_c_t" });
	fx("s16_c_t", "fade", "apos", { duracao: 0.4 });
	const concl = [
		{ ic: I.cenaClara, t: `Restam ${fmt(A.total_a_tratar_ha)} ha a adubar em ${A.dias_safra} dias; ${pct(dA.participacao_16)} do volume é 16-00-20.` },
		{ ic: I.sacoClaro, t: `Estoque e compras cobrem ${pct(dA.cobertura_16_disponivel)} do 16-00-20: faltam ${fmt(A.compra_16_t)} t (A) ou ${fmt(B.compra_16_t)} t (B).` },
		{ ic: I.alertaClaro, t: `O 46-00-00 sobra ${fmt(b46A.saldo_t)}–${fmt(b46B.saldo_t)} t e só tem uso com vinhaça ou com uma fonte de K.` },
		{ ic: I.relogioClaro, t: "Sem preços e datas de entrega, orçamento e risco de ruptura seguem em aberto." },
	];
	concl.forEach((c, i) => {
		const y = 2.12 + i * 1.02, nm = `s16_c${i}`;
		s.addImage({ data: c.ic, x: 0.6, y: y + 0.04, w: 0.52, h: 0.52, objectName: `${nm}_ic` });
		txt(s, c.t, { x: 1.32, y, w: 5.3, h: 0.9, fontSize: 16, color: C.background1, objectName: `${nm}_t` });
		grp(nm, [`${nm}_ic`, `${nm}_t`]);
		fx(nm, "fade", "clique", { duracao: 0.5 });
	});
	caixa(s, { x: 7.05, y: 1.62, w: 5.68, h: 4.95, fill: { color: "FFFFFF" }, objectName: "s16_d_bg" });
	txt(s, "Decisões prioritárias", { x: 7.35, y: 1.82, w: 5.1, h: 0.4, fontSize: 18, bold: true, color: C.text2, objectName: "s16_d_t" });
	const decs16 = [
		{ t: "Aprovar a compra de 16-00-20", d: `${fmt(A.compra_16_t)} t (A) ou ${fmt(B.compra_16_t)} t (B), com entrega antes de ~${dataBR(cv.data_esgota_disponivel)}` },
		{ t: "Definir o destino do 46-00-00", d: "reprogramar o que está a receber, guardar para a próxima safra ou usar com uma fonte de K" },
		{ t: "Cobrar cronograma e preços", d: `datas das ${fmt(dA.a_receber_total_t)} t compradas e cotações para fechar o orçamento` },
	];
	const mD16 = ["s16_d_bg", "s16_d_t"];
	decs16.forEach((d, i) => {
		const y = 2.42 + i * 1.12;
		s.addShape(pres.shapes.OVAL, { x: 7.35, y: y + 0.04, w: 0.44, h: 0.44, fill: { color: C.accent1 }, objectName: `s16_d${i}_n` });
		txt(s, String(i + 1), { x: 7.35, y: y + 0.04, w: 0.44, h: 0.44, fontSize: 16, bold: true, color: C.background1, align: "center", valign: "middle", objectName: `s16_d${i}_nt` });
		txt(s, [{ text: d.t, options: { bold: true, fontSize: 16, color: C.text2, breakLine: true } }, { text: d.d, options: { fontSize: 13, color: C.text1 } }],
			{ x: 7.95, y, w: 4.6, h: 1.0, objectName: `s16_d${i}_tx` });
		mD16.push(`s16_d${i}_n`, `s16_d${i}_nt`, `s16_d${i}_tx`);
	});
	txt(s, "Próximos passos: validar a base → escolher o cenário → cotar e comprar → programar entregas → painel semanal",
		{ x: 7.35, y: 5.82, w: 5.2, h: 0.6, fontSize: 12, color: C.accent6, objectName: "s16_prox" });
	mD16.push("s16_prox");
	grp("s16_decisoes", mD16);
	fx("s16_decisoes", "wipe", "clique", { direcao: "esquerda", duracao: 0.7 });
	s.addNotes(
		"Fechamento. As quatro conclusões entram uma a uma, no clique; depois, também no clique, o quadro com as três decisões. " +
		`Decisão 1: compra de ${fmt(A.compra_16_t, 1)} t (A) ou ${fmt(B.compra_16_t, 1)} t (B) de 16-00-20 — a escolha do cenário depende de quantos dias a vinhaça localizada ainda atende o potássio. ` +
		"Decisão 2: destino do excedente de 46-00-00. Decisão 3: cronograma de entrega das compras já feitas e preços, sem os quais não há orçamento nem controle de ruptura."
	);

	// ================================================================== 17. ANEXO
	pres.addSection({ title: "Anexo" });
	s = novoSlide("CONTEUDO", "Anexo");
	titulos(s, "ANEXO · QUALIDADE DA BASE", "Premissas e pontos a corrigir na planilha");
	const somaPartes = ql.partes_ha.reduce((a, b) => a + b, 0);
	const inc = [
		["Ambas · linhas 26–29", `Quadro de vinhaça localizada não fecha: total ${fmt(ql.total_ha, 2)} ha × ${fmt(somaPartes, 2)} ha nas linhas; B27+B28 (${fmt(ql.partes_ha[0] + ql.partes_ha[1], 2)}) ≠ B22 (${fmt(A.realizado.crv_localizada_ha, 2)}).`, "Não muda a necessidade; impede conciliar o consumo realizado."],
		[`${ABA_B} · H38`, "Valor 0 digitado; na aba A é =C14−H36 (que daria −446,63 ha neste cenário).", "Não se atualiza se os dias de K mudarem."],
		["Ambas · D38 e H36", "Constantes digitadas (11.593 + 9.354 + 8.779; 1.200) sem referência de origem.", "Perde rastreabilidade por área."],
		["Ambas · K58 × K60:K61", "Oxyfertil (142,19 t) fora dos subtotais de socaria e de plantio.", "Estoque sem destino definido."],
		["Ambas · I49:J58 × B50:F51", "21-00-00, 13-00-34, 25-00-00 e Oxyfertil estão em estoque, mas fora do balanço.", "Saldo considera só 16-00-20 e 46-00-00."],
		["Ambas · nota A42", "Plantio de 2.400 ha (1.400 fundação + 1.000 renovação) sem necessidade calculada.", "Possível compra adicional de fórmulas de plantio."],
		["Planilha inteira", "Sem preços, datas de entrega, prazo de fornecimento, demanda mensal e ritmo da adubação mineral.", "Orçamento e curva real não são calculáveis."],
		["Rótulos", "'16-00-20.' e '16-00-20' (idem 46-00-00); grafias 'Fertlizante' e 'Rítmo'.", "Pode quebrar consolidações por produto."],
	];
	s.addTable([
		[th("Onde"), th("Achado"), th("Efeito")],
		...inc.map((r) => r.map((c) => td(c))),
	], {
		x: 0.6, y: 1.6, w: 12.13, colW: [2.45, 6.38, 3.3], fontSize: 11, color: HEX.dk1, valign: "middle",
		border: { type: "solid", pt: 0.75, color: EXTRA.linhaTabela }, margin: [0.03, 0.08, 0.03, 0.08], rowH: [0.34, 0.6, 0.48, 0.48, 0.48, 0.48, 0.48, 0.48, 0.48], objectName: "s17_tab",
	});
	fx("s17_tab", "fade", "apos", { duracao: 0.5 });
	txt(s, `Premissas desta apresentação: cenários = abas da planilha; nutrientes pela garantia da fórmula; curva de abastecimento com aplicação uniforme até ${dataBR(A.termino_safra)} (ilustrativa); sensibilidade por dia com a lógica da aba A; nenhum preço assumido. ${D.checks.filter((c) => !c.item.startsWith("Sensibilidade")).length} fórmulas e totais da planilha conferidos sem divergência.`,
		{ x: 0.6, y: 6.03, w: 12.13, h: 0.55, fontSize: 11, italic: true, color: C.accent6, objectName: "s17_prem" });
	fx("s17_prem", "fade", "apos", { duracao: 0.4 });
	fonte(s, `${ARQ} — abas ${ABA_A} e ${ABA_B}. Lista completa em README.md e na aba Inconsistências da planilha auxiliar.`);
	s.addNotes("Anexo para consulta: pontos da planilha que precisam de correção ou complemento e as premissas usadas nesta apresentação. Nenhum valor da planilha de origem foi alterado.");

	// ------------------------------------------------------------ gravação
	const base = path.join(BUILD, "base.pptx");
	await pres.writeFile({ fileName: base });
	await aplicarTema(base, THEME);
	fs.writeFileSync(path.join(BUILD, "animacoes.json"), JSON.stringify(SPEC, null, 1), "utf8");
	console.log(`${n} slides -> ${base}`);
	console.log(`${SPEC.efeitos.length} efeitos, ${SPEC.grupos.length} grupos, ${SPEC.graficos.length} ajustes de gráfico`);
}

// Grava as cores do tema no theme1.xml (o pptxgenjs só escreve as fontes).
async function aplicarTema(arquivo, tema) {
	const zip = await JSZip.loadAsync(fs.readFileSync(arquivo));
	const parte = "ppt/theme/theme1.xml";
	const slots = ["dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"];
	const esquema = `<a:clrScheme name="${tema.name}">` + slots.map((k) => `<a:${k}><a:srgbClr val="${tema.colors[k]}"/></a:${k}>`).join("") + "</a:clrScheme>";
	let xml = await zip.file(parte).async("string");
	xml = xml.replace(/<a:clrScheme\b[\s\S]*?<\/a:clrScheme>/, () => esquema)
		.replace(/(<a:(?:theme|fontScheme)\b[^>]*?\bname=")[^"]*"/g, (_, h) => `${h}${tema.name}"`);
	if (!xml.includes(esquema)) throw new Error("theme1.xml sem <a:clrScheme>");
	zip.file(parte, xml);
	for (const nome of Object.keys(zip.files)) {
		if (!nome.endsWith(".xml")) continue;
		const conteudo = await zip.file(nome).async("string");
		const ruim = conteudo.match(/<a:srgbClr val="(?![0-9A-Fa-f]{6}")[^"]*"/);
		if (ruim) throw new Error(`Cor inválida em ${nome}: ${ruim[0]}`);
	}
	fs.writeFileSync(arquivo, await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));
}

main().catch((e) => { console.error(e); process.exit(1); });
