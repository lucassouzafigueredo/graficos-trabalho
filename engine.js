// =============================================================================
// Gráficos FP&A — motor de layout
// Converte (dados + especificação + estilo) em uma lista de primitivas
// desenháveis (retângulos, linhas, textos, elipses), em pontos, com origem
// no canto superior esquerdo do quadro. Os renderizadores (SVG para a
// pré-visualização, Excel para inserir na planilha) apenas desenham a lista.
// =============================================================================

import { fmtNum, fmtPct } from "./formats.js?v=1.7.0";

export const TIPOS = {
  IBCS: { nome: "Cascata IBCS", W: 420, H: 250, comTipo: false },
  IBCSM: { nome: "Cascata IBCS – múltiplos", W: 280, H: 165, comTipo: false },
  EMP: { nome: "Cascata empilhada", W: 520, H: 320, comTipo: true },
  COL: { nome: "Colunas", W: 460, H: 280, comTipo: false },
  COMP: { nome: "Comparação", W: 340, H: 260, comTipo: false },
  TAB: { nome: "Tabela de variações", W: 640, H: 0, comTipo: false },
  ROS: { nome: "Roscas proporcionais", W: 600, H: 260, comTipo: false },
  BOL: { nome: "Matriz de bolhas", W: 600, H: 0, comTipo: false },
  PVM: { nome: "Ponte Preço / Volume / Mix", W: 460, H: 260, comTipo: false },
};

// Anotações que já vêm ligadas em cada tipo de gráfico
export function anotacoesPadrao(tipo) {
  const v = () => ({ on: false, de: 1, para: 0, modo: "P", balao: false });
  const a = { v: [v(), v()], cagr: { on: false, de: 1, para: 0 } };
  if (tipo === "IBCS" || tipo === "IBCSM" || tipo === "PVM") a.v[0].on = true;
  if (tipo === "COMP") { a.v[0].on = true; a.v[0].balao = true; }
  if (tipo === "COL") a.cagr.on = true;
  return a;
}

// -----------------------------------------------------------------------------
// Leitura dos dados (equivalente ao LerDados do VBA)
// values/texts: matrizes 2D vindas do Excel (range.values / range.text)
// -----------------------------------------------------------------------------
export function lerDados(values, texts, comTipo) {
  const c0 = comTipo ? 2 : 1;
  if (!values.length || values[0].length <= c0) {
    throw new Error(comTipo
      ? "Layout esperado: Rótulo | Tipo | Valores… (selecione pelo menos 3 colunas)."
      : "Layout esperado: Rótulo | Valores… (selecione pelo menos 2 colunas).");
  }
  const isNum = (x) => typeof x === "number" && isFinite(x);
  const temCab = values[0].slice(c0).some((x) => x !== "" && x !== null && !isNum(x));
  const r0 = temCab ? 1 : 0;
  let rot = [], tipo = [], v = [], txt = [];
  let cab = values[0].slice(c0).map((x, i) => (temCab ? String(texts[0][c0 + i]) : `Série ${i + 1}`));
  for (let r = r0; r < values.length; r++) {
    const linhaVazia = values[r].every((x) => x === "" || x === null);
    if (linhaVazia) continue;
    rot.push(String(texts[r][0] ?? ""));
    tipo.push(comTipo ? String(values[r][1] ?? "").trim().toUpperCase() : "");
    v.push(values[r].slice(c0).map((x) => (isNum(x) ? x : null)));
    txt.push(texts[r].slice(c0).map((x) => String(x ?? "").trim()));
  }
  if (!rot.length) throw new Error("Não há linhas de dados.");
  // descarta colunas sem nenhum número (ex.: "Tipo" ou colunas de espaçamento)
  const manter = cab.map((_, c) => v.some((linha) => linha[c] !== null));
  if (!manter.some(Boolean)) throw new Error("Não encontrei valores numéricos no intervalo.");
  // colunas só de texto (ex.: "Comentário") ficam guardadas à parte
  const cabTextos = cab.filter((_, c) => !manter[c]);
  const textos = txt.map((linha) => linha.filter((_, c) => !manter[c]));
  cab = cab.filter((_, c) => manter[c]);
  v = v.map((linha) => linha.filter((_, c) => manter[c]));
  return { rot, cab, v, tipo, nLin: rot.length, nCol: cab.length, cabTextos, textos };
}

// -----------------------------------------------------------------------------
// Construtor de primitivas
// -----------------------------------------------------------------------------
class Tela {
  constructor(estilo) {
    this.p = [];
    this.e = estilo;
    this.vMin = 0; this.vMax = 1; this.pTop = 0; this.pBot = 1;
    this.barras = []; this.bw = 0;
  }
  Y(v) {
    if (this.vMax === this.vMin) return this.pBot;
    return this.pBot - ((v - this.vMin) / (this.vMax - this.vMin)) * (this.pBot - this.pTop);
  }
  clip(v) { return Math.min(this.vMax, Math.max(this.vMin, v)); }
  rect(x, ya, w, yb, fill, extra = {}) {
    const y = Math.min(ya, yb), h = Math.max(0.75, Math.abs(yb - ya));
    this.p.push({ t: "rect", x, y, w, h, fill, ...extra });
  }
  line(x1, y1, x2, y2, color, w, extra = {}) {
    this.p.push({ t: "line", x1, y1, x2, y2, color, w, dash: "solid", ...extra });
  }
  // ancora: "top" | "middle" | "bottom";  alinh: "left" | "center" | "right"
  text(text, x, y, ancora, w, opts = {}) {
    const size = opts.size ?? this.e.fonte;
    const maxL = opts.maxLinhas ?? 3;
    const linhas = opts.wrap ? maxL : String(text).split("\n").length;
    const h = opts.wrap ? size * 1.35 * maxL : size * (0.2 + 1.4 * linhas);
    const top = ancora === "top" ? y : ancora === "middle" ? y - h / 2 : y - h;
    const alinh = opts.align ?? "center";
    const left = alinh === "left" ? x : alinh === "right" ? x - w : x - w / 2;
    this.p.push({
      t: "text", x: left, y: top, w: Math.max(10, w), h, text: String(text), size,
      color: opts.color ?? this.e.texto, bold: !!opts.bold, italic: !!opts.italic,
      align: alinh, valign: ancora, wrap: !!opts.wrap, maxLinhas: maxL,
    });
  }
  ellipse(text, x, cy, ladoEsquerdo) {
    const size = this.e.fonte + 3;
    const linhas = String(text).split("\n");
    const maxLen = Math.max(...linhas.map((l) => l.length));
    const w = maxLen * size * 0.6 + 22, h = size * (1.3 * linhas.length + 1);
    const left = ladoEsquerdo ? x : x - w / 2;
    this.p.push({ t: "ellipse", x: left, y: cy - h / 2, w, h, text, size });
  }
  // anel entre os raios r0 e r1, do ângulo a0 ao a1 (radianos, 0 = 12h, sentido horário)
  arco(cx, cy, r0, r1, a0, a1, fill) {
    if (a1 - a0 <= 1e-4) return;
    if (a1 - a0 > Math.PI * 1.98) { const m = (a0 + a1) / 2; this.arco(cx, cy, r0, r1, a0, m, fill); this.arco(cx, cy, r0, r1, m, a1, fill); return; }
    this.p.push({ t: "arco", cx, cy, r0, r1, a0, a1, fill });
  }
  dot(cx, cy, r, color) { this.p.push({ t: "dot", cx, cy, r, color }); }
  marcador(cx, cy, r, n, color) { this.p.push({ t: "marcador", cx, cy, r, n: String(n), color, size: Math.round(r * 1.1) }); }
  arrow(x1, y1, x2, y2, color, w) { this.line(x1, y1, x2, y2, color, w, { arrow: true }); }
  quebra(bx, bw, yb) {
    const k = 5;
    this.line(bx - 5, yb + k, bx + bw + 5, yb - k, "#FFFFFF", 5);
    this.line(bx - 5, yb + k - 3, bx + bw + 5, yb - k - 3, "#000000", 1.25);
    this.line(bx - 5, yb + k + 3, bx + bw + 5, yb - k + 3, "#000000", 1.25);
  }
  fundo(L, T, W, H) { this.p.push({ t: "rect", x: L - 8, y: T - 8, w: W + 16, h: H + 16, fill: "#FFFFFF", fundo: true }); }
  iniciarBarras(bw) { this.barras = []; this.bw = bw; }
  regBarra(cx, valor, nivel) { this.barras.push({ cx, valor, nivel }); }
}


// -----------------------------------------------------------------------------
// Cenários (estilo "Orçamento x Forecast"): a aparência da barra vem do rótulo
//   Orçamento: BGT, Orçamento, Budget, Plano | Forecast: 6+6, 9+3, FC, Forecast
//   Real: o resto (LY, AC, Real, anos...)
// -----------------------------------------------------------------------------
export function cenarioDe(rotulo) {
  const r = String(rotulo).toUpperCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
  if (/\b\d{1,2}\s*\+\s*\d{1,2}\b/.test(r) || /\b(FC|FCST|FORECAST|REFORECAST|RF|PROJ|PROJECAO|PREVISTO)\b/.test(r)) return "fc";
  if (/\b(BGT|BUDGET|ORC|ORCAMENTO|ORCADO|PLANO|PLAN|PBR|META)\b/.test(r)) return "orc";
  return "real";
}

// Aparências por cenário usadas quando o estilo não define as suas
export const CENARIOS_PADRAO = {
  real: { fill: "#A9CBF5" },
  orc: { fill: "#4F8DF0", listra: "#9DC1F7" },
  fc: [{ fill: "#26408F", listra: "#6F8BD3" }, { fill: "#111A4A", listra: "#4C5E9E" }],
};

// Aparência escolhida pelo usuário para uma barra (painel > Cores das barras)
//   { modo: "auto" | "real" | "orc" | "fc1" | "fc2" | "cor", fill, listrado, listra }
function aparenciaEscolhida(e, esc) {
  if (!esc || !esc.modo || esc.modo === "auto") return null;
  const cen = e.cenarios ?? CENARIOS_PADRAO;
  switch (esc.modo) {
    case "real": return { ...cen.real };
    case "orc": return { ...cen.orc };
    case "fc1": return { ...cen.fc[0] };
    case "fc2": return { ...(cen.fc[1] ?? cen.fc[0]) };
    case "cor": return esc.listrado ? { fill: esc.fill, listra: esc.listra } : { fill: esc.fill };
    default: return null;
  }
}

// Só as escolhas explícitas do usuário (índice -> aparência ou null)
function escolhidas(e, rotulos, escolhas) {
  return rotulos.map((r) => aparenciaEscolhida(e, escolhas?.[r]));
}

// Aparência de cada barra: escolha do usuário > cenário do estilo > null (padrão do gráfico)
export function aparencias(e, rotulos, escolhas) {
  const esc = escolhidas(e, rotulos, escolhas);
  let nFc = 0;
  return rotulos.map((r, i) => {
    let auto = null;
    if (e.cenarios) {
      const c = cenarioDe(r);
      if (c === "fc") { const lista = e.cenarios.fc; auto = { ...lista[Math.min(nFc++, lista.length - 1)] }; }
      else auto = { ...e.cenarios[c] };
    }
    return esc[i] ?? auto;
  });
}
const rectAparencia = (tela, bx, ya, bw, yb, ap) => tela.rect(bx, ya, bw, yb, ap.fill, { listra: ap.listra, stroke: ap.stroke });

const corTextoSobre = (hex) => {
  const n = parseInt(hex.slice(1), 16);
  const r = (n >> 16) & 255, g = (n >> 8) & 255, b = n & 255;
  return 0.299 * r + 0.587 * g + 0.114 * b < 150 ? "#FFFFFF" : "#000000";
};
const ix = (v, n) => (v < 1 || v > n ? n : v);
const num = (x) => (x === null || x === undefined ? 0 : x);
const larguraVar = (a) => (a.balao || a.modo === "B" ? 92 : 64);

function margemDireita(anot) {
  return anot.v.reduce((m, a) => m + (a.on ? larguraVar(a) : 0), 0);
}
const reservaTopo = (anot) => (anot.cagr.on ? 48 : 0);

// -----------------------------------------------------------------------------
// API principal
// -----------------------------------------------------------------------------
export function montarGrafico(dados, specOrig, estilo) {
  const tela = new Tela(estilo);
  // título, subtítulo, fonte e nota são desenhados aqui, em volta do gráfico (padrão único)
  const moldura = { titulo: specOrig.titulo, subtitulo: specOrig.subtitulo, fonte: specOrig.fonteDados, nota: specOrig.nota };
  const spec = { ...specOrig, titulo: "" };
  tela.escalaFixa = spec.escalaFixa ?? null;
  tela.escolhas = spec.barras ?? {};
  const anot = spec.anot ?? anotacoesPadrao(spec.tipo);
  const W = spec.W ?? TIPOS[spec.tipo].W, H = spec.H ?? TIPOS[spec.tipo].H;
  let largura = W, altura = H;
  switch (spec.tipo) {
    case "IBCS":
      cascataIBCS(tela, dados, 0, 0, 0, W, H, spec.titulo, null, anot, true);
      break;
    case "IBCSM": {
      const r = multiplos(tela, dados, W, H, spec.titulo, anot);
      largura = r.W; altura = r.H;
      break;
    }
    case "EMP":
      cascataEmpilhada(tela, dados, W, H, spec.titulo, spec.corte || 0, anot);
      break;
    case "COL":
      colunas(tela, dados, W, H, spec.titulo, anot);
      break;
    case "COMP":
      comparacao(tela, dados, W, H, spec.titulo, anot);
      break;
    case "TAB": {
      const r = tabela(tela, dados, W, spec);
      altura = r.altura; largura = r.largura;
      break;
    }
    case "PVM":
      tela.zoom = spec.pvmZoom !== false;
      cascataIBCS(tela, dadosPVM(pvmCalcular(dados, spec), spec), 0, 0, 0, W, H, "", null, anot, true);
      break;
    case "ROS":
      altura = roscas(tela, dados, W, H, spec);
      break;
    case "BOL": {
      const r = bolhas(tela, dados, W, spec);
      altura = r.altura;
      break;
    }
    default:
      throw new Error("Tipo de gráfico desconhecido: " + spec.tipo);
  }
  const r = emoldurar(tela, largura, altura, moldura);
  return { largura, altura: r.altura, prims: tela.p, escala: tela.escalaNatural ?? null };
}

// Escala: guarda a natural (para o grupo "mesma escala") e aplica a fixa, se houver.
// Nunca encolhe: a escala fixa só amplia a natural, para nenhum dado ficar de fora.
function escalar(tela, nat) {
  tela.escalaNatural = nat;
  const f = tela.escalaFixa;
  if (!f) return nat;
  const out = { min: Math.min(nat.min, f.min ?? nat.min), max: Math.max(nat.max, f.max ?? nat.max) };
  if (nat.d !== undefined) out.d = Math.max(nat.d, f.d ?? 0);
  return out;
}
export function unirEscalas(lista) {
  const v = lista.filter(Boolean);
  if (!v.length) return null;
  const out = { min: Math.min(...v.map((x) => x.min)), max: Math.max(...v.map((x) => x.max)) };
  const ds = v.map((x) => x.d).filter((x) => x !== undefined);
  if (ds.length) out.d = Math.max(...ds);
  return out;
}

// Linhas estimadas de um texto quebrado na largura w
const linhasTexto = (t, w, size, max) => Math.min(max, Math.max(1, Math.ceil((String(t).length * size * 0.55) / Math.max(40, w))));

// Título-mensagem + subtítulo em cima, fonte/nota embaixo; desloca o gráfico e refaz o fundo
function emoldurar(tela, W, H, m) {
  const e = tela.e;
  const tt = (m.titulo ?? "").trim(), st = (m.subtitulo ?? "").trim(), fo = (m.fonte ?? "").trim(), no = (m.nota ?? "").trim();
  const prims = tela.p.filter((q) => !q.fundo);
  let topo = 0;
  const cab = [];
  if (tt) {
    const sz = e.fonte + 3, nl = linhasTexto(tt, W, sz, 2);
    cab.push({ t: "text", x: 0, y: topo, w: W, h: sz * 1.35 * nl, text: tt, size: sz, color: "#000000", bold: true, italic: false, align: "left", valign: "top", wrap: nl > 1, maxLinhas: nl });
    topo += sz * 1.35 * nl + 2;
  }
  if (st) {
    const sz = e.fonte + 1;
    cab.push({ t: "text", x: 0, y: topo, w: W, h: sz * 1.4, text: st, size: sz, color: "#595959", bold: false, italic: false, align: "left", valign: "top", wrap: false, maxLinhas: 1 });
    topo += sz * 1.4 + 2;
  }
  if (cab.length) { cab.push({ t: "line", x1: 0, y1: topo + 2, x2: W, y2: topo + 2, color: "#000000", w: 0.75, dash: "solid" }); topo += 14; }
  if (topo) prims.forEach((q) => deslocar(q, topo));
  let fim = topo + H;
  const rod = [];
  if (fo || no) {
    fim += 10;
    rod.push({ t: "line", x1: 0, y1: fim, x2: W, y2: fim, color: "#BFBFBF", w: 0.75, dash: "solid" });
    fim += 4;
    const sz = Math.max(7, e.fonte - 1);
    for (const [rotulo, txt] of [["Fonte: ", fo], ["Nota: ", no]]) {
      if (!txt) continue;
      const nl = linhasTexto(rotulo + txt, W, sz, 3);
      rod.push({ t: "text", x: 0, y: fim, w: W, h: sz * 1.35 * nl, text: rotulo + txt, size: sz, color: "#7F7F7F", bold: false, italic: false, align: "left", valign: "top", wrap: nl > 1, maxLinhas: nl });
      fim += sz * 1.35 * nl + 1;
    }
  }
  if (!cab.length && !rod.length) return { altura: H };
  const fundo = { t: "rect", x: -8, y: -8, w: W + 16, h: fim + 16, fill: "#FFFFFF", fundo: true };
  tela.p = [fundo, ...cab, ...prims, ...rod];
  return { altura: fim };
}

function deslocar(q, dy) {
  if ("y" in q) q.y += dy;
  if ("y1" in q) { q.y1 += dy; q.y2 += dy; }
  if ("cy" in q) q.cy += dy;
}

// ---------------------------- Cascata IBCS -----------------------------------
function calcCascata(d, col) {
  const n = d.nLin, lo = [], hi = [], ac = [];
  const ini = num(d.v[0][col]);
  let acum = ini;
  lo[0] = Math.min(0, ini); hi[0] = Math.max(0, ini); ac[0] = ini;
  for (let i = 1; i < n - 1; i++) {
    const dv = num(d.v[i][col]);
    lo[i] = Math.min(acum, acum + dv); hi[i] = Math.max(acum, acum + dv);
    acum += dv; ac[i] = acum;
  }
  const fim = d.v[n - 1][col] === null ? acum : d.v[n - 1][col];
  lo[n - 1] = Math.min(0, fim); hi[n - 1] = Math.max(0, fim); ac[n - 1] = fim;
  return { lo, hi, ac, ini, fim };
}

function cascataIBCS(tela, d, col, L, T, W, H, titulo, escala, anot, comFundo) {
  const e = tela.e, n = d.nLin;
  if (n < 2) throw new Error("A cascata precisa de pelo menos 2 linhas (inicial e final).");
  const { lo, hi, ac, ini, fim } = calcCascata(d, col);
  if (escala) { tela.vMin = escala[0]; tela.vMax = escala[1]; }
  else { const s = escalar(tela, { min: Math.min(0, ...lo), max: Math.max(0, ...hi) }); tela.vMin = s.min; tela.vMax = s.max; }
  if (tela.vMax === tela.vMin) tela.vMax = tela.vMin + 1;
  // zoom (ponte PVM): corta o eixo abaixo do menor nível para os efeitos aparecerem; totais ganham "//"
  let cortado = false;
  if (tela.zoom && !escala && Math.min(...lo) >= 0) {
    const niveis = ac.slice(1, -1).concat(lo.slice(1, -1), [ini, fim]);
    const nMin = Math.min(...niveis), amp = Math.max(...hi) - nMin;
    const base = Math.max(0, nMin - Math.max(amp * 1.2, Math.max(...hi) * 0.04));
    if (base > Math.max(...hi) * 0.15) { tela.vMin = base; cortado = true; }
  }
  const Yc = (v) => tela.Y(tela.clip(v));

  if (comFundo) tela.fundo(L, T, W, H);
  let titH = 0;
  if (titulo) { titH = 20; tela.text(titulo, L, T, "top", W, { size: e.fonte + 3, align: "left" }); }
  tela.pTop = T + titH + 16 + reservaTopo(anot);
  tela.pBot = T + H - 30 - (tela.vMin < 0 ? 16 : 0);
  const xL = L + 2, xR = L + W - 2 - margemDireita(anot);
  const catW = (xR - xL) / n, bw = catW * 0.66;
  tela.iniciarBarras(bw);
  const ap = aparencias(e, d.rot, tela.escolhas);
  const esc = escolhidas(e, d.rot, tela.escolhas);
  const rotulo = (txt, acima, cx, i) => acima
    ? tela.text(txt, cx, tela.Y(hi[i]) - 2, "bottom", catW + 16, { size: e.fonte + 1 })
    : tela.text(txt, cx, tela.Y(lo[i]) + 2, "top", catW + 16, { size: e.fonte + 1 });

  for (let i = 0; i < n; i++) {
    const cx = xL + catW * (i + 0.5), bx = cx - bw / 2;
    tela.regBarra(cx, ac[i], ac[i]);
    if (i === 0) {
      rectAparencia(tela, bx, Yc(lo[i]), bw, tela.Y(hi[i]), ap[i] ?? { fill: "#FFFFFF", listra: e.total, stroke: e.total });
      if (cortado) tela.quebra(bx, bw, tela.pBot - 16);
      rotulo(fmtNum(ini), ini >= 0, cx, i);
    } else if (i === n - 1) {
      rectAparencia(tela, bx, Yc(lo[i]), bw, tela.Y(hi[i]), ap[i] ?? { fill: e.total });
      if (cortado) tela.quebra(bx, bw, tela.pBot - 16);
      rotulo(fmtNum(fim), fim >= 0, cx, i);
    } else {
      const dv = ac[i] - ac[i - 1];
      if (esc[i]) rectAparencia(tela, bx, tela.Y(lo[i]), bw, tela.Y(hi[i]), esc[i]);
      else tela.rect(bx, tela.Y(lo[i]), bw, tela.Y(hi[i]), dv >= 0 ? e.pos : e.neg);
      if (dv !== 0) rotulo(fmtNum(dv, true), dv >= 0, cx, i);
    }
    if (i < n - 1) tela.line(bx + bw, tela.Y(ac[i]), bx + catW, tela.Y(ac[i]), "#AFAFAF", 0.75);
    tela.text(d.rot[i], cx, T + H - 26, "top", catW, { size: e.fonte + 1, wrap: true });
  }
  tela.line(xL, Yc(0), xR, Yc(0), "#404040", 1.5);
  anotacoes(tela, anot, true, xR);
}

function multiplos(tela, d, Wp, Hp, titulo, anot) {
  const GX = 14, GY = 10;
  let gMin = 0, gMax = 0;
  for (let c = 0; c < d.nCol; c++) {
    const { lo, hi } = calcCascata(d, c);
    gMin = Math.min(gMin, ...lo); gMax = Math.max(gMax, ...hi);
  }
  ({ min: gMin, max: gMax } = escalar(tela, { min: gMin, max: gMax }));
  const porLinha = Math.min(3, d.nCol), nLinhas = Math.ceil(d.nCol / porLinha);
  const titH = titulo ? 26 : 0;
  const W = porLinha * (Wp + GX) - GX, H = titH + nLinhas * (Hp + GY) - GY;
  tela.fundo(0, 0, W, H);
  if (titulo) tela.text(titulo, 0, 0, "top", W, { size: tela.e.fonte + 5, bold: true, align: "left" });
  for (let c = 0; c < d.nCol; c++) {
    const r = Math.floor(c / porLinha), k = c % porLinha;
    cascataIBCS(tela, d, c, k * (Wp + GX), titH + r * (Hp + GY), Wp, Hp, d.cab[c], [gMin, gMax], anot, false);
  }
  return { W, H };
}

// ---------------------------- Cascata empilhada -------------------------------
function cascataEmpilhada(tela, d, W, H, titulo, corte, anot) {
  const e = tela.e, n = d.nLin, m = d.nCol;
  const tp = [], lo = [], hi = [], ac = [], tot = [];
  let acum = 0;
  for (let i = 0; i < n; i++) {
    let t = d.tipo[i];
    if (!["T", "D", "F"].includes(t)) t = i === 0 || i === n - 1 ? "T" : "D";
    tp[i] = t;
    const s = d.v[i].reduce((a, x) => a + num(x), 0);
    tot[i] = s;
    if (t === "T") { lo[i] = Math.min(0, s); hi[i] = Math.max(0, s); acum = s; }
    else { lo[i] = Math.min(acum, acum + s); hi[i] = Math.max(acum, acum + s); acum += s; }
    ac[i] = acum;
  }
  ({ min: tela.vMin, max: tela.vMax } = escalar(tela, { min: Math.min(0, ...lo), max: Math.max(0, ...hi) }));
  if (corte > 0) {
    if (tela.vMax >= -tela.vMin) { if (corte < tela.vMax) tela.vMax = corte; }
    else if (corte < -tela.vMin) tela.vMin = -corte;
  }
  if (tela.vMax === tela.vMin) tela.vMax = tela.vMin + 1;

  tela.fundo(0, 0, W, H);
  let titH = 0;
  if (titulo) {
    titH = 32;
    tela.text(titulo, 0, 0, "top", W, { size: e.fonte + 2, bold: true, align: "left", color: "#000000" });
    tela.line(0, 20, W, 20, "#000000", 0.75);
  }
  const legW = m > 1 ? 90 : 0;
  tela.pTop = titH + 20 + reservaTopo(anot) - (tela.vMax <= 0 ? 12 : 0);
  tela.pBot = H - 36 - (tela.vMin < 0 ? 20 : 0);
  const xL = legW, xR = W - margemDireita(anot);
  const catW = (xR - xL) / n, bw = catW * 0.6;
  tela.iniciarBarras(bw);
  const C = (v) => tela.clip(v), Y = (v) => tela.Y(v);
  const fmtDelta = (v) => (e.sinalDelta ? fmtNum(v, true) : fmtNum(Math.abs(v)));
  let legOk = false;
  const ap = aparencias(e, d.rot, tela.escolhas);

  for (let i = 0; i < n; i++) {
    const cx = xL + catW * (i + 0.5), bx = cx - bw / 2;
    tela.regBarra(cx, tp[i] === "T" ? tot[i] : ac[i], tp[i] === "T" ? tot[i] : ac[i]);
    // corte de eixo (desenhado antes dos rótulos para não cobri-los)
    const quebraSup = hi[i] > tela.vMax && lo[i] < tela.vMax;
    const quebraInf = lo[i] < tela.vMin && hi[i] > tela.vMin;
    const desenharQuebras = () => {
      if (quebraSup) tela.quebra(bx, bw, Y(tela.vMax) + Math.min(22, (Y(C(lo[i])) - Y(tela.vMax)) * 0.35));
      if (quebraInf) tela.quebra(bx, bw, Y(tela.vMin) - Math.min(22, (Y(tela.vMin) - Y(C(hi[i]))) * 0.35));
    };
    if (tp[i] === "D") {
      const escD = escolhidas(e, [d.rot[i]], tela.escolhas)[0];
      const cor = escD ? escD.fill : e.porSinal ? (tot[i] >= 0 ? e.pos : e.neg) : e.delta;
      if (escD) rectAparencia(tela, bx, Y(C(lo[i])), bw, Y(C(hi[i])), escD);
      else tela.rect(bx, Y(C(lo[i])), bw, Y(C(hi[i])), cor);
      desenharQuebras();
      const txt = fmtDelta(tot[i]);
      if (quebraSup || quebraInf)
        tela.text(txt, cx, (Y(C(lo[i])) + Y(C(hi[i]))) / 2 + (quebraSup ? 6 : -6), "middle", bw, { size: e.fonte + 2, color: corTextoSobre(cor) });
      else if (Y(C(lo[i])) - Y(C(hi[i])) >= e.fonte + 4)
        tela.text(txt, cx, (Y(C(lo[i])) + Y(C(hi[i]))) / 2, "middle", bw, { size: e.fonte + 2, color: corTextoSobre(cor) });
      else if (tot[i] < 0) tela.text(txt, cx, Y(C(lo[i])) + 2, "top", catW, { size: e.fonte + 2 });
      else tela.text(txt, cx, Y(C(hi[i])) - 2, "bottom", catW, { size: e.fonte + 2 });
    } else {
      let cur = lo[i];
      for (let c = 0; c < m; c++) {
        const a = Math.abs(num(d.v[i][c]));
        if (a <= 0) continue;
        const cor = m === 1 ? e.total : e.pilha[c % 6];
        const y1 = Y(C(cur)), y2 = Y(C(cur + a));
        if (m === 1 && tp[i] === "T" && ap[i]) rectAparencia(tela, bx, y1, bw, y2, ap[i]);
        else tela.rect(bx, y1, bw, y2, cor);
        if (m > 1 && y1 - y2 >= e.fonte + 3)
          tela.text(fmtNum(a), cx, (y1 + y2) / 2, "middle", bw, { size: e.fonte + 1, color: corTextoSobre(cor) });
        if (m > 1 && !legOk && tp[i] === "T")
          tela.text(d.cab[c], bx - 8, (y1 + y2) / 2, "middle", legW - 8, { size: e.fonte + 1, align: "right" });
        cur += a;
      }
      if (m > 1 && tp[i] === "T") legOk = true;
      desenharQuebras();
      const txt = tp[i] === "T" ? fmtNum(tot[i]) : fmtDelta(tot[i]);
      if (tot[i] < 0 && hi[i] <= 0) tela.text(txt, cx, Y(C(lo[i])) + 2, "top", catW, { size: e.fonte + 2 });
      else tela.text(txt, cx, Y(C(hi[i])) - 2, "bottom", catW, { size: e.fonte + 2 });
    }
    if (i < n - 1) {
      const ligar = tp[i + 1] !== "T" || Math.abs(tot[i + 1] - ac[i]) <= 0.005 * Math.max(1, Math.abs(ac[i]));
      if (ligar) tela.line(bx + bw, Y(C(ac[i])), bx + catW, Y(C(ac[i])), "#404040", 0.75, { dash: "dot" });
    }
    tela.text(d.rot[i], cx, H - 31, "top", catW - 4, { size: e.fonte + 1, wrap: true, bold: i === n - 1 });
  }
  tela.line(xL, Y(0), xR, Y(0), "#404040", 1);
  anotacoes(tela, anot, false, xR);
}

// ---------------------------- Colunas ----------------------------------------
function colunas(tela, d, W, H, titulo, anot) {
  const e = tela.e, n = d.nLin;
  const tot = d.v.map((linha) => linha.reduce((a, x) => a + num(x), 0));
  ({ min: tela.vMin, max: tela.vMax } = escalar(tela, { min: 0, max: Math.max(0, ...tot) }));
  if (tela.vMax === tela.vMin) tela.vMax = 1;
  tela.fundo(0, 0, W, H);
  let titH = 0;
  if (titulo) { titH = 24; tela.text(titulo, 0, 0, "top", W, { size: e.fonte + 3, bold: true, align: "left", color: "#000000" }); }
  tela.pTop = titH + 18 + reservaTopo(anot);
  tela.pBot = H - 22;
  const xL = 0, xR = W - margemDireita(anot);
  const catW = (xR - xL) / n, bw = catW * 0.62;
  tela.iniciarBarras(bw);
  const ap = aparencias(e, d.rot, tela.escolhas);
  for (let i = 0; i < n; i++) {
    const cx = xL + catW * (i + 0.5), bx = cx - bw / 2;
    tela.regBarra(cx, tot[i], tot[i]);
    let cur = 0;
    d.v[i].forEach((x, c) => {
      const a = Math.abs(num(x));
      if (a <= 0) return;
      if (d.nCol === 1 && ap[i]) rectAparencia(tela, bx, tela.Y(cur), bw, tela.Y(cur + a), ap[i]);
      else tela.rect(bx, tela.Y(cur), bw, tela.Y(cur + a), e.colunas[c % 4]);
      cur += a;
    });
    tela.text(fmtNum(tot[i]), cx, tela.Y(tot[i]) - 3, "bottom", catW, { size: e.fonte + 4 });
    tela.text(d.rot[i], cx, tela.pBot + 4, "top", catW, { size: e.fonte + 1, wrap: true });
  }
  tela.line(xL, tela.Y(0), xR, tela.Y(0), "#404040", 0.75);
  anotacoes(tela, anot, false, xR);
}

// ---------------------------- Comparação -------------------------------------
function comparacao(tela, d, W, H, titulo, anot) {
  const e = tela.e, n = d.nLin;
  if (n < 2) throw new Error("A comparação precisa de pelo menos 2 linhas.");
  const vals = d.v.map((linha) => num(linha[0]));
  ({ min: tela.vMin, max: tela.vMax } = escalar(tela, { min: Math.min(0, ...vals), max: Math.max(0, ...vals) }));
  if (tela.vMax === tela.vMin) tela.vMax = tela.vMin + 1;
  tela.fundo(0, 0, W, H);
  let titH = 0;
  if (titulo) { titH = 24; tela.text(titulo, 0, 0, "top", W, { size: e.fonte + 3, bold: true, align: "left", color: "#000000" }); }
  tela.pTop = titH + 26 + reservaTopo(anot);
  tela.pBot = H - 22;
  const xL = 0, xR = W - margemDireita(anot);
  const catW = (xR - xL) / n, bw = catW * 0.6;
  tela.iniciarBarras(bw);
  const ap = aparencias(e, d.rot, tela.escolhas);
  vals.forEach((v, i) => {
    const cx = xL + catW * (i + 0.5), bx = cx - bw / 2;
    tela.regBarra(cx, v, v);
    rectAparencia(tela, bx, tela.Y(0), bw, tela.Y(v), ap[i] ?? { fill: i === 0 ? e.colunas[0] : e.colunas[1] });
    if (v >= 0) tela.text(fmtNum(v), cx, tela.Y(v) - 3, "bottom", catW, { size: e.fonte + 6 });
    else tela.text(fmtNum(v), cx, tela.Y(v) + 3, "top", catW, { size: e.fonte + 6 });
    tela.text(d.rot[i], cx, tela.pBot + 4, "top", catW, { size: e.fonte + 1, wrap: true });
  });
  tela.line(xL, tela.Y(0), xR, tela.Y(0), "#404040", 0.75);
  anotacoes(tela, anot, false, xR);
}


// ---------------------------- Tabela de variações (estilo IBCS) ---------------
// Dados: Conta | Referência (BGT/FC/PY) | Atual. Colunas desenhadas:
//   referência (número) | atual (barra) | Δ absoluta (barra) | Δ % (alfinete)
// Cada linha pode ser: invertida (custo: queda é favorável), resultado (subtotal),
// negrito ou oculta. O padrão vem do nome da conta e pode ser corrigido no painel.

const normalizar = (t) => String(t).toUpperCase().normalize("NFD").replace(/[̀-ͯ]/g, "");
const RX_CUSTO = /(^|[^A-Z])(CPV|CMV|CSP|CUSTOS?|DESPESAS?|DEDUC(AO|OES)|PDD|PERDAS?|IMPOSTOS?|TRIBUTOS?|IR\s*\/?\s*CS|IRPJ|CSLL|D\s*&\s*A|DEPRECIAC|AMORTIZAC|SG\s*&\s*A|OPEX|COGS|TAXES?|EXPENSES?|COSTS?)([^A-Z]|$)/;
const RX_RESULTADO = /^(\W*)(LUCRO|PREJUIZO|EBITDA|EBIT|MARGEM|RECEITA LIQUIDA|RESULTADO (OPERACIONAL|LIQUIDO|ANTES|DO EXERCICIO|BRUTO)|TOTAL|GROSS PROFIT|OPERATING PROFIT|NET (PROFIT|INCOME|SALES))/;

export function flagsAutomaticas(rotulo, ref, ac) {
  const r = normalizar(rotulo);
  const res = RX_RESULTADO.test(r);
  // só inverte se o custo estiver lançado positivo (se já vier negativo, o sinal resolve)
  const inv = !res && RX_CUSTO.test(r) && (ref > 0 || ac > 0);
  return { inv, res, neg: res, ocultar: false };
}

export function flagsLinha(spec, rotulo, ref, ac) {
  return { ...flagsAutomaticas(rotulo, ref, ac), ...(spec?.linhas?.[rotulo] ?? {}) };
}

// Unidade única para a tabela inteira (automática: pelo maior valor)
export function unidadeTabela(spec, valores) {
  if (spec?.unidade && spec.unidade !== "auto") return { unidade: spec.unidade, casas: spec.casas, sufixo: "" };
  const max = Math.max(0, ...valores.map(Math.abs));
  const inteiros = valores.every((v) => Number.isInteger(v));
  if (max < 1e5) return { unidade: "un", casas: spec?.casas ?? (inteiros ? 0 : 1), sufixo: "" };
  if (max < 1e8) return { unidade: "mil", casas: spec?.casas ?? 1, sufixo: "K" };
  return { unidade: "mi", casas: spec?.casas ?? 1, sufixo: "M" };
}

export function fmtTab(v, spec, comSinal = false) {
  const u = spec?.unidade ?? "un";
  const div = { un: 1, mil: 1e3, mi: 1e6, bi: 1e9 }[u] ?? 1;
  const casas = spec?.casas ?? (u === "un" ? 0 : 1);
  const s = new Intl.NumberFormat("pt-BR", { minimumFractionDigits: casas, maximumFractionDigits: casas }).format(Math.abs(v) / div) + (spec?.sufixo ?? "");
  if (v < 0 && /[1-9]/.test(s)) return "-" + s;
  if (comSinal && v > 0) return "+" + s;
  return s;
}

// Quais colunas numéricas são o "atual" e quais são as referências.
//   automático: atual = a coluna "Real/AC" (pelo nome) ou a última; referências = as demais (até 3)
export function colunasTabela(spec, d) {
  const n = d.nCol;
  let iAc = Number.isInteger(spec?.colAtual) && spec.colAtual >= 0 && spec.colAtual < n ? spec.colAtual : -1;
  if (iAc < 0) {
    const reais = d.cab.map((c, i) => (cenarioDe(c) === "real" ? i : -1)).filter((i) => i >= 0);
    iAc = reais.length === 1 ? reais[0] : n === 2 ? 1 : reais.length ? reais[reais.length - 1] : n - 1;
  }
  let iRefs = Array.isArray(spec?.refs) ? spec.refs.filter((i) => i !== iAc && i >= 0 && i < n) : [];
  if (!iRefs.length) iRefs = [...Array(n).keys()].filter((i) => i !== iAc).slice(0, 3);
  return { iAc, iRefs };
}

export const ORDENS_TABELA = { planilha: "Ordem da planilha", abs: "Maior desvio (|Δ|)", fav: "Mais favorável → mais desfavorável", pct: "Maior desvio % (|Δ%|)" };

// Ordena as linhas dentro de cada bloco entre subtotais (a DRE continua fechando)
function ordenarLinhas(linhas, ordem) {
  if (!ordem || ordem === "planilha") return linhas;
  const chave = (l) => {
    const dv = l.ac - l.ref;
    if (ordem === "fav") return -(l.f.inv ? -dv : dv);
    if (ordem === "pct") return l.ref !== 0 ? -Math.abs(dv / l.ref) : 0;
    return -Math.abs(dv);
  };
  const out = [];
  let bloco = [];
  const fechar = () => {
    bloco.sort((a, b) => (a.outros ? 1 : 0) - (b.outros ? 1 : 0) || chave(a) - chave(b));
    out.push(...bloco); bloco = [];
  };
  linhas.forEach((l) => { if (l.f.res) { fechar(); out.push(l); } else bloco.push(l); });
  fechar();
  return out;
}

function tabela(tela, d, W, spec) {
  const e = tela.e;
  if (d.nCol < 2) throw new Error("A tabela precisa de duas colunas numéricas: Referência (ex.: BGT) e Atual (ex.: Real).");
  const { iAc, iRefs } = colunasTabela(spec, d);
  const nR = iRefs.length;
  const refNomes = iRefs.map((i) => d.cab[i] || "Ref"), acNome = d.cab[iAc] || "AC";
  const vis = { ref: true, ac: true, abs: true, pct: true, ...(spec.colunas ?? {}) };
  if (!vis.ac && !vis.abs && !vis.pct) throw new Error("Deixe visível ao menos uma coluna: Atual, Δ ou Δ%.");
  const iCom = (d.cabTextos ?? []).findIndex((c) => /coment|obs|nota|explic|justif/i.test(c));
  const colCom = iCom >= 0 ? iCom : (d.cabTextos?.length ? 0 : -1);
  let todas = d.rot.map((r, i) => {
    const refs = iRefs.map((k) => num(d.v[i][k])), ac = num(d.v[i][iAc]);
    return { r, refs, ref: refs[0], ac, f: flagsLinha(spec, r, refs[0], ac), com: colCom >= 0 ? (d.textos[i][colCom] || "") : "" };
  });

  // Top N: mantém as N linhas (não-subtotal) com maior |Δ| (1ª referência) e agrupa o resto em "Outros"
  const topN = +spec.topN || 0;
  if (topN > 0) {
    const cand = todas.filter((l) => !l.f.res && !l.f.ocultar);
    if (cand.length > topN) {
      const manter = new Set([...cand].sort((a, b) => Math.abs(b.ac - b.ref) - Math.abs(a.ac - a.ref)).slice(0, topN));
      const fora = cand.filter((l) => !manter.has(l));
      const ultimo = todas.indexOf(cand[cand.length - 1]);
      const novas = [];
      todas.forEach((l, i) => {
        if (!fora.includes(l)) novas.push(l);
        if (i === ultimo && spec.agruparOutros !== false && fora.length) {
          const refs = iRefs.map((_, k) => fora.reduce((a, x) => a + x.refs[k], 0));
          novas.push({ r: `Outros (${fora.length})`, refs, ref: refs[0], ac: fora.reduce((a, x) => a + x.ac, 0), outros: true,
            f: { inv: fora.every((x) => x.f.inv), res: false, neg: false, ocultar: false }, com: "" });
        }
      });
      todas = novas;
    }
  }
  todas = ordenarLinhas(todas.filter((l) => !l.f.ocultar), spec.ordem);

  // Cascata: subtotais começam do zero; demais linhas somam (custos invertidos subtraem)
  let acum = 0;
  todas.forEach((l) => {
    if (l.f.res) { l.c0 = 0; l.c1 = l.ac; acum = l.ac; }
    else { const c = l.f.inv ? -l.ac : l.ac; l.c0 = acum; l.c1 = acum + c; acum += c; }
  });
  const linhas = todas.map((l) => ({
    ...l,
    dvs: l.refs.map((x) => l.ac - x),
    pcts: l.refs.map((x) => (x !== 0 ? (l.ac - x) / Math.abs(x) : null)),
  }));
  linhas.forEach((l) => { l.dv = l.dvs[0]; l.pct = l.pcts[0]; });
  if (!linhas.length) throw new Error("Todas as linhas estão ocultas.");
  const cascata = !!spec.cascata && vis.ac;
  const un = unidadeTabela(spec, linhas.flatMap((l) => [...l.refs, l.ac, ...l.dvs]));
  const fmt = (v, sinal) => fmtTab(v, un, sinal);
  const fav = (l, k = 0) => (l.f.inv ? l.dvs[k] <= 0 : l.dvs[k] >= 0);

  // ----- comentários (caixa à direita, fora da largura da tabela)
  const comentarios = spec.comentarios === false ? [] : linhas.filter((l) => l.com);
  const comW = comentarios.length ? 250 : 0;
  comentarios.forEach((l, k) => (l.nCom = k + 1));

  // ----- layout: rótulo | refs (números) | atual | [Δ | Δ%] por referência
  const tam = e.fonte + 1, rowH = Math.round(tam * 2.4);
  const maxRot = Math.max(...linhas.map((l) => l.r.length));
  const lw = Math.min(220, Math.max(90, maxRot * tam * 0.56 + 14));
  const refW = 64, refsW = vis.ref ? nR * refW : 0, mCom = comW ? 18 : 0;
  const pAc = vis.ac ? (cascata ? 0.42 : 0.38) : 0, pD = vis.abs ? (cascata ? 0.30 : 0.32) : 0, pP = vis.pct ? (cascata ? 0.28 : 0.30) : 0;
  const pesos = pAc + nR * (pD + pP);
  // com várias referências a tabela se alarga para cada bloco continuar legível
  const unid = Math.max((W - lw - refsW - mCom) / pesos, nR > 1 ? 400 : 0);
  const acW = unid * pAc, dW = unid * pD, pW = unid * pP;
  const Wt = Math.round(lw + refsW + unid * pesos + mCom);
  const xRefs = iRefs.map((_, k) => lw + k * refW);
  const xAc = lw + refsW;
  const xDs = iRefs.map((_, k) => xAc + acW + k * (dW + pW)), xPs = xDs.map((x) => x + dW);

  let titH = 0;
  const larguraTotal = Wt + comW;
  tela.fundo(0, 0, larguraTotal, 10);
  const fundo = tela.p[tela.p.length - 1];
  if (spec.titulo) {
    titH = 30;
    tela.text(spec.titulo, 0, 0, "top", larguraTotal, { size: e.fonte + 3, bold: true, align: "left", color: "#000000" });
    tela.line(0, 21, larguraTotal, 21, "#000000", 0.75);
  }
  const cabH = 22, y0 = titH + cabH;
  let H = y0 + linhas.length * rowH + 6;
  const yLinha = (i) => y0 + i * rowH + rowH / 2;
  const barH = rowH * 0.62;
  const cinzaEixo = "#8C8C8C";

  // ----- escalas (Δ e Δ% na MESMA escala para todas as referências: dá para comparar)
  const pontosAc = cascata ? linhas.flatMap((l) => [l.c0, l.c1]) : linhas.map((l) => l.ac);
  const dNat = Math.max(...linhas.flatMap((l) => l.dvs.map(Math.abs))) || 1;
  const esc = escalar(tela, { min: Math.min(0, ...pontosAc), max: Math.max(0, ...pontosAc), d: dNat });
  const acMin = esc.min, acMax = esc.max;
  const acI1 = xAc + (acMin < 0 ? 46 : 6), acI2 = xAc + acW - 46;
  const xA = (v) => acI1 + ((v - acMin) / ((acMax - acMin) || 1)) * (acI2 - acI1);
  const dMax = esc.d;
  const dCs = xDs.map((x) => x + dW / 2), dR = dW / 2 - 44;
  const pcts = linhas.flatMap((l) => l.pcts).filter((p) => p !== null).map(Math.abs);
  let pMax = Math.max(0, ...pcts) || 1;
  if (spec.suprimirPct !== false && pcts.length > 2) {
    const ord = [...pcts].sort((a, b) => a - b), med = ord[Math.floor(ord.length / 2)] || ord[ord.length - 1];
    const normais = ord.filter((x) => x <= Math.max(med * 4, 0.05));
    if (normais.length) pMax = Math.max(...normais) || pMax;
  }
  const pCs = xPs.map((x) => x + pW / 2), pR = pW / 2 - 40;

  // ----- cabeçalhos (marca IBCS sob a referência: real = cheio, orçamento = contorno, forecast = hachurado)
  const cab = { size: e.fonte, color: "#595959" };
  iRefs.forEach((_, k) => {
    const nome = refNomes[k];
    if (vis.ref) {
      tela.text(nome, xRefs[k] + refW - 6, titH + 4, "top", refW, { ...cab, align: "right" });
      const c = cenarioDe(nome), x1 = xRefs[k] + 14, x2 = xRefs[k] + refW - 4;
      if (c === "real") tela.line(x1, titH + 18, x2, titH + 18, "#A6A6A6", 2);
      else tela.rect(x1, titH + 16, x2 - x1, titH + 20, "#FFFFFF", { stroke: "#595959", ...(c === "fc" ? { listra: "#8C8C8C" } : {}) });
    }
    if (vis.abs) tela.text("Δ" + nome, dCs[k], titH + 4, "top", dW, cab);
    if (vis.pct) tela.text("Δ" + nome + "%", pCs[k], titH + 4, "top", pW, { ...cab, italic: true });
  });
  if (vis.ac) tela.text(acNome, xA(0) + 4, titH + 4, "top", 80, { ...cab, align: "left" });

  // ----- linhas
  linhas.forEach((l, i) => {
    const y = yLinha(i), bold = l.f.neg || l.f.res;
    if (l.f.res && i > 0) tela.line(0, y - rowH / 2, Wt - mCom, y - rowH / 2, "#BFBFBF", 0.75);
    const pref = l.f.inv && !/^\s*[(\-–]/.test(l.r) ? "– " : l.f.res ? "= " : "";
    tela.text(pref + l.r, 0, y, "middle", lw - 8, { size: tam, bold, align: "left" });
    if (vis.ref) l.refs.forEach((x, k) =>
      tela.text(fmt(x), xRefs[k] + refW - 6, y, "middle", refW, { size: tam, bold, align: "right", color: "#404040" }));
    // atual: barra simples ou degrau da cascata
    if (vis.ac) {
      const v0 = cascata ? l.c0 : 0, v1 = cascata ? l.c1 : l.ac;
      const a0 = xA(v0), a1 = xA(v1);
      tela.rect(Math.min(a0, a1), y - barH / 2, Math.max(0.75, Math.abs(a1 - a0)), y + barH / 2, e.total);
      if (cascata && i < linhas.length - 1) {
        const prox = linhas[i + 1];
        if (!prox.f.res) tela.line(xA(prox.c0), y + barH / 2, xA(prox.c0), yLinha(i + 1) - barH / 2, "#7F7F7F", 0.75);
      }
      // rótulo do lado da ponta da barra; se não couber à esquerda, vai para a direita
      const larguraRot = fmt(l.ac).length * tam * 0.6 + 6;
      const cabeEsq = Math.min(a0, a1) - larguraRot >= xA(0) + 2;
      if (v1 >= v0 || !cabeEsq) tela.text(fmt(l.ac), Math.max(a0, a1) + 4, y, "middle", 60, { size: tam, bold, align: "left" });
      else tela.text(fmt(l.ac), Math.min(a0, a1) - 4, y, "middle", 60, { size: tam, bold, align: "right" });
    }
    l.dvs.forEach((dv, k) => {
      const cor = fav(l, k) ? e.pos : e.neg;
      // Δ absoluta
      if (vis.abs) {
        const dC = dCs[k], dx = (dv / dMax) * dR;
        if (dv !== 0) tela.rect(Math.min(dC, dC + dx), y - barH / 2, Math.max(0.75, Math.abs(dx)), y + barH / 2, cor);
        if (dv >= 0) tela.text(fmt(dv, true), dC + Math.max(dx, 0) + 4, y, "middle", 60, { size: tam, bold, align: "left" });
        else tela.text(fmt(dv, true), dC + dx - 4, y, "middle", 60, { size: tam, bold, align: "right" });
      }
      // Δ %: alfinete (seta quando passa da escala)
      if (vis.pct) {
        const pC = pCs[k], pct = l.pcts[k];
        if (pct === null) tela.text("–", pC + 6, y, "middle", 30, { size: tam, align: "left", color: "#808080" });
        else {
          const fora = Math.abs(pct) > pMax * 1.0001;
          const px = Math.sign(pct) * Math.min(Math.abs(pct) / pMax, 1) * pR;
          if (fora) tela.arrow(pC, y, pC + px + Math.sign(px) * 6, y, cor, 2);
          else {
            if (Math.abs(px) > 0.5) tela.line(pC, y, pC + px, y, cor, 2);
            tela.dot(pC + px, y, 3, "#262626");
          }
          const xt = pC + px + Math.sign(px || 1) * (fora ? 10 : 6);
          tela.text(fmtPct(pct, Math.abs(pct) >= 1 ? 0 : 1), xt, y, "middle", 60, { size: tam, bold, italic: true, align: px < 0 ? "right" : "left" });
        }
      }
    });
    // marcador do comentário
    if (l.nCom) tela.marcador(Wt - mCom + 6, y, 7, l.nCom, "#1A53FF");
  });

  // ----- eixos verticais e separadores entre as referências
  const yA = y0 - 2, yB = y0 + linhas.length * rowH + 2;
  if (vis.ac) tela.line(xA(0), yA, xA(0), yB, cinzaEixo, 1.5);
  iRefs.forEach((_, k) => {
    if (vis.abs) tela.line(dCs[k], yA, dCs[k], yB, cinzaEixo, 1.5);
    if (vis.pct) tela.line(pCs[k], yA, pCs[k], yB, cinzaEixo, 1.5);
    if (nR > 1 && (k > 0 || vis.ac)) tela.line(xDs[k] - 2, titH + 2, xDs[k] - 2, yB, "#E0E0E0", 0.75);
  });

  // ----- caixa de comentários (variação contra a 1ª referência)
  if (comW) {
    const xC = Wt + 6, wC = comW - 6, tc = e.fonte;
    let cursor = y0 - 4;
    tela.line(Wt - 2, y0 - 6, Wt - 2, yB, "#E0E0E0", 0.75);
    comentarios.forEach((l) => {
      const y = Math.max(cursor, yLinha(linhas.indexOf(l)) - rowH / 2);
      tela.marcador(xC + 8, y + 8, 7, l.nCom, "#1A53FF");
      const seta = l.dv >= 0 ? "▲" : "▼"; // direção = sinal; cor = favorável/desfavorável
      const modo = spec.comModo ?? "B";
      const pctTxt = l.pct === null ? "" : fmtPct(l.pct, Math.abs(l.pct) >= 1 ? 0 : 1);
      const absTxt = fmt(l.dv, true);
      let varTxt = modo === "P" ? (pctTxt || absTxt) : modo === "A" ? absTxt : [pctTxt, absTxt].filter(Boolean).join(" | ");
      if (nR > 1) varTxt += " vs " + refNomes[0];
      const cab1 = `${l.r} ${fmt(l.ac)}`;
      tela.text(cab1, xC + 20, y + 8, "middle", wC - 20, { size: tc, bold: true, align: "left" });
      const xSeta = xC + 20 + cab1.length * tc * 0.6 + 4;
      tela.text(seta, xSeta, y + 8, "middle", 12, { size: tc, align: "left", color: fav(l) ? e.pos : e.neg });
      tela.text(varTxt, xSeta + 12, y + 8, "middle", wC, { size: tc, align: "left" });
      const nLin = Math.min(4, Math.max(1, Math.ceil((l.com.length * tc * 0.55) / (wC - 20))));
      tela.text(l.com, xC + 20, y + 17, "top", wC - 20, { size: tc, align: "left", wrap: true, maxLinhas: nLin, color: "#404040" });
      cursor = y + 17 + tc * 1.35 * nLin + 8;
    });
    H = Math.max(H, cursor + 4);
  }
  fundo.h = H + 16;
  return { altura: H, largura: larguraTotal };
}

// ---------------------------- Roscas proporcionais ---------------------------
// Uma rosca por coluna (ex.: 2023 | 2024 | BGT), fatias = linhas (ex.: segmentos).
// A ÁREA de cada rosca é proporcional ao total da coluna; a fatia destacada fica
// na cor do estilo e as demais em cinza; legenda única à esquerda.

// mistura uma cor com branco (t = 0 -> cor, t = 1 -> branco)
export function clarear(hex, t) {
  const n = parseInt(hex.slice(1), 16);
  const c = [(n >> 16) & 255, (n >> 8) & 255, n & 255].map((x) => Math.round(x + (255 - x) * t));
  return "#" + c.map((x) => x.toString(16).padStart(2, "0")).join("").toUpperCase();
}
const indiceDestaque = (d, destaque) => (destaque === undefined || destaque === null || destaque === "" ? -1 : d.rot.indexOf(destaque));

function roscas(tela, d, W, H, spec) {
  const e = tela.e, n = d.nCol, m = d.nLin;
  const iD = indiceDestaque(d, spec.destaque);
  const totais = d.cab.map((_, j) => d.rot.reduce((a, _, i) => a + Math.max(0, num(d.v[i][j])), 0));
  const tMax = escalar(tela, { min: 0, max: Math.max(...totais) }).max;
  if (!(tMax > 0)) throw new Error("As roscas precisam de valores positivos.");
  tela.fundo(0, 0, W, H);
  const fundo = tela.p[tela.p.length - 1];
  let titH = 0;
  if (spec.titulo) { titH = 24; tela.text(spec.titulo, 0, 0, "top", W, { size: e.fonte + 3, bold: true, align: "left", color: "#000000" }); }
  // cores: destaque x cinzas, ou a paleta do estilo
  const cinzas = ["#BFBFBF", "#D9D9D9"];
  const paleta = e.pilha?.length ? e.pilha : [e.total];
  const cor = (i) => (iD >= 0 ? (i === iD ? e.total : cinzas[(i < iD ? i : i - 1) % 2]) : paleta[i % paleta.length]);
  // legenda única
  const itens = iD >= 0 ? [[d.rot[iD], e.total], ["Demais", cinzas[0]]] : d.rot.map((r, i) => [r, cor(i)]);
  const legW = Math.min(150, Math.max(70, Math.max(...itens.map(([r]) => r.length)) * e.fonte * 0.56 + 22));
  const yLeg0 = titH + 14;
  itens.forEach(([r, c], k) => {
    const y = yLeg0 + k * (e.fonte * 1.9);
    tela.rect(0, y, 10, y + 10, c);
    tela.text(r, 15, y + 5, "middle", legW - 15, { size: e.fonte, align: "left" });
  });
  // roscas
  const colW = (W - legW) / n;
  const rotH = e.fonte * 1.4 * (iD >= 0 ? 3 : 2) + 8;
  const rMax = Math.max(12, Math.min(colW / 2 - (iD >= 0 && spec.separar !== false ? 39 : 36), (H - titH - 16 - rotH) / 2 - 18));
  const cy = titH + 16 + 18 + rMax;
  let partAnt = null;
  // total no centro de TODAS as roscas, ou embaixo de todas (nunca misturado)
  const totNoCentro = totais.every((T) => rMax * Math.sqrt(T / tMax) * 0.58 > fmtNum(T).length * (e.fonte + 1) * 0.3);
  d.cab.forEach((nome, j) => {
    const cx = legW + colW * (j + 0.5), T = totais[j];
    const r1 = rMax * Math.sqrt(T / tMax), r0 = r1 * 0.58;
    if (T > 0) {
      const folga = Math.min(0.03, 1.2 / ((r0 + r1) / 2)); // fresta branca entre as fatias
      let a = 0;
      const postos = []; // rótulos já colocados nesta rosca (evita sobreposição)
      d.rot.forEach((_, i) => {
        const v = Math.max(0, num(d.v[i][j]));
        if (!v) return;
        const ang = (v / T) * 2 * Math.PI;
        const fatias = d.rot.filter((_, k) => num(d.v[k][j]) > 0).length;
        const pct = v / T, meio = a + ang / 2;
        // fatia destacada "explodida": deslocada para fora, isolada do resto da rosca
        const desl = i === iD && spec.separar !== false && fatias > 1 ? Math.max(5, r1 * 0.16) : 0;
        const ox = desl * Math.sin(meio), oy = -desl * Math.cos(meio);
        tela.arco(cx + ox, cy + oy, r0, r1, a + (fatias > 1 ? folga / 2 : 0), a + ang - (fatias > 1 ? folga / 2 : 0), cor(i));
        // rótulo de %: só a fatia destacada (ou, sem destaque, as fatias de 6% ou mais)
        if ((iD >= 0 && i === iD) || (iD < 0 && pct >= 0.06)) {
          const rr = r1 + 11 + desl, tx = cx + rr * Math.sin(meio);
          let ty = cy - rr * Math.cos(meio);
          const dir = Math.sin(meio) >= 0 ? "left" : "right";
          const passo = Math.cos(meio) > 0 ? -1 : 1; // afasta para fora da rosca
          for (let k = 0; k < 6 && postos.some(([x, y]) => Math.abs(x - tx) < 26 && Math.abs(y - ty) < e.fonte * 1.25); k++) ty += passo * e.fonte * 0.7;
          postos.push([tx, ty]);
          tela.text(fmtPct(pct, 0).replace("+", ""), tx + (dir === "left" ? -2 : 2), ty, "middle", 40,
            { size: e.fonte, bold: i === iD, align: Math.abs(Math.sin(meio)) < 0.25 ? "center" : dir });
        }
        a += ang;
      });
    }
    // total no centro (ou embaixo, se a rosca for pequena demais)
    const totTxt = fmtNum(T);
    if (totNoCentro) tela.text(totTxt, cx, cy, "middle", r0 * 2, { size: e.fonte + 1, bold: true });
    let yRot = cy + rMax + 18;
    tela.text(nome, cx, yRot, "top", colW - 6, { size: e.fonte + 1, bold: true });
    yRot += e.fonte * 1.5;
    if (!totNoCentro) { tela.text(totTxt, cx, yRot, "top", colW - 6, { size: e.fonte }); yRot += e.fonte * 1.4; }
    // participação do destaque e variação em p.p. contra a rosca anterior
    if (iD >= 0 && T > 0) {
      const part = Math.max(0, num(d.v[iD][j])) / T;
      if (partAnt !== null) {
        const pp = (part - partAnt) * 100;
        const txt = (pp > 0.05 ? "+" : pp < -0.05 ? "-" : "") + new Intl.NumberFormat("pt-BR", { minimumFractionDigits: 1, maximumFractionDigits: 1 }).format(Math.abs(pp)) + " p.p.";
        tela.text(txt, cx, yRot, "top", colW - 6, { size: e.fonte, italic: true, color: pp > 0.05 ? e.pos : pp < -0.05 ? e.neg : e.texto });
      }
      partAnt = part;
    }
  });
  // altura usada (a rosca é limitada pela largura; não sobra espaço vazio embaixo)
  const usada = Math.min(H, Math.max(cy + rMax + 18 + rotH, yLeg0 + itens.length * e.fonte * 1.9));
  fundo.h = usada + 16;
  return usada;
}

// ---------------------------- Matriz de bolhas --------------------------------
// Linhas x colunas; área da bolha = volume; cor = faixa. Se os cabeçalhos se
// repetirem (Jan Fev Mar | Jan Fev Mar), o 1º bloco dá o tamanho e o 2º a cor
// (ex.: receita e margem %). Uma linha pode ser destacada.
export function blocosBolhas(d) {
  const n = d.nCol;
  if (n % 2 === 0 && n >= 2) {
    const h = n / 2;
    if (d.cab.slice(0, h).every((c, j) => String(c).trim().toUpperCase() === String(d.cab[j + h]).trim().toUpperCase()))
      return { cols: [...Array(h).keys()], corDe: (j) => j + h, separada: true };
  }
  return { cols: [...Array(n).keys()], corDe: (j) => j, separada: false };
}

function bolhas(tela, d, W, spec) {
  const e = tela.e, m = d.nLin;
  const { cols, corDe, separada } = blocosBolhas(d);
  const nc = cols.length, iD = indiceDestaque(d, spec.destaque);
  const vol = (i, j) => Math.abs(num(d.v[i][j]));
  const vc = (i, j) => d.v[i][corDe(j)];
  const vMax = escalar(tela, { min: 0, max: Math.max(...d.rot.flatMap((_, i) => cols.map((j) => vol(i, j)))) || 1 }).max;
  const cores = d.rot.flatMap((_, i) => cols.map((j) => vc(i, j))).filter((x) => x !== null);
  const cMin = Math.min(...cores), cMax = Math.max(...cores);
  const ehPct = cores.every((x) => Math.abs(x) <= 1.5) && cores.some((x) => !Number.isInteger(x));
  const fmtC = (x) => (ehPct ? fmtPct(x, Math.abs(x) >= 0.1 ? 0 : 1).replace("+", "") : fmtNum(x));
  // faixas: com negativos -> divergente (vermelho/verde); senão 4 faixas iguais na cor do estilo
  let faixas;
  if (cMin < 0 && cMax > 0) {
    const a = Math.max(-cMin, cMax);
    faixas = [
      { ate: -a / 2, cor: e.neg }, { ate: 0, cor: clarear(e.neg, 0.55) },
      { ate: a / 2, cor: clarear(e.pos, 0.5) }, { ate: Infinity, cor: e.pos },
    ];
    faixas.forEach((f, k) => (f.rot = k === 0 ? `< ${fmtC(-a / 2)}` : k === 1 ? `${fmtC(-a / 2)} a 0` : k === 2 ? `0 a ${fmtC(a / 2)}` : `> ${fmtC(a / 2)}`));
  } else {
    const passo = (cMax - cMin) / 4 || 1;
    faixas = [0.75, 0.5, 0.25, 0].map((t, k) => ({
      ate: k === 3 ? Infinity : cMin + passo * (k + 1), cor: clarear(e.total, t),
      rot: k === 3 ? `> ${fmtC(cMin + passo * 3)}` : `${fmtC(cMin + passo * k)} a ${fmtC(cMin + passo * (k + 1))}`,
    }));
  }
  const corDaFaixa = (x) => (x === null ? "#D9D9D9" : faixas.find((f) => x <= f.ate).cor);

  // layout
  const tam = e.fonte;
  const lw = Math.min(180, Math.max(80, Math.max(...d.rot.map((r) => r.length)) * (tam + 1) * 0.56 + 12));
  const cw = (W - lw) / nc;
  const rowH = Math.max(26, Math.min(48, cw * 0.55));
  const rMax = Math.min(rowH / 2 - 2, cw / 2 - 18);
  let titH = 0;
  tela.fundo(0, 0, W, 10);
  const fundo = tela.p[tela.p.length - 1];
  if (spec.titulo) { titH = 24; tela.text(spec.titulo, 0, 0, "top", W, { size: tam + 3, bold: true, align: "left", color: "#000000" }); }
  // legenda das faixas
  const yLeg = titH + 4;
  let xl = 0;
  tela.text(separada ? "Cor:" : "Faixas:", xl, yLeg + 6, "middle", 40, { size: tam, align: "left", color: "#595959" });
  xl += separada ? 26 : 38;
  faixas.forEach((f) => {
    tela.dot(xl + 5, yLeg + 6, 5, f.cor);
    tela.text(f.rot, xl + 13, yLeg + 6, "middle", 90, { size: tam, align: "left" });
    xl += 13 + f.rot.length * tam * 0.56 + 14;
  });
  tela.text("Tamanho = |valor|", W, yLeg + 6, "middle", 120, { size: tam, align: "right", color: "#595959", italic: true });
  const yCab = yLeg + 22, y0 = yCab + 18;
  cols.forEach((j, k) => tela.text(d.cab[j], lw + cw * (k + 0.5), yCab, "top", cw - 4, { size: tam, bold: true, color: "#595959" }));
  tela.line(0, y0 - 2, W, y0 - 2, "#BFBFBF", 0.75);
  d.rot.forEach((r, i) => {
    const y = y0 + i * rowH + rowH / 2, dest = i === iD;
    if (dest) tela.rect(0, y - rowH / 2 + 1, W, y + rowH / 2 - 1, clarear(e.total, 0.88));
    else if (i > 0) tela.line(0, y - rowH / 2, W, y - rowH / 2, "#EDEDED", 0.75);
    tela.text(r, 0, y, "middle", lw - 8, { size: tam + 1, bold: dest, align: "left" });
    cols.forEach((j, k) => {
      const v = d.v[i][j];
      if (v === null) return;
      const cx = lw + cw * k + rMax + 6;
      const rr = v === 0 ? 0 : Math.max(1.5, rMax * Math.sqrt(vol(i, j) / vMax));
      if (rr) tela.dot(cx, y, rr, corDaFaixa(vc(i, j)));
      if (spec.valores !== false) {
        const txt = separada ? `${fmtNum(v)} · ${vc(i, j) === null ? "–" : fmtC(vc(i, j))}` : fmtNum(v);
        tela.text(txt, cx + Math.max(rr, 3) + 3, y, "middle", cw - (rr + 8), { size: tam, bold: dest, align: "left", color: "#404040" });
      }
    });
  });
  const H = y0 + m * rowH + 4;
  fundo.h = H + 16;
  return { altura: H };
}


// =============================================================================
// Ponte Preço / Volume / Mix
//   Layout: Produto | Volume REF | Preço REF | Volume ATUAL | Preço ATUAL
//   (Receita no lugar de Preço também serve: o preço vira Receita ÷ Volume)
//   Volume = (V1 − V0) × P0 médio          (quanto a mais se vendeu, ao preço médio de referência)
//   Mix    = Σ v1·p0 − V1 × P0 médio       (vendeu-se mais dos produtos caros ou dos baratos)
//   Preço  = Σ (p1 − p0) · v1              (preço praticado contra o de referência)
//   Volume + Mix + Preço = Receita ATUAL − Receita REF (fecha exato)
// Produto novo (sem volume na referência) entra sem efeito preço: p0 = p1.
// =============================================================================
const RX_PRECO = /PRECO|PRICE|\bPM\b|TICKET|UNITARIO/;
const RX_RECEITA = /RECEITA|FATURAMENTO|VENDAS?\b|REVENUE|SALES|\bROL\b|\bROB\b/;
const RX_PALAVRAS_PVM = /\b(VOLUMES?|VOL\.?|QTDE?\.?|QUANTIDADES?|TONELADAS?|TON|UNIDADES?|PRECOS?( MEDIOS?)?|PRICE|PM|TICKET( MEDIO)?|RECEITAS?|FATURAMENTO|VENDAS?|REVENUE|SALES|ROL|ROB|R\$\s*\/\s*\S+|UNITARIO)\b/g;

export function pvmCalcular(d, spec = {}) {
  if (d.nCol < 4) throw new Error("Layout da ponte PVM: Produto | Volume BGT | Preço BGT | Volume Real | Preço Real (ou Receita no lugar de Preço).");
  const tipoCol = (c) => { const n = normalizar(c); return RX_PRECO.test(n) ? "p" : RX_RECEITA.test(n) ? "r" : "v"; };
  const par = (a, b) => {
    let ka = tipoCol(d.cab[a]), kb = tipoCol(d.cab[b]);
    if (ka === "v" && kb === "v") kb = "p"; // sem pista no nome: Volume, Preço
    if (ka !== "v" && kb !== "v") ka = "v";
    return ka === "v" ? { iv: a, ix: b, k: kb } : { iv: b, ix: a, k: ka };
  };
  const p0c = par(0, 1), p1c = par(2, 3);
  const nomeCen = (i, padrao) => {
    const t = normalizar(d.cab[i]).replace(RX_PALAVRAS_PVM, " ").replace(/[^A-Z0-9+]+/g, " ").trim();
    if (!t) return padrao;
    // devolve o trecho original correspondente (mantém acentos/maiúsculas do usuário)
    const orig = String(d.cab[i]).split(/\s+/).filter((w) => t.split(" ").includes(normalizar(w).replace(/[^A-Z0-9+]/g, "")));
    return orig.join(" ") || padrao;
  };
  const refNome = nomeCen(p0c.iv, "Ref") , acNome = nomeCen(p1c.iv, "Real");
  const precoDe = (i, c) => {
    const v = num(d.v[i][c.iv]), x = num(d.v[i][c.ix]);
    return c.k === "r" ? (v ? x / v : null) : x;
  };
  const prods = [];
  d.rot.forEach((r, i) => {
    if (RX_RESULTADO.test(normalizar(r))) return; // linha de total não entra (contaria em dobro)
    const v0 = num(d.v[i][p0c.iv]), v1 = num(d.v[i][p1c.iv]);
    let p0 = precoDe(i, p0c), p1 = precoDe(i, p1c);
    if (!v0 || p0 === null) p0 = p1 ?? 0;
    if (!v1 || p1 === null) p1 = p0;
    const r0 = receitaDe(i, p0c, v0, p0), r1 = receitaDe(i, p1c, v1, p1);
    prods.push({ r, v0, v1, p0, p1, r0, r1 });
  });
  function receitaDe(i, c, v, p) { return c.k === "r" ? num(d.v[i][c.ix]) : v * p; }
  if (!prods.length) throw new Error("Não há produtos na ponte PVM.");
  const R0 = prods.reduce((a, x) => a + x.r0, 0), R1 = prods.reduce((a, x) => a + x.r1, 0);
  const V0 = prods.reduce((a, x) => a + x.v0, 0), V1 = prods.reduce((a, x) => a + x.v1, 0);
  const P0 = V0 ? R0 / V0 : 0;
  prods.forEach((x) => { x.preco = (x.p1 - x.p0) * x.v1; x.volMix = (x.v1 - x.v0) * x.p0; });
  const preco = prods.reduce((a, x) => a + x.preco, 0);
  let volume, mix;
  if (spec.pvmMix === false) { volume = prods.reduce((a, x) => a + x.volMix, 0); mix = 0; }
  else { volume = (V1 - V0) * P0; mix = prods.reduce((a, x) => a + x.v1 * x.p0, 0) - V1 * P0; }
  // resíduo de arredondamento / receitas informadas sem volume vai para o preço (fecha a ponte)
  const residuo = R1 - R0 - (volume + mix + preco);
  return { refNome, acNome, R0, R1, V0, V1, volume, mix: mix, preco: preco + residuo, prods, comMix: spec.pvmMix !== false };
}

// Dados sintéticos para a cascata IBCS: REF → Volume → Mix → Preço (ou preço por produto) → ATUAL
function dadosPVM(c, spec) {
  const rot = [c.refNome, "Volume"], v = [[c.R0], [c.volume]];
  if (c.comMix) { rot.push("Mix"); v.push([c.mix]); }
  const n = Math.max(0, +spec.pvmDetalhe || 0);
  if (n > 0) {
    const ord = [...c.prods].sort((a, b) => Math.abs(b.preco) - Math.abs(a.preco));
    const top = ord.slice(0, n), resto = ord.slice(n);
    top.forEach((x) => { rot.push(`Preço ${x.r}`); v.push([x.preco]); });
    if (resto.length) { rot.push(`Preço – demais (${resto.length})`); v.push([c.preco - top.reduce((a, x) => a + x.preco, 0)]); }
  } else { rot.push("Preço"); v.push([c.preco]); }
  rot.push(c.acNome); v.push([c.R1]);
  return { rot, cab: ["Receita"], v, tipo: [], nLin: rot.length, nCol: 1, cabTextos: [], textos: rot.map(() => []) };
}

// =============================================================================
// Título-mensagem: a conclusão do gráfico em uma frase (padrão de consultoria).
// É uma SUGESTÃO calculada dos dados; o usuário edita no painel.
// =============================================================================
const pct1 = (p) => fmtPct(p, Math.abs(p) >= 1 ? 0 : 1).replace(/^-/, "−");
const pctSemSinal = (p) => pct1(p).replace(/^[+\-−]/, "");
const comSinal = (v) => fmtNum(v, true).replace(/^-/, "−");
const verbo = (xs, um, varios) => (xs.length > 1 ? varios : um);
const lista = (xs) => (xs.length <= 1 ? xs.join("") : xs.slice(0, -1).join(", ") + " e " + xs[xs.length - 1]);

// frase "X% acima/abaixo de REF: drivers" a partir de impactos (positivo = favorável)
function fraseDesvio(nome, ref, ac, refNome, drivers, inverso = false) {
  const dv = ac - ref;
  const fav = inverso ? dv <= 0 : dv >= 0;
  const p = ref !== 0 ? dv / Math.abs(ref) : null;
  const pos = drivers.filter((x) => x.imp > 0).sort((a, b) => b.imp - a.imp);
  const neg = drivers.filter((x) => x.imp < 0).sort((a, b) => a.imp - b.imp);
  const fmtD = (x) => `${x.r} (${comSinal(x.imp)})`;
  if (p !== null && Math.abs(p) < 0.0005) return `${nome} em linha com ${refNome}` + (pos[0] && neg[0] ? `: ${fmtD(pos[0])} compensou ${fmtD(neg[0])}` : "");
  const quanto = p === null ? comSinal(dv) : pctSemSinal(p);
  const lado = dv >= 0 ? "acima" : "abaixo";
  let t = `${nome} ${quanto} ${lado} ${/^(o|a)\s/i.test(refNome) ? "d" + refNome : "do " + refNome}`;
  const P = pos.slice(0, 2), N = neg.slice(0, 2);
  if (fav) {
    if (P.length) t += `: ${lista(P.map(fmtD))} ` + (neg.length ? `${verbo(P, "compensou", "compensaram")} ${fmtD(neg[0])}` : verbo(P, "explica o resultado", "explicam o resultado"));
  } else if (N.length) {
    t += `: ${lista(N.map(fmtD))} ` + (pos.length ? `${verbo(N, "anulou", "anularam")} ${fmtD(pos[0])}` : verbo(N, "explica o desvio", "explicam o desvio"));
  }
  return t;
}

export function sugerirTitulo(d, spec) {
  if (!d || !d.nLin) return "";
  const nome0 = (d.cab[0] && !/^S[ée]rie \d+$/.test(d.cab[0]) ? d.cab[0] : "Total");
  switch (spec.tipo) {
    case "TAB": {
      const { iAc, iRefs } = colunasTabela(spec, d);
      const ls = d.rot.map((r, i) => {
        const ref = num(d.v[i][iRefs[0]]), ac = num(d.v[i][iAc]);
        return { r, ref, ac, f: flagsLinha(spec, r, ref, ac) };
      }).filter((l) => !l.f.ocultar);
      const res = ls.filter((l) => l.f.res);
      const principal = res.length ? res[res.length - 1]
        : { r: "Total", ref: ls.reduce((a, l) => a + (l.f.inv ? -l.ref : l.ref), 0), ac: ls.reduce((a, l) => a + (l.f.inv ? -l.ac : l.ac), 0), f: {} };
      const drivers = ls.filter((l) => !l.f.res).map((l) => ({ r: l.r, imp: l.f.inv ? -(l.ac - l.ref) : l.ac - l.ref }));
      return fraseDesvio(principal.r, principal.ref, principal.ac, d.cab[iRefs[0]] || "referência", drivers);
    }
    case "PVM": {
      const c = pvmCalcular(d, spec);
      const drivers = [{ r: "volume", imp: c.volume }, { r: "preço", imp: c.preco }];
      if (c.comMix) drivers.push({ r: "mix", imp: c.mix });
      return fraseDesvio("Receita", c.R0, c.R1, c.refNome, drivers.filter((x) => Math.abs(x.imp) > 1e-9));
    }
    case "IBCS": {
      const { ini, fim } = calcCascata(d, 0);
      const drivers = d.rot.slice(1, -1).map((r, i) => ({ r, imp: num(d.v[i + 1][0]) }));
      return fraseDesvio(nome0, ini, fim, d.rot[0], drivers, spec.custo);
    }
    case "IBCSM": {
      const vs = d.cab.map((c, j) => { const { ini, fim } = calcCascata(d, j); return { c, p: ini ? (fim - ini) / Math.abs(ini) : 0 }; })
        .sort((a, b) => b.p - a.p);
      const melhor = vs[0], pior = vs[vs.length - 1];
      return vs.length < 2 ? `${melhor.c} ${pct1(melhor.p)} vs ${d.rot[0]}`
        : `${melhor.c} lidera (${pct1(melhor.p)} vs ${d.rot[0]}); ${pior.c} fica para trás (${pct1(pior.p)})`;
    }
    case "EMP": {
      const tot = d.rot.map((_, i) => d.v[i].reduce((a, x) => a + num(x), 0));
      const iT = d.tipo.map((t, i) => (t === "T" ? i : -1)).filter((i) => i >= 0);
      const a = iT.length ? iT[0] : 0, b = iT.length > 1 ? iT[iT.length - 1] : d.nLin - 1;
      const drivers = d.rot.map((r, i) => ({ r, imp: tot[i], t: d.tipo[i] })).filter((x, i) => i > a && i < b && x.t !== "T")
        .sort((x, y) => Math.abs(y.imp) - Math.abs(x.imp));
      const p = tot[a] > 0 && tot[b] >= 0 ? (tot[b] - tot[a]) / tot[a] : null; // % só faz sentido com base positiva
      const D = drivers.slice(0, 2);
      return `De ${d.rot[a]} (${fmtNum(tot[a]).replace(/^-/, "−")}) a ${d.rot[b]} (${fmtNum(tot[b]).replace(/^-/, "−")})` + (p !== null ? `, ${pct1(p)}` : "") +
        (D.length ? `: ${lista(D.map((x) => `${x.r} (${comSinal(x.imp)})`))} ${verbo(D, "explica", "explicam")} a maior parte` : "");
    }
    case "COL": {
      const tot = d.rot.map((_, i) => d.v[i].reduce((a, x) => a + num(x), 0));
      const n = tot.length, a = tot[0], b = tot[n - 1];
      const nome = d.nCol > 1 ? "Total" : nome0;
      if (n < 2 || a <= 0 || b <= 0) return `${nome}: ${fmtNum(b)} em ${d.rot[n - 1]}`;
      // CAGR só quando os rótulos são anos; senão (LY, BGT, 6+6…) variação simples
      if (!d.rot.every((r) => /^(19|20)\d{2}\b/.test(String(r).trim())))
        return `${nome} ${pct1((b - a) / a)} de ${d.rot[0]} para ${d.rot[n - 1]} (${fmtNum(a)} → ${fmtNum(b)})`;
      const g = Math.pow(b / a, 1 / (n - 1)) - 1;
      const verbo = g >= 0 ? "cresce" : "cai";
      let t = `${nome} ${verbo} ${pctSemSinal(g)} a.a. entre ${d.rot[0]} e ${d.rot[n - 1]} (${fmtNum(a)} → ${fmtNum(b)})`;
      if (d.nCol > 1) {
        const cres = d.cab.map((c, j) => ({ c, dv: num(d.v[n - 1][j]) - num(d.v[0][j]) })).sort((x, y) => y.dv - x.dv);
        t += `, com ${cres[0].c} respondendo por ${pctSemSinal(cres[0].dv / (b - a || 1))} do aumento`;
      }
      return t;
    }
    case "COMP": {
      const n = d.nLin, a = num(d.v[0][0]), b = num(d.v[n - 1][0]);
      const p = a ? (b - a) / Math.abs(a) : null;
      return `${nome0} ${p === null ? comSinal(b - a) : pct1(p)} de ${d.rot[0]} para ${d.rot[n - 1]} (${fmtNum(a)} → ${fmtNum(b)})`;
    }
    case "ROS": {
      const tot = d.cab.map((_, j) => d.rot.reduce((a, _, i) => a + Math.max(0, num(d.v[i][j])), 0));
      const n = d.nCol;
      const part = (i, j) => (tot[j] ? Math.max(0, num(d.v[i][j])) / tot[j] : 0);
      let i = d.rot.indexOf(spec.destaque);
      if (i < 0) i = d.rot.map((_, k) => k).sort((x, y) => Math.abs(part(y, n - 1) - part(y, 0)) - Math.abs(part(x, n - 1) - part(x, 0)))[0];
      const p0 = part(i, 0), p1 = part(i, n - 1);
      const verbo = p1 > p0 + 0.0005 ? "ganha participação" : p1 < p0 - 0.0005 ? "perde participação" : "mantém a participação";
      return n < 2 ? `${d.rot[i]} responde por ${pctSemSinal(p1)} do total`
        : `${d.rot[i]} ${verbo}: de ${pctSemSinal(p0)} para ${pctSemSinal(p1)} do total (${d.cab[0]} → ${d.cab[n - 1]})`;
    }
    case "BOL": {
      const { cols } = blocosBolhas(d);
      const somas = d.rot.map((r, i) => ({ r, s: cols.reduce((a, j) => a + num(d.v[i][j]), 0) }));
      const temNeg = d.rot.some((_, i) => cols.some((j) => num(d.v[i][j]) < 0));
      if (temNeg) {
        const ord = [...somas].sort((a, b) => a.s - b.s), pior = ord[0], melhor = ord[ord.length - 1];
        return `${pior.r} concentra o maior desvio negativo (${comSinal(pior.s)} no período); ${melhor.r} o maior positivo (${comSinal(melhor.s)})`;
      }
      const total = somas.reduce((a, x) => a + x.s, 0) || 1;
      const top = [...somas].sort((a, b) => b.s - a.s)[0];
      return `${top.r} responde por ${pctSemSinal(top.s / total)} do total`;
    }
    default:
      return "";
  }
}

// ---------------------------- Anotações --------------------------------------
function anotacoes(tela, anot, estiloIBCS, xR) {
  const nb = tela.barras.length;
  if (!nb) return;
  let xs = xR + 14;
  anot.v.forEach((a) => {
    if (!a.on) return;
    const i = ix(a.de, nb), j = ix(a.para, nb);
    if (i !== j) indicador(tela, tela.barras[i - 1], tela.barras[j - 1], i > j, a, estiloIBCS, xs, i - 1, j - 1);
    xs += larguraVar(a);
  });
  if (anot.cagr.on) linhaCAGR(tela, ix(anot.cagr.de, nb), ix(anot.cagr.para, nb));
}

function textoVariacao(va, vb, modo) {
  const abs = fmtNum(vb - va, true);
  let pct = "";
  if (va !== 0) { const p = (vb - va) / Math.abs(va); pct = fmtPct(p, Math.abs(p) >= 1 ? 0 : 1); }
  if (modo === "A") return abs;
  if (modo === "B") return pct ? `${abs}\n(${pct})` : abs;
  return pct || abs;
}

// Linha de referência tracejada que sai da barra de origem (idx) até xFim, na altura y.
// Fica POR CIMA das barras (contínua) e só se interrompe onde passaria sobre um texto
// já desenhado (rótulos de valor, inclusive os de dentro das barras) ou sobre balões
// e setas de outras variações.
function linhaReferencia(tela, idx, y, xFim, cor) {
  const folga = 3;
  const x0 = tela.barras[idx].cx + tela.bw / 2;
  const bloqueios = [];
  for (const q of tela.p) {
    if (q.t !== "text" || !q.text) continue;
    const linhas = String(q.text).split("\n");
    const larg = Math.min(q.w, Math.max(...linhas.map((l) => l.length)) * q.size * 0.56 + 4);
    const alt = q.size * 1.25 * (q.wrap ? 3 : linhas.length);
    const topo = q.valign === "top" ? q.y : q.valign === "middle" ? q.y + q.h / 2 - alt / 2 : q.y + q.h - alt;
    if (y < topo - 1 || y > topo + alt + 1) continue;
    const xi = q.align === "left" ? q.x : q.align === "right" ? q.x + q.w - larg : q.x + q.w / 2 - larg / 2;
    bloqueios.push([xi - folga, xi + larg + folga]);
  }
  for (const [bx1, bx2, by1, by2] of tela.caixasAnot ?? [])
    if (y >= by1 - 1 && y <= by2 + 1) bloqueios.push([bx1 - folga, bx2 + folga]);
  bloqueios.sort((p, q) => p[0] - q[0]);
  let x = x0;
  for (const [a, z] of bloqueios) {
    if (z <= x) continue;
    if (a >= xFim) break;
    if (a - x > 4) tela.line(x, y, Math.min(a, xFim), y, cor, 1, { dash: "dash" });
    x = Math.max(x, z);
  }
  if (xFim - x > 4) tela.line(x, y, xFim, y, cor, 1, { dash: "dash" });
}

function indicador(tela, A, B, aDireita, a, estiloIBCS, xs, ia, ib) {
  const e = tela.e;
  const ya = tela.Y(tela.clip(A.nivel)), yb = tela.Y(tela.clip(B.nivel));
  if (estiloIBCS) {
    const xBorda = (aDireita ? A.cx : B.cx) + tela.bw / 2;
    tela.line(xBorda, ya, xs + 3, ya, "#AFAFAF", 0.75);
    tela.line(xBorda, yb, xs + 3, yb, "#AFAFAF", 0.75);
    tela.line(xs, ya, xs, yb, B.valor >= A.valor ? e.pos : e.neg, 3.5);
    tela.dot(xs, ya, 2.5, "#000000");
    tela.dot(xs, yb, 2.5, "#000000");
  } else {
    linhaReferencia(tela, ia, ya, xs + 10, e.texto);
    linhaReferencia(tela, ib, yb, xs + 10, e.texto);
    tela.arrow(xs, ya, xs, yb, "#000000", 1.75);
  }
  const caixas = (tela.caixasAnot ??= []);
  caixas.push([xs - 4, xs + 4, Math.min(ya, yb), Math.max(ya, yb)]); // seta / barra vertical
  const txt = textoVariacao(A.valor, B.valor, a.modo);
  const n0 = tela.p.length;
  if (a.balao) tela.ellipse(txt, xs + 10, (ya + yb) / 2, true);
  else tela.text(txt, xs + 8, (ya + yb) / 2, "middle", 84, { size: e.fonte + 1, italic: true, align: "left" });
  const q = tela.p[n0];
  const largTxt = q.t === "ellipse" ? q.w : Math.max(...String(txt).split("\n").map((l) => l.length)) * q.size * 0.6;
  caixas.push([q.x, q.x + largTxt, q.y, q.y + q.h]);
}

function linhaCAGR(tela, a, b) {
  if (a === b) return;
  if (a > b) [a, b] = [b, a];
  const A = tela.barras[a - 1], B = tela.barras[b - 1];
  if (A.valor <= 0 || B.valor <= 0) return;
  const g = Math.pow(B.valor / A.valor, 1 / (b - a)) - 1;
  const x1 = A.cx, x2 = B.cx + tela.bw * 0.45;
  let y1 = tela.Y(tela.clip(A.nivel)) - 34, y2 = tela.Y(tela.clip(B.nivel)) - 34;
  let sh = 0;
  for (let i = a; i <= b; i++) {
    const bi = tela.barras[i - 1];
    const yl = y1 + ((y2 - y1) * (bi.cx - x1)) / (x2 - x1);
    const lim = tela.Y(tela.clip(Math.max(bi.nivel, 0))) - 28;
    sh = Math.max(sh, yl - lim);
  }
  y1 -= sh; y2 -= sh;
  tela.arrow(x1, y1, x2, y2, "#000000", 1.5);
  tela.ellipse(fmtPct(g, 0) + tela.e.sufCagr, (x1 + x2) / 2, (y1 + y2) / 2, false);
}
