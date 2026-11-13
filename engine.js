// =============================================================================
// Gráficos FP&A — motor de layout
// Converte (dados + especificação + estilo) em uma lista de primitivas
// desenháveis (retângulos, linhas, textos, elipses), em pontos, com origem
// no canto superior esquerdo do quadro. Os renderizadores (SVG para a
// pré-visualização, Excel para inserir na planilha) apenas desenham a lista.
// =============================================================================

import { fmtNum, fmtPct } from "./formats.js?v=1.0.4";

export const TIPOS = {
  IBCS: { nome: "Cascata IBCS", W: 420, H: 250, comTipo: false },
  IBCSM: { nome: "Cascata IBCS – múltiplos", W: 280, H: 165, comTipo: false },
  EMP: { nome: "Cascata empilhada", W: 520, H: 320, comTipo: true },
  COL: { nome: "Colunas", W: 460, H: 280, comTipo: false },
  COMP: { nome: "Comparação", W: 340, H: 260, comTipo: false },
};

// Anotações que já vêm ligadas em cada tipo de gráfico
export function anotacoesPadrao(tipo) {
  const v = () => ({ on: false, de: 1, para: 0, modo: "P", balao: false });
  const a = { v: [v(), v()], cagr: { on: false, de: 1, para: 0 } };
  if (tipo === "IBCS" || tipo === "IBCSM") a.v[0].on = true;
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
  let rot = [], tipo = [], v = [];
  let cab = values[0].slice(c0).map((x, i) => (temCab ? String(texts[0][c0 + i]) : `Série ${i + 1}`));
  for (let r = r0; r < values.length; r++) {
    const linhaVazia = values[r].every((x) => x === "" || x === null);
    if (linhaVazia) continue;
    rot.push(String(texts[r][0] ?? ""));
    tipo.push(comTipo ? String(values[r][1] ?? "").trim().toUpperCase() : "");
    v.push(values[r].slice(c0).map((x) => (isNum(x) ? x : null)));
  }
  if (!rot.length) throw new Error("Não há linhas de dados.");
  // descarta colunas sem nenhum número (ex.: "Tipo" ou colunas de espaçamento)
  const manter = cab.map((_, c) => v.some((linha) => linha[c] !== null));
  if (!manter.some(Boolean)) throw new Error("Não encontrei valores numéricos no intervalo.");
  cab = cab.filter((_, c) => manter[c]);
  v = v.map((linha) => linha.filter((_, c) => manter[c]));
  return { rot, cab, v, tipo, nLin: rot.length, nCol: cab.length };
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
    const linhas = opts.wrap ? 3 : String(text).split("\n").length;
    const h = opts.wrap ? size * 1.35 * 3 : size * (0.2 + 1.4 * linhas);
    const top = ancora === "top" ? y : ancora === "middle" ? y - h / 2 : y - h;
    const alinh = opts.align ?? "center";
    const left = alinh === "left" ? x : alinh === "right" ? x - w : x - w / 2;
    this.p.push({
      t: "text", x: left, y: top, w: Math.max(10, w), h, text: String(text), size,
      color: opts.color ?? this.e.texto, bold: !!opts.bold, italic: !!opts.italic,
      align: alinh, valign: ancora, wrap: !!opts.wrap,
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
  dot(cx, cy, r, color) { this.p.push({ t: "dot", cx, cy, r, color }); }
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
export function montarGrafico(dados, spec, estilo) {
  const tela = new Tela(estilo);
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
    default:
      throw new Error("Tipo de gráfico desconhecido: " + spec.tipo);
  }
  return { largura, altura, prims: tela.p };
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
  else { tela.vMin = Math.min(0, ...lo); tela.vMax = Math.max(0, ...hi); }
  if (tela.vMax === tela.vMin) tela.vMax = tela.vMin + 1;

  if (comFundo) tela.fundo(L, T, W, H);
  let titH = 0;
  if (titulo) { titH = 20; tela.text(titulo, L, T, "top", W, { size: e.fonte + 3, align: "left" }); }
  tela.pTop = T + titH + 16 + reservaTopo(anot);
  tela.pBot = T + H - 30 - (tela.vMin < 0 ? 16 : 0);
  const xL = L + 2, xR = L + W - 2 - margemDireita(anot);
  const catW = (xR - xL) / n, bw = catW * 0.66;
  tela.iniciarBarras(bw);
  const rotulo = (txt, acima, cx, i) => acima
    ? tela.text(txt, cx, tela.Y(hi[i]) - 2, "bottom", catW + 16, { size: e.fonte + 1 })
    : tela.text(txt, cx, tela.Y(lo[i]) + 2, "top", catW + 16, { size: e.fonte + 1 });

  for (let i = 0; i < n; i++) {
    const cx = xL + catW * (i + 0.5), bx = cx - bw / 2;
    tela.regBarra(cx, ac[i], ac[i]);
    if (i === 0) {
      tela.rect(bx, tela.Y(lo[i]), bw, tela.Y(hi[i]), e.total, { hatch: true, stroke: e.total });
      rotulo(fmtNum(ini), ini >= 0, cx, i);
    } else if (i === n - 1) {
      tela.rect(bx, tela.Y(lo[i]), bw, tela.Y(hi[i]), e.total);
      rotulo(fmtNum(fim), fim >= 0, cx, i);
    } else {
      const dv = ac[i] - ac[i - 1];
      tela.rect(bx, tela.Y(lo[i]), bw, tela.Y(hi[i]), dv >= 0 ? e.pos : e.neg);
      if (dv !== 0) rotulo(fmtNum(dv, true), dv >= 0, cx, i);
    }
    if (i < n - 1) tela.line(bx + bw, tela.Y(ac[i]), bx + catW, tela.Y(ac[i]), "#AFAFAF", 0.75);
    tela.text(d.rot[i], cx, T + H - 26, "top", catW, { size: e.fonte + 1, wrap: true });
  }
  tela.line(xL, tela.Y(0), xR, tela.Y(0), "#404040", 1.5);
  anotacoes(tela, anot, true, xR);
}

function multiplos(tela, d, Wp, Hp, titulo, anot) {
  const GX = 14, GY = 10;
  let gMin = 0, gMax = 0;
  for (let c = 0; c < d.nCol; c++) {
    const { lo, hi } = calcCascata(d, c);
    gMin = Math.min(gMin, ...lo); gMax = Math.max(gMax, ...hi);
  }
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
  tela.vMin = Math.min(0, ...lo); tela.vMax = Math.max(0, ...hi);
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
      const cor = e.porSinal ? (tot[i] >= 0 ? e.pos : e.neg) : e.delta;
      tela.rect(bx, Y(C(lo[i])), bw, Y(C(hi[i])), cor);
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
        tela.rect(bx, y1, bw, y2, cor);
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
  tela.vMin = 0; tela.vMax = Math.max(0, ...tot);
  if (tela.vMax === tela.vMin) tela.vMax = 1;
  tela.fundo(0, 0, W, H);
  let titH = 0;
  if (titulo) { titH = 24; tela.text(titulo, 0, 0, "top", W, { size: e.fonte + 3, bold: true, align: "left", color: "#000000" }); }
  tela.pTop = titH + 18 + reservaTopo(anot);
  tela.pBot = H - 22;
  const xL = 0, xR = W - margemDireita(anot);
  const catW = (xR - xL) / n, bw = catW * 0.62;
  tela.iniciarBarras(bw);
  for (let i = 0; i < n; i++) {
    const cx = xL + catW * (i + 0.5), bx = cx - bw / 2;
    tela.regBarra(cx, tot[i], tot[i]);
    let cur = 0;
    d.v[i].forEach((x, c) => {
      const a = Math.abs(num(x));
      if (a > 0) { tela.rect(bx, tela.Y(cur), bw, tela.Y(cur + a), e.colunas[c % 4]); cur += a; }
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
  tela.vMin = Math.min(0, ...vals); tela.vMax = Math.max(0, ...vals);
  if (tela.vMax === tela.vMin) tela.vMax = tela.vMin + 1;
  tela.fundo(0, 0, W, H);
  let titH = 0;
  if (titulo) { titH = 24; tela.text(titulo, 0, 0, "top", W, { size: e.fonte + 3, bold: true, align: "left", color: "#000000" }); }
  tela.pTop = titH + 26 + reservaTopo(anot);
  tela.pBot = H - 22;
  const xL = 0, xR = W - margemDireita(anot);
  const catW = (xR - xL) / n, bw = catW * 0.6;
  tela.iniciarBarras(bw);
  vals.forEach((v, i) => {
    const cx = xL + catW * (i + 0.5), bx = cx - bw / 2;
    tela.regBarra(cx, v, v);
    tela.rect(bx, tela.Y(0), bw, tela.Y(v), i === 0 ? e.colunas[0] : e.colunas[1]);
    if (v >= 0) tela.text(fmtNum(v), cx, tela.Y(v) - 3, "bottom", catW, { size: e.fonte + 6 });
    else tela.text(fmtNum(v), cx, tela.Y(v) + 3, "top", catW, { size: e.fonte + 6 });
    tela.text(d.rot[i], cx, tela.pBot + 4, "top", catW, { size: e.fonte + 1, wrap: true });
  });
  tela.line(xL, tela.Y(0), xR, tela.Y(0), "#404040", 0.75);
  anotacoes(tela, anot, false, xR);
}

// ---------------------------- Anotações --------------------------------------
function anotacoes(tela, anot, estiloIBCS, xR) {
  const nb = tela.barras.length;
  if (!nb) return;
  let xs = xR + 14;
  anot.v.forEach((a) => {
    if (!a.on) return;
    const i = ix(a.de, nb), j = ix(a.para, nb);
    if (i !== j) indicador(tela, tela.barras[i - 1], tela.barras[j - 1], i > j, a, estiloIBCS, xs);
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

function indicador(tela, A, B, aDireita, a, estiloIBCS, xs) {
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
    tela.line(A.cx + tela.bw / 2, ya, xs + 10, ya, e.texto, 1, { dash: "dash" });
    tela.line(B.cx + tela.bw / 2, yb, xs + 10, yb, e.texto, 1, { dash: "dash" });
    tela.arrow(xs, ya, xs, yb, "#000000", 1.75);
  }
  const txt = textoVariacao(A.valor, B.valor, a.modo);
  if (a.balao) tela.ellipse(txt, xs + 10, (ya + yb) / 2, true);
  else tela.text(txt, xs + 8, (ya + yb) / 2, "middle", 84, { size: e.fonte + 1, italic: true, align: "left" });
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
