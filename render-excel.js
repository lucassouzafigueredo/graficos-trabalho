// =============================================================================
// Desenha o gráfico no Excel (ExcelApi 1.9). Dois modos:
//
//  • "imagem" (Rápido): o gráfico inteiro vira UMA imagem vetorial (SVG). Uma única
//    ida e volta ao Excel. No PowerPoint dá para desagrupar com "Converter em Forma".
//  • "formas" (Editável): cada barra/texto/linha é uma forma nativa, agrupadas.
//
// Desempenho do modo "formas": cada ida e volta ao Excel (context.sync) é cara. O
// add-in memoriza o que a versão do Excel do usuário aceita (espessuras, tracejados,
// setas, SVG) e, a partir do 2º gráfico, aplica tudo já na primeira ida e volta.
// Se a memória estiver errada (ex.: Excel atualizado), ela é apagada e o desenho
// é refeito com tentativas.
// =============================================================================
import { svgListrado, segmentosListras, ESPESSURA_LISTRA } from "./listras.js?v=1.5.0";
import { renderSVG } from "./render-svg.js?v=1.5.0";

const DASH = { dash: ["Dash", "SystemDash", "LongDash"], dot: ["RoundDot", "SquareDot", "SystemDot", "Dash"] };
const SETAS = ["Triangle", "Stealth", "Open"];
const HALIGN = { left: "Left", center: "Center", right: "Right" };
const VALIGN = { top: "Top", middle: "Middle", bottom: "Bottom" };
const MARGEM = 8; // o fundo branco do gráfico começa 8 pt antes do quadro
export const PREFIXO = "GFPA_";

// Descreve um erro do Office.js com o local exato (ex.: "ShapeLineFormat.dashStyle")
export function descreverErro(e) {
  const local = e && e.debugInfo && (e.debugInfo.errorLocation || e.debugInfo.message);
  return (e && e.message ? e.message : String(e)) + (local ? ` [${local}]` : "");
}

// ----------------------------------------------------------------- memória do que o Excel aceita
const CHAVE_CAP = "gfpa.capacidades.v1";
function lerCap() {
  try { const c = JSON.parse(localStorage.getItem(CHAVE_CAP)); if (c && c.ok && c.ruim) return c; } catch { /* vazio */ }
  return { ok: {}, ruim: {} };
}
let cap = lerCap();
function gravarCap() { try { localStorage.setItem(CHAVE_CAP, JSON.stringify(cap)); } catch { /* sem armazenamento */ } }
export function esquecerCapacidades() { cap = { ok: {}, ruim: {} }; gravarCap(); }
const conhecido = (chave) => cap.ok[chave];

// Tenta valores alternativos (cada tentativa = 1 ida e volta) e memoriza o resultado
async function aplicarComAlternativas(context, alvos, aplicar, alternativas, chave) {
  if (!alvos.length) return;
  const ruins = new Set(cap.ruim[chave] ?? []);
  for (const valor of alternativas.filter((v) => !ruins.has(String(v)))) {
    try {
      alvos.forEach((a) => aplicar(a, valor));
      await context.sync();
      cap.ok[chave] = valor; gravarCap();
      return;
    } catch (e) {
      console.warn("Valor recusado pelo Excel:", chave, valor, descreverErro(e));
      (cap.ruim[chave] ??= []).push(String(valor)); gravarCap();
    }
  }
}

// ----------------------------------------------------------------- metadados
function lerMetaDoArquivo(nome) {
  try { return Office.context.document.settings.get("gfpa:" + nome); } catch { return null; }
}
function salvarMetaNoArquivo(nome, meta) {
  return new Promise((ok) => {
    try {
      const cfg = Office.context.document.settings;
      cfg.set("gfpa:" + nome, meta);
      cfg.saveAsync(() => ok());
    } catch { ok(); }
  });
}

// Grava nome + Texto Alt (+ posicionamento) numa forma/grupo e devolve o id
async function finalizar(context, forma, meta) {
  const nome = PREFIXO + Date.now() + Math.floor(Math.random() * 1000);
  forma.name = nome;
  if (conhecido("alt") === true) {
    forma.altTextDescription = JSON.stringify(meta);
    forma.placement = "Move";
    forma.load("id");
    try { await context.sync(); return forma.id; } catch { esquecerCapacidades(); forma.name = nome; }
  }
  forma.load("id");
  await context.sync();
  try {
    forma.altTextDescription = JSON.stringify(meta);
    await context.sync();
    cap.ok.alt = true; gravarCap();
  } catch (e) {
    console.warn("Texto Alt recusado, usando as configurações do arquivo:", descreverErro(e));
    cap.ok.alt = false; gravarCap();
    await salvarMetaNoArquivo(nome, meta);
  }
  try { forma.placement = "Move"; await context.sync(); } catch { /* opcional */ }
  return forma.id;
}

// ----------------------------------------------------------------- ponto de entrada
export async function desenharNoExcel(context, sheet, grafico, L, T, meta) {
  const modo = meta?.spec?.modo ?? "formas";
  if (modo === "imagem") {
    try { return await inserirComoImagem(context, sheet, grafico, L, T, meta); }
    catch (e) { console.warn("Imagem recusada; desenhando com formas:", descreverErro(e)); }
  }
  try { return await inserirComFormas(context, sheet, grafico, L, T, meta, true); }
  catch (e) {
    // a memória pode estar desatualizada (ex.: Excel atualizado): apaga e tenta de novo
    if (Object.keys(cap.ok).length || Object.keys(cap.ruim).length) {
      esquecerCapacidades();
      return await inserirComFormas(context, sheet, grafico, L, T, meta, false);
    }
    throw e;
  }
}
export { desenharNoExcel as inserirNoExcel }; // compatibilidade

// ----------------------------------------------------------------- modo Rápido (1 imagem)
const temFuncao = (obj, nome) => { try { return typeof obj[nome] === "function"; } catch { return false; } };

// Desenha o SVG num canvas do painel e devolve um PNG em base64 (alta resolução)
export async function svgParaPngBase64(svg, largura, altura, escala = 3) {
  const img = new Image();
  await new Promise((ok, falha) => {
    img.onload = ok;
    img.onerror = () => falha(new Error("Não foi possível gerar a imagem do gráfico."));
    img.src = "data:image/svg+xml;charset=utf-8," + encodeURIComponent(svg);
  });
  const c = document.createElement("canvas");
  c.width = Math.round(largura * escala); c.height = Math.round(altura * escala);
  const g = c.getContext("2d");
  g.fillStyle = "#FFFFFF"; g.fillRect(0, 0, c.width, c.height);
  g.drawImage(img, 0, 0, c.width, c.height);
  return c.toDataURL("image/png").split(",")[1];
}

async function inserirComoImagem(context, sheet, grafico, L, T, meta) {
  const svg = renderSVG(grafico, { margem: MARGEM });
  const W = grafico.largura + 2 * MARGEM, H = grafico.altura + 2 * MARGEM;
  let s;
  if (temFuncao(sheet.shapes, "addSvg") && conhecido("imagem") !== "png") {
    s = sheet.shapes.addSvg(svg);
  } else if (temFuncao(sheet.shapes, "addImage")) {
    s = sheet.shapes.addImage(await svgParaPngBase64(svg, W, H));
    if (conhecido("imagem") !== "png") { cap.ok.imagem = "png"; gravarCap(); }
  } else {
    throw new Error("Esta versão do Excel não permite inserir imagens pelo add-in.");
  }
  s.left = L - MARGEM; s.top = T - MARGEM; s.width = W; s.height = H;
  return await finalizar(context, s, meta);
}

// ----------------------------------------------------------------- modo Editável (formas)
function prepararTexto(s, p, fonte) {
  const tf = s.textFrame;
  tf.autoSizeSetting = "AutoSizeNone";
  tf.leftMargin = 0; tf.rightMargin = 0; tf.topMargin = 0; tf.bottomMargin = 0;
  tf.horizontalAlignment = HALIGN[p.align] ?? "Center";
  tf.verticalAlignment = VALIGN[p.valign] ?? "Middle";
  const f = tf.textRange.font;
  f.name = fonte; f.size = p.size; f.color = p.color;
  f.bold = !!p.bold; f.italic = !!p.italic;
}

async function inserirComFormas(context, sheet, grafico, L, T, meta, usarMemoria, fonte = "Segoe UI") {
  const shapes = sheet.shapes;
  const criadas = [];
  // detalhes cujo valor aceito ainda não é conhecido (resolvidos na fase 2)
  const pendentes = new Map(); // chave -> { alvos, aplicar, alternativas }
  const pendente = (chave, alvo, aplicar, alternativas) => {
    if (!pendentes.has(chave)) pendentes.set(chave, { alvos: [], aplicar, alternativas });
    pendentes.get(chave).alvos.push(alvo);
  };
  // aplica já (se a memória souber o valor aceito) ou deixa para a fase 2
  const detalhe = (chave, alvo, aplicar, alternativas) => {
    const v = usarMemoria ? conhecido(chave) : undefined;
    if (v !== undefined) aplicar(alvo, v); else pendente(chave, alvo, aplicar, alternativas);
  };
  const peso = (s, w) => detalhe(`peso:${w}`, s, (x, v) => (x.lineFormat.weight = v), [...new Set([w, Math.ceil(w), 1])]);
  const pos = (s, x, y, w, h) => { s.left = L + x; s.top = T + y; s.width = Math.max(0.75, w); s.height = Math.max(0.75, h); };
  const svgDisponivel = temFuncao(shapes, "addSvg");
  let modoListras = usarMemoria ? conhecido("listras") : undefined;
  if (!svgDisponivel) modoListras = "linhas"; // Excel sem addSvg: listras sempre como linhas nativas
  const listradas = [];
  const linhasDeListra = (p) => segmentosListras(p.w, p.h).map(([x1, y1, x2, y2]) => {
    const l = shapes.addLine(L + p.x + x1, T + p.y + y1, L + p.x + x2, T + p.y + y2, "Straight");
    l.lineFormat.color = p.listra;
    detalhe("peso:listra", l, (x, v) => (x.lineFormat.weight = v), [ESPESSURA_LISTRA, 2, 1.5]);
    return l;
  });

  try {
    // ---------------- fase 1: tudo o que já se sabe que o Excel aceita, numa única ida e volta
    for (const p of grafico.prims) {
      let s;
      switch (p.t) {
        case "rect":
          if (p.listra && modoListras === "svg") {
            s = shapes.addSvg(svgListrado(p.w, p.h, p.fill, p.listra, p.stroke));
            pos(s, p.x, p.y, p.w, p.h);
            break;
          }
          s = shapes.addGeometricShape("Rectangle");
          s.fill.setSolidColor(p.fill);
          if (p.stroke) { s.lineFormat.color = p.stroke; peso(s, 1); } else s.lineFormat.visible = false;
          pos(s, p.x, p.y, p.w, p.h);
          if (p.listra) {
            if (modoListras === "linhas") { criadas.push(s); criadas.push(...linhasDeListra(p)); continue; }
            listradas.push({ s, p });
          }
          break;
        case "line":
          if (Math.hypot(p.x2 - p.x1, p.y2 - p.y1) < 0.2) continue;
          s = shapes.addLine(L + p.x1, T + p.y1, L + p.x2, T + p.y2, "Straight");
          s.lineFormat.color = p.color;
          peso(s, p.w);
          if (p.dash === "dash" || p.dash === "dot") detalhe(p.dash, s, (x, v) => (x.lineFormat.dashStyle = v), DASH[p.dash]);
          if (p.arrow) detalhe("seta", s, (x, v) => { x.line.endArrowheadStyle = v; }, SETAS);
          break;
        case "text":
          if (!p.text) continue;
          s = shapes.addTextBox(p.text);
          pos(s, p.x, p.y, p.w, p.h);
          s.fill.clear();
          s.lineFormat.visible = false;
          prepararTexto(s, p, fonte);
          break;
        case "ellipse":
          s = shapes.addGeometricShape("Ellipse");
          pos(s, p.x, p.y, p.w, p.h);
          s.fill.setSolidColor("#FFFFFF");
          s.lineFormat.color = "#000000";
          peso(s, 1.25);
          s.textFrame.textRange.text = p.text;
          prepararTexto(s, { ...p, color: "#000000", bold: true, align: "center", valign: "middle" }, fonte);
          break;
        case "marcador":
          s = shapes.addGeometricShape("Ellipse");
          pos(s, p.cx - p.r, p.cy - p.r, 2 * p.r, 2 * p.r);
          s.fill.setSolidColor("#FFFFFF");
          s.lineFormat.color = p.color;
          peso(s, 1);
          s.textFrame.textRange.text = p.n;
          prepararTexto(s, { size: p.size, color: p.color, align: "center", valign: "middle" }, fonte);
          break;
        case "arco": {
          // sem forma de arco ajustável no Office.js: traços grossos curtos ao longo do raio médio
          const rm = (p.r0 + p.r1) / 2, esp = Math.max(1, Math.round(p.r1 - p.r0));
          const passos = Math.max(2, Math.ceil((p.a1 - p.a0) / (Math.PI / 24)));
          const da = (p.a1 - p.a0) / passos, sobra = Math.min(da * 0.25, 0.6 / rm);
          for (let k = 0; k < passos; k++) {
            const a = p.a0 + k * da - (k ? sobra : 0), b = p.a0 + (k + 1) * da + (k < passos - 1 ? sobra : 0);
            const l = shapes.addLine(L + p.cx + rm * Math.sin(a), T + p.cy - rm * Math.cos(a), L + p.cx + rm * Math.sin(b), T + p.cy - rm * Math.cos(b), "Straight");
            l.lineFormat.color = p.fill;
            peso(l, esp);
            criadas.push(l);
          }
          continue;
        }
        case "dot":
          s = shapes.addGeometricShape("Ellipse");
          pos(s, p.cx - p.r, p.cy - p.r, 2 * p.r, 2 * p.r);
          s.fill.setSolidColor(p.color);
          s.lineFormat.visible = false;
          break;
        default:
          continue;
      }
      criadas.push(s);
    }
    criadas.forEach((s) => s.load("id"));
    await context.sync();

    // ---------------- fase 2: só o que ainda não se sabe (1ª vez neste Excel)
    for (const [chave, t] of pendentes) {
      if (chave === "peso:listra") continue; // resolvido junto com as listras
      await aplicarComAlternativas(context, t.alvos, t.aplicar, t.alternativas, chave);
    }

    // ---------------- fase 2b: listras ainda sem modo conhecido -> testa SVG (todas de uma vez)
    if (listradas.length) {
      let imgs = [];
      try {
        imgs = listradas.map(({ p }) => {
          const img = shapes.addSvg(svgListrado(p.w, p.h, p.fill, p.listra, p.stroke));
          pos(img, p.x, p.y, p.w, p.h);
          img.load("id");
          return img;
        });
        await context.sync();
        listradas.forEach(({ s }, i) => { criadas.splice(criadas.indexOf(s), 1, imgs[i]); s.delete(); });
        await context.sync();
        cap.ok.listras = "svg"; gravarCap();
      } catch (e) {
        console.warn("SVG recusado; listras como linhas nativas:", descreverErro(e));
        if (imgs.length) { try { imgs.forEach((i) => i.delete()); await context.sync(); } catch { /* não chegaram a existir */ } }
        const linhas = listradas.flatMap(({ p }) => linhasDeListra(p));
        linhas.forEach((l) => l.load("id"));
        try {
          await context.sync();
          criadas.push(...linhas);
          cap.ok.listras = "linhas"; gravarCap();
        } catch (e2) { console.warn("Listras recusadas; barras sólidas:", descreverErro(e2)); }
      }
    }
    const pl = pendentes.get("peso:listra");
    if (pl) await aplicarComAlternativas(context, pl.alvos, pl.aplicar, pl.alternativas, "peso:listra");

    // ---------------- fase 3: agrupar e gravar os metadados
    const grupo = shapes.addGroup(criadas.map((s) => s.id));
    criadas.length = 0; // a partir daqui as formas pertencem ao grupo
    return await finalizar(context, grupo, meta);
  } catch (e) {
    try {
      for (const s of criadas) { try { s.delete(); } catch { /* já removida */ } }
      await context.sync();
    } catch { /* nada a fazer */ }
    throw new Error("O Excel recusou o desenho: " + descreverErro(e));
  }
}

// ----------------------------------------------------------------- listagem
function extrair(sheetName, s, out) {
  if (!s.name.startsWith(PREFIXO)) return;
  let meta = null;
  try { meta = JSON.parse(s.altTextDescription); } catch { meta = lerMetaDoArquivo(s.name); }
  if (meta && meta.gfpa) out.push({ id: s.id, left: s.left, top: s.top, meta, sheet: sheetName });
}

// Lista os gráficos do add-in numa planilha
export async function listarGraficos(context, sheet) {
  sheet.shapes.load("items/id,items/name,items/altTextDescription,items/left,items/top");
  await context.sync();
  const out = [];
  sheet.shapes.items.forEach((s) => extrair(undefined, s, out));
  return out;
}

// Lista os gráficos de TODAS as abas em 2 idas e voltas (não importa quantas abas)
export async function listarGraficosTodas(context) {
  const wss = context.workbook.worksheets;
  wss.load("items/name,items/id");
  await context.sync();
  wss.items.forEach((ws) => ws.shapes.load("items/id,items/name,items/altTextDescription,items/left,items/top"));
  await context.sync();
  const out = [], abas = {};
  wss.items.forEach((ws) => {
    abas[ws.id] = ws.name;
    ws.shapes.items.forEach((s) => extrair(ws.name, s, out));
  });
  return { graficos: out, abas };
}
