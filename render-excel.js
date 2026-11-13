// =============================================================================
// Desenha as primitivas do motor como formas nativas do Excel (ExcelApi 1.9)
// e agrupa tudo num único objeto, com os metadados no Texto Alt.
//
// Robustez: o desenho é feito em fases. A fase 1 (formas, cores, textos) usa só
// propriedades básicas. Os detalhes (tracejado, pontas de seta) vêm depois, cada
// um com alternativas: se a versão do Excel recusar um valor, tenta o próximo e,
// no pior caso, mantém a linha contínua em vez de derrubar o gráfico inteiro.
// Se algo essencial falhar, as formas parciais são apagadas.
// =============================================================================
import { svgHachura } from "./render-svg.js?v=1.0.4";

const DASH = { dash: ["Dash", "SystemDash", "LongDash"], dot: ["RoundDot", "SquareDot", "SystemDot", "Dash"] };
const HALIGN = { left: "Left", center: "Center", right: "Right" };
const VALIGN = { top: "Top", middle: "Middle", bottom: "Bottom" };
export const PREFIXO = "GFPA_";

// Descreve um erro do Office.js com o local exato (ex.: "ShapeLineFormat.dashStyle")
export function descreverErro(e) {
  const local = e && e.debugInfo && (e.debugInfo.errorLocation || e.debugInfo.message);
  return (e && e.message ? e.message : String(e)) + (local ? ` [${local}]` : "");
}

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

// Tenta aplicar um valor com alternativas; cada tentativa é sincronizada isoladamente
async function aplicarComAlternativas(context, alvos, aplicar, alternativas) {
  if (!alvos.length) return;
  for (const valor of alternativas) {
    try {
      alvos.forEach((a) => aplicar(a, valor));
      await context.sync();
      return;
    } catch (e) {
      console.warn("Valor recusado pelo Excel:", valor, descreverErro(e));
    }
  }
}

/**
 * @param sheet    planilha de destino
 * @param grafico  resultado de montarGrafico()
 * @param L,T      canto superior esquerdo do quadro (pontos)
 * @param meta     objeto gravado no Texto Alt do grupo (JSON)
 * @returns id do grupo criado
 */
export async function inserirNoExcel(context, sheet, grafico, L, T, meta, fonte = "Segoe UI") {
  const shapes = sheet.shapes;
  const criadas = [];
  const tracejadas = { dash: [], dot: [] };
  const setas = [];
  const pesos = new Map(); // espessura -> formas (aplicadas na fase 2)
  const peso = (s, w) => { if (!pesos.has(w)) pesos.set(w, []); pesos.get(w).push(s); };
  const pos = (s, x, y, w, h) => { s.left = L + x; s.top = T + y; s.width = Math.max(0.75, w); s.height = Math.max(0.75, h); };

  try {
    // ---------------- fase 1: formas, cores e textos
    for (const p of grafico.prims) {
      let s;
      switch (p.t) {
        case "rect":
          if (p.hatch) {
            try {
              s = shapes.addSvg(svgHachura(Math.round(p.w * 4 / 3), Math.round(p.h * 4 / 3), p.fill));
            } catch { s = null; }
            if (!s) { s = shapes.addGeometricShape("Rectangle"); s.fill.setSolidColor("#FFFFFF"); s.lineFormat.color = p.fill; peso(s, 1.5); }
          } else {
            s = shapes.addGeometricShape("Rectangle");
            s.fill.setSolidColor(p.fill);
            if (p.stroke) { s.lineFormat.color = p.stroke; peso(s, 1); } else s.lineFormat.visible = false;
          }
          pos(s, p.x, p.y, p.w, p.h);
          break;
        case "line":
          if (Math.hypot(p.x2 - p.x1, p.y2 - p.y1) < 0.2) continue;
          s = shapes.addLine(L + p.x1, T + p.y1, L + p.x2, T + p.y2, "Straight");
          s.lineFormat.color = p.color;
          peso(s, p.w);
          if (p.dash === "dash" || p.dash === "dot") tracejadas[p.dash].push(s);
          if (p.arrow) setas.push(s);
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

    // ---------------- fase 2: detalhes (com alternativas; nunca derrubam o gráfico)
    for (const [w, lista] of pesos) {
      await aplicarComAlternativas(context, lista, (s, v) => (s.lineFormat.weight = v), [...new Set([w, Math.ceil(w), 1])]);
    }
    await aplicarComAlternativas(context, tracejadas.dash, (s, v) => (s.lineFormat.dashStyle = v), DASH.dash);
    await aplicarComAlternativas(context, tracejadas.dot, (s, v) => (s.lineFormat.dashStyle = v), DASH.dot);
    await aplicarComAlternativas(context, setas, (s, v) => { s.line.endArrowheadStyle = v; }, ["Triangle", "Stealth", "Open"]);

    // ---------------- fase 3: agrupar e gravar os metadados
    const grupo = shapes.addGroup(criadas.map((s) => s.id));
    const nome = PREFIXO + Date.now();
    grupo.name = nome;
    grupo.load("id");
    await context.sync();
    criadas.length = 0; // a partir daqui as formas pertencem ao grupo
    try {
      grupo.altTextDescription = JSON.stringify(meta);
      await context.sync();
    } catch (e) {
      // Texto Alt recusado (ex.: limite de tamanho): guarda as configurações no próprio arquivo
      console.warn("Texto Alt recusado, usando as configurações do arquivo:", descreverErro(e));
      await salvarMetaNoArquivo(nome, meta);
      try { grupo.altTextDescription = "GFPA"; await context.sync(); } catch { /* opcional */ }
    }
    try { grupo.placement = "Move"; await context.sync(); } catch { /* opcional */ }
    return grupo.id;
  } catch (e) {
    // apaga o que foi desenhado pela metade
    try {
      for (const s of criadas) { try { s.delete(); } catch { /* já removida */ } }
      await context.sync();
    } catch { /* nada a fazer */ }
    throw new Error("O Excel recusou o desenho: " + descreverErro(e));
  }
}

// Configurações guardadas no arquivo (Office document settings), por nome do grupo
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

// Lista os gráficos do add-in numa planilha
export async function listarGraficos(context, sheet) {
  const shapes = sheet.shapes;
  shapes.load("items/id,items/name,items/altTextDescription,items/left,items/top");
  await context.sync();
  const out = [];
  for (const s of shapes.items) {
    if (!s.name.startsWith(PREFIXO)) continue;
    let meta = null;
    try { meta = JSON.parse(s.altTextDescription); } catch { meta = lerMetaDoArquivo(s.name); }
    if (meta && meta.gfpa) out.push({ id: s.id, left: s.left, top: s.top, meta });
  }
  return out;
}
