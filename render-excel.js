// =============================================================================
// Desenha as primitivas do motor como formas nativas do Excel (ExcelApi 1.9)
// e agrupa tudo num único objeto, com os metadados no Texto Alt.
// =============================================================================
import { svgHachura } from "./render-svg.js";

const DASH = { solid: "Solid", dash: "Dash", dot: "RoundDot" };
const HALIGN = { left: "Left", center: "Center", right: "Right" };
const VALIGN = { top: "Top", middle: "Middle", bottom: "Bottom" };
export const PREFIXO = "GFPA_";

function semBorda(s) { s.lineFormat.visible = false; }

function prepararTexto(s, p, fonte) {
  const tf = s.textFrame;
  tf.autoSizeSetting = "AutoSizeNone";
  tf.leftMargin = 0; tf.rightMargin = 0; tf.topMargin = 0; tf.bottomMargin = 0;
  tf.horizontalAlignment = HALIGN[p.align] ?? "Center";
  tf.verticalAlignment = VALIGN[p.valign] ?? "Middle";
  tf.horizontalOverflow = "Overflow";
  tf.verticalOverflow = "Overflow";
  const f = tf.textRange.font;
  f.name = fonte; f.size = p.size; f.color = p.color;
  f.bold = !!p.bold; f.italic = !!p.italic;
}

/**
 * @param {Excel.Worksheet} sheet  planilha de destino
 * @param grafico  resultado de montarGrafico()
 * @param L,T      canto superior esquerdo do quadro (pontos)
 * @param meta     objeto gravado no Texto Alt do grupo (JSON)
 * @returns id do grupo criado
 */
export async function inserirNoExcel(context, sheet, grafico, L, T, meta, fonte = "Segoe UI") {
  const shapes = sheet.shapes;
  const criadas = [];
  const pos = (s, x, y, w, h) => { s.left = L + x; s.top = T + y; s.width = Math.max(0.75, w); s.height = Math.max(0.75, h); };

  for (const p of grafico.prims) {
    let s;
    switch (p.t) {
      case "rect":
        if (p.hatch) {
          s = shapes.addSvg(svgHachura(Math.round(p.w * 4 / 3), Math.round(p.h * 4 / 3), p.fill));
        } else {
          s = shapes.addGeometricShape("Rectangle");
          s.fill.setSolidColor(p.fill);
          if (p.stroke) { s.lineFormat.color = p.stroke; s.lineFormat.weight = 1; } else semBorda(s);
        }
        pos(s, p.x, p.y, p.w, p.h);
        break;
      case "line": {
        if (Math.hypot(p.x2 - p.x1, p.y2 - p.y1) < 0.2) continue;
        s = shapes.addLine(L + p.x1, T + p.y1, L + p.x2, T + p.y2, "Straight");
        s.lineFormat.color = p.color;
        s.lineFormat.weight = p.w;
        s.lineFormat.dashStyle = DASH[p.dash] ?? "Solid";
        if (p.arrow) { s.line.endArrowheadStyle = "Triangle"; s.line.endArrowheadLength = "Medium"; s.line.endArrowheadWidth = "Medium"; }
        break;
      }
      case "text":
        if (!p.text) continue;
        s = shapes.addTextBox(p.text);
        pos(s, p.x, p.y, p.w, p.h);
        s.fill.clear();
        semBorda(s);
        prepararTexto(s, p, fonte);
        break;
      case "ellipse":
        s = shapes.addGeometricShape("Ellipse");
        pos(s, p.x, p.y, p.w, p.h);
        s.fill.setSolidColor("#FFFFFF");
        s.lineFormat.color = "#000000";
        s.lineFormat.weight = 1.25;
        s.textFrame.textRange.text = p.text;
        prepararTexto(s, { ...p, color: "#000000", bold: true, align: "center", valign: "middle" }, fonte);
        break;
      case "dot":
        s = shapes.addGeometricShape("Ellipse");
        pos(s, p.cx - p.r, p.cy - p.r, 2 * p.r, 2 * p.r);
        s.fill.setSolidColor(p.color);
        semBorda(s);
        break;
      default:
        continue;
    }
    criadas.push(s);
  }
  criadas.forEach((s) => s.load("id"));
  await context.sync();

  const grupo = shapes.addGroup(criadas.map((s) => s.id));
  grupo.name = PREFIXO + Date.now();
  grupo.altTextDescription = JSON.stringify(meta);
  grupo.placement = "Move";
  grupo.load("id");
  await context.sync();
  return grupo.id;
}

// Lista os gráficos do add-in numa planilha
export async function listarGraficos(context, sheet) {
  const shapes = sheet.shapes;
  shapes.load("items/id,items/name,items/altTextDescription,items/left,items/top");
  await context.sync();
  const out = [];
  for (const s of shapes.items) {
    if (!s.name.startsWith(PREFIXO)) continue;
    try {
      const meta = JSON.parse(s.altTextDescription);
      if (meta && meta.gfpa) out.push({ id: s.id, left: s.left, top: s.top, meta });
    } catch { /* forma com nome parecido, mas sem metadados */ }
  }
  return out;
}
