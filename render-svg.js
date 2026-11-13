// Renderiza as primitivas do motor como SVG (pré-visualização no painel)
import { poligonosListras } from "./listras.js?v=1.5.1";
const esc = (s) => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
const DASH = { solid: "", dash: "4 3", dot: "1 2" };
let contadorSVG = 0; // ids únicos entre vários SVGs na mesma página

function quebrarLinhas(texto, largura, size, maxLinhas = 3) {
  const max = Math.max(1, Math.floor(largura / (size * 0.55)));
  const palavras = texto.split(/\s+/);
  const linhas = [];
  let atual = "";
  for (const p of palavras) {
    if (!atual) atual = p;
    else if ((atual + " " + p).length <= max) atual += " " + p;
    else { linhas.push(atual); atual = p; }
  }
  if (atual) linhas.push(atual);
  if (linhas.length > maxLinhas) { linhas.length = maxLinhas; linhas[maxLinhas - 1] = linhas[maxLinhas - 1].replace(/\s*\S*$/, "") + "…"; }
  return linhas;
}

function textoSVG(p, fonte) {
  const linhas = p.wrap ? quebrarLinhas(p.text, p.w, p.size, p.maxLinhas ?? 3) : p.text.split("\n");
  const lh = p.size * 1.25;
  const anchor = p.align === "left" ? "start" : p.align === "right" ? "end" : "middle";
  const x = p.align === "left" ? p.x : p.align === "right" ? p.x + p.w : p.x + p.w / 2;
  const blocoH = lh * linhas.length;
  let y0;
  if (p.valign === "top") y0 = p.y;
  else if (p.valign === "middle") y0 = p.y + p.h / 2 - blocoH / 2;
  else y0 = p.y + p.h - blocoH;
  const attrs = `font-family="${fonte}" font-size="${p.size}" fill="${p.color}" text-anchor="${anchor}"` +
    (p.bold ? ' font-weight="bold"' : "") + (p.italic ? ' font-style="italic"' : "");
  return linhas.map((l, i) =>
    `<text x="${x.toFixed(1)}" y="${(y0 + lh * (i + 0.8)).toFixed(1)}" ${attrs}>${esc(l)}</text>`).join("");
}

export function renderSVG(grafico, opts = {}) {
  const fonte = opts.fonte ?? "Segoe UI, Arial, sans-serif";
  const { largura, altura, prims } = grafico;
  let defs = "";
  const pref = `g${++contadorSVG}`;
  const partes = prims.map((p) => {
    switch (p.t) {
      case "rect": {
        const stroke = p.stroke ? ` stroke="${p.stroke}" stroke-width="1"` : "";
        const base = `<rect x="${p.x.toFixed(1)}" y="${p.y.toFixed(1)}" width="${p.w.toFixed(1)}" height="${p.h.toFixed(1)}" fill="${p.fill}"${p.listra ? "" : stroke}/>`;
        if (!p.listra) return base;
        const polys = poligonosListras(p.w, p.h).map((poli) =>
          `<polygon points="${poli.map(([x, y]) => `${(p.x + x).toFixed(1)},${(p.y + y).toFixed(1)}`).join(" ")}" fill="${p.listra}"/>`).join("");
        const contorno = p.stroke ? `<rect x="${p.x.toFixed(1)}" y="${p.y.toFixed(1)}" width="${p.w.toFixed(1)}" height="${p.h.toFixed(1)}" fill="none"${stroke}/>` : "";
        return base + polys + contorno;
      }
      case "line": {
        // seta desenhada como triângulo (o Office não suporta <marker> em SVG)
        let { x1, y1, x2, y2 } = p, cabeca = "";
        if (p.arrow) {
          const dx = x2 - x1, dy = y2 - y1, len = Math.hypot(dx, dy) || 1;
          const ux = dx / len, uy = dy / len, comp = Math.min(8, len * 0.6), larg = 3.5 + p.w;
          const bx = x2 - ux * comp, by = y2 - uy * comp;
          cabeca = `<polygon points="${x2.toFixed(1)},${y2.toFixed(1)} ${(bx - uy * larg).toFixed(1)},${(by + ux * larg).toFixed(1)} ${(bx + uy * larg).toFixed(1)},${(by - ux * larg).toFixed(1)}" fill="${p.color}"/>`;
          x2 = bx + ux * 0.5; y2 = by + uy * 0.5;
        }
        const dash = DASH[p.dash] ? ` stroke-dasharray="${DASH[p.dash]}"` : "";
        return `<line x1="${x1.toFixed(1)}" y1="${y1.toFixed(1)}" x2="${x2.toFixed(1)}" y2="${y2.toFixed(1)}" stroke="${p.color}" stroke-width="${p.w}"${dash}/>` + cabeca;
      }
      case "text":
        return textoSVG(p, fonte);
      case "ellipse":
        return `<ellipse cx="${(p.x + p.w / 2).toFixed(1)}" cy="${(p.y + p.h / 2).toFixed(1)}" rx="${(p.w / 2).toFixed(1)}" ry="${(p.h / 2).toFixed(1)}" fill="#fff" stroke="#000" stroke-width="1.25"/>` +
          textoSVG({ ...p, color: "#000", bold: true, align: "center", valign: "middle", wrap: false }, fonte);
      case "marcador":
        return `<circle cx="${p.cx.toFixed(1)}" cy="${p.cy.toFixed(1)}" r="${p.r}" fill="#fff" stroke="${p.color}" stroke-width="1.1"/>` +
          `<text x="${p.cx.toFixed(1)}" y="${(p.cy + p.size * 0.36).toFixed(1)}" font-family="${fonte}" font-size="${p.size}" fill="${p.color}" text-anchor="middle">${esc(p.n)}</text>`;
      case "arco": {
        const P = (a, r) => `${(p.cx + r * Math.sin(a)).toFixed(2)},${(p.cy - r * Math.cos(a)).toFixed(2)}`;
        const g = p.a1 - p.a0 > Math.PI ? 1 : 0;
        return `<path d="M${P(p.a0, p.r1)} A${p.r1.toFixed(2)},${p.r1.toFixed(2)} 0 ${g} 1 ${P(p.a1, p.r1)} L${P(p.a1, p.r0)} A${p.r0.toFixed(2)},${p.r0.toFixed(2)} 0 ${g} 0 ${P(p.a0, p.r0)} Z" fill="${p.fill}"/>`;
      }
      case "dot":
        return `<circle cx="${p.cx.toFixed(1)}" cy="${p.cy.toFixed(1)}" r="${p.r}" fill="${p.color}"/>`;
      default:
        return "";
    }
  });
  const m = opts.margem ?? 10;
  const W = largura + 2 * m, H = altura + 2 * m;
  const tam = opts.width ? `width="${opts.width}"` : `width="${W}" height="${H}"`;
  return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="${-m} ${-m} ${W} ${H}" ${tam} preserveAspectRatio="xMidYMin meet">` +
    (defs ? `<defs>${defs}</defs>` : "") + partes.join("") + `</svg>`;
}
