// Renderiza as primitivas do motor como SVG (pré-visualização no painel)
const esc = (s) => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
const DASH = { solid: "", dash: "4 3", dot: "1 2" };

function quebrarLinhas(texto, largura, size) {
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
  return linhas.slice(0, 3);
}

function textoSVG(p, fonte) {
  const linhas = p.wrap ? quebrarLinhas(p.text, p.w, p.size) : p.text.split("\n");
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
  let defs = "", n = 0;
  const partes = prims.map((p) => {
    switch (p.t) {
      case "rect": {
        let fill = p.fill;
        if (p.hatch) {
          const id = `h${n++}`;
          defs += `<pattern id="${id}" width="5" height="5" patternUnits="userSpaceOnUse" patternTransform="rotate(45)">` +
            `<rect width="5" height="5" fill="#fff"/><line x1="0" y1="0" x2="0" y2="5" stroke="${p.fill}" stroke-width="2"/></pattern>`;
          fill = `url(#${id})`;
        }
        const stroke = p.stroke ? ` stroke="${p.stroke}" stroke-width="1"` : "";
        return `<rect x="${p.x.toFixed(1)}" y="${p.y.toFixed(1)}" width="${p.w.toFixed(1)}" height="${p.h.toFixed(1)}" fill="${fill}"${stroke}/>`;
      }
      case "line": {
        const marker = p.arrow ? ` marker-end="url(#seta)"` : "";
        const dash = DASH[p.dash] ? ` stroke-dasharray="${DASH[p.dash]}"` : "";
        return `<line x1="${p.x1.toFixed(1)}" y1="${p.y1.toFixed(1)}" x2="${p.x2.toFixed(1)}" y2="${p.y2.toFixed(1)}" stroke="${p.color}" stroke-width="${p.w}"${dash}${marker}/>`;
      }
      case "text":
        return textoSVG(p, fonte);
      case "ellipse":
        return `<ellipse cx="${(p.x + p.w / 2).toFixed(1)}" cy="${(p.y + p.h / 2).toFixed(1)}" rx="${(p.w / 2).toFixed(1)}" ry="${(p.h / 2).toFixed(1)}" fill="#fff" stroke="#000" stroke-width="1.25"/>` +
          textoSVG({ ...p, color: "#000", bold: true, align: "center", valign: "middle", wrap: false }, fonte);
      case "dot":
        return `<circle cx="${p.cx.toFixed(1)}" cy="${p.cy.toFixed(1)}" r="${p.r}" fill="${p.color}"/>`;
      default:
        return "";
    }
  });
  defs += `<marker id="seta" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="5" markerHeight="5" orient="auto-start-reverse"><path d="M0,0 L10,5 L0,10 z" fill="#000"/></marker>`;
  const m = 10;
  return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="${-m} ${-m} ${largura + 2 * m} ${altura + 2 * m}" ` +
    `width="${opts.width ?? largura + 2 * m}" preserveAspectRatio="xMidYMin meet"><defs>${defs}</defs>${partes.join("")}</svg>`;
}

// SVG mínimo de um retângulo hachurado (o Excel não aceita preenchimento com padrão via API)
export function svgHachura(w, h, cor) {
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}" viewBox="0 0 ${w} ${h}">` +
    `<defs><pattern id="p" width="5" height="5" patternUnits="userSpaceOnUse" patternTransform="rotate(45)">` +
    `<rect width="5" height="5" fill="#ffffff"/><line x1="0" y1="0" x2="0" y2="5" stroke="${cor}" stroke-width="2"/></pattern></defs>` +
    `<rect x="0.5" y="0.5" width="${w - 1}" height="${h - 1}" fill="url(#p)" stroke="${cor}" stroke-width="1"/></svg>`;
}
