// =============================================================================
// Listras diagonais (hachura) calculadas como geometria pura.
// O Office não aceita preenchimento com padrão via API e ignora <pattern> em SVG,
// então cada listra vira um polígono recortado no retângulo da barra.
// =============================================================================

const ESPACO = 6;       // distância entre listras (pt), medida na horizontal
const ESPESSURA = 2.4;  // largura da listra (pt), medida na horizontal

// Recorta um polígono convexo pelo semiplano a*x + b*y <= c (Sutherland–Hodgman)
function recortar(poli, a, b, c) {
  const dentro = (p) => a * p[0] + b * p[1] <= c + 1e-9;
  const cruza = (p, q) => {
    const fp = a * p[0] + b * p[1] - c, fq = a * q[0] + b * q[1] - c;
    const t = fp / (fp - fq);
    return [p[0] + t * (q[0] - p[0]), p[1] + t * (q[1] - p[1])];
  };
  const out = [];
  for (let i = 0; i < poli.length; i++) {
    const p = poli[i], q = poli[(i + 1) % poli.length];
    if (dentro(p)) { out.push(p); if (!dentro(q)) out.push(cruza(p, q)); }
    else if (dentro(q)) out.push(cruza(p, q));
  }
  return out;
}

// Polígonos das listras "/" dentro do retângulo [0,w] x [0,h] (y para baixo)
export function poligonosListras(w, h) {
  const caixa = [[0, 0], [w, 0], [w, h], [0, h]];
  const out = [];
  for (let c = 0; c < w + h; c += ESPACO) {
    // faixa c <= x + y <= c + ESPESSURA
    let p = recortar(caixa, 1, 1, c + ESPESSURA);
    p = recortar(p, -1, -1, -c);
    if (p.length >= 3) out.push(p);
  }
  return out;
}

// Segmentos centrais das listras (para desenhar como linhas nativas do Excel)
export function segmentosListras(w, h) {
  const out = [];
  for (let c = ESPESSURA / 2; c < w + h; c += ESPACO) {
    const x1 = Math.max(0, c - h), x2 = Math.min(w, c);
    if (x2 - x1 < 0.5) continue;
    out.push([x1, c - x1, x2, c - x2]);
  }
  return out;
}

export const ESPESSURA_LISTRA = ESPESSURA / Math.SQRT2; // espessura perpendicular

// SVG autônomo (sem <pattern>, sem transformações) de um retângulo listrado
export function svgListrado(w, h, fundo, listra, borda) {
  const pts = (poli) => poli.map(([x, y]) => `${x.toFixed(2)},${y.toFixed(2)}`).join(" ");
  const polys = poligonosListras(w, h).map((p) => `<polygon points="${pts(p)}" fill="${listra}"/>`).join("");
  const contorno = borda ? `<rect x="0.5" y="0.5" width="${(w - 1).toFixed(2)}" height="${(h - 1).toFixed(2)}" fill="none" stroke="${borda}" stroke-width="1"/>` : "";
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${w.toFixed(2)}" height="${h.toFixed(2)}" viewBox="0 0 ${w.toFixed(2)} ${h.toFixed(2)}">` +
    `<rect x="0" y="0" width="${w.toFixed(2)}" height="${h.toFixed(2)}" fill="${fundo}"/>${polys}${contorno}</svg>`;
}
