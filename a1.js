// Utilitários de endereço A1 (sem depender do Excel)
export function colNum(letras) {
  let n = 0;
  for (const ch of letras.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}
export function colLetras(n) {
  let s = "";
  while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

// "Plan1!$A$1:$C$8" -> { sheet: "Plan1", r1, c1, r2, c2 }
export function parseA1(endereco) {
  let sheet = null, ref = endereco;
  const i = endereco.lastIndexOf("!");
  if (i >= 0) { sheet = endereco.slice(0, i).replace(/^'(.*)'$/, "$1").replace(/''/g, "'"); ref = endereco.slice(i + 1); }
  const partes = ref.replace(/\$/g, "").split(":");
  const um = (p) => {
    const m = /^([A-Z]+)?(\d+)?$/i.exec(p);
    if (!m) throw new Error("Endereço inválido: " + endereco);
    return { c: m[1] ? colNum(m[1]) : null, r: m[2] ? +m[2] : null };
  };
  const a = um(partes[0]), b = um(partes[1] ?? partes[0]);
  return {
    sheet,
    c1: a.c ?? 1, r1: a.r ?? 1,
    c2: b.c ?? 16384, r2: b.r ?? 1048576,
  };
}

export const montarA1 = (r1, c1, r2, c2) => `${colLetras(c1)}${r1}:${colLetras(c2)}${r2}`;

export function intersecta(a, b) {
  return a.r1 <= b.r2 && b.r1 <= a.r2 && a.c1 <= b.c2 && b.c1 <= a.c2;
}
