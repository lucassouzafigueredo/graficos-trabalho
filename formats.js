// Formatação numérica pt-BR (vírgula decimal) com sufixos K/M/B configuráveis
export const config = { sufMil: "K", sufMi: "M", sufBi: "B" };

const f1 = new Intl.NumberFormat("pt-BR", { minimumFractionDigits: 1, maximumFractionDigits: 1 });
const f0 = new Intl.NumberFormat("pt-BR", { maximumFractionDigits: 0 });

export function fmtNum(v, comSinal = false) {
  const a = Math.abs(v);
  let s;
  if (a >= 1e9) s = f1.format(a / 1e9) + config.sufBi;
  else if (a >= 1e6) s = f1.format(a / 1e6) + config.sufMi;
  else if (a >= 1e3) s = f1.format(a / 1e3) + config.sufMil;
  else if (Number.isInteger(a)) s = f0.format(a);
  else s = f1.format(a);
  if (v < 0) return "-" + s;
  if (comSinal && v > 0) return "+" + s;
  return s;
}

export function fmtPct(p, casas) {
  const nf = new Intl.NumberFormat("pt-BR", { minimumFractionDigits: casas, maximumFractionDigits: casas });
  const s = nf.format(Math.abs(p) * 100) + "%";
  if (p > 0) return "+" + s;
  if (p < 0) return "-" + s;
  return s;
}
