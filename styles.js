// =============================================================================
// Estilos: 5 prontos + estilos do usuário (guardados no navegador do add-in
// e exportáveis/importáveis como arquivo .json para compartilhar via OneDrive)
// =============================================================================

const base = { fonte: 9, texto: "#262626", delta: "#BFBFBF", pos: "#8CC63F", neg: "#E62020", porSinal: false, sinalDelta: false, sufCagr: " a.a." };

export const ESTILOS_PRONTOS = [
  { ...base, id: "cons-azul", nome: "Consultoria – azul", pronto: true, total: "#1A53FF",
    pilha: ["#051C2C", "#1E3CBE", "#2251FF", "#609BFF", "#A0C3FF", "#CDE1FF"],
    colunas: ["#1A53FF", "#599BFF", "#A6C9FF", "#D2E4FF"] },
  { ...base, id: "cons-navy", nome: "Consultoria – navy", pronto: true, total: "#051C2C",
    pilha: ["#051C2C", "#203E5C", "#466E96", "#78A0C8", "#AAC8E1", "#D7E6F0"],
    colunas: ["#051C2C", "#78A0C8", "#AAC8E1", "#D7E6F0"] },
  { ...base, id: "ibcs", nome: "IBCS (verde/vermelho)", pronto: true, total: "#3B5770", porSinal: true,
    pilha: ["#3B5770", "#627C94", "#8CA0B4", "#B4C3D2", "#D2DCE6", "#EBF0F5"],
    colunas: ["#3B5770", "#8CA0B4", "#B4C3D2", "#D2DCE6"] },
  { ...base, id: "ibcs-custo", nome: "IBCS – custo (invertido)", pronto: true, total: "#3B5770", porSinal: true,
    pos: "#E62020", neg: "#8CC63F",
    pilha: ["#3B5770", "#627C94", "#8CA0B4", "#B4C3D2", "#D2DCE6", "#EBF0F5"],
    colunas: ["#3B5770", "#8CA0B4", "#B4C3D2", "#D2DCE6"] },
  { ...base, id: "cinza", nome: "Cinza executivo", pronto: true, total: "#404040", delta: "#CDCDCD",
    pilha: ["#262626", "#595959", "#808080", "#A6A6A6", "#C8C8C8", "#E1E1E1"],
    colunas: ["#404040", "#A6A6A6", "#C8C8C8", "#E1E1E1"] },
];

const CHAVE = "gfpa.estilos.v1";
const CHAVE_ATUAL = "gfpa.estiloAtual";

function lerLocal(chave, padrao) {
  try { const s = localStorage.getItem(chave); return s ? JSON.parse(s) : padrao; } catch { return padrao; }
}
function gravarLocal(chave, valor) {
  try { localStorage.setItem(chave, JSON.stringify(valor)); } catch { /* sem armazenamento */ }
}

export function estilosUsuario() { return lerLocal(CHAVE, []); }
export function todosEstilos() { return [...ESTILOS_PRONTOS, ...estilosUsuario()]; }
export function estiloPorId(id) { return todosEstilos().find((e) => e.id === id) ?? ESTILOS_PRONTOS[0]; }
export function estiloAtualId() { return lerLocal(CHAVE_ATUAL, ESTILOS_PRONTOS[0].id); }
export function definirEstiloAtual(id) { gravarLocal(CHAVE_ATUAL, id); }

export function salvarEstiloUsuario(estilo) {
  const lista = estilosUsuario();
  if (ESTILOS_PRONTOS.some((p) => p.nome.toLowerCase() === estilo.nome.toLowerCase()))
    throw new Error("Esse é o nome de um estilo pronto. Escolha outro nome.");
  const existente = lista.findIndex((e) => e.nome.toLowerCase() === estilo.nome.toLowerCase());
  const novo = { ...base, ...estilo, pronto: false, id: existente >= 0 ? lista[existente].id : "u-" + Date.now() };
  if (existente >= 0) lista[existente] = novo; else lista.push(novo);
  gravarLocal(CHAVE, lista);
  return novo;
}

export function excluirEstiloUsuario(id) {
  gravarLocal(CHAVE, estilosUsuario().filter((e) => e.id !== id));
}

export function exportarEstilos() {
  return JSON.stringify({ gfpaEstilos: 1, estilos: estilosUsuario() }, null, 2);
}

export function importarEstilos(texto) {
  const obj = JSON.parse(texto);
  const lista = Array.isArray(obj) ? obj : obj.estilos;
  if (!Array.isArray(lista)) throw new Error("Arquivo de estilos inválido.");
  let n = 0;
  for (const e of lista) {
    if (!e || !e.nome || !e.total) continue;
    salvarEstiloUsuario(e); n++;
  }
  return n;
}
