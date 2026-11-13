// =============================================================================
// Gráficos FP&A — painel lateral (Office Add-in)
// =============================================================================
import { TIPOS, lerDados, montarGrafico, anotacoesPadrao, CENARIOS_PADRAO, aparencias } from "./engine.js?v=1.1.1";
import { renderSVG } from "./render-svg.js?v=1.1.1";
import { desenharNoExcel, listarGraficosTodas, descreverErro } from "./render-excel.js?v=1.1.1";
import {
  ESTILOS_PRONTOS, todosEstilos, estiloPorId, estiloAtualId, definirEstiloAtual, salvarEstiloUsuario,
  excluirEstiloUsuario, exportarEstilos, importarEstilos,
} from "./styles.js?v=1.1.1";
import { parseA1, montarA1, intersecta } from "./a1.js?v=1.1.1";
import { DEMOS, EXEMPLOS, textos as textosDemo } from "./demo.js?v=1.1.1";

const VERSAO = "1.1.1";
const $ = (id) => document.getElementById(id);
const clone = (o) => JSON.parse(JSON.stringify(o));

const st = {
  noExcel: false,
  fonte: null,          // { sheet, address, tabela, autoLinhas }
  valores: null, textos: null, dados: null,
  spec: { tipo: "IBCS", titulo: "", corte: 0, W: TIPOS.IBCS.W, H: TIPOS.IBCS.H, anot: anotacoesPadrao("IBCS"), modo: null },
  editando: null,       // item da lista de gráficos sendo editado
  estiloGrafico: null,  // estilo gravado no gráfico (quando não existe mais na lista)
  graficos: [], abas: {},
  pendentes: [], timer: null, ocupado: false,
};

// ----------------------------------------------------------------- mensagens
function msg(texto, tipo = "") {
  const el = $("msg");
  if (!texto) { el.hidden = true; return; }
  el.hidden = false; el.className = "msg " + tipo; el.textContent = texto;
}
async function seguro(fn) {
  try { msg(""); await fn(); }
  catch (e) { console.error(e); msg(descreverErro(e), "erro"); }
}

// ----------------------------------------------------------------- estilos
function estiloAtual() {
  const v = $("estilo").value;
  if (v === "__grafico" && st.estiloGrafico) return st.estiloGrafico;
  return estiloPorId(v);
}
function popularEstilos(selecionar) {
  const sel = $("estilo");
  sel.innerHTML = "";
  const lista = todosEstilos();
  const grp = (rotulo, itens) => {
    if (!itens.length) return;
    const og = document.createElement("optgroup"); og.label = rotulo;
    itens.forEach((e) => og.appendChild(new Option(e.nome, e.id)));
    sel.appendChild(og);
  };
  grp("Prontos", lista.filter((e) => e.pronto));
  grp("Meus estilos", lista.filter((e) => !e.pronto));
  if (st.estiloGrafico && !lista.some((e) => e.id === st.estiloGrafico.id))
    sel.appendChild(new Option(`(estilo do gráfico: ${st.estiloGrafico.nome})`, "__grafico"));
  const alvo = selecionar ?? estiloAtualId();
  sel.value = [...sel.options].some((o) => o.value === alvo) ? alvo : lista[0].id;
}

function abrirEditor() {
  const e = estiloAtual();
  $("editorEstilo").hidden = false;
  $("estNome").value = e.pronto ? `${e.nome} (meu)` : e.nome;
  $("estTotal").value = e.total; $("estDelta").value = e.delta;
  $("estPos").value = e.pos; $("estNeg").value = e.neg;
  $("estPorSinal").checked = !!e.porSinal;
  $("estPilha").innerHTML = e.pilha.map((c, i) => `<input type="color" value="${c}" title="Segmento ${i + 1}">`).join("");
  $("estColunas").innerHTML = e.colunas.map((c, i) => `<input type="color" value="${c}" title="Série ${i + 1}">`).join("");
  const cen = e.cenarios ?? ESTILOS_PRONTOS.find((x) => x.cenarios).cenarios;
  $("estCenarios").checked = !!e.cenarios;
  $("cenReal").value = cen.real.fill;
  $("cenOrcF").value = cen.orc.fill; $("cenOrcL").value = cen.orc.listra ?? cen.orc.fill;
  $("cenFc1F").value = cen.fc[0].fill; $("cenFc1L").value = cen.fc[0].listra ?? cen.fc[0].fill;
  const fc2 = cen.fc[1] ?? cen.fc[0];
  $("cenFc2F").value = fc2.fill; $("cenFc2L").value = fc2.listra ?? fc2.fill;
  $("estCenariosBox").classList.toggle("desligada", !e.cenarios);
}
function estiloDoEditor() {
  return {
    nome: $("estNome").value.trim(),
    total: $("estTotal").value, delta: $("estDelta").value, pos: $("estPos").value, neg: $("estNeg").value,
    porSinal: $("estPorSinal").checked,
    pilha: [...$("estPilha").querySelectorAll("input")].map((i) => i.value.toUpperCase()),
    colunas: [...$("estColunas").querySelectorAll("input")].map((i) => i.value.toUpperCase()),
    cenarios: $("estCenarios").checked ? {
      real: { fill: $("cenReal").value },
      orc: { fill: $("cenOrcF").value, listra: $("cenOrcL").value },
      fc: [{ fill: $("cenFc1F").value, listra: $("cenFc1L").value }, { fill: $("cenFc2F").value, listra: $("cenFc2L").value }],
    } : undefined,
  };
}



// ----------------------------------------------------------------- modo de desenho
const CHAVE_MODO = "gfpa.modoDesenho";
function modoPadrao() { try { return localStorage.getItem(CHAVE_MODO) || "imagem"; } catch { return "imagem"; } }
function lembrarModo(m) { try { localStorage.setItem(CHAVE_MODO, m); } catch { /* sem armazenamento */ } }

// ----------------------------------------------------------------- cores das barras
const MODOS_BARRA = [
  ["auto", "Automática"], ["real", "Como Real"], ["orc", "Como Orçamento"],
  ["fc1", "Como Forecast 1"], ["fc2", "Como Forecast 2"], ["cor", "Cor personalizada"],
];
const escHtml = (t) => String(t).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/"/g, "&quot;");

// Cor que a barra vai ter de fato (para o quadradinho ao lado do nome)
function corEfetiva(rotulo, esc) {
  const e = estiloAtual();
  const ap = aparencias(e, [rotulo], esc && esc.modo !== "auto" ? { [rotulo]: esc } : {})[0];
  return ap ? ap.fill : e.total;
}

function desenharCoresBarras() {
  const box = $("coresBarras");
  if (!st.dados) { box.innerHTML = '<p class="muted">Selecione os dados primeiro.</p>'; return; }
  const escolhas = st.spec.barras ?? {};
  const padraoCor = estiloAtual().total;
  box.innerHTML = st.dados.rot.map((r, i) => {
    const esc = escolhas[r] ?? { modo: "auto" };
    const fill = esc.fill ?? padraoCor, listra = esc.listra ?? "#FFFFFF";
    const opts = MODOS_BARRA.map(([v, t]) => `<option value="${v}"${v === esc.modo ? " selected" : ""}>${t}</option>`).join("");
    return `<div class="cor-barra" data-rot="${escHtml(r)}">
      <div class="topo"><span class="rot" title="${escHtml(r)}"><span class="amostra" style="background:${corEfetiva(r, { ...esc, fill })}"></span>${i + 1}. ${escHtml(r)}</span>
      <select class="modo">${opts}</select></div>
      <div class="pers"${esc.modo === "cor" ? "" : " hidden"}>
        <label>Cor <input type="color" class="fill" value="${fill}"></label>
        <label><input type="checkbox" class="listrado"${esc.listrado ? " checked" : ""}> Listras</label>
        <label>Cor da listra <input type="color" class="listra" value="${listra}"${esc.listrado ? "" : " disabled"}></label>
      </div></div>`;
  }).join("");
}

function lerCoresBarras() {
  const out = {};
  document.querySelectorAll("#coresBarras .cor-barra").forEach((el) => {
    const modo = el.querySelector(".modo").value;
    const pers = el.querySelector(".pers");
    pers.hidden = modo !== "cor";
    const listrado = el.querySelector(".listrado").checked;
    el.querySelector(".listra").disabled = !listrado;
    const fill = el.querySelector(".fill").value.toUpperCase();
    el.querySelector(".amostra").style.background = corEfetiva(el.dataset.rot, { modo, fill });
    if (modo === "auto") return;
    out[el.dataset.rot] = modo === "cor"
      ? { modo, fill, listrado, listra: el.querySelector(".listra").value.toUpperCase() }
      : { modo };
  });
  st.spec.barras = out;
}

// ----------------------------------------------------------------- formulário
function opcoesBarras(select, valor) {
  const rot = st.dados ? st.dados.rot : [];
  select.innerHTML = rot.map((r, i) => `<option value="${i + 1}">${i + 1}. ${r.replace(/</g, "&lt;")}</option>`).join("");
  if (rot.length) select.value = String(valor < 1 || valor > rot.length ? rot.length : valor);
}

function escreverForm() {
  const s = st.spec;
  $("tipo").value = s.tipo;
  $("titulo").value = s.titulo ?? "";
  $("corte").value = s.corte ? s.corte : "";
  $("lblCorte").hidden = s.tipo !== "EMP";
  $("largura").value = s.W; $("altura").value = s.H;
  const modo = s.modo ?? (st.editando ? "formas" : modoPadrao());
  document.querySelectorAll("input[name=modoDesenho]").forEach((r) => (r.checked = r.value === modo));
  if (st.fonte) { $("endereco").value = `${st.fonte.sheet}!${st.fonte.address}`; $("autoLinhas").checked = st.fonte.autoLinhas !== false; }
  document.querySelectorAll(".anot[data-k]").forEach((el) => {
    const a = s.anot.v[+el.dataset.k];
    el.querySelector(".on").checked = a.on;
    opcoesBarras(el.querySelector(".de"), a.de);
    opcoesBarras(el.querySelector(".para"), a.para);
    el.querySelectorAll("input[type=radio]").forEach((r) => (r.checked = r.value === a.modo));
    el.querySelector(".balao").checked = a.balao;
    el.classList.toggle("desligada", !a.on);
  });
  desenharCoresBarras();
  $("cagrOn").checked = s.anot.cagr.on;
  opcoesBarras($("cagrDe"), s.anot.cagr.de);
  opcoesBarras($("cagrPara"), s.anot.cagr.para);
  $("anotCagr").classList.toggle("desligada", !s.anot.cagr.on);
}

function lerForm() {
  const s = st.spec;
  s.titulo = $("titulo").value.trim();
  s.corte = Math.abs(parseFloat($("corte").value)) || 0;
  s.W = Math.max(150, +$("largura").value || TIPOS[s.tipo].W);
  s.H = Math.max(100, +$("altura").value || TIPOS[s.tipo].H);
  s.modo = document.querySelector("input[name=modoDesenho]:checked")?.value ?? "imagem";
  document.querySelectorAll(".anot[data-k]").forEach((el) => {
    const a = s.anot.v[+el.dataset.k];
    a.on = el.querySelector(".on").checked;
    a.de = +el.querySelector(".de").value || 1;
    a.para = +el.querySelector(".para").value || 0;
    a.modo = el.querySelector("input[type=radio]:checked")?.value ?? "P";
    a.balao = el.querySelector(".balao").checked;
    el.classList.toggle("desligada", !a.on);
  });
  lerCoresBarras();
  s.anot.cagr = { on: $("cagrOn").checked, de: +$("cagrDe").value || 1, para: +$("cagrPara").value || 0 };
  $("anotCagr").classList.toggle("desligada", !s.anot.cagr.on);
  if (st.fonte) st.fonte.autoLinhas = $("autoLinhas").checked;
}

function processarValores() {
  st.dados = null;
  if (!st.valores) return;
  st.dados = lerDados(st.valores, st.textos, TIPOS[st.spec.tipo].comTipo);
  const n = st.dados.nLin, m = st.dados.nCol;
  $("resumoDados").textContent = `${n} linha(s) · ${m} série(s)` + (st.fonte?.tabela ? ` · Tabela "${st.fonte.tabela}"` : "");
}

function preview() {
  const alvo = $("preview");
  if (!st.dados) return;
  try {
    const g = montarGrafico(st.dados, st.spec, estiloAtual());
    alvo.innerHTML = renderSVG(g);
  } catch (e) { alvo.innerHTML = `<p class="muted">${e.message}</p>`; }
}

// Sugere o tipo pela forma dos dados (coluna B só com T/D/F => cascata empilhada)
function sugerirTipo(valores) {
  const colB = valores.slice(1).map((l) => String(l[1] ?? "").trim().toUpperCase()).filter(Boolean);
  if (colB.length && colB.every((x) => ["T", "D", "F"].includes(x))) return "EMP";
  return null;
}

function trocarTipo(tipo, manterAnot = false) {
  st.spec.tipo = tipo;
  st.spec.W = TIPOS[tipo].W; st.spec.H = TIPOS[tipo].H;
  if (!manterAnot) st.spec.anot = anotacoesPadrao(tipo);
  try { processarValores(); } catch (e) { msg(e.message, "erro"); }
  escreverForm(); preview();
}

// ----------------------------------------------------------------- Excel: leitura
async function lerFonte(ctx, fonte) {
  const ws = ctx.workbook.worksheets.getItem(fonte.sheet);
  let r = null;
  if (fonte.tabela) {
    const t = ctx.workbook.tables.getItemOrNullObject(fonte.tabela);
    await ctx.sync();
    if (!t.isNullObject) r = t.getRange();
  }
  if (!r) {
    const a = parseA1(fonte.address);
    let r2 = a.r2;
    if (fonte.autoLinhas !== false) {
      // segue a 1ª coluna para baixo até a primeira célula vazia (inclui/exclui linhas)
      const jan = ws.getRange(montarA1(a.r1, a.c1, Math.min(a.r2 + 500, 1048576), a.c1));
      jan.load("values");
      await ctx.sync();
      let ultimo = 0;
      for (let i = 1; i < jan.values.length; i++) {
        const x = jan.values[i][0];
        if (x === "" || x === null) break;
        ultimo = i;
      }
      if (ultimo > 0) r2 = a.r1 + ultimo;
    }
    r = ws.getRange(montarA1(a.r1, a.c1, r2, a.c2));
  }
  r.load("values,text,address,left,top,width");
  await ctx.sync();
  fonte.address = r.address.split("!").pop().replace(/\$/g, "");
  return r;
}

async function usarSelecao() {
  if (!st.noExcel) { msg("Disponível dentro do Excel."); return; }
  await Excel.run(async (ctx) => {
    let r = ctx.workbook.getSelectedRange();
    r.load("address,cellCount");
    r.worksheet.load("name");
    await ctx.sync();
    const aba = r.worksheet.name;
    if (r.cellCount === 1) {
      try { const reg = r.getSurroundingRegion(); reg.load("address"); await ctx.sync(); r = reg; } catch { /* API indisponível */ }
    }
    const inter = r.getIntersectionOrNullObject(ctx.workbook.worksheets.getItem(aba).getUsedRange());
    inter.load("address");
    await ctx.sync();
    if (inter.isNullObject) throw new Error("A seleção está vazia.");
    // se a seleção é exatamente uma Tabela do Excel, o gráfico passa a seguir a tabela
    let tabela = null;
    try {
      const tbs = inter.getTables(false);
      tbs.load("items/name");
      await ctx.sync();
      if (tbs.items.length === 1) {
        const tr = tbs.items[0].getRange(); tr.load("address"); await ctx.sync();
        if (tr.address === inter.address) tabela = tbs.items[0].name;
      }
    } catch { /* versão sem getTables */ }

    st.fonte = { sheet: aba, address: inter.address.split("!").pop().replace(/\$/g, ""), tabela, autoLinhas: $("autoLinhas").checked };
    const rr = await lerFonte(ctx, st.fonte);
    st.valores = rr.values; st.textos = rr.text;
  });
  if (!st.editando) {
    const sug = sugerirTipo(st.valores);
    if (sug && sug !== st.spec.tipo) { trocarTipo(sug); return; }
  }
  processarValores();
  escreverForm(); preview();
  if (st.editando) msg("Novo intervalo carregado. Clique em “Aplicar” para atualizar o gráfico.");
}

// ----------------------------------------------------------------- Excel: escrita
function montarMeta(fonte, spec, estilo) {
  return { gfpa: 1, versao: VERSAO, fonte: clone(fonte), spec: clone(spec), estilo: clone(estilo) };
}

// Redesenha um gráfico existente (mesma posição) e apaga o antigo
async function redesenhar(ctx, item, meta) {
  const r = await lerFonte(ctx, meta.fonte);
  const dados = lerDados(r.values, r.text, TIPOS[meta.spec.tipo].comTipo);
  const g = montarGrafico(dados, meta.spec, meta.estilo);
  const ws = ctx.workbook.worksheets.getItem(item.sheet);
  const novoId = await desenharNoExcel(ctx, ws, g, item.left + 8, item.top + 8, meta);
  ws.shapes.getItem(item.id).delete();
  await ctx.sync();
  return novoId;
}

async function inserir() {
  if (!st.noExcel) { msg("Disponível dentro do Excel. Aqui é só a pré-visualização."); return; }
  if (!st.fonte) throw new Error("Selecione os dados e clique em “Usar seleção”.");
  lerForm();
  let idNovo;
  await Excel.run(async (ctx) => {
    const r = await lerFonte(ctx, st.fonte);
    const dados = lerDados(r.values, r.text, TIPOS[st.spec.tipo].comTipo);
    const estilo = estiloAtual();
    const g = montarGrafico(dados, st.spec, estilo);
    const ws = ctx.workbook.worksheets.getItem(st.fonte.sheet);
    idNovo = await desenharNoExcel(ctx, ws, g, r.left + r.width + 30, r.top + 8, montarMeta(st.fonte, st.spec, estilo));
  });
  await recarregarLista(idNovo);
  msg("Gráfico inserido. Ele está na lista acima para edição.", "ok");
}

async function aplicar() {
  if (!st.editando) return;
  lerForm();
  let idNovo;
  await Excel.run(async (ctx) => {
    idNovo = await redesenhar(ctx, st.editando, montarMeta(st.fonte, st.spec, estiloAtual()));
  });
  await recarregarLista(idNovo);
  msg("Gráfico atualizado.", "ok");
}

async function recarregarLista(selecionarId) {
  if (!st.noExcel) return;
  await Excel.run(async (ctx) => {
    const r = await listarGraficosTodas(ctx); // 2 idas e voltas, com qualquer número de abas
    st.graficos = r.graficos; st.abas = r.abas;
  });
  const sel = $("listaGraficos");
  sel.innerHTML = '<option value="">— novo gráfico —</option>';
  st.graficos.forEach((g) => {
    const t = g.meta.spec.titulo ? ` – ${g.meta.spec.titulo}` : "";
    sel.appendChild(new Option(`${g.sheet} · ${TIPOS[g.meta.spec.tipo]?.nome ?? g.meta.spec.tipo}${t}`, g.id));
  });
  if (selecionarId) { sel.value = selecionarId; await selecionarGrafico(selecionarId, true); }
  else if (st.editando) {
    const ainda = st.graficos.find((g) => g.id === st.editando.id);
    if (ainda) { sel.value = ainda.id; st.editando = ainda; } else { st.editando = null; modoEdicao(); }
  }
}

function modoEdicao() {
  const ed = !!st.editando;
  $("btnAplicar").hidden = !ed;
  $("btnInserir").textContent = ed ? "Inserir como novo gráfico" : "Inserir gráfico";
}

async function selecionarGrafico(id, silencioso = false) {
  const item = st.graficos.find((g) => g.id === id);
  st.editando = item ?? null;
  modoEdicao();
  if (!item) { st.estiloGrafico = null; popularEstilos(); return; }
  st.fonte = clone(item.meta.fonte);
  st.spec = clone(item.meta.spec);
  st.estiloGrafico = item.meta.estilo;
  popularEstilos(item.meta.estilo.id);
  if ($("estilo").value !== item.meta.estilo.id) $("estilo").value = "__grafico";
  await Excel.run(async (ctx) => {
    const r = await lerFonte(ctx, st.fonte);
    st.valores = r.values; st.textos = r.text;
  });
  processarValores();
  escreverForm(); preview();
  if (!silencioso) msg("Editando o gráfico. Mude o que quiser e clique em “Aplicar”.");
}

async function atualizarTodos() {
  await recarregarLista();
  let n = 0;
  for (const item of st.graficos) {
    await Excel.run(async (ctx) => { await redesenhar(ctx, item, item.meta); });
    n++;
  }
  await recarregarLista();
  msg(`${n} gráfico(s) atualizado(s).`, "ok");
}


// ----------------------------------------------------------------- planilha de exemplos
async function gerarExemplos() {
  if (!st.noExcel) { msg("Disponível dentro do Excel. Aqui use a lista de exemplos acima."); return; }
  let total = 0, nomeAba = "";
  await Excel.run(async (ctx) => {
    const wss = ctx.workbook.worksheets;
    wss.load("items/name");
    await ctx.sync();
    const existentes = new Set(wss.items.map((w) => w.name));
    nomeAba = "Exemplos Gráficos FP&A";
    for (let k = 2; existentes.has(nomeAba); k++) nomeAba = `Exemplos Gráficos FP&A (${k})`;
    const ws = wss.add(nomeAba);
    ws.activate();
    try { ws.showGridlines = false; } catch { /* versão antiga */ }
    ws.getRange("A:A").format.columnWidth = 170;
    const cab = ws.getRange("A1");
    cab.values = [["Exemplos do add-in Gráficos FP&A – mude os números e veja os gráficos se atualizarem. Para editar, escolha o gráfico na lista do painel."]];
    cab.format.font.bold = true;
    const colGraf = ws.getRange("H1");
    colGraf.load("left");
    await ctx.sync();

    let linha = 3;
    for (const ex of EXEMPLOS) {
      const demo = DEMOS[ex.demo];
      const nl = demo.valores.length, nc = Math.max(...demo.valores.map((l) => l.length));
      const matriz = demo.valores.map((l) => [...l, ...Array(nc - l.length).fill("")].map((x) => (x === null ? "" : x)));
      const desc = ws.getRange(`A${linha}`);
      desc.values = [[ex.descricao]];
      desc.format.font.bold = true;
      desc.format.font.color = "#1A53FF";
      const r0 = linha + 1;
      const endereco = montarA1(r0, 1, r0 + nl - 1, nc);
      ws.getRange(montarA1(r0, 1, r0 + nl - 1, 1)).numberFormat = matriz.map(() => ["@"]); // rótulos como texto (6+6, 2024…)
      const bloco = ws.getRange(endereco);
      bloco.values = matriz.map((l) => l.map((x, c) => (c === 0 ? String(x) : x)));
      ws.getRange(montarA1(r0, 1, r0, nc)).format.font.bold = true;
      if (nl > 1 && nc > 1)
        ws.getRange(montarA1(r0 + 1, 2, r0 + nl - 1, nc)).numberFormat = matriz.slice(1).map((l) => l.slice(1).map(() => "#,##0.0;-#,##0.0"));
      desc.load("top");
      await ctx.sync();

      const fonte = { sheet: nomeAba, address: endereco, tabela: null, autoLinhas: true };
      const r = await lerFonte(ctx, fonte);
      const dados = lerDados(r.values, r.text, TIPOS[demo.tipo].comTipo);
      const estilo = estiloPorId(ex.estilo);
      const spec = {
        tipo: demo.tipo, titulo: demo.titulo, corte: ex.corte ?? demo.corte ?? 0,
        W: TIPOS[demo.tipo].W, H: TIPOS[demo.tipo].H,
        anot: ex.anot ? clone(ex.anot) : anotacoesPadrao(demo.tipo), barras: ex.barras ?? {},
        modo: modoPadrao(),
      };
      const g = montarGrafico(dados, spec, estilo);
      await desenharNoExcel(ctx, ws, g, colGraf.left, desc.top + 8, montarMeta(fonte, spec, estilo));
      total++;
      // próximo bloco abaixo do gráfico (linhas de ~15 pt) e sempre depois dos dados
      linha = Math.max(r0 + nl + 2, linha + Math.ceil((g.altura + 40) / 15));
    }
  });
  await recarregarLista();
  msg(`${total} exemplos criados na aba “${nomeAba}”. Escolha qualquer um na lista acima para ver como foi configurado.`, "ok");
}

// ----------------------------------------------------------------- atualização automática
function aoMudar(ev) {
  if (!$("autoAtualizar").checked) return;
  st.pendentes.push({ abaId: ev.worksheetId, address: ev.address });
  clearTimeout(st.timer);
  st.timer = setTimeout(() => seguro(processarMudancas), 900);
}

async function processarMudancas() {
  if (st.ocupado) { st.timer = setTimeout(() => seguro(processarMudancas), 900); return; }
  st.ocupado = true;
  try {
    const mud = st.pendentes.splice(0);
    await recarregarLista();
    const afetados = st.graficos.filter((g) => {
      const f = g.meta.fonte;
      let a;
      try { a = parseA1(f.address); } catch { return false; }
      const zona = { ...a, r2: f.autoLinhas !== false || f.tabela ? a.r2 + 500 : a.r2 };
      return mud.some((m) => st.abas[m.abaId] === f.sheet && intersecta(parseA1(m.address), zona));
    });
    if (!afetados.length) return;
    let idEditado = null;
    for (const item of afetados) {
      await Excel.run(async (ctx) => {
        const novo = await redesenhar(ctx, item, item.meta);
        if (st.editando && st.editando.id === item.id) idEditado = novo;
      });
    }
    await recarregarLista(idEditado ?? undefined);
    msg(`${afetados.length} gráfico(s) atualizado(s) automaticamente.`, "ok");
  } finally { st.ocupado = false; }
}

async function registrarEventos() {
  await Excel.run(async (ctx) => {
    ctx.workbook.worksheets.onChanged.add(aoMudar);
    await ctx.sync();
  });
}

// ----------------------------------------------------------------- demonstração (fora do Excel)
function carregarDemo(i) {
  const d = DEMOS[i];
  st.valores = d.valores; st.textos = textosDemo(d.valores);
  st.fonte = null;
  st.spec = { tipo: d.tipo, titulo: d.titulo, corte: d.corte ?? 0, W: TIPOS[d.tipo].W, H: TIPOS[d.tipo].H, anot: anotacoesPadrao(d.tipo) };
  if (d.semCagr) st.spec.anot.cagr.on = false;
  if (d.estilo) $("estilo").value = d.estilo;
  processarValores();
  $("endereco").value = "(exemplo)";
  escreverForm(); preview();
}



// ----------------------------------------------------------------- abrir com o arquivo
// Grava na pasta de trabalho a configuração padrão do Office que reabre este painel
// sempre que o arquivo for aberto (útil para relatórios recorrentes com gráficos).
function lerAutoAbrir() {
  try { $("autoAbrir").checked = !!Office.context.document.settings.get("Office.AutoShowTaskpaneWithDocument"); } catch { /* fora do Excel */ }
}
function gravarAutoAbrir(ligar) {
  return new Promise((ok, falha) => {
    const cfg = Office.context.document.settings;
    if (ligar) cfg.set("Office.AutoShowTaskpaneWithDocument", true); else cfg.remove("Office.AutoShowTaskpaneWithDocument");
    cfg.saveAsync((r) => (r.status === Office.AsyncResultStatus.Succeeded ? ok() : falha(r.error)));
  });
}

// ----------------------------------------------------------------- aviso de nova versão
// Compara a versão carregada com a publicada (versao.json, sem cache). Se houver uma
// mais nova, mostra um aviso com botão que recarrega o painel já na versão nova.
const numVersao = (v) => v.split(".").map(Number).reduce((a, x) => a * 1000 + x, 0);
async function verificarAtualizacao() {
  try {
    const r = await fetch(`versao.json?t=${Date.now()}`, { cache: "no-store" });
    if (!r.ok) return;
    const { versao } = await r.json();
    if (numVersao(versao) <= numVersao(VERSAO)) return;
    const el = $("avisoVersao");
    el.hidden = false;
    el.querySelector("span").textContent = `Nova versão disponível (v${versao}).`;
    el.querySelector("button").onclick = () => {
      const u = new URL(location.href);
      u.searchParams.set("v", versao);
      location.replace(u.toString());
    };
  } catch { /* sem rede ou rodando localmente */ }
}

// ----------------------------------------------------------------- inicialização
function ligarEventos() {
  $("btnSelecao").onclick = () => seguro(usarSelecao);
  $("btnInserir").onclick = () => seguro(inserir);
  $("btnAplicar").onclick = () => seguro(aplicar);
  $("btnRecarregarLista").onclick = () => seguro(() => recarregarLista());
  $("btnAtualizarTodos").onclick = () => seguro(atualizarTodos);
  $("btnExemplos").onclick = () => seguro(gerarExemplos);
  $("listaGraficos").onchange = (e) => seguro(() => selecionarGrafico(e.target.value));
  $("tipo").onchange = (e) => trocarTipo(e.target.value);
  $("estilo").onchange = () => { if ($("estilo").value !== "__grafico") definirEstiloAtual($("estilo").value); lerCoresBarras(); preview(); };

  // qualquer mudança no formulário atualiza a pré-visualização
  const reagir = () => { lerForm(); preview(); };
  ["titulo", "corte", "largura", "altura", "cagrOn", "cagrDe", "cagrPara", "autoLinhas"].forEach((id) => $(id).addEventListener("input", reagir));
  document.querySelectorAll(".anot[data-k]").forEach((el) => {
    el.addEventListener("input", reagir);
    el.addEventListener("change", (e) => {
      // escolher De/Para/Tipo liga a variação automaticamente
      if (e.target.matches(".de,.para,input[type=radio]")) { el.querySelector(".on").checked = true; reagir(); }
    });
  });
  $("cagrDe").addEventListener("change", () => { $("cagrOn").checked = true; reagir(); });
  $("cagrPara").addEventListener("change", () => { $("cagrOn").checked = true; reagir(); });

  document.querySelectorAll("input[name=modoDesenho]").forEach((r) => r.addEventListener("change", () => { lembrarModo(r.value); reagir(); }));
  $("coresBarras").addEventListener("input", reagir);
  $("coresBarras").addEventListener("change", reagir);
  $("btnCoresAuto").onclick = () => { st.spec.barras = {}; desenharCoresBarras(); preview(); };

  // editor de estilos
  $("btnEditarEstilo").onclick = abrirEditor;
  $("btnFecharEditor").onclick = () => ($("editorEstilo").hidden = true);
  $("editorEstilo").addEventListener("input", () => {
    // pré-visualiza o estilo em edição sem salvar
    const tmp = { ...estiloAtual(), ...estiloDoEditor() };
    if (st.dados) $("preview").innerHTML = renderSVG(montarGrafico(st.dados, st.spec, tmp));
  });
  $("btnSalvarEstilo").onclick = () => seguro(async () => {
    const novo = salvarEstiloUsuario(estiloDoEditor());
    definirEstiloAtual(novo.id); popularEstilos(novo.id); preview();
    msg(`Estilo “${novo.nome}” salvo e selecionado.`, "ok");
  });
  $("btnExcluirEstilo").onclick = () => seguro(async () => {
    const e = estiloAtual();
    if (e.pronto || $("estilo").value === "__grafico") throw new Error("Selecione na lista um estilo criado por você.");
    excluirEstiloUsuario(e.id); popularEstilos(); preview();
    $("editorEstilo").hidden = true;
    msg(`Estilo “${e.nome}” excluído.`, "ok");
  });
  $("btnExportar").onclick = () => {
    const blob = new Blob([exportarEstilos()], { type: "application/json" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob); a.download = "estilos-graficos-fpa.json"; a.click();
    msg("Arquivo de estilos gerado. Salve-o numa pasta do OneDrive/SharePoint para o time importar.", "ok");
  };
  $("arqImportar").onchange = (ev) => seguro(async () => {
    const arq = ev.target.files[0]; if (!arq) return;
    const n = importarEstilos(await arq.text());
    popularEstilos(); preview();
    msg(`${n} estilo(s) importado(s).`, "ok");
    ev.target.value = "";
  });
}

async function iniciar(noExcel) {
  st.noExcel = noExcel;
  $("versao").textContent = "v" + VERSAO;
  verificarAtualizacao();
  popularEstilos();
  ligarEventos();
  escreverForm();
  modoEdicao();
  if (!noExcel) {
    $("avisoDemo").hidden = false;
    const sel = $("demoSel");
    DEMOS.forEach((d, i) => sel.appendChild(new Option(`${TIPOS[d.tipo].nome} – ${d.titulo}`, i)));
    sel.onchange = () => carregarDemo(+sel.value);
    carregarDemo(0);
    return;
  }
  lerAutoAbrir();
  $("autoAbrir").onchange = (e) => seguro(async () => {
    await gravarAutoAbrir(e.target.checked);
    msg(e.target.checked ? "Pronto: salve o arquivo (Ctrl+S) e o painel abrirá junto com ele." : "O painel não abrirá mais automaticamente com este arquivo.", "ok");
  });
  await seguro(async () => { await recarregarLista(); await registrarEventos(); });
}

if (typeof Office !== "undefined" && Office.onReady) {
  Office.onReady((info) => iniciar(info.host === Office.HostType.Excel));
} else {
  iniciar(false);
}
