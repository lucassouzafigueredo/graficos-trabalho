// Conjuntos de dados de exemplo (mesmos do botão "Demonstração" do VBA).
// Formato: matriz como a que o Excel devolve em range.values.
export const DEMOS = [
  { tipo: "IBCS", titulo: "Receita – Forecast vs Real", valores: [
    ["Período", "Receita"], ["FC", 2250000], ["Q1", 98400], ["Q2", 125600], ["Q3", -130700], ["Q4", 143000], ["AC", 2486300]] },
  { tipo: "IBCSM", titulo: "Receita por unidade de negócio", valores: [
    ["Período", "Speakers", "Hair Care BU", "Tablets", "Smartphones", "Accessories"],
    ["FC", 912800, 877400, 756800, 666600, 351800], ["Q1", -52800, 13200, 31700, 30100, 63200],
    ["Q2", -24500, -67800, -44900, -5100, -88300], ["Q3", -12300, -51400, -15000, 15300, -26200],
    ["Q4", 3900, 55400, 91200, 21900, 7900], ["AC", "", "", "", "", ""]] },
  { tipo: "EMP", titulo: "Escopo do CSC, FTE", valores: [
    ["Etapa", "Tipo", "Contabilidade", "Atendimento", "Compras", "Suporte a vendas"],
    ["Escopo benchmark", "T", 89, 45, 34, 25], ["Padronização de processos", "D", -21, "", "", ""],
    ["Importância estratégica", "D", -21, "", "", ""], ["Limitações de sistema", "D", -22, "", "", ""],
    ["Escopo aprovado", "T", 66, 31, 19, 13], ["Transferido p/ dados", "F", -45, -23, -19, -8],
    ["A transferir", "T", 21, 8, 5, 0]] },
  { tipo: "EMP", titulo: "DRE resumida, R$ mi (eixo cortado em 120)", corte: 120, valores: [
    ["Etapa", "Tipo", "Valor"], ["Receita bruta", "T", 167], ["Deduções", "D", -62],
    ["Receita líquida", "T", 105], ["Custos", "D", -25], ["Lucro bruto", "T", 80]] },
  { tipo: "EMP", titulo: "PDD – BGT vs FC 9+3", valores: [
    ["Etapa", "Tipo", "Valor"], ["BGT Operacional", "T", -44.34], ["Retorno Carteira Valora", "D", -113],
    ["Ajuste % da Carteira Original", "D", -21.62], ["Incremento da Carteira", "D", -35.67],
    ["Ajuste % Do Aumento da Carteira", "D", -15.29], ["FC 9+3 Operacional", "T", -229.93]] },
  { tipo: "COL", titulo: "Receita, R$ bi", valores: [
    ["Ano", "Serviços", "Produtos"], [2019, 4, 4], [2020, 7, 6], [2021, 6, 6], [2022, 9, 7], [2023, 11, 8], [2024, 15, 11]] },
  { tipo: "COL", titulo: "Receita – Orçamento x Forecast", estilo: "orc-fc", semCagr: true, valores: [
    ["Cenário", "Receita"], ["LY", 1180], ["BGT", 1057], ["6+6", 993], ["9+3", 989]] },
  { tipo: "COMP", titulo: "EBITDA, R$ mi", valores: [["Ano", "EBITDA"], [2023, 13.7], [2024, 15.0]] },
  { tipo: "TAB", titulo: "DRE – Real x Orçamento, R$ mil", estilo: "cinza", valores: [
    ["Conta", "BGT", "Real", "Comentário"],
    ["Receita bruta", 12500, 13100, "Volume 6% acima do plano no varejo, puxado pela linha nova."],
    ["Deduções", 2100, 2250, ""], ["Receita líquida", 10400, 10850, ""],
    ["CPV", 6200, 6650, "Frete e insumo X mais caros no 2º trimestre."], ["Lucro bruto", 4200, 4200, ""],
    ["Despesas comerciais", 1100, 1020, ""], ["Despesas administrativas", 900, 960, ""], ["EBITDA", 2200, 2220, ""],
    ["D&A", 400, 410, ""], ["Resultado financeiro", -300, -380, "Juros maiores pela antecipação de recebíveis."],
    ["IR/CS", 510, 480, ""], ["Lucro líquido", 990, 950, ""]] },
  { tipo: "TAB", titulo: "Receita por cliente – Real x BGT, R$ mil (Top 5)", estilo: "orc-fc", topN: 5, valores: [
    ["Cliente", "BGT", "Real"], ["Cliente A", 3200, 3550], ["Cliente B", 2800, 2610], ["Cliente C", 2100, 2390],
    ["Cliente D", 1500, 1480], ["Cliente E", 1300, 1020], ["Cliente F", 900, 940], ["Cliente G", 800, 805],
    ["Cliente H", 600, 720], ["Cliente I", 500, 470], ["Cliente J", 400, 395], ["Total", 14100, 14380]] },
  { tipo: "TAB", titulo: "DRE – Real x Orçamento x Forecast 9+3, R$ mil", estilo: "cinza", valores: [
    ["Conta", "BGT", "FC 9+3", "Real", "Comentário"],
    ["Receita bruta", 12500, 12900, 13100, "Linha nova acima do FC; o BGT não previa o lançamento."],
    ["Deduções", 2100, 2200, 2250, ""], ["Receita líquida", 10400, 10700, 10850, ""],
    ["CPV", 6200, 6550, 6650, "Insumo X: o FC já capturava a alta; o desvio residual é o frete."],
    ["Lucro bruto", 4200, 4150, 4200, ""],
    ["Despesas comerciais", 1100, 1050, 1020, ""], ["Despesas administrativas", 900, 940, 960, ""], ["EBITDA", 2200, 2160, 2220, ""],
    ["D&A", 400, 405, 410, ""], ["Resultado financeiro", -300, -350, -380, ""],
    ["IR/CS", 510, 470, 480, ""], ["Lucro líquido", 990, 935, 950, ""]] },
];

// Texto exibido de cada célula (o Excel devolve range.text formatado)
export const textos = (valores) => valores.map((l) => l.map((x) => (x === null ? "" : String(x))));

// Roteiro da "planilha de exemplos": cada item mostra uma possibilidade do add-in.
// anot: [variação 1, variação 2, cagr] em forma resumida; barras: cores por barra.
const v = (de, para, modo = "P", balao = false) => ({ on: true, de, para, modo, balao });
const off = { on: false, de: 1, para: 0, modo: "P", balao: false };
export const EXEMPLOS = [
  { demo: 0, estilo: "ibcs", descricao: "1. Cascata IBCS – FC hachurado, variações verde/vermelho e indicador % FC→AC" },
  { demo: 1, estilo: "ibcs", descricao: "2. Múltiplos IBCS – uma cascata por unidade, todas na mesma escala" },
  { demo: 2, estilo: "cons-azul", descricao: "3. Cascata empilhada – totais por área, variações em cinza e legenda à esquerda" },
  { demo: 3, estilo: "cons-navy", corte: 120, descricao: "4. Corte de eixo – barras grandes cortadas com // (Corte do eixo = 120)" },
  { demo: 4, estilo: "orc-fc", anot: { v: [v(1, 0, "B", true), off], cagr: { on: false, de: 1, para: 0 } },
    descricao: "5. Ponte BGT → FC – estilo Orçamento x Forecast e variação em valor e % com balão" },
  { demo: 5, estilo: "cons-azul", anot: { v: [v(5, 6, "P", true), off], cagr: { on: true, de: 1, para: 6 } },
    descricao: "6. Colunas empilhadas – linha de CAGR e variação do último ano com balão" },
  { demo: 6, estilo: "orc-fc", anot: { v: [v(2, 4, "A"), off], cagr: { on: false, de: 1, para: 0 } },
    barras: { LY: { modo: "cor", fill: "#7F7F7F", listrado: false } },
    descricao: "7. Orçamento x Forecast – cenário reconhecido pelo nome; LY com cor escolhida em Cores das barras" },
  { demo: 7, estilo: "cons-azul", descricao: "8. Comparação – seta e balão de variação entre dois períodos" },
  { demo: 8, estilo: "cinza", descricao: "9. Tabela de variações – DRE Real x BGT: custos invertidos (queda = verde), subtotais e comentários numerados" },
  { demo: 8, estilo: "cinza", cascata: true, comentarios: false, descricao: "10. Tabela com cascata – a coluna Real vira uma ponte que zera nos subtotais" },
  { demo: 9, estilo: "orc-fc", topN: 5, descricao: "11. Top N – só os 5 clientes com maior desvio; o resto agrupado em Outros" },
  { demo: 10, estilo: "cinza", descricao: "12. Várias referências – Real x BGT e Real x FC 9+3 lado a lado, na mesma escala" },
  { demo: 9, estilo: "orc-fc", ordem: "fav", titulo: "Receita por cliente – do mais favorável ao mais desfavorável, R$ mil", colunas: { ref: false, ac: true, abs: true, pct: true },
    descricao: "13. Ordenada – clientes do mais favorável ao mais desfavorável, sem a coluna de referência" },
];
