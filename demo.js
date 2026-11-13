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
  { tipo: "COMP", titulo: "EBITDA, R$ mi", valores: [["Ano", "EBITDA"], [2023, 13.7], [2024, 15.0]] },
];

// Texto exibido de cada célula (o Excel devolve range.text formatado)
export const textos = (valores) => valores.map((l) => l.map((x) => (x === null ? "" : String(x))));
