// Modelos de página: conjuntos de cartões prontos, com dados de exemplo na planilha.
// "PADRÃO BELA FP&A_fechamento": apresentação de fechamento LY | BGT | 6+6 | Actual.
// Cada cartão vira um gráfico normal (editável pelo painel); os comentários ficam
// numa coluna "Comentário" ao lado dos dados, então mudam todo mês direto na planilha.

const P = ["LY", "BGT", "6+6", "Actual"];
// monta a matriz: cabeçalho + uma linha por período; comentários entram na coluna de texto
function tabela(cab, colunas, comentarios = []) {
  const temCom = comentarios.length > 0;
  const linhas = [["Período", ...cab, ...(temCom ? ["Comentário"] : [])]];
  P.forEach((p, i) => linhas.push([p, ...colunas.map((c) => c[i]), ...(temCom ? [comentarios[i] ?? ""] : [])]));
  return linhas;
}
const pc = (...xs) => xs.map((x) => x / 100);

export const ESTILO_MODELOS = "bela-fechamento";

export const MODELOS = [
  // Números e comentários FICTÍCIOS, só para mostrar o layout. Substitua pelos seus na aba gerada.
  {
    nome: "PADRÃO BELA FP&A_fechamento_Consolidado", aba: "BELA_fechamento_Consolidado", subtitulo: "Consolidado",
    cartoes: [
      { tipo: "CARD", titulo: "Receita Líquida, R$M", valores: tabela(["Receita Líquida"], [[1450, 620, 590, 548]], [
        "**Unidade A**, R$-18,0M vs 6+6, postergação de entregas para o próximo mês.",
        "**Unidade B**, R$-24,0M vs 6+6, volume abaixo do previsto (exemplo fictício)"]) },
      { tipo: "CARD", titulo: "Margem Bruta, R$M", valores: tabela(["Margem Bruta", "Margem %"], [[132, 81, 95, 91], pc(9.1, 13.1, 16.1, 16.6)], [
        "**Unidade A:** melhor preço médio no trimestre",
        "**Unidade B**: redução de custo logístico (exemplo fictício)"]) },
      { tipo: "CARD", titulo: "SG&A & Provisão AR, R$M", valores: tabela(["SG&A", "Prov. AR", "Margem %"],
        [[-96, -71, -68, -63], [-18, -6, -14, -9], pc(-7.9, -12.4, -13.9, -13.1)], [
        "**SG&A**: vagas em aberto e menor gasto com viagens (exemplo fictício)",
        "**Prov. AR**: efeito de calendário, regulariza até o fim do ano."]) },
      { tipo: "CARD", titulo: "EBITDA, R$M", valores: tabela(["EBITDA", "Margem %"], [[18, 4, 13, 19], pc(1.2, 0.6, 2.2, 3.5)], [
        "**Destaque**, melhora puxada por margem e redução de despesas (exemplo fictício)"]) },
    ],
  },
  {
    nome: "PADRÃO BELA FP&A_fechamento_Insumos", aba: "BELA_fechamento_Insumos", subtitulo: "Unidade de Insumos – Operacional (exemplo)",
    cartoes: [
      { tipo: "CARD", titulo: "Receita Líquida, R$M", valores: tabela(["Receita Líquida"], [[640, 310, 284, 262]], [
        "**Insumos**, volume faturado abaixo do previsto por atraso nas entregas (exemplo fictício).",
        "**Mix na Receita (Atual):** Linha A 45% · Linha B 25% · Linha C 20% · Linha D 10%"]) },
      { tipo: "CARD", titulo: "Margem Bruta, R$M", valores: tabela(["Margem Bruta", "Margem %"], [[84, 46, 38, 35], pc(13.1, 14.8, 13.4, 13.4)], [
        "**Destaque**, menor volume e mudança de mix para produtos de menor margem (exemplo fictício)"]) },
      { tipo: "CARDM", titulo: "Receita Líquida por Mix (Resultado Comercial)", W: 520, porLinha: 2, valores: tabela(
        ["Linha A", "Margem Linha A", "Linha B", "Margem Linha B", "Linha C", "Margem Linha C", "Linha D", "Margem Linha D"],
        [[310, 140, 128, 118], pc(9.8, 8.6, 7.1, 7.4), [36, 41, 30, 26], pc(31.5, 28.0, 30.4, 29.8),
         [168, 72, 70, 66], pc(18.2, 15.9, 19.5, 20.1), [126, 57, 56, 52], pc(12.4, 16.8, 15.1, 15.9)]) },
    ],
  },
  {
    nome: "PADRÃO BELA FP&A_fechamento_Grãos", aba: "BELA_fechamento_Grãos", subtitulo: "Grãos (exemplo)",
    cartoes: [
      { tipo: "CARD", titulo: "Receita Líquida, R$M", valores: tabela(["Receita Líquida"], [[1210, 305, 288, 270]], [
        "**Destaque**, menor volume embarcado no mês; parte dos contratos foi renegociada (exemplo fictício)."]) },
      { tipo: "CARD", titulo: "Despesas Operacionais, R$M", valores: tabela(["Despesas Operacionais", "Margem %"], [[-42, -16, -14, -15], pc(-3.5, -5.2, -4.9, -5.6)]) },
      { tipo: "CARD", titulo: "Margem Bruta, R$M", valores: tabela(["Margem Bruta", "Margem %"], [[-28, 9, 22, 25], pc(-2.3, 3.0, 7.6, 9.3)], [
        "**Destaque**, ganho de marcação a mercado nas posições a fixar (exemplo fictício)",
        "Phasing: parte do resultado fica para os próximos meses",
        "**Importante:** valores ilustrativos, substitua pelos do fechamento"]) },
      { tipo: "CARDM", titulo: "Grãos Originado", W: 250, porLinha: 1, valores: tabela(["Produto A em kTon", "Produto B em kTon"], [[540, 96, 92, 90], [410, 62, 60, 35]]) },
    ],
  },
];
