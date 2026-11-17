// Modelos de página: conjuntos de cartões prontos, com dados de exemplo na planilha.
// "PADRÃO BELA FP&A_fechamento": apresentação de fechamento LY | BGT | 6+6 | Actual.
// Cada cartão vira um gráfico normal (editável pelo painel); os comentários ficam
// numa coluna "Comentário" ao lado dos dados, então mudam todo mês direto na planilha.
// Números e comentários FICTÍCIOS, só para mostrar o layout. Idiomas: "pt" | "en".

const P = ["LY", "BGT", "6+6", "Actual"];
const pc = (...xs) => xs.map((x) => x / 100);

export const ESTILO_MODELOS = "bela-fechamento";

export function modelosNoIdioma(idioma = "pt") {
  const en = idioma === "en";
  const L = (pt, us) => (en ? us : pt);
  // cabeçalho + uma linha por período; comentários entram na coluna de texto
  function tabela(cab, colunas, comentarios = []) {
    const temCom = comentarios.length > 0;
    const linhas = [[L("Período", "Period"), ...cab, ...(temCom ? [L("Comentário", "Comment")] : [])]];
    P.forEach((p, i) => linhas.push([p, ...colunas.map((c) => c[i]), ...(temCom ? [comentarios[i] ?? ""] : [])]));
    return linhas;
  }
  const RL = L("Receita Líquida", "Net Revenue"), MB = L("Margem Bruta", "Gross Margin"), MG = L("Margem %", "Margin %");
  const ex = L("exemplo fictício", "illustrative example");
  return [
    {
      nome: "PADRÃO BELA FP&A_fechamento_Consolidado", aba: L("BELA_fechamento_Consolidado", "BELA_closing_Consolidated"),
      subtitulo: L("Consolidado", "Consolidated"),
      cartoes: [
        { tipo: "CARD", titulo: `${RL}, R$M`, valores: tabela([RL], [[1450, 620, 590, 548]], [
          L(`**Unidade A**, R$-18,0M vs 6+6, postergação de entregas para o próximo mês.`, `**Unit A**, R$-18.0M vs 6+6, deliveries postponed to next month.`),
          L(`**Unidade B**, R$-24,0M vs 6+6, volume abaixo do previsto (${ex})`, `**Unit B**, R$-24.0M vs 6+6, volume below plan (${ex})`)]) },
        { tipo: "CARD", titulo: `${MB}, R$M`, valores: tabela([MB, MG], [[132, 81, 95, 91], pc(9.1, 13.1, 16.1, 16.6)], [
          L("**Unidade A:** melhor preço médio no trimestre", "**Unit A:** better average price in the quarter"),
          L(`**Unidade B**: redução de custo logístico (${ex})`, `**Unit B**: lower logistics cost (${ex})`)]) },
        { tipo: "CARD", titulo: L("SG&A & Provisão AR, R$M", "SG&A & AR Provision, R$M"), valores: tabela(["SG&A", L("Prov. AR", "AR Prov."), MG],
          [[-96, -71, -68, -63], [-18, -6, -14, -9], pc(-7.9, -12.4, -13.9, -13.1)], [
          L(`**SG&A**: vagas em aberto e menor gasto com viagens (${ex})`, `**SG&A**: open positions and lower travel expenses (${ex})`),
          L("**Prov. AR**: efeito de calendário, regulariza até o fim do ano.", "**AR Prov.**: calendar effect, normalizes by year-end.")]) },
        { tipo: "CARD", titulo: "EBITDA, R$M", valores: tabela(["EBITDA", MG], [[18, 4, 13, 19], pc(1.2, 0.6, 2.2, 3.5)], [
          L(`**Destaque**, melhora puxada por margem e redução de despesas (${ex})`, `**Highlight**, improvement driven by margin and lower expenses (${ex})`)]) },
      ],
    },
    {
      nome: "PADRÃO BELA FP&A_fechamento_Insumos", aba: L("BELA_fechamento_Insumos", "BELA_closing_Inputs"),
      subtitulo: L("Unidade de Insumos – Operacional (exemplo)", "Inputs Unit – Operating (example)"),
      cartoes: [
        { tipo: "CARD", titulo: `${RL}, R$M`, valores: tabela([RL], [[640, 310, 284, 262]], [
          L(`**Insumos**, volume faturado abaixo do previsto por atraso nas entregas (${ex}).`, `**Inputs**, invoiced volume below plan due to delivery delays (${ex}).`),
          L("**Mix na Receita (Atual):** Linha A 45% · Linha B 25% · Linha C 20% · Linha D 10%", "**Revenue mix (Actual):** Line A 45% · Line B 25% · Line C 20% · Line D 10%")]) },
        { tipo: "CARD", titulo: `${MB}, R$M`, valores: tabela([MB, MG], [[84, 46, 38, 35], pc(13.1, 14.8, 13.4, 13.4)], [
          L(`**Destaque**, menor volume e mudança de mix para produtos de menor margem (${ex})`, `**Highlight**, lower volume and mix shift to lower-margin products (${ex})`)]) },
        { tipo: "CARDM", titulo: L("Receita Líquida por Mix (Resultado Comercial)", "Net Revenue by Mix (Commercial Result)"), W: 520, porLinha: 2, valores: tabela(
          [L("Linha A", "Line A"), L("Margem Linha A", "Margin Line A"), L("Linha B", "Line B"), L("Margem Linha B", "Margin Line B"),
           L("Linha C", "Line C"), L("Margem Linha C", "Margin Line C"), L("Linha D", "Line D"), L("Margem Linha D", "Margin Line D")],
          [[310, 140, 128, 118], pc(9.8, 8.6, 7.1, 7.4), [36, 41, 30, 26], pc(31.5, 28.0, 30.4, 29.8),
           [168, 72, 70, 66], pc(18.2, 15.9, 19.5, 20.1), [126, 57, 56, 52], pc(12.4, 16.8, 15.1, 15.9)]) },
      ],
    },
    {
      nome: "PADRÃO BELA FP&A_fechamento_Grãos", aba: L("BELA_fechamento_Grãos", "BELA_closing_Grains"),
      subtitulo: L("Grãos (exemplo)", "Grains (example)"),
      cartoes: [
        { tipo: "CARD", titulo: `${RL}, R$M`, valores: tabela([RL], [[1210, 305, 288, 270]], [
          L(`**Destaque**, menor volume embarcado no mês; parte dos contratos foi renegociada (${ex}).`, `**Highlight**, lower volume shipped in the month; part of the contracts was renegotiated (${ex}).`)]) },
        { tipo: "CARD", titulo: L("Despesas Operacionais, R$M", "Operating Expenses, R$M"), valores: tabela([L("Despesas Operacionais", "Operating Expenses"), MG], [[-42, -16, -14, -15], pc(-3.5, -5.2, -4.9, -5.6)]) },
        { tipo: "CARD", titulo: `${MB}, R$M`, valores: tabela([MB, MG], [[-28, 9, 22, 25], pc(-2.3, 3.0, 7.6, 9.3)], [
          L(`**Destaque**, ganho de marcação a mercado nas posições a fixar (${ex})`, `**Highlight**, mark-to-market gain on unpriced positions (${ex})`),
          L("Phasing: parte do resultado fica para os próximos meses", "Phasing: part of the result moves to the coming months"),
          L("**Importante:** valores ilustrativos, substitua pelos do fechamento", "**Important:** illustrative values, replace with the closing figures")]) },
        { tipo: "CARDM", titulo: L("Grãos Originado", "Grains Sourced"), W: 250, porLinha: 1,
          valores: tabela([L("Produto A em kTon", "Product A in kTon"), L("Produto B em kTon", "Product B in kTon")], [[540, 96, 92, 90], [410, 62, 60, 35]]) },
      ],
    },
  ];
}

export const MODELOS = modelosNoIdioma("pt");
