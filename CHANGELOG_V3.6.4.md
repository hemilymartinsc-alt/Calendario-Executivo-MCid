# Calendário Executivo MCid — V3.6.4.0

## Alterações

- Incluídos dois novos campos de dados no visual:
  - `Valor de Investimento` → `Base[Valor de Investimento]`
  - `Valor de Repasse` → `Base[Valor de Repasse]`
- O bloco expandido do evento passa a seguir esta ordem:
  1. Secretaria
  2. Fonte
  3. Proponente
  4. Programa
  5. Investimento
  6. Repasse
  7. UH
  8. Observações
- `Investimento` e `Repasse` são exibidos somente quando houver valor e são formatados em Real (pt-BR) quando o dado recebido for numérico.
- `Executor` deixa de ser exibido no bloco expandido. O papel de dados foi preservado em `capabilities.json` para evitar quebra desnecessária de compatibilidade com relatórios já configurados.
- `runtime.js` permanece intocado.

## Base protegida

Evolução direta da V3.6.3.0 validada no Power BI Desktop.

- Texto explicativo de previsões encurtado para evitar corte pela área da legenda, preservando legibilidade.
