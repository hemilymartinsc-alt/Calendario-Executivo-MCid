# Calendário Executivo MCid — V3.5.0 / Opção 2

Branch alvo: `v3-opcao2`. A `main` permanece como baseline funcional 1.2.1.

## Arquitetura
- `src/runtime.js`: preservado byte a byte da versão funcional 1.2.1.
- `src/visual.ts`: camada V3, com regras adicionais, interface Opção 2 e tratamento visível de erros.
- `style/visual.less`: preservado da versão funcional.
- `style/v3-opcao2.less`: camada visual Full HD 1920×1080.

## Build oficial
O GitHub Actions da branch usa Node 22 e `npx pbiviz package`. O artefato gerado se chama `Calendario-Executivo-MCid-v3.5.0-OPCAO2`.

## Campos adicionais da V3
- Fonte / Subfonte → `Base[Subfonte]`
- Minha Casa, Minha Vida → `Base[Minha casa minha vida]`
- Novo PAC → `Base[Novo PAC (sim/não)]`
- Quantidade de UH → `Base[UH]`

API: 5.10.0.
Power BI Desktop alvo: 2.157.879.0 64-bit (agosto/2026).
