# Contexto e Progresso — Portfólio Matheus Bianco

**Última atualização:** setembro 2026

---

## Objetivo

Portfólio profissional de projetos em dados, sistemas e políticas públicas educacionais. Serve à verificação de experiências e competências técnicas em processos seletivos (UNESCO/SEDUC-RS, PNUD/Consórcio Nordeste e outros). Destaca a atuação no Nordeste: Prefeitura de São Luís/MA e governo estadual de Sergipe (FGV). Não apresenta candidaturas em curso como projeto entregue.

---

## Estrutura do Projeto

| Item | Descrição |
|------|------------|
| **Site** | https://matheus-bianco.github.io/portfolio/ |
| **Repositório** | https://github.com/Matheus-Bianco/portfolio |
| **PDF** | `portfolio-relatorio.pdf` (gerado via `npm run generate-pdf`) |
| **Template PDF** | `portfolio-report.html` |

### Páginas principais
- `index.html` — Início (escopo ~2.660 escolas / ~820 mil matrículas)
- `about.html` — Sobre Mim
- `skills.html` — Competências (inclui IBGE, APIs oficiais, dados territoriais)
- `projects.html` — Grid de projetos
- `projects/project1.html` a `project12.html` — Detalhes de cada projeto
- `contact.html` — Contato

---

## Projetos (16, ordem PNUD)

1. **Plataforma de Indicadores de Risco Educacional** — SEED Sergipe/FGV
2. **Painel de Monitoramento de Estudantes Imigrantes** — Joinville / BID
3. **Sistema de Acompanhamento de Dados GA/SE** — SEED Sergipe/FGV
4. **Sistema de Gestão de Avaliações Diagnósticas** — SEED Sergipe/FGV
5. **Painel de Recomposição de Aprendizagens** — Joinville
6. **Sistema de Monitoramento Temporal de Risco** — SEED Sergipe/FGV
7. **Sistema de Priorização de Escolas para o SAEB** — SEED Sergipe/FGV
8. **Painel de Indicadores Educacionais — IDEB & Censo** — SEED Sergipe/FGV
9. **Sistema de Monitoramento Estratégico SED Joinville** — Joinville (PEI 2025-2029)
10. **Painel de Indicadores Educacionais — SEDUC-RS** — UNESCO / SEDUC-RS (contrato ED00585/2026)
10b. **Painel de Governança da Educação** — UNESCO / SEDUC-RS (painel gerencial; arquivo `project17.html`)
11. **Painel de Dados Abertos — Educação Joinville** — SME Joinville (IBGE + INEP)
12. **Gestão de fármacos em São Luís — ciência de dados e território** — TCC USP/ESALQ (mapa Folium + IBGE; resumo na Revista E&S)
13. **Painel de Obras / Infra** — SED/UIN Joinville (Leaflet + bairros PMJ)
14. **Gestão contrato BID / Plano de Expansão** — Joinville
15. **Força Tarefa Prefeita** — processos licitatórios
16. **Auxiliares pedagógicos** — solicitações NEE

Ordem de exibição no site (cards 01–16) prioriza território, Nordeste e georreferenciamento para o TR 21/2026.

### Escopo agregado (home)
- **~2.660 escolas:** 2.294 RS + 202 Sergipe + 166 Joinville
- **~820 mil matrículas:** 654 mil RS + 90 mil Sergipe + 79 mil Joinville

---

## Funcionalidades do PDF

- Capa com data dinâmica
- Sumário com hiperlinks e números de página
- Paginação no rodapé (X / Y)
- Seções: Sobre Mim, Competências, Projetos, Contato
- Cada projeto: Contexto, Solução Desenvolvida, Impacto, 2 imagens
- `page-break-after: avoid` em títulos para evitar títulos órfãos

---

## Código-fonte incluído no repositório

- `sistema-planejamento-jvl/` — Código do Sistema de Monitoramento Estratégico (Projeto 09)

---

## Comandos úteis

```bash
npm run generate-pdf   # Gera portfolio-relatorio.pdf
```

---

## Prints do Projeto 10

Prints em `assets/images/projects/project10/` capturados de https://indicadores.educacao.rs.gov.br/ (hub, acesso/matrículas, fluxo e visão por escola). Regenerar: `node scripts/capture-project10.js`.

Prints em `assets/images/projects/project11/` capturados de https://matheus-bianco.github.io/painel-indicadores-joinville/ (hub, demografia IBGE, acesso e visão por escola). Regenerar: `node scripts/capture-project11.js`.

---

## Arquivos de referência (Trabalhos)

- Pasta `06. UNESCO/04. Produto 4_Indicadores Educacionais/`
- `01. Editais/PNUD-TR-21-2026/` — materiais da candidatura TR 21/2026 (não versionar dados pessoais)
