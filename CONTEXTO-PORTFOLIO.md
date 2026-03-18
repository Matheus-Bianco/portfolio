# Contexto e Progresso — Portfólio Matheus Bianco

**Última atualização:** Março 2026

---

## Objetivo

Portfólio profissional para o **Processo Seletivo UNESCO — Edital nº 03/2026** (Projeto 914BRZ1153 — Fortalecimento da Gestão Educacional do Estado do Rio Grande do Sul). O material é utilizado para verificação de experiências, competências técnicas e aderência aos requisitos do Edital.

---

## Estrutura do Projeto

| Item | Descrição |
|------|------------|
| **Site** | https://matheus-bianco.github.io/portfolio/ |
| **Repositório** | https://github.com/Matheus-Bianco/portfolio |
| **PDF** | `portfolio-relatorio.pdf` (gerado via `npm run generate-pdf`) |
| **Template PDF** | `portfolio-report.html` |

### Páginas principais
- `index.html` — Início
- `about.html` — Sobre Mim
- `skills.html` — Competências
- `projects.html` — Grid de projetos
- `projects/project1.html` a `project9.html` — Detalhes de cada projeto
- `contact.html` — Contato

---

## Projetos (9)

1. **Plataforma de Indicadores de Risco Educacional** — SED Sergipe/FGV
2. **Painel de Monitoramento de Estudantes Imigrantes** — Joinville
3. **Sistema de Acompanhamento de Dados GA/SE** — SED Sergipe/FGV
4. **Sistema de Gestão de Avaliações Diagnósticas** — SED Sergipe/FGV
5. **Painel de Recomposição de Aprendizagens** — Joinville
6. **Sistema de Monitoramento Temporal de Risco** — SED Sergipe/FGV
7. **Sistema de Priorização de Escolas para o SAEB** — SED Sergipe/FGV
8. **Painel de Indicadores Educacionais — IDEB & Censo** — SED Sergipe/FGV
9. **Sistema de Monitoramento Estratégico SED Joinville** — Joinville (PEI 2025-2029)

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

## Requisitos do Edital 03/2026 (referência)

- **Formação:** Estatística, Economia, Ciência de Dados, Engenharia, Computação, Políticas Públicas ou correlatas
- **Experiência:** 3+ anos em dashboards de políticas públicas (preferencialmente educacional); 3+ anos em avaliação de políticas educacionais; 1+ ano com bases educacionais (Censo, SAEB, registros administrativos)
- **Produtos esperados:** 7 painéis (Todo Jovem na Escola, Atendimento, Metas e Resultados, Indicadores Educacionais, Proteção à Trajetória, Integridade/Clima, Pé-de-Meia RS)

---

## Arquivos de referência (Trabalhos)

- `Edital PF - ToR 03_2026 (2).pdf`
- `TR 032026 - Analista de Dados (sem valor) (4).pdf`
- `PORTFOLIO_UNESCO_Edital_03_2026.md` — Versão expandida em Markdown
