# Reordena cards e números de hero para a vaga PNUD (território, Nordeste, georref).
from pathlib import Path
import re

ROOT = Path(__file__).resolve().parents[1]

# file_stem, display, icon, img, alt, meta, title, blurb, tags
CARDS = [
    ("project10", "01", "fa-map-marked-alt", "assets/images/projects/project10/1.jpg", "Painel SEDUC-RS",
     "UNESCO / SEDUC-RS · 2026", "Painel de Indicadores Educacionais — SEDUC-RS",
     "Dados abertos da rede estadual: microdados do INEP em mapas e gráficos para gestores, escolas e o público.",
     ["Python", "INEP / Censo", "Leaflet.js", "Chart.js"]),
    ("project17", "02", "fa-landmark", "assets/images/projects/project17/2.jpg", "Painel de Governança SEDUC-RS",
     "UNESCO / SEDUC-RS · 2026", "Painel de Governança da Educação — SEDUC-RS",
     "Painel mensal para o Governador: frequência, aulas dadas, notas e execução do Agiliza.",
     ["Chart.js", "Leaflet.js", "Cloudflare", "Governança"]),
    ("project8", "03", "fa-chart-area", "assets/images/projects/project8/1.png", "Painel IDEB Sergipe",
     "SEED Sergipe / FGV", "Painel de Indicadores Educacionais — IDEB & Censo",
     "Comparações de IDEB, SAEB e aprovação de Sergipe frente às demais unidades da federação e ao Brasil.",
     ["Python", "Dados INEP", "IDEB", "Pipeline de Dados"]),
    ("project12", "04", "fa-map", "assets/images/projects/project12/1.png", "Mapa de fármacos em São Luís",
     "USP/ESALQ · São Luís/MA", "Gestão de fármacos em São Luís — ciência de dados e território",
     "TCC com mapa Folium, Censo IBGE e 23 mil registros administrativos da rede municipal.",
     ["Python", "Folium", "IBGE / Censo", "Streamlit"]),
    ("project13", "05", "fa-hard-hat", "assets/images/projects/project13/1.jpg", "Mapa de obras escolares de Joinville",
     "SED Joinville / UIN · 2026", "Painel de Obras, Licitações e Terrenos",
     "Sistema territorial da infraestrutura escolar: Leaflet, bairros oficiais da PMJ e pins de obras, licitações e terrenos.",
     ["Leaflet.js", "GeoJSON", "Google Apps Script", "Python"]),
    ("project11", "06", "fa-city", "assets/images/projects/project11/1.jpg", "Painel Dados Abertos Joinville",
     "Secretaria de Educação de Joinville · 2026", "Painel de Dados Abertos — Educação Joinville",
     "Plataforma pública com demografia IBGE, indicadores do INEP e mapa por escola.",
     ["Python", "IBGE", "INEP / Censo", "Leaflet.js"]),
    ("project1", "07", "fa-exclamation-triangle", "assets/images/projects/project1/1.jpg", "Indicadores de risco",
     "SEED Sergipe / FGV", "Plataforma de Indicadores de Risco Educacional",
     "Risco multidimensional para mais de 200 escolas da rede estadual de Sergipe.",
     ["Python", "Pandas", "ETL", "SAEB"]),
    ("project2", "08", "fa-globe-americas", "assets/images/projects/project2/1.jpg", "Painel de imigrantes",
     "Secretaria de Educação de Joinville / BID", "Painel de Monitoramento de Estudantes Imigrantes",
     "Mapa de mais de 3.900 alunos imigrantes de 47 nacionalidades em 170 escolas.",
     ["Python", "Google Apps Script", "Leaflet.js", "Chart.js"]),
    ("project14", "09", "fa-handshake", "assets/images/projects/project14/1.jpg", "Gestão do contrato BID",
     "Prefeitura de Joinville / BID", "Apoio à gestão do contrato BID e Plano de Expansão",
     "Operação BR-L1665: evidência territorial, monitoramento de migrantes e materiais ao TCE.",
     ["Leaflet.js", "Python", "Google Apps Script"]),
    ("project3", "10", "fa-tachometer-alt", "assets/images/projects/project3/1.jpg", "Sistema GA/SE",
     "SEED Sergipe / FGV", "Sistema de Acompanhamento de Dados GA/SE",
     "Movimentação, reprovação e frequência da rede estadual sergipana.",
     ["Python", "API", "Google Apps Script"]),
    ("project4", "11", "fa-clipboard-check", "assets/images/projects/project4/1.jpg", "Avaliações diagnósticas",
     "SEED Sergipe / FGV", "Sistema de Gestão de Avaliações Diagnósticas",
     "Logística e resultados das avaliações diagnósticas em cerca de 201 escolas e 10 DREs.",
     ["Python", "Google Apps Script"]),
    ("project6", "12", "fa-chart-line", "assets/images/projects/project6/1.jpg", "Risco temporal",
     "SEED Sergipe / FGV", "Sistema de Monitoramento Temporal de Risco",
     "Série temporal do risco de reprovação na rede estadual de Sergipe.",
     ["Python", "Chart.js"]),
    ("project7", "13", "fa-sort-amount-down", "assets/images/projects/project7/1.jpg", "Priorização SAEB",
     "SEED Sergipe / FGV", "Sistema de Priorização de Escolas para o SAEB",
     "Classificação multicritério das escolas com maior necessidade de apoio ao SAEB.",
     ["Python", "Google Apps Script"]),
    ("project9", "14", "fa-bullseye", "assets/images/projects/project9/1.jpg", "PEI Joinville",
     "Secretaria de Educação de Joinville · 2025–2026", "Sistema de Monitoramento Estratégico SED Joinville",
     "Metas, OKRs e 77 indicadores do PEI 2025–2029.",
     ["Google Apps Script", "Chart.js", "D3.js"]),
    ("project15", "15", "fa-gavel", "assets/images/projects/project15/1.png", "Força Tarefa Prefeita",
     "Prefeitura de Joinville · 2026", "Força Tarefa da Educação — painel da Prefeita",
     "Monitoramento de processos licitatórios e aquisições, inclusive itens do programa BID.",
     ["Google Apps Script", "Chart.js"]),
    ("project16", "16", "fa-hands-helping", "assets/images/projects/project16/1.png", "Auxiliares pedagógicos",
     "SED Joinville · NEE · 2026", "Monitoramento de solicitações de auxiliares pedagógicos",
     "Cruza ~3,9 mil processos SED com alocação EVN e RH em 166 unidades.",
     ["Google Apps Script", "Python"]),
    ("project5", "17", "fa-book-reader", "assets/images/projects/project5/1.jpg", "Recomposição",
     "Secretaria de Educação de Joinville", "Painel de Recomposição de Aprendizagens",
     "Acompanhamento da participação e do progresso no programa de recomposição.",
     ["Google Apps Script", "Chart.js"]),
]


def card(c):
    stem, num, icon, img, alt, meta, title, blurb, tags = c
    tags_html = "\n".join(f'              <span class="tag">{t}</span>' for t in tags)
    return f'''        <div class="project-card reveal">
          <div class="project-card__image">
            <img src="{img}" alt="{alt}">
            <div class="project-card__image-overlay">
              <div class="project-number">{num}</div>
              <i class="fas {icon}"></i>
            </div>
          </div>
          <div class="project-card__body">
            <p class="project-card__meta">{meta}</p>
            <h3>{title}</h3>
            <p>{blurb}</p>
            <div class="project-card__tags">
{tags_html}
            </div>
            <div class="project-card__buttons">
              <a href="projects/{stem}.html">Ver Projeto <i class="fas fa-arrow-right"></i></a>
            </div>
          </div>
        </div>
'''


def rebuild_projects_html():
    path = ROOT / "projects.html"
    text = path.read_text(encoding="utf-8")
    grid = "      <div class=\"projects-grid\">\n\n" + "\n".join(card(c) for c in CARDS) + "      </div>\n"
    text = re.sub(
        r'      <div class="projects-grid">.*?</div>\n    </div>\n  </section>',
        grid + "    </div>\n  </section>",
        text,
        count=1,
        flags=re.S,
    )
    path.write_text(text, encoding="utf-8")


def update_heroes():
    for stem, num, *_ in CARDS:
        path = ROOT / "projects" / f"{stem}.html"
        html = path.read_text(encoding="utf-8")
        html = re.sub(
            r'<div class="project-hero__number">\d+</div>',
            f'<div class="project-hero__number">{num}</div>',
            html,
            count=1,
        )
        path.write_text(html, encoding="utf-8")


if __name__ == "__main__":
    rebuild_projects_html()
    update_heroes()
    print("cards:", len(CARDS))
