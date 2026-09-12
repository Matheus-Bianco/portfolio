# Gera capa territorial do Painel de Obras e capas institucionais (FT / Auxiliares).
import json
import os
import re
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import Polygon as MplPolygon
from PIL import Image, ImageDraw, ImageFont

ROOT = Path(__file__).resolve().parents[1]
INFRA = Path(r"C:\Users\mathe\OneDrive\Desktop\Trabalhos\02. Joinville\13. Automatização de Processos\01. Infra")
NAVY = (30, 58, 95)
GOLD = (196, 149, 58)
WHITE = (255, 255, 255)


def extract_geojson():
    text = (INFRA / "BairrosData.html").read_text(encoding="utf-8")
    m = re.search(r"var BAIRROS_GEOJSON = (\{.*\});", text)
    if not m:
        raise SystemExit("GeoJSON de bairros não encontrado")
    return json.loads(m.group(1))


def extract_points():
    text = (INFRA / "PopularDados.js").read_text(encoding="utf-8")
    obras = re.findall(r"\['OEX\d+','[^']+','[^']+','[^']+',[^,]+,'[^']*','[^']*',(-26\.\d+),(-48\.\d+)", text)
    lics = re.findall(r"-26\.\d{2,},-48\.\d{2,}", text)
    licitacoes = []
    for pair in lics:
        lat, lon = pair.split(",")
        licitacoes.append((float(lat), float(lon)))
    return [(float(lat), float(lon)) for lat, lon in obras], licitacoes[:22]


def draw_map(geo, obras, dest):
    dest.parent.mkdir(parents=True, exist_ok=True)
    fig, ax = plt.subplots(figsize=(12, 9), dpi=140)
    fig.patch.set_facecolor("#f4f6f8")
    ax.set_facecolor("#e8eef3")
    for feat in geo["features"]:
        geom = feat["geometry"]
        rings = geom["coordinates"] if geom["type"] == "Polygon" else []
        if geom["type"] == "MultiPolygon":
            rings = [r for poly in geom["coordinates"] for r in poly]
        for ring in rings:
            xs = [c[0] for c in ring]
            ys = [c[1] for c in ring]
            ax.add_patch(MplPolygon(list(zip(xs, ys)), closed=True,
                                    facecolor="#c5d4e0", edgecolor="#1e3a5f",
                                    linewidth=0.4, alpha=0.85))
    if obras:
        ax.scatter([p[1] for p in obras], [p[0] for p in obras],
                   s=36, c="#c4953a", edgecolors="#1e3a5f", linewidths=0.6, zorder=5,
                   label="Obras em execução (cadastro)")
    ax.set_aspect("equal")
    ax.set_title("Painel de Obras SED/UIN — Joinville\nBairros oficiais (PMJ) e pontos georreferenciados",
                 fontsize=13, color="#1e3a5f", pad=12)
    ax.set_xlabel("Longitude")
    ax.set_ylabel("Latitude")
    ax.legend(loc="lower right", frameon=True)
    ax.margins(0.02)
    fig.tight_layout()
    fig.savefig(dest, bbox_inches="tight")
    plt.close(fig)


def cover(path, title, subtitle):
    path.parent.mkdir(parents=True, exist_ok=True)
    img = Image.new("RGB", (1400, 788), NAVY)
    draw = ImageDraw.Draw(img)
    draw.rectangle([0, 0, 12, 788], fill=GOLD)
    try:
        font_t = ImageFont.truetype("arial.ttf", 42)
        font_s = ImageFont.truetype("arial.ttf", 22)
    except OSError:
        font_t = ImageFont.load_default()
        font_s = font_t
    draw.text((64, 280), title, fill=WHITE, font=font_t)
    draw.text((64, 360), subtitle, fill=(220, 226, 232), font=font_s)
    img.save(path, quality=90)


def main():
    geo = extract_geojson()
    obras, _ = extract_points()
    draw_map(geo, obras, ROOT / "assets/images/projects/project13/1.png")
    cover(ROOT / "assets/images/projects/project15/1.jpg",
          "Força Tarefa da Educação",
          "Monitoramento de processos licitatórios para a Prefeita de Joinville")
    cover(ROOT / "assets/images/projects/project16/1.jpg",
          "Auxiliares Pedagógicos",
          "Monitoramento das solicitações de educação especial — SED Joinville")
    print("obras:", len(obras), "bairros:", len(geo["features"]))


if __name__ == "__main__":
    main()
