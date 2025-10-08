from io import BytesIO
from docx import Document
from docx.shared import Inches
import matplotlib.pyplot as plt
import pandas as pd

def _make_bar_image(series: pd.Series, title: str, xlabel: str) -> BytesIO:
    counts = series.value_counts().sort_index()
    fig, ax = plt.subplots(figsize=(8, 4.5))
    ax.bar(counts.index.astype(str), counts.values)
    ax.set_title(title)
    ax.set_xlabel(xlabel)
    ax.set_ylabel("Anzahl")
    for c, v in zip(ax.patches, counts.values):
        ax.annotate(f"{int(v)}", (c.get_x()+c.get_width()/2, v), ha="center", va="bottom", fontsize=9)
    for t in ax.get_xticklabels():
        t.set_rotation(0); t.set_ha("center")
    fig.tight_layout()
    buf = BytesIO()
    fig.savefig(buf, format="png", dpi=200, bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    return buf

def build_word_report(df: pd.DataFrame) -> BytesIO:
    """Erzeugt einen Word-Report mit den Grafiken (PNG eingebettet)."""
    doc = Document()
    doc.add_heading("Pflegeheim-Datenanalyse (Grafikreport)", 0)
    doc.add_paragraph(
        "Automatisch generierter Bericht aus der hochgeladenen Excel-Datei. "
        "Die folgenden Abbildungen zeigen die wichtigsten Verteilungen."
    )

    # Charts je nach vorhandenen Spalten bauen
    if "Alter" in df.columns and not df["Alter"].empty:
        img = _make_bar_image(df["Alter"], "Altersverteilung", "Alter")
        doc.add_paragraph("Altersverteilung").bold = True
        doc.add_picture(img, width=Inches(6.5))

    if "Betreuungsbedarf" in df.columns and not df["Betreuungsbedarf"].empty:
        img = _make_bar_image(df["Betreuungsbedarf"], "Betreuungsbedarf", "Kategorie")
        doc.add_paragraph("Betreuungsbedarf").bold = True
        doc.add_picture(img, width=Inches(6.5))

    if "Abteilung" in df.columns and not df["Abteilung"].empty:
        img = _make_bar_image(df["Abteilung"], "Abteilungen", "Abteilung")
        doc.add_paragraph("Abteilungen").bold = True
        doc.add_picture(img, width=Inches(6.5))

    mem = BytesIO()
    doc.save(mem)
    mem.seek(0)
    return mem
