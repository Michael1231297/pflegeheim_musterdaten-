import streamlit as st
import pandas as pd
import altair as alt
from docx import Document
import io

st.set_page_config(page_title="Pflegeheim-Auswertung", layout="wide")
st.title("🏥 Pflegeheim Musterdaten Analyse")

# ---- Datei-Upload (eindeutiger key!) ----
uploaded_file = st.file_uploader(
    "📂 Zieh deine anonymisierte Excel-Datei hier rein",
    type=["xlsx"],
    key="file_upload_main"
)

df = None

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)
        st.success("✅ Datei erfolgreich geladen!")

        # ---- Vorschau ----
        st.subheader("📋 Datenvorschau (anonymisiert)")
        st.dataframe(df)

        # ---- Altersverteilung ----
        if "Alter" in df.columns:
            st.subheader("📊 Altersverteilung")
            chart_age = (
                alt.Chart(df)
                .mark_bar()
                .encode(
                    x=alt.X("Alter:O", sort="ascending", axis=alt.Axis(labelAngle=0)),
                    y=alt.Y("count():Q", title="Anzahl"),
                    tooltip=["Alter", "count()"]
                )
                .properties(height=400)
            )
            st.altair_chart(chart_age, use_container_width=True)

        # ---- Betreuungsbedarf ----
        if "Betreuungsbedarf" in df.columns:
            st.subheader("🧠 Verteilung Betreuungsbedarf")
            chart_bedarf = (
                alt.Chart(df)
                .mark_bar()
                .encode(
                    x=alt.X("Betreuungsbedarf:N", axis=alt.Axis(labelAngle=0)),
                    y=alt.Y("count():Q", title="Anzahl"),
                    tooltip=["Betreuungsbedarf", "count()"]
                )
                .properties(height=400)
            )
            st.altair_chart(chart_bedarf, use_container_width=True)

        # ---- Abteilungen ----
        if "Abteilung" in df.columns:
            st.subheader("🏥 Verteilung nach Abteilungen")
            chart_abt = (
                alt.Chart(df)
                .mark_bar()
                .encode(
                    x=alt.X("Abteilung:N", axis=alt.Axis(labelAngle=0)),
                    y=alt.Y("count():Q", title="Anzahl"),
                    tooltip=["Abteilung", "count()"]
                )
                .properties(height=400)
            )
            st.altair_chart(chart_abt, use_container_width=True)

        # ---- Filter: Einzelzimmer ----
        if "Einzelzimmer" in df.columns:
            st.subheader("🛏️ Filter")
            only_single = st.checkbox("🔘 Nur Einzelzimmer zeigen", key="filter_single_room")
            if only_single:
                df = df[df["Einzelzimmer"] == "Ja"]
                st.dataframe(df)

        # ---- Word-Export ----
        if df is not None and not df.empty:
            doc = Document()
            doc.add_heading("Pflegeheim-Datenanalyse", 0)
            doc.add_paragraph(
                "Diese Analyse wurde automatisch aus der hochgeladenen Excel-Datei generiert."
            )
            doc.add_paragraph("Tabellarische Übersicht:")

            table = doc.add_table(rows=1, cols=len(df.columns))
            table.style = 'Table Grid'
            hdr = table.rows[0].cells
            for i, col in enumerate(df.columns):
                hdr[i].text = str(col)
            for _, row in df.iterrows():
                cells = table.add_row().cells
                for i, val in enumerate(row):
                    cells[i].text = str(val)

            mem = io.BytesIO()
            doc.save(mem)
            mem.seek(0)

            st.download_button(
                label="📄 Auswertung als Word-Datei herunterladen",
                data=mem,
                file_name="pflegeheim_auswertung.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

    except Exception as e:
        st.error(f"❌ Fehler beim Einlesen der Datei: {e}")
else:
    st.info("Bitte lade eine Excel-Datei hoch, um die Auswertung zu starten.")
