import streamlit as st
import pandas as pd
import altair as alt
from report_export import build_word_report  # ← Export-Funktion

st.set_page_config(page_title="Pflegeheim-Auswertung", layout="wide")
st.title("🏥 Pflegeheim Musterdaten Analyse")

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

        st.subheader("📋 Datenvorschau (anonymisiert)")
        st.dataframe(df)

        # 📊 Altersverteilung
        if "Alter" in df.columns:
            st.subheader("📊 Altersverteilung")
            chart_age = (
                alt.Chart(df)
                .mark_bar()
                .encode(
                    x=alt.X("Alter:O", sort="ascending", axis=alt.Axis(labelAngle=0)),
                    y=alt.Y("count():Q", title="Anzahl"),
                    tooltip=["Alter", "count()"],
                )
            )
            st.altair_chart(chart_age, use_container_width=True)

        # 🧠 Betreuungsbedarf
        if "Betreuungsbedarf" in df.columns:
            st.subheader("🧠 Verteilung Betreuungsbedarf")
            chart_bedarf = (
                alt.Chart(df)
                .mark_bar()
                .encode(
                    x=alt.X("Betreuungsbedarf:N", axis=alt.Axis(labelAngle=0)),
                    y=alt.Y("count():Q", title="Anzahl"),
                    tooltip=["Betreuungsbedarf", "count()"],
                )
            )
            st.altair_chart(chart_bedarf, use_container_width=True)

        # 🏥 Abteilungen
        if "Abteilung" in df.columns:
            st.subheader("🏥 Verteilung nach Abteilungen")
            chart_abt = (
                alt.Chart(df)
                .mark_bar()
                .encode(
                    x=alt.X("Abteilung:N", axis=alt.Axis(labelAngle=0)),
                    y=alt.Y("count():Q", title="Anzahl"),
                    tooltip=["Abteilung", "count()"],
                )
            )
            st.altair_chart(chart_abt, use_container_width=True)

        # 🛏️ Filter
        if "Einzelzimmer" in df.columns:
            st.subheader("🛏️ Filter")
            if st.checkbox("🔘 Nur Einzelzimmer zeigen", key="filter_single_room"):
                df = df[df["Einzelzimmer"] == "Ja"]
                st.dataframe(df)

        # 📄 Word-Grafikreport (kein Tabellenschrott)
        if df is not None and not df.empty:
            word_bytes = build_word_report(df)
            st.download_button(
                "📄 Grafikreport als Word herunterladen",
                data=word_bytes,
                file_name="pflegeheim_grafikreport.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                key="download_word_report",
            )

    except Exception as e:
        st.error(f"❌ Fehler beim Einlesen/Verarbeiten: {e}")
else:
    st.info("Bitte lade eine Excel-Datei hoch, um die Auswertung zu starten.")
