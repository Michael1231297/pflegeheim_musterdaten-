import streamlit as st
import pandas as pd
import altair as alt

st.set_page_config(page_title="Pflegeheim-Auswertung", layout="wide")

st.title("🏥 Pflegeheim Musterdaten Analyse")

# Datei-Upload
uploaded_file = st.file_uploader("📂 Zieh deine anonymisierte Excel-Datei hier rein", type=["xlsx"])

if uploaded_file:
    try:
        # Excel einlesen
        df = pd.read_excel(uploaded_file)

        st.success("✅ Datei erfolgreich geladen!")

        # Vorschau
        st.subheader("📋 Datenvorschau (anonymisiert)")
        st.dataframe(df)

        # 📈 Altersverteilung
        if "Alter" in df.columns:
            st.subheader("📊 Altersverteilung")
            chart = (
                alt.Chart(df)
                .mark_bar()
                .encode(
                    x=alt.X("Alter:O", sort="ascending", axis=alt.Axis(labelAngle=0)),
                    y=alt.Y("count():Q", title="Anzahl"),
                    tooltip=["Alter", "count()"]
                )
                .properties(width=600, height=400)
            )
            st.altair_chart(chart, use_container_width=True)

        # 🧠 Betreuungsbedarf
        if "Betreuungsbedarf" in df.columns:
            st.subheader("🧠 Verteilung Betreuungsbedarf")
            chart = (
                alt.Chart(df)
                .mark_bar()
                .encode(
                    x=alt.X("Betreuungsbedarf:N", axis=alt.Axis(labelAngle=0)),
                    y="count()",
                    tooltip=["Betreuungsbedarf", "count()"]
                )
                .properties(width=600, height=400)
            )
            st.altair_chart(chart, use_container_width=True)

        # 🏥 Abteilungen
        if "Abteilung" in df.columns:
            st.subheader("🏥 Verteilung nach Abteilungen")
            chart = (
                alt.Chart(df)
                .mark_bar()
                .encode(
                    x=alt.X("Abteilung:N", axis=alt.Axis(labelAngle=0)),
                    y="count()",
                    tooltip=["Abteilung", "count()"]
                )
                .properties(width=600, height=400)
            )
            st.altair_chart(chart, use_container_width=True)

        # 🛏️ Einzelzimmer-Filter
        if "Einzelzimmer" in df.columns:
            st.subheader("🛏️ Filter: Nur Einzelzimmer?")
            if st.checkbox("🔘 Ja, nur Einzelzimmer anzeigen"):
                df_filtered = df[df["Einzelzimmer"] == "Ja"]
                st.dataframe(df_filtered)

    except Exception as e:
        st.error(f"❌ Fehler beim Einlesen der Datei: {e}")
else:
    st.info("Bitte lade eine Excel-Datei hoch, um die Auswertung zu starten.")
