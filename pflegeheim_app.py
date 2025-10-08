import streamlit as st
import pandas as pd

st.set_page_config(page_title="Pflegeheim-Auswertung", layout="wide")

st.title("🏥 Pflegeheim Musterdaten Analyse")

# Datei-Upload
uploaded_file = st.file_uploader("📂 Zieh deine Excel-Datei hier rein", type=["xlsx"])

if uploaded_file:
    try:
        # Excel einlesen
        df = pd.read_excel(uploaded_file)

        st.success("✅ Datei erfolgreich geladen!")

        # Vorschau
        st.subheader("📋 Datenvorschau")
        st.dataframe(df)

        # Altersverteilung
        if "Alter" in df.columns:
            st.subheader("📈 Altersverteilung")
            st.bar_chart(df["Alter"].value_counts().sort_index())

        # Betreuungsbedarf
        if "Betreuungsbedarf" in df.columns:
            st.subheader("🧠 Betreuungsbedarf")
            st.bar_chart(df["Betreuungsbedarf"].value_counts())

        # Abteilungen
        if "Abteilung" in df.columns:
            st.subheader("🧭 Verteilung nach Abteilungen")
            st.bar_chart(df["Abteilung"].value_counts())

        # Einzelzimmer-Filter
        if "Einzelzimmer" in df.columns:
            st.subheader("🏡 Nur Einzelzimmer?")
            if st.checkbox("🔘 Ja, nur Einzelzimmer zeigen"):
                df_filtered = df[df["Einzelzimmer"] == "Ja"]
                st.dataframe(df_filtered)

    except Exception as e:
        st.error(f"❌ Fehler beim Einlesen der Datei: {e}")
else:
    st.info("Bitte lade eine Excel-Datei hoch, um die Auswertung zu starten.")
