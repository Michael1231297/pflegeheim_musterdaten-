import streamlit as st
import pandas as pd
import altair as alt
from report_export import build_word_report

# === Konfiguration ===
st.set_page_config(page_title="Pflegeheim-Auswertung", layout="wide")

# === AWO Corporate Design ===
AWO_ROT = "#e2001A"
AWO_ROT_HELL = "#ff4d5e"  # Hellere Abstufung
AWO_GRAU = "#666666"

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
        st.dataframe(df, use_container_width=True)
        
        # === KPI-Karten oben ===
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("👥 Bewohner gesamt", len(df))
        with col2:
            durchschnittsalter = df["Alter"].mean() if "Alter" in df.columns else 0
            st.metric("📅 Durchschnittsalter", f"{durchschnittsalter:.1f} Jahre")
        with col3:
            if "Betreuungsbedarf" in df.columns:
                hoher_bedarf = len(df[df["Betreuungsbedarf"] == "hoch"])
                st.metric("🔴 Hoher Betreuungsbedarf", hoher_bedarf)
        with col4:
            if "Einzelzimmer" in df.columns:
                einzelzimmer = len(df[df["Einzelzimmer"] == "Ja"])
                st.metric("🛏️ Einzelzimmer", einzelzimmer)
        
        st.markdown("---")
        
        # === Altersverteilung mit Gruppen ===
        if "Alter" in df.columns:
            st.subheader("📊 Altersverteilung (gruppiert)")
            
            # Altersgruppen erstellen
            df_age = df.copy()
            bins = [70, 75, 80, 85, 90, 95, 100]
            labels = ["70-74", "75-79", "80-84", "85-89", "90-94", "95+"]
            df_age["Altersgruppe"] = pd.cut(df_age["Alter"], bins=bins, labels=labels, right=False)
            
            # Gruppierung zählen
            age_counts = df_age["Altersgruppe"].value_counts().sort_index().reset_index()
            age_counts.columns = ["Altersgruppe", "Anzahl"]
            
            chart_age = (
                alt.Chart(age_counts)
                .mark_bar(color=AWO_ROT, cornerRadiusTopLeft=3, cornerRadiusTopRight=3)
                .encode(
                    x=alt.X("Altersgruppe:N", title="Altersgruppe", axis=alt.Axis(labelAngle=0)),
                    y=alt.Y("Anzahl:Q", title="Anzahl Bewohner", axis=alt.Axis(tickMinStep=1)),
                    tooltip=[
                        alt.Tooltip("Altersgruppe:N", title="Altersgruppe"),
                        alt.Tooltip("Anzahl:Q", title="Anzahl")
                    ]
                )
                .properties(height=400)
            )
            st.altair_chart(chart_age, use_container_width=True)
        
        # === Betreuungsbedarf ===
        if "Betreuungsbedarf" in df.columns:
            st.subheader("🧠 Verteilung Betreuungsbedarf")
            
            bedarf_counts = df["Betreuungsbedarf"].value_counts().reset_index()
            bedarf_counts.columns = ["Betreuungsbedarf", "Anzahl"]
            
            # Farben nach Priorität
            color_scale = alt.Scale(
                domain=["hoch", "mittel", "niedrig"],
                range=[AWO_ROT, AWO_ROT_HELL, AWO_GRAU]
            )
            
            chart_bedarf = (
                alt.Chart(bedarf_counts)
                .mark_bar(cornerRadiusTopLeft=3, cornerRadiusTopRight=3)
                .encode(
                    x=alt.X("Betreuungsbedarf:N", title="Betreuungsbedarf", axis=alt.Axis(labelAngle=0)),
                    y=alt.Y("Anzahl:Q", title="Anzahl Bewohner", axis=alt.Axis(tickMinStep=1)),
                    color=alt.Color("Betreuungsbedarf:N", scale=color_scale, legend=None),
                    tooltip=[
                        alt.Tooltip("Betreuungsbedarf:N", title="Betreuungsbedarf"),
                        alt.Tooltip("Anzahl:Q", title="Anzahl")
                    ]
                )
                .properties(height=400)
            )
            st.altair_chart(chart_bedarf, use_container_width=True)
        
        # === Abteilungen ===
        if "Abteilung" in df.columns:
            st.subheader("🏥 Verteilung nach Abteilungen")
            
            abt_counts = df["Abteilung"].value_counts().reset_index()
            abt_counts.columns = ["Abteilung", "Anzahl"]
            
            chart_abt = (
                alt.Chart(abt_counts)
                .mark_bar(color=AWO_ROT, cornerRadiusTopLeft=3, cornerRadiusTopRight=3)
                .encode(
                    x=alt.X("Abteilung:N", title="Abteilung", axis=alt.Axis(labelAngle=-15)),
                    y=alt.Y("Anzahl:Q", title="Anzahl Bewohner", axis=alt.Axis(tickMinStep=1)),
                    tooltip=[
                        alt.Tooltip("Abteilung:N", title="Abteilung"),
                        alt.Tooltip("Anzahl:Q", title="Anzahl")
                    ]
                )
                .properties(height=400)
            )
            st.altair_chart(chart_abt, use_container_width=True)
        
        st.markdown("---")
        
        # === Filter ===
        if "Einzelzimmer" in df.columns:
            st.subheader("🛏️ Filter")
            if st.checkbox("🔘 Nur Einzelzimmer zeigen", key="filter_single_room"):
                df_filtered = df[df["Einzelzimmer"] == "Ja"]
                st.dataframe(df_filtered, use_container_width=True)
        
        # === Word-Export ===
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
