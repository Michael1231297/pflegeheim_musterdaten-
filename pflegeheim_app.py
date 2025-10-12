import streamlit as st
import pandas as pd
import altair as alt
from report_export import build_word_report

# === Konfiguration ===
st.set_page_config(
    page_title="AWO Pflegeheim-Auswertung",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# === AWO Corporate Design ===
AWO_ROT = "#e2001A"
AWO_GRAU_DUNKEL = "#333333"
AWO_GRAU_HELL = "#f5f5f5"

# === Custom CSS für Premium-Look ===
st.markdown("""
<style>
    /* Haupttitel */
    h1 {
        color: #e2001A !important;
        font-weight: 700 !important;
        padding-bottom: 1rem !important;
        border-bottom: 3px solid #e2001A !important;
        margin-bottom: 2rem !important;
    }
    
    /* Untertitel */
    h2, h3 {
        color: #333333 !important;
        font-weight: 600 !important;
        margin-top: 2rem !important;
        margin-bottom: 1rem !important;
    }
    
    /* Metriken-Cards verschönern */
    [data-testid="stMetricValue"] {
        font-size: 2rem !important;
        font-weight: 700 !important;
        color: #e2001A !important;
    }
    
    [data-testid="stMetricLabel"] {
        font-size: 1rem !important;
        font-weight: 500 !important;
        color: #666666 !important;
    }
    
    /* File Uploader */
    [data-testid="stFileUploader"] {
        background-color: #f8f9fa;
        border: 2px dashed #e2001A;
        border-radius: 10px;
        padding: 2rem;
    }
    
    /* Buttons */
    .stDownloadButton button {
        background-color: #e2001A !important;
        color: white !important;
        border: none !important;
        border-radius: 8px !important;
        padding: 0.75rem 2rem !important;
        font-weight: 600 !important;
        transition: all 0.3s ease !important;
    }
    
    .stDownloadButton button:hover {
        background-color: #c20017 !important;
        box-shadow: 0 4px 12px rgba(226, 0, 26, 0.3) !important;
    }
    
    /* Dataframe */
    [data-testid="stDataFrame"] {
        border-radius: 10px;
        overflow: hidden;
        box-shadow: 0 2px 8px rgba(0,0,0,0.08);
    }
    
    /* Success Message */
    .stSuccess {
        background-color: #d4edda !important;
        border-left: 4px solid #28a745 !important;
        border-radius: 8px !important;
    }
    
    /* Trennlinien */
    hr {
        border: none;
        height: 2px;
        background: linear-gradient(to right, #e2001A, transparent);
        margin: 2rem 0;
    }
    
    /* Container-Padding */
    .block-container {
        padding-top: 2rem !important;
        padding-bottom: 3rem !important;
    }
</style>
""", unsafe_allow_html=True)

# === Header ===
st.markdown("<h1>🏥 AWO Pflegeheim – Datenanalyse</h1>", unsafe_allow_html=True)

uploaded_file = st.file_uploader(
    "Ziehen Sie Ihre anonymisierte Excel-Datei hier hinein oder klicken Sie zum Auswählen",
    type=["xlsx"],
    key="file_upload_main"
)

df = None

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)
        st.success("✅ Datei erfolgreich geladen und verarbeitet")
        
        # === KPI-Dashboard ===
        st.markdown("### 📊 Kennzahlen auf einen Blick")
        
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.metric(
                label="👥 Bewohner gesamt",
                value=f"{len(df)}"
            )
        
        with col2:
            if "Alter" in df.columns:
                durchschnittsalter = df["Alter"].mean()
                st.metric(
                    label="📅 Durchschnittsalter",
                    value=f"{durchschnittsalter:.1f} Jahre"
                )
        
        with col3:
            if "Betreuungsbedarf" in df.columns:
                hoher_bedarf = len(df[df["Betreuungsbedarf"] == "hoch"])
                anteil = (hoher_bedarf / len(df) * 100) if len(df) > 0 else 0
                st.metric(
                    label="🔴 Hoher Betreuungsbedarf",
                    value=f"{hoher_bedarf}",
                    delta=f"{anteil:.1f}%"
                )
        
        with col4:
            if "Einzelzimmer" in df.columns:
                einzelzimmer = len(df[df["Einzelzimmer"] == "Ja"])
                anteil_ez = (einzelzimmer / len(df) * 100) if len(df) > 0 else 0
                st.metric(
                    label="🛏️ Einzelzimmer",
                    value=f"{einzelzimmer}",
                    delta=f"{anteil_ez:.1f}%"
                )
        
        st.markdown("---")
        
        # === Visualisierungen ===
        st.markdown("### 📈 Detaillierte Auswertungen")
        
        # === Altersverteilung ===
        if "Alter" in df.columns:
            st.markdown("#### 📊 Altersverteilung")
            
            df_age = df.copy()
            bins = [70, 75, 80, 85, 90, 95, 100]
            labels = ["70-74", "75-79", "80-84", "85-89", "90-94", "95+"]
            df_age["Altersgruppe"] = pd.cut(df_age["Alter"], bins=bins, labels=labels, right=False)
            
            age_counts = df_age["Altersgruppe"].value_counts().sort_index().reset_index()
            age_counts.columns = ["Altersgruppe", "Anzahl"]
            
            chart_age = (
                alt.Chart(age_counts)
                .mark_bar(
                    color=AWO_ROT,
                    cornerRadiusTopLeft=8,
                    cornerRadiusTopRight=8,
                    opacity=0.9
                )
                .encode(
                    x=alt.X(
                        "Altersgruppe:N",
                        title="Altersgruppe",
                        axis=alt.Axis(
                            labelAngle=0,
                            labelFontSize=12,
                            titleFontSize=14,
                            titlePadding=15,
                            labelPadding=10
                        )
                    ),
                    y=alt.Y(
                        "Anzahl:Q",
                        title="Anzahl Bewohner",
                        axis=alt.Axis(
                            tickMinStep=1,
                            labelFontSize=12,
                            titleFontSize=14,
                            titlePadding=15,
                            grid=True,
                            gridOpacity=0.3
                        )
                    ),
                    tooltip=[
                        alt.Tooltip("Altersgruppe:N", title="Altersgruppe"),
                        alt.Tooltip("Anzahl:Q", title="Anzahl Bewohner")
                    ]
                )
                .properties(height=450)
                .configure_view(strokeWidth=0)
            )
            st.altair_chart(chart_age, use_container_width=True)
        
        # === Zwei Charts nebeneinander ===
        col_left, col_right = st.columns(2)
        
        # === Betreuungsbedarf ===
        with col_left:
            if "Betreuungsbedarf" in df.columns:
                st.markdown("#### 🧠 Betreuungsbedarf")
                
                bedarf_counts = df["Betreuungsbedarf"].value_counts().reset_index()
                bedarf_counts.columns = ["Betreuungsbedarf", "Anzahl"]
                
                chart_bedarf = (
                    alt.Chart(bedarf_counts)
                    .mark_bar(
                        color=AWO_ROT,
                        cornerRadiusTopLeft=8,
                        cornerRadiusTopRight=8,
                        opacity=0.9
                    )
                    .encode(
                        x=alt.X(
                            "Betreuungsbedarf:N",
                            title="Betreuungsbedarf",
                            axis=alt.Axis(
                                labelAngle=0,
                                labelFontSize=12,
                                titleFontSize=14,
                                titlePadding=15,
                                labelPadding=10
                            )
                        ),
                        y=alt.Y(
                            "Anzahl:Q",
                            title="Anzahl",
                            axis=alt.Axis(
                                tickMinStep=1,
                                labelFontSize=12,
                                titleFontSize=14,
                                titlePadding=15,
                                grid=True,
                                gridOpacity=0.3
                            )
                        ),
                        tooltip=[
                            alt.Tooltip("Betreuungsbedarf:N", title="Betreuungsbedarf"),
                            alt.Tooltip("Anzahl:Q", title="Anzahl")
                        ]
                    )
                    .properties(height=400)
                    .configure_view(strokeWidth=0)
                )
                st.altair_chart(chart_bedarf, use_container_width=True)
        
        # === Abteilungen ===
        with col_right:
            if "Abteilung" in df.columns:
                st.markdown("#### 🏥 Abteilungen")
                
                abt_counts = df["Abteilung"].value_counts().reset_index()
                abt_counts.columns = ["Abteilung", "Anzahl"]
                
                chart_abt = (
                    alt.Chart(abt_counts)
                    .mark_bar(
                        color=AWO_ROT,
                        cornerRadiusTopLeft=8,
                        cornerRadiusTopRight=8,
                        opacity=0.9
                    )
                    .encode(
                        x=alt.X(
                            "Abteilung:N",
                            title="Abteilung",
                            axis=alt.Axis(
                                labelAngle=0,
                                labelFontSize=11,
                                titleFontSize=14,
                                titlePadding=15,
                                labelPadding=10,
                                labelLimit=120
                            )
                        ),
                        y=alt.Y(
                            "Anzahl:Q",
                            title="Anzahl",
                            axis=alt.Axis(
                                tickMinStep=1,
                                labelFontSize=12,
                                titleFontSize=14,
                                titlePadding=15,
                                grid=True,
                                gridOpacity=0.3
                            )
                        ),
                        tooltip=[
                            alt.Tooltip("Abteilung:N", title="Abteilung"),
                            alt.Tooltip("Anzahl:Q", title="Anzahl")
                        ]
                    )
                    .properties(height=400)
                    .configure_view(strokeWidth=0)
                )
                st.altair_chart(chart_abt, use_container_width=True)
        
        st.markdown("---")
        
        # === Datenvorschau ===
        with st.expander("📋 Vollständige Datentabelle anzeigen"):
            st.dataframe(df, use_container_width=True, height=400)
        
        # === Filter ===
        st.markdown("### 🔍 Filterfunktionen")
        
        col_filter1, col_filter2 = st.columns([1, 3])
        
        with col_filter1:
            if "Einzelzimmer" in df.columns:
                show_einzelzimmer = st.checkbox("🛏️ Nur Einzelzimmer", key="filter_single_room")
        
        with col_filter2:
            if show_einzelzimmer:
                df_filtered = df[df["Einzelzimmer"] == "Ja"]
                st.info(f"📊 Gefiltert: {len(df_filtered)} von {len(df)} Bewohnern in Einzelzimmern")
                st.dataframe(df_filtered, use_container_width=True, height=300)
        
        st.markdown("---")
        
        # === Export ===
        st.markdown("### 📥 Export")
        
        if df is not None and not df.empty:
            word_bytes = build_word_report(df)
            st.download_button(
                label="📄 Grafikreport als Word herunterladen",
                data=word_bytes,
                file_name="awo_pflegeheim_report.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                key="download_word_report",
            )
    
    except Exception as e:
        st.error(f"❌ Fehler beim Verarbeiten der Datei: {e}")

else:
    st.info("👆 Bitte laden Sie eine Excel-Datei hoch, um die Analyse zu starten")
