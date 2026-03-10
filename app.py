import streamlit as st
import pandas as pd
import plotly.express as px
import os

from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, Table
from reportlab.lib.styles import getSampleStyleSheet

st.set_page_config(layout="wide")

st.title("Stationsdashboard")

monate = [
"Januar","Februar","März","April","Mai","Juni",
"Juli","August","September","Oktober","November","Dezember"
]

uploaded_file = st.file_uploader("Excel Datei hochladen", type=["xlsx"])

if uploaded_file:

    personal = pd.read_excel(uploaded_file, sheet_name="Personalkennzahlen")
    leistung = pd.read_excel(uploaded_file, sheet_name="Leistungskennzahlen_PPBV")
    wirtschaft = pd.read_excel(uploaded_file, sheet_name="Wirtschaftskennzahlen")

    stations = personal["Station"].dropna().unique()

    col1, col2 = st.columns(2)

    with col1:
        station = st.selectbox("Station auswählen", stations)

    with col2:
        selected_monate = st.multiselect(
            "Monate auswählen",
            monate,
            default=["Januar"]
        )

    show_year = st.checkbox("Jahresverlauf anzeigen")

    p = personal[personal["Station"] == station]
    l = leistung[leistung["Station"] == station]
    w = wirtschaft[wirtschaft["Station"] == station]

    # -----------------------
    # KPI-Kacheln
    # -----------------------

    def show_kpis(df, name_column):

        kpis = df[[name_column] + selected_monate]

        cols = st.columns(3)

        for i, (_, row) in enumerate(kpis.iterrows()):

            values = [row[m] for m in selected_monate if pd.notna(row[m])]

            if len(values) > 0:
                value = round(sum(values) / len(values), 2)
            else:
                value = "-"

            cols[i % 3].metric(row[name_column], value)

    # -----------------------
    # Diagramm erstellen
    # -----------------------

    def create_bar_chart(df, name_column):

        df_melt = df.melt(
            id_vars=[name_column],
            value_vars=selected_monate,
            var_name="Monat",
            value_name="Wert"
        )

        fig = px.bar(
            df_melt,
            x=name_column,
            y="Wert",
            color="Monat",
            barmode="group"
        )

        return fig

    # -----------------------
    # Jahresdiagramme
    # -----------------------

    def show_year_charts(df, name_column):

        rows = list(df.iterrows())

        for i in range(0, len(rows), 2):

            col1, col2 = st.columns(2)

            with col1:

                index, row = rows[i]

                df_chart = pd.DataFrame({
                    "Monat": monate,
                    "Wert": [row[m] for m in monate]
                })

                fig = px.line(
                    df_chart,
                    x="Monat",
                    y="Wert",
                    title=row[name_column],
                    markers=True
                )

                st.plotly_chart(fig, use_container_width=True)

            if i + 1 < len(rows):

                with col2:

                    index, row = rows[i+1]

                    df_chart = pd.DataFrame({
                        "Monat": monate,
                        "Wert": [row[m] for m in monate]
                    })

                    fig = px.line(
                        df_chart,
                        x="Monat",
                        y="Wert",
                        title=row[name_column],
                        markers=True
                    )

                    st.plotly_chart(fig, use_container_width=True)

    # -----------------------
    # PERSONALKENNZAHLEN
    # -----------------------

    st.header("Personalkennzahlen")

    show_kpis(p, "KPI")

    st.dataframe(p)

    fig_p = create_bar_chart(p, "KPI")
    st.plotly_chart(fig_p, use_container_width=True)

    if show_year:
        st.subheader("Jahresverlauf")
        show_year_charts(p, "KPI")

    # -----------------------
    # LEISTUNGSKENNZAHLEN
    # -----------------------

    st.header("Leistungskennzahlen")

    show_kpis(l, "PPBV")

    st.dataframe(l)

    fig_l = create_bar_chart(l, "PPBV")
    st.plotly_chart(fig_l, use_container_width=True)

    if show_year:
        st.subheader("Jahresverlauf")
        show_year_charts(l, "PPBV")

    # -----------------------
    # WIRTSCHAFTSKENNZAHLEN
    # -----------------------

    st.header("Wirtschaftskennzahlen")

    show_kpis(w, "PPBV")

    st.dataframe(w)

    fig_w = create_bar_chart(w, "PPBV")
    st.plotly_chart(fig_w, use_container_width=True)

    if show_year:
        st.subheader("Jahresverlauf")
        show_year_charts(w, "PPBV")

    # -----------------------
    # PDF GENERIERUNG
    # -----------------------

    def create_pdf():

        styles = getSampleStyleSheet()
        elements = []

        elements.append(Paragraph("Stationsdashboard Bericht", styles["Title"]))
        elements.append(Spacer(1,20))
        elements.append(Paragraph(f"Station: {station}", styles["Normal"]))
        elements.append(Paragraph(f"Monate: {', '.join(selected_monate)}", styles["Normal"]))
        elements.append(Spacer(1,20))

        # Diagramme speichern
        fig_p.write_image("personal_chart.png")
        fig_l.write_image("leistung_chart.png")
        fig_w.write_image("wirtschaft_chart.png")

        elements.append(Paragraph("Personalkennzahlen", styles["Heading2"]))
        elements.append(Image("personal_chart.png", width=500, height=300))
        elements.append(Spacer(1,20))

        elements.append(Paragraph("Leistungskennzahlen", styles["Heading2"]))
        elements.append(Image("leistung_chart.png", width=500, height=300))
        elements.append(Spacer(1,20))

        elements.append(Paragraph("Wirtschaftskennzahlen", styles["Heading2"]))
        elements.append(Image("wirtschaft_chart.png", width=500, height=300))

        doc = SimpleDocTemplate("stationsbericht.pdf")
        doc.build(elements)

        return "stationsbericht.pdf"

    st.divider()

    if st.button("Monatsbericht erstellen"):

        pdf_file = create_pdf()

        with open(pdf_file, "rb") as f:

            st.download_button(
                "PDF herunterladen",
                f,
                file_name="stationsbericht.pdf",
                mime="application/pdf"
            )