import streamlit as st
import pandas as pd
import datetime as dt

from pathlib import Path

import Dictionary as Dict
import Functions as Func

# Configure app
st.set_page_config(page_title="Krudtuglerne", layout="wide")

# Load dataset
project_dir = Path(__file__).resolve().parent
dataset_path = project_dir / "Complete_dataset.xlsx"
df = pd.read_excel(dataset_path)

st.header("Krudtuglerne: The App")

st.divider()

st.subheader("Vælg bryggeri")
bryggeri = st.selectbox(
    "",
    options=sorted(df['Brewery'].unique()),
    index=None
)

df_filtered = df[df['Brewery'] == bryggeri]

st.dataframe(df_filtered)

st.image("poxycat.jpeg")

with st.sidebar:
    st.header("De vigtige detaljer")
    # Tema
    tema = st.text_input(
        "Aftenens tema",
        width='stretch'
    )
    runder = st.number_input(
        "Antal runder",
        step=1,
        value=3
    )
    # Dato
    dato = st.date_input(
        "Dato",
        value=dt.datetime.today(),
        width='stretch'
    )
    # Lokation
    lokation = st.text_input(
        "Lokation",
        width='stretch'
    )
    # Deltagere
    deltagere = st.multiselect(
        "Deltagere",
        options=Dict.deltagere,
        default=Dict.deltagere,
        accept_new_options=True,
        width='stretch'
    )

    # ---- Validation ----
    missing_fields = []
    if not tema:
        missing_fields.append("tema")
    if not lokation:
        missing_fields.append("lokation")
    if dato is None:
        missing_fields.append("dato")
    if runder is None or runder <= 0:
        missing_fields.append("antal runder")

    download_disabled = len(missing_fields) > 0

    if download_disabled:
        st.warning(
            "Du mangler at indtaste følgende: "
            + ", ".join(missing_fields)
        )

    # Kun generér data hvis alle felter er udfyldt
    if not download_disabled:
        excel_data = Func.setup_participants_and_rounds(
            tema,
            lokation,
            dato,
            deltagere,
            runder
        )
    else:
        excel_data = b""  # placeholder, knappen er alligevel disabled

    # Knap til Excel
    download = st.download_button(
        label='Download template',
        data=excel_data,
        file_name=f'Krudtuglerne - {tema or "uden_tema"} - {dato}.xlsx',
        mime='application/octet-stream',
        disabled=download_disabled
    )

    if download:
        st.success("I maltets rige.... er alle lige! SKÅL!")

