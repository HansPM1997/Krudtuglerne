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
    st.header("Smagningens karakteristika")
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
    # Knap til Excel
    excel = st.button(
        label="Lav Excel fil",
        width='stretch',
        type='primary',
        icon="😃"
    )

    if excel:
        path = Func.setup_participants_and_rounds(
            tema,
            lokation,
            dato,
            deltagere,
            runder
        )
        st.success(f'Ja tak, chef. Filen er uploadet til: {path}', icon="✅")

