import streamlit as st
import pandas as pd

st.title("RE-Excel Sheet Converter")

file = st.file_uploader("Upload Excel or CSV", type=["xlsx","xls","csv"])

if file:
    if file.name.endswith(".csv"):
        df = pd.read_csv(file)
    else:
        df = pd.read_excel(file)

    st.success("File loaded")
    st.dataframe(df)
