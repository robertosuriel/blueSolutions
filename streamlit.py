import streamlit as st
import os

st.title("Executar Script de Web Scraping")

if st.button("Rodar Script"):
    os.system("python atualizaSheets.py")
    st.success("Script executado com sucesso!")
