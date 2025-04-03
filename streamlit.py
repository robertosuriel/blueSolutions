import streamlit as st
import os
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
import pandas as pd
import time
import requests

st.title("Executar Script de Web Scraping")

if st.button("Rodar Script"):
    os.system("python atualizaSheets.py")
    st.success("Script executado com sucesso!")
