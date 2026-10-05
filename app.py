import streamlit as st
import pandas as pd
import os

EXCEL_FILE = "hijsmiddelen_database.xlsx"

# Dit wist direct de database online zodra de app ververst!
if os.path.exists(EXCEL_FILE):
    os.remove(EXCEL_FILE)

# Maak direct een 100% lege, schone database aan
df_actueel = pd.DataFrame(columns=['id', 'type', 'locatie', 'keuringsdatum', 'beproevingsdatum', 'status'])
df_historie = pd.DataFrame(columns=['datum', 'object_id', 'type', 'actie', 'details'])
with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl') as writer:
    df_actueel.to_excel(writer, sheet_name='actueel', index=False)
    df_historie.to_excel(writer, sheet_name='historie', index=False)

st.set_page_config(layout="wide")
st.title("🏗️ Database Succesvol Gereset!")
st.success("🎉 De database is volledig leeggemaakt en hersteld naar de fabrieksinstellingen! Je kunt deze code nu weer vervangen door de echte code.")
