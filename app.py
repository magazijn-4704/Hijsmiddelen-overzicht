import streamlit as st
import os

EXCEL_FILE = "hijsmiddelen_database.xlsx"

st.title("🚨 Database Harde Reset")

# Fysiek de Excel-sheet verwijderen van de schijf
if os.path.exists(EXCEL_FILE):
    os.remove(EXCEL_FILE)
    st.warning("🔥 Het Excel-bestand is succesvol verwijderd.")

# De cache van de browser leegmaken
st.session_state.clear()
st.success("♻️ De Streamlit-cache is volledig schoongemaakt!")
st.info("Je kunt dit bestand nu sluiten. De database is helemaal leeg.")
