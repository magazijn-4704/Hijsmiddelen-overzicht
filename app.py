#Deel 1

import streamlit as st
import pandas as pd
import os
import io
import plotly.express as px
from datetime import datetime, timedelta

EXCEL_FILE = "hijsmiddelen_database.xlsx"

# 1. INITIALISATIE: Maak een SCHONE Excel aan als deze nog niet bestaat
if not os.path.exists(EXCEL_FILE):
    df_actueel = pd.DataFrame(columns=['id', 'type', 'locatie', 'laatste_keuring', 'volgende_keuring', 'laatste_beproeving', 'volgende_beproeving', 'status'])
    df_historie = pd.DataFrame(columns=['datum', 'object_id', 'type', 'actie', 'details'])
    with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl') as writer:
        df_actueel.to_excel(writer, sheet_name='actueel', index=False)
        df_historie.to_excel(writer, sheet_name='historie', index=False)

def laad_data(sheet):
    try:
        df = pd.read_excel(EXCEL_FILE, sheet_name=sheet, dtype=object)
        return df.fillna("")
    except:
        if sheet == 'actueel':
            return pd.DataFrame(columns=['id', 'type', 'locatie', 'laatste_keuring', 'volgende_keuring', 'laatste_beproeving', 'volgende_beproeving', 'status'])
        else:
            return pd.DataFrame(columns=['datum', 'object_id', 'type', 'actie', 'details'])

def sla_data_op(df_actuel, df_hist):
    with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl') as writer:
        df_actuel.to_excel(writer, sheet_name='actueel', index=False)
        df_hist.to_excel(writer, sheet_name='historie', index=False)

# Session state initialiseren voor data
if 'df_actueel' not in st.session_state:
    st.session_state['df_actueel'] = laad_data('actueel')
if 'df_historie' not in st.session_state:
    st.session_state['df_historie'] = laad_data('historie')

df_actueel = st.session_state['df_actueel']
df_historie = st.session_state['df_historie']

# Veiligheidsnet tegen missende kolommen
for col in ['id', 'type', 'locatie', 'laatste_keuring', 'volgende_keuring', 'laatste_beproeving', 'volgende_beproeving', 'status']:
    if col not in df_actueel.columns:
        df_actueel[col] = ""

vandaag = datetime.now().date()

#Deel 2

def check_status(row, is_bep=False):
    if 'status' not in row or row['status'] in ["Afgekeurd (Gearchiveerd)", "Niet gevonden (Vermist)"]: 
        return "⚫ Gearchiveerd"
    
    dt_str = row['volgende_beproeving'] if is_bep else row['volgende_keuring']
    if not dt_str or str(dt_str).strip() in ["", "N.v.t.", "-", "nan"]: 
        return ""
    
    try:
        parsed_dt = pd.to_datetime(str(dt_str).strip(), dayfirst=True, errors='raise')
        v_dt = parsed_dt.date()
        if v_dt < vandaag: return "🔴 Verlopen"
        if v_dt <= vandaag + timedelta(days=30): return "🟠 Binnenkort"
        return "🟢 OK"
    except: 
        return "⚪ Fout"

if not df_actueel.empty and 'status' in df_actueel.columns:
    df_actueel['Keur_Status'] = df_actueel.apply(lambda r: check_status(r, False), axis=1)
    df_actueel['Beproef_Status'] = df_actueel.apply(lambda r: check_status(r, True), axis=1)
else:
    df_actueel['Keur_Status'] = pd.Series(dtype='str', index=df_actueel.index)
    df_actueel['Beproef_Status'] = pd.Series(dtype='str', index=df_actueel.index)

st.set_page_config(layout="wide", page_title="Hijsmiddelen Beheer")
st.title("🏗️ Centraal Hijsmiddelen Dashboard")

# KPI KAARTEN BOVENIN HET SCHERM
df_ct = df_actueel[~df_actueel['status'].isin(["Afgekeurd (Gearchiveerd)", "Niet gevonden (Vermist)"]) & (df_actueel['id'] != "")] if not df_actueel.empty and 'status' in df_actueel.columns else pd.DataFrame()

col1, col2, col3, col4 = st.columns(4)
col1.metric("Totaal Actief Materieel", len(df_ct))
col2.metric("🔴 Verlopen Keuringen", len(df_ct[df_ct['Keur_Status'] == "🔴 Verlopen"]) if not df_ct.empty else 0)
col3.metric("🔴 Verlopen Beproevingen", len(df_ct[df_ct['Beproef_Status'] == "🔴 Verlopen"]) if not df_ct.empty else 0)
col4.metric("⚠️ Vermist Materieel", len(df_actueel[df_actueel['status'] == "Niet gevonden (Vermist)"]) if not df_actueel.empty and 'status' in df_actueel.columns else 0)

#Deel 3

# ==================== DE WATERDICHTE HIER-SORTERING ====================
ruwe_types = list(set(df_actueel['type'].dropna().tolist())) if not df_actueel.empty and 'type' in df_actueel.columns else []
bekende_types = sorted([str(t).strip() for t in ruwe_types if str(t).strip() not in ["", "nan", "Kies een type hijsmiddel... ", "--- Handmatig nieuw type invoeren ---"]])

ruwe_locaties = list(set(df_actueel['locatie'].dropna().tolist())) if not df_actueel.empty and 'locatie' in df_actueel.columns else []
database_locaties = [str(l).strip() for l in ruwe_locaties if str(l).strip() not in ["", "nan", "Kies een locatie... ", "--- Handmatig nieuwe locatie invoeren ---", "Gearchiveerd", "Vermist"]]

# Voeg vaste basissen toe
for standaard_loc in ['Magazijn A', 'Magazijn D', 'Voertuig 316', 'Voertuig 317', 'Werkplaats']:
    if standaard_loc not in database_locaties:
        database_locaties.append(standaard_loc)

# Sorteer de locatielijst nu rotsvast van A tot Z
bekende_locaties = sorted(database_locaties)
# =======================================================================

gesorteerde_ids = sorted(df_actueel['id'].dropna().tolist()) if not df_actueel.empty and 'id' in df_actueel.columns else []
gesorteerde_ids = [str(x) for x in gesorteerde_ids if str(x).strip() != ""]

links, rechts = st.columns([1, 2.3])

with links:
    st.subheader("🔒 Toegangsbeheer")
    
    if 'admin_wachtwoord' not in st.session_state: st.session_state['admin_wachtwoord'] = ""
    if 'form_id' not in st.session_state: st.session_state['form_id'] = ""
    if 'form_handmat_type' not in st.session_state: st.session_state['form_handmat_type'] = ""
    if 'form_handmat_loc' not in st.session_state: st.session_state['form_handmat_loc'] = ""
    
    if 'reset_counter' not in st.session_state: st.session_state['reset_counter'] = 0
    if 'bew_ids' not in st.session_state: st.session_state['bew_ids'] = []

    wachtwoord_invoer = st.text_input("Voer admin-wachtwoord in voor wijzigingen:", value=st.session_state['admin_wachtwoord'], type="password")
    st.session_state['admin_wachtwoord'] = wachtwoord_invoer
    is_admin = (st.session_state['admin_wachtwoord'] == "HijsBeheer2026!")

    if is_admin:
        st.success("🔓 Admin-modus actief.")
        if st.button("🔒 Uitloggen / Admin-sessie sluiten"):
            st.session_state['admin_wachtwoord'] = ""
            st.rerun()
                
        st.write("---")
        st.subheader("🛠️ Acties")
        modus = st.radio("Wat wil je doen?", ["Bestaand Object Bewerken", "Nieuw Object Toevoegen", "Object Gegevens Wijzigen of Verwijderen"])
        
        INTERVALS = ["Geen beproeving (Textiel)", "Jaarlijks (1 jaar)", "Om de 4 jaar (Staal)"]
        STATUS_OPTIES = ["Actief", "Afgekeurd (Gearchiveerd)", "Niet gevonden (Vermist)"]
        
        if modus == "Bestaand Object Bewerken" and not df_actueel.empty and 'id' in df_actueel.columns:
            sel_ids = st.multiselect("Stap 1: Kies ID-nummers:", gesorteerde_ids, default=st.session_state['bew_ids'], key="m_sel")
            n_loc_keuze = st.selectbox("Verplaats naar locatie:", ["Geen wijziging"] + bekende_locaties + ["Gearchiveerd", "Vermist"])
            n_stat = st.selectbox("Wijzig status naar:", ["Geen wijziging"] + STATUS_OPTIES)
            
            st.write("**Stap 2: Nieuwe Keuring / Beproeving registreren**")
            u_dt = st.date_input("Uitvoerdatum (dd-mm-jjjj):", vandaag, format="DD-MM-YYYY", key="u_dt_bew")
            k_opt = st.checkbox("🔄 Jaarlijkse Keuring uitgevoerd (+1 jaar)", key="k_opt_bew")
            b_opt = st.checkbox("⚖️ Beproeving uitgevoerd", key="b_opt_bew")
            
            g_int = "Geen beproeving (Textiel)"
            if b_opt: 
                g_int = st.selectbox("Kies beproevingsinterval:", INTERVALS, key="g_int_bew")
                
            if st.button("Wijzigingen toepassen"):
                if sel_ids:
                    if b_opt and g_int == "Geen beproeving (Textiel)":
                        st.error("❌ Je hebt aangegeven dat het object beproefd is, maar geen interval (1 of 4 jaar) gekozen!")
                    else:
                        for oid in sel_ids:
                            masker = df_actueel['id'] == oid
                            if masker.any():
                                h_loc = df_actueel.loc[masker, 'locatie'].values
                                h_stat = df_actueel.loc[masker, 'status'].values
                                h_type = df_actueel.loc[masker, 'type'].values
                                h_l_kr = df_actueel.loc[masker, 'laatste_keuring'].values
                                h_v_kr = df_actueel.loc[masker, 'volgende_keuring'].values
                                h_l_bp = df_actueel.loc[masker, 'laatste_beproeving'].values
                                h_v_bp = df_actueel.loc[masker, 'volgende_beproeving'].values
                                
                                v_loc = n_loc_keuze if n_loc_keuze != "Geen wijziging" else h_loc
                                v_stat = n_stat if n_stat != "Geen wijziging" else h_stat
                                
                                if v_stat == "Afgekeurd (Gearchiveerd)": v_loc = "Gearchiveerd"
                                elif v_stat == "Niet gevonden (Vermist)": v_loc = "Vermist"
                                
                                v_l_kr = u_dt.strftime("%d-%m-%Y") if k_opt else h_l_kr
                                v_v_kr = (u_dt + timedelta(days=365)).strftime("%d-%m-%Y") if k_opt else h_v_kr
                                
                                if b_opt:
                                    v_l_bp = u_dt.strftime("%d-%m-%Y")
                                    v_v_bp = (u_dt + timedelta(days=365 if "1 jaar" in g_int else 4*365)).strftime("%d-%m-%Y")
                                else:
                                    v_l_bp, v_v_bp = h_l_bp, h_v_bp
                                    
                                if "Geen" in str(v_v_bp) or str(v_v_bp).strip() in ["", "-", "nan"]:
                                    v_l_bp, v_v_bp = "-", "-"
                                
                                idx = df_actueel[masker].index
                                df_actueel.loc[idx, 'locatie'] = v_loc
                                df_actueel.loc[idx, 'status'] = v_stat
                                df_actueel.loc[idx, 'laatste_keuring'] = v_l_kr
                                df_actueel.loc[idx, 'volgende_keuring'] = v_v_kr
                                df_actueel.loc[idx, 'laatste_beproeving'] = v_l_bp
                                df_actueel.loc[idx, 'volgende_beproeving'] = v_v_bp
                                
                                nu_log = pd.DataFrame([{'datum': datetime.now().strftime("%d-%m-%Y %H:%M"), 'object_id': oid, 'type': h_type, 'actie': f"Bewerkt (Status: {v_stat} | Locatie: {v_loc})", 'details': 'Admin'}])
                                df_historie = pd.concat([df_historie, nu_log], ignore_index=True)
                        
                        st.session_state['df_actueel'] = df_actueel
                        st.session_state['df_historie'] = df_historie
                        sla_data_op(df_actueel, df_historie)
                        st.session_state['bew_ids'] = []
                        st.success("✅ Wijzigingen succesvol doorgevoerd!")
                        st.rerun()
                else: st.error("❌ Kies minimaal één ID.")

#Deel 4a

        elif modus == "Nieuw Object Toevoegen":
            st.write("**📝 Voer de gegevens van het nieuwe object in:**")
            n_id = st.text_input("Uniek ID Nummer (bijv. PL-001):", value=st.session_state['form_id']).strip()
            
            type_opties = ["Kies een type hijsmiddel... ", "--- Handmatig nieuw type invoeren ---"] + bekende_types
            gekozen_type = st.selectbox("Type selecteren:", type_opties, key=f"sb_type_{st.session_state['reset_counter']}")
            handmatig_type = st.text_input("Indien nieuw type, typ hier de benaming:", value=st.session_state['form_handmat_type'])
            def_type = handmatig_type.strip() if gekozen_type == "--- Handmatig nieuw type invoeren ---" else gekozen_type
            if gekozen_type == "Kies een type hijsmiddel... ": def_type = ""
            
            loc_opties = ["Kies een locatie... ", "--- Handmatig nieuwe locatie invoeren ---"] + bekende_locaties
            gekozen_loc = st.selectbox("Locatie / Vlootnummer selecteren:", loc_opties, key=f"sb_loc_{st.session_state['reset_counter']}")
            handmatig_loc = st.text_input("Indien nieuwe locatie, typ hier het vlootnummer:", value=st.session_state['form_handmat_loc'])
            def_loc = handmatig_loc.strip() if gekozen_loc == "--- Handmatig nieuwe locatie invoeren ---" else gekozen_loc
            if gekozen_loc == "Kies een locatie... ": def_loc = ""
            
            st.write("---")
            st.write("**📅 Datums invoeren:**")
            
            k_l_dt = st.date_input("Laatste Keuringsdatum:", vandaag, format="DD-MM-YYYY")
            v_keur_calc = (k_l_dt + timedelta(days=365)).strftime("%d-%m-%Y")
            st.info(f"💡 Volgende keuringsdatum wordt automatisch: **{v_keur_calc}**")
            
            heeft_bep = st.checkbox("⚖️ Dit object heeft ook een Beproeving (bijv. Ketting/Staal)", key="chk_bep_nieuw")
            def_l_bep, def_v_bep = "-", "-"
            beproevings_fout = False
            
            if heeft_bep:
                b_l_dt = st.date_input("Laatste Beprevingsdatum:", vandaag, format="DD-MM-YYYY")
                i_keuze = st.selectbox("Beproevingsinterval:", INTERVALS)
                def_l_bep = b_l_dt.strftime("%d-%m-%Y")
                if "Geen" in i_keuze:
                    beproevings_fout = True
                else:
                    dagen = 365 if "1 jaar" in i_keuze else 4*365
                    def_v_bep = (b_l_dt + timedelta(days=dagen)).strftime("%d-%m-%Y")
                    st.info(f"💡 Volgende beproevingsdatum wordt automatisch: **{def_v_bep}**")
            
            st.session_state['form_id'] = n_id
            st.session_state['form_handmat_type'] = handmatig_type
            st.session_state['form_handmat_loc'] = handmatig_loc
            
            if st.button("💾 Object Opslaan"):
                id_bestaat_al = False
                if not df_actueel.empty and 'id' in df_actueel.columns:
                    id_bestaat_al = n_id in df_actueel['id'].astype(str).str.strip().tolist()
                
                if not (n_id and def_type and def_loc):
                    st.error("❌ ID, Type en Locatie zijn verplicht!")
                elif id_bestaat_al:
                    st.error(f"❌ Het ID nummer '{n_id}' bestaat al!")
                elif beproevings_fout:
                    st.error("❌ Je hebt aangegeven dat het object beproefd is, maar geen interval gekozen!")
                else:
                    nieuwe_rij = pd.DataFrame([{'id': n_id, 'type': def_type, 'locatie': def_loc, 'laatste_keuring': k_l_dt.strftime("%d-%m-%Y"), 'volgende_keuring': v_keur_calc, 'laatste_beproeving': def_l_bep, 'volgende_beproeving': def_v_bep, 'status': 'Actief'}])
                    df_actueel = pd.concat([df_actueel, nieuwe_rij], ignore_index=True)
                    nu_log = pd.DataFrame([{'datum': datetime.now().strftime("%d-%m-%Y %H:%M"), 'object_id': n_id, 'type': def_type, 'actie': 'Aangemaakt', 'details': 'Admin'}])
                    df_historie = pd.concat([df_historie, nu_log], ignore_index=True)
                    
                    st.session_state['df_actueel'] = df_actueel
                    st.session_state['df_historie'] = df_historie
                    sla_data_op(df_actueel, df_historie)
                    
                    st.session_state['form_id'] = ""
                    st.session_state['form_handmat_type'] = ""
                    st.session_state['form_handmat_loc'] = ""
                    st.session_state['reset_counter'] += 1
                    
                    st.success("🎉 Succesvol toegevoegd!")
                    st.rerun()


#Deel 4b

        elif modus == "Nieuw Object Toevoegen":
            st.write("**📝 Voer de gegevens van het nieuwe object in:**")
            n_id = st.text_input("Uniek ID Nummer (bijv. PL-001):", value=st.session_state['form_id']).strip()
            
            type_opties = ["Kies een type hijsmiddel... ", "--- Handmatig nieuw type invoeren ---"] + bekende_types
            gekozen_type = st.selectbox("Type selecteren:", type_opties, key=f"sb_type_{st.session_state['reset_counter']}")
            handmatig_type = st.text_input("Indien nieuw type, typ hier de benaming:", value=st.session_state['form_handmat_type'])
            def_type = handmatig_type.strip() if gekozen_type == "--- Handmatig nieuw type invoeren ---" else gekozen_type
            if gekozen_type == "Kies een type hijsmiddel... ": def_type = ""
            
            loc_opties = ["Kies een locatie... ", "--- Handmatig nieuwe locatie invoeren ---"] + bekende_locaties
            gekozen_loc = st.selectbox("Locatie / Vlootnummer selecteren:", loc_opties, key=f"sb_loc_{st.session_state['reset_counter']}")
            handmatig_loc = st.text_input("Indien nieuwe locatie, typ hier het vlootnummer:", value=st.session_state['form_handmat_loc'])
            def_loc = handmatig_loc.strip() if gekozen_loc == "--- Handmatig nieuwe locatie invoeren ---" else gekozen_loc
            if gekozen_loc == "Kies een locatie... ": def_loc = ""
            
            st.write("---")
            st.write("**📅 Datums invoeren:**")
            
            k_l_dt = st.date_input("Laatste Keuringsdatum:", vandaag, format="DD-MM-YYYY")
            v_keur_calc = (k_l_dt + timedelta(days=365)).strftime("%d-%m-%Y")
            st.info(f"💡 Volgende keuringsdatum wordt automatisch: **{v_keur_calc}**")
            
            heeft_bep = st.checkbox("⚖️ Dit object heeft ook een Beproeving (bijv. Ketting/Staal)", key="chk_bep_nieuw")
            def_l_bep, def_v_bep = "-", "-"
            beproevings_fout = False
            
            if heeft_bep:
                b_l_dt = st.date_input("Laatste Beprevingsdatum:", vandaag, format="DD-MM-YYYY")
                i_keuze = st.selectbox("Beproevingsinterval:", INTERVALS)
                def_l_bep = b_l_dt.strftime("%d-%m-%Y")
                if "Geen" in i_keuze:
                    beproevings_fout = True
                else:
                    dagen = 365 if "1 jaar" in i_keuze else 4*365
                    def_v_bep = (b_l_dt + timedelta(days=dagen)).strftime("%d-%m-%Y")
                    st.info(f"💡 Volgende beproevingsdatum wordt automatisch: **{def_v_bep}**")
            
            st.session_state['form_id'] = n_id
            st.session_state['form_handmat_type'] = handmatig_type
            st.session_state['form_handmat_loc'] = handmatig_loc
            
            if st.button("💾 Object Opslaan"):
                id_bestaat_al = False
                if not df_actueel.empty and 'id' in df_actueel.columns:
                    id_bestaat_al = n_id in df_actueel['id'].astype(str).str.strip().tolist()
                
                if not (n_id and def_type and def_loc):
                    st.error("❌ ID, Type en Locatie zijn verplicht!")
                elif id_bestaat_al:
                    st.error(f"❌ Het ID nummer '{n_id}' bestaat al!")
                elif beproevings_fout:
                    st.error("❌ Je hebt aangegeven dat het object beproefd is, maar geen interval gekozen!")
                else:
                    nieuwe_rij = pd.DataFrame([{'id': n_id, 'type': def_type, 'locatie': def_loc, 'laatste_keuring': k_l_dt.strftime("%d-%m-%Y"), 'volgende_keuring': v_keur_calc, 'laatste_beproeving': def_l_bep, 'volgende_beproeving': def_v_bep, 'status': 'Actief'}])
                    df_actueel = pd.concat([df_actueel, nieuwe_rij], ignore_index=True)
                    nu_log = pd.DataFrame([{'datum': datetime.now().strftime("%d-%m-%Y %H:%M"), 'object_id': n_id, 'type': def_type, 'actie': 'Aangemaakt', 'details': 'Admin'}])
                    df_historie = pd.concat([df_historie, nu_log], ignore_index=True)
                    
                    st.session_state['df_actueel'] = df_actueel
                    st.session_state['df_historie'] = df_historie
                    sla_data_op(df_actueel, df_historie)
                    
                    st.session_state['form_id'] = ""
                    st.session_state['form_handmat_type'] = ""
                    st.session_state['form_handmat_loc'] = ""
                    st.session_state['reset_counter'] += 1
                    
                    st.success("🎉 Succesvol toegevoegd!")
                    st.rerun()


