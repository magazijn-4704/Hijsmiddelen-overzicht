#deel1

import streamlit as st
import pandas as pd
import os
import io
import plotly.express as px
from datetime import datetime, timedelta

EXCEL_FILE = "hijsmiddelen_database.xlsx"

# 1. INITIALISATIE: Maak Excel aan met alle kolommen als deze nog niet bestaat
if not os.path.exists(EXCEL_FILE):
    df_actueel = pd.DataFrame(columns=['id', 'type', 'locatie', 'laatste_keuring', 'volgende_keuring', 'laatste_beproeving', 'volgende_beproeving', 'status'])
    df_historie = pd.DataFrame(columns=['datum', 'object_id', 'type', 'actie', 'details'])
    with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl') as writer:
        df_actueel.to_excel(writer, sheet_name='actueel', index=False)
        df_historie.to_excel(writer, sheet_name='historie', index=False)

def laad_data(sheet):
    return pd.read_excel(EXCEL_FILE, sheet_name=sheet).fillna("")

def sla_data_op(df_actuel, df_hist):
    with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl') as writer:
        df_actuel.to_excel(writer, sheet_name='actueel', index=False)
        df_hist.to_excel(writer, sheet_name='historie', index=False)

df_actueel = laad_data('actueel')
df_historie = laad_data('historie')

for col in ['id', 'type', 'locatie', 'laatste_keuring', 'volgende_keuring', 'laatste_beproeving', 'volgende_beproeving', 'status']:
    if col not in df_actueel.columns:
        df_actueel[col] = ""

vandaag = datetime.now().date()

def check_status(row, is_bep=False):
    if row['status'] in ["Afgekeurd (Gearchiveerd)", "Niet gevonden (Vermist)"]: return "⚫ Gearchiveerd"
    dt_str = row['volgende_beproeving'] if is_bep else row['volgende_keuring']
    if not dt_str or str(dt_str).strip() == "" or str(dt_str).strip() == "N.v.t." or str(dt_str).strip() == "-": return ""
    try:
        v_dt = datetime.strptime(str(dt_str).strip(), "%d-%m-%Y").date()
        if v_dt < vandaag: return "🔴 Verlopen"
        if v_dt <= vandaag + timedelta(days=30): return "🟠 Binnenkort"
        return "🟢 OK"
    except: return "⚪ Fout"

if not df_actueel.empty and len(df_actueel.dropna()) > 0:
    df_actueel['Keur_Status'] = df_actueel.apply(lambda r: check_status(r, False), axis=1)
    df_actueel['Beproef_Status'] = df_actueel.apply(lambda r: check_status(r, True), axis=1)
else:
    df_actueel['Keur_Status'], df_actueel['Beproef_Status'] = pd.Series(dtype='str', index=df_actueel.index), pd.Series(dtype='str', index=df_actueel.index)

st.set_page_config(layout="wide", page_title="Hijsmiddelen Beheer")

# Gewoon de titel over de volle breedte, zonder logo-gedoe
st.title("🏗️ Centraal Hijsmiddelen Dashboard")


# KPI KAARTEN BOVENIN HET SCHERM
df_ct = df_actueel[~df_actueel['status'].isin(["Afgekeurd (Gearchiveerd)", "Niet gevonden (Vermist)"])] if not df_actueel.empty else pd.DataFrame()
col1, col2, col3, col4 = st.columns(4)
col1.metric("Totaal Actief Materieel", len(df_ct))
col2.metric("🔴 Verlopen Keuringen", len(df_ct[df_ct['Keur_Status'] == "🔴 Verlopen"]) if not df_ct.empty else 0)
col3.metric("🔴 Verlopen Beproevingen", len(df_ct[df_ct['Beproef_Status'] == "🔴 Verlopen"]) if not df_ct.empty else 0)
col4.metric("⚠️ Vermist Materieel", len(df_actueel[df_actueel['status'] == "Niet gevonden (Vermist)"]) if not df_actueel.empty else 0)

bekende_types = sorted(list(set(df_actueel['type'].dropna().tolist()))) if not df_actueel.empty else []
bekende_types = [t for t in bekende_types if str(t).strip() != ""]

bekende_locaties = sorted(list(set(df_actueel['locatie'].dropna().tolist()))) if not df_actueel.empty else []
bekende_locaties = [l for l in bekende_locaties if str(l).strip() != ""]

for standaard_loc in ['Magazijn A', 'Auto 314', 'Auto 316', 'Auto 317', 'Werkplaats']:
    if standaard_loc not in bekende_locaties:
        bekende_locaties.append(standaard_loc)
bekende_locaties = sorted(bekende_locaties)

links, rechts = st.columns([1, 2.3])

# Deel2A

with links:
    st.subheader("🔒 Toegangsbeheer")
    
    # Gebruik session_state om invoervelden geforceerd te kunnen resetten
    if 'admin_wachtwoord' not in st.session_state: st.session_state['admin_wachtwoord'] = ""
    if 'form_id' not in st.session_state: st.session_state['form_id'] = ""
    if 'form_handmat_type' not in st.session_state: st.session_state['form_handmat_type'] = ""
    if 'form_handmat_loc' not in st.session_state: st.session_state['form_handmat_loc'] = ""
    if 'sel_type_idx' not in st.session_state: st.session_state['sel_type_idx'] = 0
    if 'sel_loc_idx' not in st.session_state: st.session_state['sel_loc_idx'] = 0
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
        
        if modus == "Bestaand Object Bewerken" and not df_actueel.empty:
            sel_ids = st.multiselect("Stap 1: Kies ID-nummers:", df_actueel['id'].tolist(), default=st.session_state['bew_ids'], key="m_sel")
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
                                h_loc = df_actueel.loc[masker, 'locatie'].values[0]
                                h_stat = df_actueel.loc[masker, 'status'].values[0]
                                h_type = df_actueel.loc[masker, 'type'].values[0]
                                h_l_kr = df_actueel.loc[masker, 'laatste_keuring'].values[0]
                                h_v_kr = df_actueel.loc[masker, 'volgende_keuring'].values[0]
                                h_l_bp = df_actueel.loc[masker, 'laatste_beproeving'].values[0]
                                h_v_bp = df_actueel.loc[masker, 'volgende_beproeving'].values[0]
                                
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
                                    
                                if "Geen" in str(v_v_bp) or str(v_v_bp).strip() == "" or str(v_v_bp).strip() == "-":
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
                        sla_data_op(df_actueel, df_historie)
                        st.session_state['bew_ids'] = [] 
                        st.success("✅ Wijzigingen succesvol doorgevoerd!"); st.rerun()
                else: st.error("❌ Kies minimaal één ID.")

#Deel 2b

        elif modus == "Nieuw Object Toevoegen":
            st.write("**📝 Voer de gegevens van het nieuwe object in:**")
            n_id = st.text_input("Uniek ID Nummer (bijv. PL-001):", value=st.session_state['form_id']).strip()
            
            type_opties = ["Kies een type hijsmiddel... ", "--- Handmatig nieuw type invoeren ---"] + bekende_types
            gekozen_type = st.selectbox("Type selecteren:", type_opties, index=st.session_state['sel_type_idx'])
            handmatig_type = st.text_input("Indien nieuw type, typ hier de benaming:")
            def_type = handmatig_type.strip() if gekozen_type == "--- Handmatig nieuw type invoeren ---" else gekozen_type
            if gekozen_type == "Kies een type hijsmiddel... ": def_type = ""
            
            loc_opties = ["Kies een locatie... ", "--- Handmatig nieuwe locatie invoeren ---"] + bekende_locaties
            gekozen_loc = st.selectbox("Locatie / Vlootnummer selecteren:", loc_opties, index=st.session_state['sel_loc_idx'])
            handmatig_loc = st.text_input("Indien nieuwe locatie, typ hier het vlootnummer:")
            def_loc = handmatig_loc.strip() if gekozen_loc == "--- Handmatig nieuwe locatie invoeren ---" else gekozen_loc
            if gekozen_loc == "Kies een locatie... ": def_loc = ""
            
            st.write("---")
            st.write("**📅 Datums invoeren (App rekent vervaldatum zelf uit):**")
            
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
            try: st.session_state['sel_type_idx'] = type_opties.index(gekozen_type)
            except: st.session_state['sel_type_idx'] = 0
            try: st.session_state['sel_loc_idx'] = loc_opties.index(gekozen_loc)
            except: st.session_state['sel_loc_idx'] = 0
            
            if st.button("💾 Object Opslaan"):
                if not (n_id and def_type and def_loc):
                    st.error("❌ ID, Type en Locatie zijn verplicht! Kies aub geldige waardes. Jouw ingevulde tekst is bewaard.")
                elif beproevings_fout:
                    st.error("❌ Je hebt aangegeven dat het object beproefd is, maar geen interval (1 of 4 jaar) gekozen!")
                else:
                    nieuwe_rij = pd.DataFrame([{'id': n_id, 'type': def_type, 'locatie': def_loc, 'laatste_keuring': k_l_dt.strftime("%d-%m-%Y"), 'volgende_keuring': v_keur_calc, 'laatste_beproeving': def_l_bep, 'volgende_beproeving': def_v_bep, 'status': 'Actief'}])
                    df_actueel = pd.concat([df_actueel, nieuwe_rij], ignore_index=True)
                    nu_log = pd.DataFrame([{'datum': datetime.now().strftime("%d-%m-%Y %H:%M"), 'object_id': n_id, 'type': def_type, 'actie': 'Aangemaakt', 'details': 'Admin'}])
                    df_historie = pd.concat([df_historie, nu_log], ignore_index=True)
                    sla_data_op(df_actueel, df_historie)
                    
                    st.session_state['form_id'] = ""
                    st.session_state['form_handmat_type'] = ""
                    st.session_state['form_handmat_loc'] = ""
                    st.session_state['sel_type_idx'] = 0
                    st.session_state['sel_loc_idx'] = 0
                    st.success("🎉 Succesvol toegevoegd!"); st.rerun()

#Deel 3a

        elif modus == "Object Gegevens Wijzigen of Verwijderen" and not df_actueel.empty:
            st.write("✏️ **Corrigeer typefouten of verwijder een object permanent**")
            
            id_opties = ["Kies een ID nummer..."] + df_actueel['id'].tolist()
            id_keuze = st.selectbox("Selecteer het te corrigeren ID nummer:", id_opties, index=0)
            
            if id_keuze != "Kies een ID nummer...":
                masker = df_actueel['id'] == id_keuze
                
                if masker.any():
                    h_id = str(df_actueel.loc[masker, 'id'].values[0])
                    h_type = str(df_actueel.loc[masker, 'type'].values[0])
                    h_loc = str(df_actueel.loc[masker, 'locatie'].values[0])
                    
                    c_id = st.text_input("Aanpassen ID Nummer:", h_id)
                    c_type = st.text_input("Aanpassen Type omschrijving:", h_type)
                    c_loc = st.text_input("Aanpassen Locatie / Vlootnummer:", h_loc)
                    
                    st.write("---")
                    bevestig = st.checkbox("⚠️ Ik weet zeker dat ik dit object wilt wijzigen of permanent wissen.")
                    
                    kol1, kol2 = st.columns(2)
                    with kol1:
                        if st.button("💾 Wijzigingen Opslaan"):
                            if bevestig:
                                df_actueel.loc[masker, 'id'] = c_id
                                df_actueel.loc[masker, 'type'] = c_type
                                df_actueel.loc[masker, 'locatie'] = c_loc
                                nu_log = pd.DataFrame([{'datum': datetime.now().strftime("%d-%m-%Y %H:%M"), 'object_id': c_id, 'type': c_type, 'actie': 'Gegevens handmatig gecorrigeerd', 'details': 'Admin'}])
                                df_historie = pd.concat([df_historie, nu_log], ignore_index=True)
                                sla_data_op(df_actueel, df_historie)
                                st.success("✅ Gegevens succesvol aangepast!"); st.rerun()
                            else:
                                st.error("❌ Vink eerst het vakje 'Ik weet zeker...' aan.")
                    with kol2:
                        if st.button("🗑️ Permanent VERWIJDEREN"):
                            if bevestig:
                                df_actueel = df_actueel[df_actueel['id'] != id_keuze]
                                nu_log = pd.DataFrame([{'datum': datetime.now().strftime("%d-%m-%Y %H:%M"), 'object_id': id_keuze, 'type': h_type, 'actie': 'Object permanent gewist', 'details': 'Admin'}])
                                df_historie = pd.concat([df_historie, nu_log], ignore_index=True)
                                sla_data_op(df_actueel, df_historie)
                                st.warning("🗑️ Object permanent verwijderd!"); st.rerun()
                            else:
                                st.error("❌ Vink eerst het vakje 'Ik weet zeker...' aan.")
            else:
                st.info("ℹ️ Selecteer hierboven een ID nummer om de bijbehorende velden te laden.")
    else:
        st.info("ℹ️ **Alleen-lezen modus actief.** Voer bovenaan het admin-wachtwoord in om mutaties, correcties of nieuwe objecten toe te voegen.")

with rechts:

#Deel 3b

    st.subheader("📊 Overzichten (Alleen Lezen)")
    zoek = st.text_input("🔍 Snel zoeken (typ ID, type of auto):", key="z_uniek")
    df_g = df_actueel.copy()
    if zoek and not df_actueel.empty:
        df_g = df_g[df_g['id'].astype(str).str.contains(zoek,case=False) | df_g['type'].astype(str).str.contains(zoek,case=False) | df_g['locatie'].astype(str).str.contains(zoek,case=False)]
    
    t1, t2, t3, t4, t5 = st.tabs(["Actuele Status", "Archief / Vermist", "Locatie Grafiek (Interactief)", "Aantallen per Type", "Volledige Historie"])
    
    vaste_kolommen = ['id', 'type', 'locatie', 'laatste_keuring', 'volgende_keuring', 'Keur_Status', 'laatste_beproeving', 'volgende_beproeving', 'Beproef_Status', 'status']

    # VERBETERD (Punt 2): HTML-tabel dwingt nu oranje koppen, 1 regel (nowrap) én een verticale schuifbalk af!
    def toon_oranje_tabel(df_tabel):
        html = df_tabel.to_html(index=False, classes='table table-striped')
        html = html.replace('<thead>', '<thead style="background-color: #ff9800; color: white; white-space: nowrap; position: sticky; top: 0;">')
        html = html.replace('<td>', '<td style="white-space: nowrap; padding: 8px;">')
        html = html.replace('<th>', '<th style="white-space: nowrap; padding: 8px; text-align: left;">')
        html = html.replace('<tr>', '<tr style="text-align: left;">', 1)
        st.markdown(f'<div style="overflow-x:auto; overflow-y:auto; max-height:400px; width:100%; border:1px solid #ddd; border-radius:5px;">{html}</div>', unsafe_allow_html=True)

    with t1:
        df_r = df_g[~df_g['status'].isin(["Afgekeurd (Gearchiveerd)", "Niet gevonden (Vermist)"])] if not df_g.empty else pd.DataFrame()
        df_r = df_r[df_r['id'] != ""]
        if not df_r.empty:
            toon_oranje_tabel(df_r[vaste_kolommen])
            st.write("")
            out_stream = io.BytesIO()
            with pd.ExcelWriter(out_stream, engine='openpyxl') as w: df_actueel.to_excel(w, sheet_name='actueel', index=False); df_historie.to_excel(w, sheet_name='historie', index=False)
            st.download_button(label="📥 Download Volledige Database Back-up (Excel .xlsx)", data=out_stream.getvalue(), file_name=f"hijsmiddelen_backup_{datetime.now().strftime('%d-%m-%Y')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="btn_xlsx_backup")
        else: st.info("Geen actieve objecten gevonden.")
        
    with t2:
        df_a = df_g[df_g['status'].isin(["Afgekeurd (Gearchiveerd)", "Niet gevonden (Vermist)"])] if not df_g.empty else pd.DataFrame()
        if not df_a.empty:
            toon_oranje_tabel(df_a[vaste_kolommen])
            st.write("---")
            st.write("📊 **Verdeling van niet-actief materiaal (Archief & Vermist):**")
            
            # VERBETERD (Punt 3): Gekoppeld aan de losse regels zodat je ID en Type ziet bij eroverheen bewegen!
            fig_a = px.bar(df_a, x='type', color='status', hover_data=['id', 'type', 'status'], labels={'type': 'Type Hijsmiddel', 'status': 'Status', 'count': 'Aantal middelen'}, title="Overzicht van afgekeurde en vermiste middelen (Beweeg muis over de staven voor ID details)")
            st.plotly_chart(fig_a, use_container_width=True)
        else: st.info("Het archief is momenteel leeg.")
        
    with t3:
        if not df_actueel.empty:
            st.write("**Aantal actieve hijsmiddelen per locatie (Vlootnummer):**")
            df_actief = df_actueel[~df_actueel['status'].isin(["Afgekeurd (Gearchiveerd)", "Niet gevonden (Vermist)"]) & (df_actueel['id'] != "")]
            if not df_actief.empty:
                fig = px.bar(df_actief, x='locatie', color='type', hover_data=['id', 'type'], labels={'locatie': 'Locatie / Vlootnummer', 'count': 'Aantal middelen', 'type': 'Producttype'}, title="Materiële bezetting per auto (Beweeg muis over de staven voor details)")
                st.plotly_chart(fig, use_container_width=True)
                st.write(df_actief['locatie'].value_counts())
        else: st.info("Geen data beschikbaar voor de grafiek.")
        
    with t4:
        st.write("📋 **Totaaloverzicht van alle unieke types hijsmiddelen in het bedrijf:**")
        df_types_filtered = df_actueel[df_actueel['id'] != ""]
        if not df_types_filtered.empty:
            df_type_counts = df_types_filtered['type'].value_counts().reset_index()
            df_type_counts.columns = ['Type Omschrijving', 'Totaal in bezit (Aantal)']
            toon_oranje_tabel(df_type_counts)
        else: st.info("Geen data beschikbaar.")
            
    with t5:
        st.dataframe(df_historie.sort_index(ascending=False), use_container_width=True, hide_index=True)

