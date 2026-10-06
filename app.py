#Deel 1: Basisinstellingen & Titelverdeling

import streamlit as st
import pandas as pd
import os
import io
import plotly.express as px
from datetime import datetime, timedelta

EXCEL_FILE = "hijsmiddelen_database.xlsx"

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
    if row['status'] == "Afgekeurd (Gearchiveerd)": return "⚫ Gearchiveerd"
    dt_str = row['volgende_beproeving'] if is_bep else row['volgende_keuring']
    if not dt_str or str(dt_str).strip() == "" or str(dt_str).strip() == "N.v.t.": return ""
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

# INDELING VOOR DE TITEL EN LOGO
kol_titel, kol_logo = st.columns([5.5, 1])
with kol_titel:
    st.title("🏗️ Centraal Hijsmiddelen Dashboard")
with kol_logo:
    pass
    # st.image("logo.png", width=130)

# KPI KAARTEN BOVENIN HET SCHERM
df_ct = df_actueel[df_actueel['status'] != "Afgekeurd (Gearchiveerd)"] if not df_actueel.empty else pd.DataFrame()
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

#Deel 2:

with links:
    st.subheader("🔒 Toegangsbeheer")
    wachtwoord_invoer = st.text_input("Voer admin-wachtwoord in voor wijzigingen:", type="password")
    is_admin = (wachtwoord_invoer == "HijsBeheer2026!")

    if is_admin:
        st.success("🔓 Admin-modus actief. Je kunt nu mutaties doorvoeren.")
        st.write("---")
        st.subheader("🛠️ Acties")
        modus = st.radio("Wat wil je doen?", ["Bestaand Object Bewerken", "Nieuw Object Toevoegen", "Object Gegevens Wijzigen of Verwijderen"])
        
        INTERVALS = ["Geen beproeving (Textiel)", "Jaarlijks (1 jaar)", "Om de 4 jaar (Staal)"]
        STATUS_OPTIES = ["Actief", "Afgekeurd (Gearchiveerd)", "Niet gevonden (Vermist)"]
        
        if modus == "Bestaand Object Bewerken" and not df_actueel.empty:
            sel_ids = st.multiselect("Stap 1: Kies ID-nummers:", df_actueel['id'].tolist(), key="m_sel")
            n_loc_keuze = st.selectbox("Verplaats naar locatie:", ["Geen wijziging"] + bekende_locaties + ["Gearchiveerd", "Vermist"])
            n_stat = st.selectbox("Wijzig status naar:", ["Geen wijziging"] + STATUS_OPTIES)
            
            st.write("**Stap 2: Keuring / Beproeving registreren**")
            u_dt = st.date_input("Uitvoerdatum (dd-mm-jjjj):", vandaag, format="DD-MM-YYYY")
            k_opt = st.checkbox("🔄 Jaarlijkse Keuring uitgevoerd (+1 jaar)")
            b_opt = st.checkbox("⚖️ Beproeving uitgevoerd")
            if b_opt: g_int = st.selectbox("Kies beproevingsinterval:", INTERVALS)
            
            if st.button("Wijzigingen toepassen"):
                if sel_ids:
                    for oid in sel_ids:
                        masker = df_actueel['id'] == oid
                        if masker.any():
                            idx = df_actueel[masker].index
                            h_loc = df_actueel.at[idx, 'locatie']
                            h_stat = df_actueel.at[idx, 'status']
                            h_type = df_actueel.at[idx, 'type']
                            h_l_kr = df_actueel.at[idx, 'laatste_keuring']
                            h_v_kr = df_actueel.at[idx, 'volgende_keuring']
                            h_l_bp = df_actueel.at[idx, 'laatste_beproeving']
                            h_v_bp = df_actueel.at[idx, 'volgende_beproeving']
                            
                            v_loc = n_loc_keuze if n_loc_keuze != "Geen wijziging" else h_loc
                            v_stat = n_stat if n_stat != "Geen wijziging" else h_stat
                            
                            if v_stat == "Afgekeurd (Gearchiveerd)": v_loc = "Gearchiveerd"
                            elif v_stat == "Niet gevonden (Vermist)": v_loc = "Vermist"
                            
                            v_l_kr = u_dt.strftime("%d-%m-%Y") if k_opt else h_l_kr
                            v_v_kr = (u_dt + timedelta(days=365)).strftime("%d-%m-%Y") if k_opt else h_v_kr
                            
                            if b_opt:
                                v_l_bp = u_dt.strftime("%d-%m-%Y")
                                v_v_bp = (u_dt + timedelta(days=365 if "1 jaar" in g_int else 4*365)).strftime("%d-%m-%Y") if "Geen" not in g_int else "N.v.t."
                            else:
                                v_l_bp, v_v_bp = h_l_bp, h_v_bp
                                
                            if v_stat == "Afgekeurd (Gearchiveerd)":
                                v_v_kr, v_v_bp = "N.v.t.", "N.v.t."
                                
                            df_actueel.at[idx, 'locatie'] = v_loc
                            df_actueel.at[idx, 'status'] = v_stat
                            df_actueel.at[idx, 'laatste_keuring'] = v_l_kr
                            df_actueel.at[idx, 'volgende_keuring'] = v_v_kr
                            df_actueel.at[idx, 'laatste_beproeving'] = v_l_bp
                            df_actueel.at[idx, 'volgende_beproeving'] = v_v_bp
                            
                            nu_log = pd.DataFrame([{'datum': datetime.now().strftime("%d-%m-%Y %H:%M"), 'object_id': oid, 'type': h_type, 'actie': f"Bewerkt (Status: {v_stat} | Locatie: {v_loc})", 'details': 'Admin'}])
                            df_historie = pd.concat([df_historie, nu_log], ignore_index=True)
                    sla_data_op(df_actueel, df_historie)
                    st.success("✅ Wijzigingen succesvol doorgevoerd!"); st.rerun()
                else: st.error("❌ Kies minimaal één ID.")
                
        elif modus == "Nieuw Object Toevoegen":
            with st.form("i_form", clear_on_submit=True):
                n_id = st.text_input("Uniek ID Nummer (bijv. PL-001):").strip()
                
                type_opties = ["Kies een type hijsmiddel... ", "--- Handmatig nieuw type invoeren ---"] + bekende_types
                gekozen_type = st.selectbox("Type selecteren:", type_opties, index=0)
                handmatig_type = st.text_input("Indien nieuw type, typ hier de benaming:")
                
                if gekozen_type == "Kies een type hijsmiddel... ": def_type = ""
                elif gekozen_type == "--- Handmatig nieuw type invoeren ---": def_type = handmatig_type.strip()
                else: def_type = gekozen_type
                
                loc_opties = ["Kies een locatie... ", "--- Handmatig nieuwe locatie invoeren ---"] + bekende_locaties
                gekozen_loc = st.selectbox("Locatie / Vlootnummer selecteren:", loc_opties, index=0)
                handmatig_loc = st.text_input("Indien nieuwe locatie, typ hier het vlootnummer:")
                
                if gekozen_loc == "Kies een locatie... ": def_loc = ""
                elif gekozen_loc == "--- Handmatig nieuwe locatie invoeren ---": def_loc = handmatig_loc.strip()
                else: def_loc = gekozen_loc
                
                st.write("---")
                k_init = st.date_input("Volgende Keuringsdatum:", vandaag + timedelta(days=365), format="DD-MM-YYYY")
                i_keuze = st.selectbox("Beproevingsinterval:", INTERVALS)
                
                if st.form_submit_button("Object Opslaan"):
                    if n_id and def_type and def_loc:
                        b_str = (vandaag + timedelta(days=365 if "1 jaar" in i_keuze else 4*365)).strftime("%d-%m-%Y") if "Geen" not in i_keuze else "N.v.t."
                        nieuwe_rij = pd.DataFrame([{'id': n_id, 'type': def_type, 'locatie': def_loc, 'laatste_keuring': 'Nieuw', 'volgende_keuring': k_init.strftime("%d-%m-%Y"), 'laatste_beproeving': 'Nieuw', 'volgende_beproeving': b_str, 'status': 'Actief'}])
                        df_actueel = pd.concat([df_actueel, nieuwe_rij], ignore_index=True)
                        nu_log = pd.DataFrame([{'datum': datetime.now().strftime("%d-%m-%Y %H:%M"), 'object_id': n_id, 'type': def_type, 'actie': 'Aangemaakt', 'details': 'Admin'}])
                        df_historie = pd.concat([df_historie, nu_log], ignore_index=True)
                        sla_data_op(df_actueel, df_historie)
                        st.success("🎉 Succesvol toegevoegd!"); st.rerun()
                    else: st.error("❌ ID, Type en Locatie zijn verplicht! Kies aub geldige waardes.")

#Deel 3: Correctiemenu & Tabellenoverzichten

        elif modus == "Object Gegevens Wijzigen of Verwijderen" and not df_actueel.empty:
            st.write("✏️ **Corrigeer typefouten of verwijder een object permanent**")
            id_keuze = st.selectbox("Selecteer het te corrigeren ID nummer:", df_actueel['id'].tolist())
            
            rij_data = df_actueel[df_actueel['id'] == id_keuze]
            
            if not rij_data.empty:
                c_id = st.text_input("Aanpassen ID Nummer:", str(rij_data['id'].iloc[0]))
                c_type = st.text_input("Aanpassen Type omschrijving:", str(rij_data['type'].iloc[0]))
                c_loc = st.text_input("Aanpassen Locatie / Vlootnummer:", str(rij_data['locatie'].iloc[0]))
                
                kol1, kol2 = st.columns(2)
                with kol1:
                    if st.button("💾 Wijzigingen Opslaan"):
                        masker = df_actueel['id'] == id_keuze
                        df_actueel.loc[masker, 'id'] = c_id
                        df_actueel.loc[masker, 'type'] = c_type
                        df_actueel.loc[masker, 'locatie'] = c_loc
                        nu_log = pd.DataFrame([{'datum': datetime.now().strftime("%d-%m-%Y %H:%M"), 'object_id': c_id, 'type': c_type, 'actie': 'Gegevens handmatig gecorrigeerd', 'details': 'Admin'}])
                        df_historie = pd.concat([df_historie, nu_log], ignore_index=True)
                        sla_data_op(df_actueel, df_historie)
                        st.success("✅ Gegevens succesvol aangepast!"); st.rerun()
                with kol2:
                    if st.button("🗑️ Permanent VERWIJDEREN"):
                        df_actueel = df_actueel[df_actueel['id'] != id_keuze]
                        nu_log = pd.DataFrame([{'datum': datetime.now().strftime("%d-%m-%Y %H:%M"), 'object_id': id_keuze, 'type': c_type, 'actie': 'Object permanent gewist', 'details': 'Admin'}])
                        df_historie = pd.concat([df_historie, nu_log], ignore_index=True)
                        sla_data_op(df_actueel, df_historie)
                        st.warning("🗑️ Object permanent verwijderd!"); st.rerun()
    else:
        st.info("ℹ️ **Alleen-lezen modus actief.** Voer bovenaan het admin-wachtwoord in om mutaties, correcties of nieuwe objecten toe te voegen.")

with rechts:
    st.subheader("📊 Overzichten (Alleen Lezen)")
    zoek = st.text_input("🔍 Snel zoeken (typ ID, type of auto):", key="z_uniek")
    df_g = df_actueel.copy()
    if zoek and not df_actueel.empty:
        df_g = df_g[df_g['id'].astype(str).str.contains(zoek,case=False) | df_g['type'].astype(str).str.contains(zoek,case=False) | df_g['locatie'].astype(str).str.contains(zoek,case=False)]
    
    t1, t2, t3, t4, t5 = st.tabs(["Actuele Status", "Archief", "Locatie Grafiek (Interactief)", "Aantallen per Type", "Volledige Historie"])
    
    with t1:
        df_r = df_g[df_g['status'] != "Afgekeurd (Gearchiveerd)"] if not df_g.empty else pd.DataFrame()
        df_r = df_r[df_r['id'] != ""]
        if not df_r.empty:
            st.dataframe(df_r[['id', 'type', 'locatie', 'laatste_keuring', 'volgende_keuring', 'Keur_Status', 'laatste_beproeving', 'volgende_beproeving', 'Beproef_Status', 'status']], hide_index=True, use_container_width=True)
            
            out_stream = io.BytesIO()
            with pd.ExcelWriter(out_stream, engine='openpyxl') as w: df_actueel.to_excel(w, sheet_name='actueel', index=False); df_historie.to_excel(w, sheet_name='historie', index=False)
            st.download_button(label="📥 Download Volledige Database Back-up (Excel .xlsx)", data=out_stream.getvalue(), file_name=f"hijsmiddelen_backup_{datetime.now().strftime('%d-%m-%Y')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="btn_xlsx_backup")
        else: st.info("Geen actieve objecten gevonden.")
        
    with t2:
        df_a = df_g[df_g['status'] == "Afgekeurd (Gearchiveerd)"] if not df_g.empty else pd.DataFrame()
        if not df_a.empty:
            st.dataframe(df_a[['id', 'type', 'locatie', 'laatste_keuring', 'laatste_beproeving', 'status']], hide_index=True, use_container_width=True)
            st.write("---")
            st.write("📊 **Verdeling van afgekeurd materiaal in het archief:**")
            df_a_counts = df_a['type'].value_counts().reset_index()
            df_a_counts.columns = ['Type Hijsmiddel', 'Aantal Afgekeurd']
            fig_a = px.bar(df_a_counts, x='Type Hijsmiddel', y='Aantal Afgekeurd', text='Aantal Afgekeurd', title="Afgekeurd materiaal per producttype")
            st.plotly_chart(fig_a, use_container_width=True)
        else: st.info("Het archief is momenteel leeg.")
        
    with t3:
        if not df_actueel.empty:
            st.write("**Aantal actieve hijsmiddelen per locatie (Vlootnummer):**")
            df_actief = df_actueel[(df_actueel['status'] != "Afgekeurd (Gearchiveerd)") & (df_actueel['id'] != "")]
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
            st.dataframe(df_type_counts, hide_index=True, use_container_width=True)
        else: st.info("Geen data beschikbaar.")
            
    with t5:
        st.dataframe(df_historie.sort_index(ascending=False), use_container_width=True, hide_index=True)
