import streamlit as st
import pandas as pd
import os
import io
from datetime import datetime, timedelta

EXCEL_FILE = "hijsmiddelen_database.xlsx"

if not os.path.exists(EXCEL_FILE):
    df_actueel = pd.DataFrame(columns=['id', 'type', 'locatie', 'keuringsdatum', 'beproevingsdatum', 'status'])
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
vandaag = datetime.now().date()

def check_status(row, is_bep=False):
    if row['status'] == "Afgekeurd (Gearchiveerd)": return "⚫ Gearchiveerd"
    dt_str = row['beproevingsdatum'] if is_bep else row['keuringsdatum']
    if not dt_str or str(dt_str).strip() == "": return ""
    try:
        v_dt = datetime.strptime(str(dt_str).strip(), "%d-%m-%Y").date()
        if v_dt < vandaag: return "🔴 Verlopen"
        if v_dt <= vandaag + timedelta(days=30): return "🟠 Binnenkort"
        return "🟢 OK"
    except: return "⚪ Fout"

if not df_actueel.empty:
    df_actueel['Keur_Status'] = df_actueel.apply(lambda r: check_status(r, False), axis=1)
    df_actueel['Beproef_Status'] = df_actueel.apply(lambda r: check_status(r, True), axis=1)
else:
    df_actueel['Keur_Status'], df_actueel['Beproef_Status'] = pd.Series(dtype='str'), pd.Series(dtype='str')

st.set_page_config(layout="wide", page_title="Hijsmiddelen Beheer")
st.title("🏗️ Hijsmiddelen Dashboard")

df_ct = df_actueel[df_actueel['status'] != "Afgekeurd (Gearchiveerd)"] if not df_actueel.empty else pd.DataFrame()
col1, col2, col3, col4 = st.columns(4)
col1.metric("Totaal Actief", len(df_ct))
col2.metric("🔴 Verlopen Keuringen", len(df_ct[df_ct['Keur_Status'] == "🔴 Verlopen"]) if not df_ct.empty else 0)
col3.metric("🔴 Verlopen Beproevingen", len(df_ct[df_ct['Beproef_Status'] == "🔴 Verlopen"]) if not df_ct.empty else 0)
col4.metric("⚠️ Vermist", len(df_actueel[df_actueel['status'] == "Niet gevonden (Vermist)"]) if not df_actueel.empty else 0)

links, rechts = st.columns([1, 2.3])

with links:
    st.subheader("🛠️ Acties")
    modus = st.radio("Wat wil je doen?", ["Objecten Bewerken", "Nieuw Object Toevoegen"])
    INTERVALS = ["Geen beproeving (Textiel)", "Jaarlijks (1 jaar)", "Om de 4 jaar (Staal)"]
    STATUS_OPTIES = ["Actief", "Afgekeurd (Gearchiveerd)", "Niet gevonden (Vermist)"]
    basis_locs = sorted(list(set(['Magazijn A', 'Auto 314', 'Auto 316', 'Auto 317', 'Werkplaats'] + df_actueel['locatie'].tolist() if not df_actueel.empty else [])))
    
    if modus == "Objecten Bewerken" and not df_actueel.empty:
        sel_ids = st.multiselect("Stap 1: Kies ID-nummers:", df_actueel['id'].tolist(), key="m_sel")
        n_loc = st.selectbox("Verplaats naar:", ["Geen wijziging"] + basis_locs)
        n_stat = st.selectbox("Wijzig status naar:", ["Geen wijziging"] + STATUS_OPTIES)
        u_dt = st.date_input("Uitvoerdatum van keuring:", vandaag)
        k_opt = st.checkbox("🔄 Jaarlijkse Keuring uitgevoerd (+1 jaar)")
        b_opt = st.checkbox("⚖️ Beproeving uitgevoerd")
        if b_opt: g_int = st.selectbox("Kies beproevingsinterval:", INTERVALS)
        
        if st.button("Wijzigingen toepassen"):
            if sel_ids:
                for oid in sel_ids:
                    idx = df_actueel[df_actueel['id'] == oid].index
                    acties = []
                    if n_loc != "Geen wijziging" and n_loc != df_actueel.at[idx, 'locatie'].values[0]:
                        df_actueel.at[idx, 'locatie'] = n_loc; acties.append(f"Naar {n_loc}")
                    if n_stat != "Geen wijziging" and n_stat != df_actueel.at[idx, 'status'].values[0]:
                        df_actueel.at[idx, 'status'] = n_stat; acties.append(f"Status: {n_stat}")
                    if k_opt and n_stat != "Afgekeurd (Gearchiveerd)":
                        n_k = (u_dt + timedelta(days=365)).strftime("%d-%m-%Y")
                        df_actueel.at[idx, 'keuringsdatum'] = n_k; acties.append(f"Gekeurd tot {n_k}")
                    if b_opt and n_stat != "Afgekeurd (Gearchiveerd)":
                        n_b = (u_dt + timedelta(days=365 if "1 jaar" in g_int else 4*365)).strftime("%d-%m-%Y") if "Geen" not in g_int else ""
                        df_actueel.at[idx, 'beproevingsdatum'] = n_b; acties.append(f"Beproefd tot {n_b}")
                    if acties:
                        nu_log = pd.DataFrame([{'datum': datetime.now().strftime("%d-%m-%Y %H:%M"), 'object_id': oid, 'type': df_actueel.at[idx, 'type'].values[0], 'actie': " / ".join(acties), 'details': 'Online'}])
                        df_historie = pd.concat([df_historie, nu_log], ignore_index=True)
                sla_data_op(df_actueel, df_historie)
                st.success("✅ Succesvol bijgewerkt!")
                st.rerun()
            else: st.error("❌ Kies minimaal één ID.")
            
    elif modus == "Nieuw Object Toevoegen":
        with st.form("i_form", clear_on_submit=True):
            n_id = st.text_input("Uniek ID Nummer:").strip()
            n_type = st.text_input("Type (bijv. Pullift, Hijsband):").strip()
            n_l = st.text_input("Locatie:").strip()
            k_init = st.date_input("Volgende Keuringsdatum:", vandaag + timedelta(days=365))
            i_keuze = st.selectbox("Beproevingsinterval:", INTERVALS)
            if st.form_submit_button("Object Opslaan"):
                if n_id and n_type and n_l:
                    b_str = (vandaag + timedelta(days=365 if "1 jaar" in i_keuze else 4*365)).strftime("%d-%m-%Y") if "Geen" not in i_keuze else ""
                    nieuwe_rij = pd.DataFrame([{'id': n_id, 'type': n_type, 'locatie': n_l, 'keuringsdatum': k_init.strftime("%d-%m-%Y"), 'beproevingsdatum': b_str, 'status': 'Actief'}])
                    df_actueel = pd.concat([df_actueel, nieuwe_rij], ignore_index=True)
                    nu_log = pd.DataFrame([{'datum': datetime.now().strftime("%d-%m-%Y %H:%M"), 'object_id': n_id, 'type': n_type, 'actie': 'Aangemaakt', 'details': 'Online'}])
                    df_historie = pd.concat([df_historie, nu_log], ignore_index=True)
                    sla_data_op(df_actueel, df_historie)
                    st.success("🎉 Toegevoegd!"); st.rerun()
                else: st.error("❌ Vul alle velden in!")

with rechts:
    st.subheader("📊 Overzichten")
    zoek = st.text_input("🔍 Snel zoeken:", key="z_uniek")
    df_g = df_actueel.copy()
    if zoek and not df_actueel.empty:
        df_g = df_g[df_g['id'].astype(str).str.contains(zoek,case=False) | df_g['type'].astype(str).str.contains(zoek,case=False) | df_g['locatie'].astype(str).str.contains(zoek,case=False)]
    
    t1, t2, t3, t4 = st.tabs(["Actuele Status", "Archief", "Locatie Grafiek", "Volledige Historie"])
    
    with t1:
        df_r = df_g[df_g['status'] != "Afgekeurd (Gearchiveerd)"] if not df_g.empty else pd.DataFrame()
        if not df_r.empty:
            st.dataframe(df_r[['id', 'type', 'locatie', 'keuringsdatum', 'Keur_Status', 'beproevingsdatum', 'Beproef_Status', 'status']], hide_index=True, use_container_width=True)
        else: st.info("Geen actieve objecten.")
        
    with t2:
        df_a = df_g[df_g['status'] == "Afgekeurd (Gearchiveerd)"] if not df_g.empty else pd.DataFrame()
        if not df_a.empty:
            st.dataframe(df_a[['id', 'type', 'locatie', 'keuringsdatum', 'beproevingsdatum', 'status']], hide_index=True, use_container_width=True)
            
            # APART STATUS OVERZICHT: Grafiek speciaal voor wat er in het archief ligt!
            st.write("---")
            st.write("📊 **Aantal afgekeurde hijsmiddelen per type product in het archief:**")
            type_counts_archief = df_a['type'].value_counts()
            st.bar_chart(type_counts_archief)
            st.write(type_counts_archief)
        else: st.info("Archief is leeg.")
        
    with t3:
        if not df_actueel.empty:
            st.write("**Aantal actieve hijsmiddelen per locatie (Vlootnummer):**")
            df_actief = df_actueel[df_actueel['status'] != "Afgekeurd (Gearchiveerd)"]
            if not df_actief.empty:
                loc_counts = df_actief['locatie'].value_counts()
                st.bar_chart(loc_counts)
                st.write(loc_counts)
        else: st.info("Geen data voor grafiek.")
            
    with t4:
        st.dataframe(df_historie.sort_index(ascending=False), use_container_width=True, hide_index=True)
