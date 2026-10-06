import streamlit as st
import pandas as pd
import io
import numpy as np

st.set_page_config(page_title="BTT Měsíční Reporty", layout="wide")

# --- POMOCNÁ FUNKCE PRO TVORBU KONTINGENČNÍ TABULKY ---
def vytvor_kontingencni_tabulku(df, row_cols, agg_dict, ma_service_fee=True):
    if df.empty:
        return pd.DataFrame()
        
    # Agregace dat
    df_grouped = df.groupby(row_cols).agg(agg_dict).reset_index()
    
    if "Service fee" in df_grouped.columns and ma_service_fee:
        df_grouped["Průměr z Service fee2"] = np.where(df_grouped["Poč.dokl."] > 0, df_grouped["Service fee"] / df_grouped["Poč.dokl."], 0)
    if "Provize" in df_grouped.columns:
        df_grouped["Průměr z Provize2"] = np.where(df_grouped["Poč.dokl."] > 0, df_grouped["Provize"] / df_grouped["Poč.dokl."], 0)

    vysledne_radky = []
    
    if len(row_cols) == 1:
        for _, row in df_grouped.iterrows():
            zaznam = {"Popisky řádků": str(row[row_cols[0]])}
            for col in agg_dict.keys(): zaznam[f"Součet z {col}"] = row[col]
            if ma_service_fee and "Service fee" in row: zaznam["Průměr z Service fee2"] = row["Průměr z Service fee2"]
            if "Provize" in row: zaznam["Průměr z Provize2"] = row["Průměr z Provize2"]
            vysledne_radky.append(zaznam)
            
    elif len(row_cols) == 2:
        hlavni_skupiny = df.groupby(row_cols[0]).agg(agg_dict).reset_index()
        if ma_service_fee and "Service fee" in hlavni_skupiny.columns:
            hlavni_skupiny["Průměr z Service fee2"] = np.where(hlavni_skupiny["Poč.dokl."] > 0, hlavni_skupiny["Service fee"] / hlavni_skupiny["Poč.dokl."], 0)
        if "Provize" in hlavni_skupiny.columns:
            hlavni_skupiny["Průměr z Provize2"] = np.where(hlavni_skupiny["Poč.dokl."] > 0, hlavni_skupiny["Provize"] / hlavni_skupiny["Poč.dokl."], 0)

        for _, subtotal_row in hlavni_skupiny.iterrows():
            skupina_nazev = str(subtotal_row[row_cols[0]])
            zaznam_subtotal = {"Popisky řádků": skupina_nazev}
            for col in agg_dict.keys(): zaznam_subtotal[f"Součet z {col}"] = subtotal_row[col]
            if ma_service_fee and "Service fee" in subtotal_row: zaznam_subtotal["Průměr z Service fee2"] = subtotal_row["Průměr z Service fee2"]
            if "Provize" in subtotal_row: zaznam_subtotal["Průměr z Provize2"] = subtotal_row["Průměr z Provize2"]
            vysledne_radky.append(zaznam_subtotal)
            
            detaily = df_grouped[df_grouped[row_cols[0]] == skupina_nazev]
            for _, detail_row in detaily.iterrows():
                zaznam_detail = {"Popisky řádků": f"   {detail_row[row_cols[1]]}"}
                for col in agg_dict.keys(): zaznam_detail[f"Součet z {col}"] = detail_row[col]
                if ma_service_fee and "Service fee" in detail_row: zaznam_detail["Průměr z Service fee2"] = detail_row["Průměr z Service fee2"]
                if "Provize" in detail_row: zaznam_detail["Průměr z Provize2"] = detail_row["Průměr z Provize2"]
                vysledne_radky.append(zaznam_detail)

    celkem_dict = df.agg(agg_dict).to_dict()
    celkem_row = {"Popisky řádků": "Celkový součet"}
    for col in agg_dict.keys(): celkem_row[f"Součet z {col}"] = celkem_dict[col]
    
    if ma_service_fee and "Service fee" in celkem_dict:
        celkem_row["Průměr z Service fee2"] = celkem_dict["Service fee"] / celkem_dict["Poč.dokl."] if celkem_dict["Poč.dokl."] > 0 else 0
    if "Provize" in celkem_dict:
        celkem_row["Průměr z Provize2"] = celkem_dict["Provize"] / celkem_dict["Poč.dokl."] if celkem_dict["Poč.dokl."] > 0 else 0
        
    vysledne_radky.append(celkem_row)
    
    return pd.DataFrame(vysledne_radky)


# --- HLAVNÍ APLIKACE ---
heslo = st.sidebar.text_input("Zadejte heslo", type="password")

if heslo == "TajneHeslo2026":
    st.title("📊 BTT Měsíční Výkaz (Přesná kopie Excelu)")
    
    uploaded_file = st.file_uploader("Nahrajte zdrojová data (.xlsx)", type=["xlsx"])
    
    if uploaded_file is not None:
        df = pd.read_excel(uploaded_file)
        
        # Ošetření filtrů 
        df_btt = df[df["Název org."].astype(str).str.contains("BTT", na=False)] if "Název org." in df.columns else df
        df_helpdesk = df[df["Služba"].astype(str).str.contains("Helpdesk", case=False, na=False)] if "Služba" in df.columns else df
        
        # Filtrování obchodníků podle jmen z Vašeho excelu (Můžete přidat/ubrat další lidi sem)
        obchodaci = ['Jandošová Petra', 'Matějková Ivona', 'Nekola Tomáš', 'Třebický Tomáš']
        if "Jméno referenta" in df.columns:
            df_obchodaci = df[df["Jméno referenta"].isin(obchodaci)]
        else:
            df_obchodaci = df # Fallback

        # Agregace
        base_agg = {"Poč.dokl.": "sum", "Celkem": "sum"}
        agg_with_provize = {**base_agg, "Provize": "sum"}
        agg_with_service = {**base_agg, "Service fee": "sum"}
        agg_full = {**base_agg, "Service fee": "sum", "Provize": "sum"}
        
        # Definice 7 tabulek
        tabulky_def = []
        if "Služba" in df.columns:
            tabulky_def.append({"nadpis": "1. SRPEN 2026", "filtry": [], "df": vytvor_kontingencni_tabulku(df, ["Služba"], agg_with_provize, ma_service_fee=False)})
            tabulky_def.append({"nadpis": "2. BTT", "filtry": [("Název org.", "(Vše)"), ("Jméno referenta", "(Vše)")], "df": vytvor_kontingencni_tabulku(df_btt, ["Služba"], agg_full)})
            
            if "Jméno referenta" in df.columns:
                tabulky_def.append({"nadpis": "3. REFERENTI", "filtry": [("Název org.", "(Vše)")], "df": vytvor_kontingencni_tabulku(df, ["Jméno referenta", "Služba"], agg_full)})
                # OPRAVA TÉTO TABULKY -> POUŽIJEME 'Jméno referenta' ale pouze pro vyfiltrované obchodníky
                tabulky_def.append({"nadpis": "5. OBCHOĎÁCI", "filtry": [], "df": vytvor_kontingencni_tabulku(df_obchodaci, ["Jméno referenta"], agg_with_provize, ma_service_fee=False)})
                
            if "Název org." in df.columns:
                tabulky_def.insert(3, {"nadpis": "4. KONSOLIDÁTOŘI", "filtry": [], "df": vytvor_kontingencni_tabulku(df, ["Název org."], agg_full)})
                tabulky_def.append({"nadpis": "6. HELPDESK", "filtry": [], "df": vytvor_kontingencni_tabulku(df_helpdesk, ["Název org."], agg_with_service, ma_service_fee=False)})
                tabulky_def.append({"nadpis": "7. KLIENTI BTT", "filtry": [("Jméno referenta", "(Vše)")], "df": vytvor_kontingencni_tabulku(df_btt, ["Název org.", "Služba"], agg_full)})

        st.success("Data byla úspěšně zpracována.")
        
        # Zobrazení
        for t in tabulky_def:
            st.subheader(t["nadpis"])
            for f_name, f_val in t["filtry"]: st.caption(f"_{f_name}: {f_val}_")
            if not t["df"].empty: st.dataframe(t["df"], use_container_width=True)

        # Export (Na 1 list)
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            wb = writer.book
            ws = wb.add_worksheet('Výsledky BTT')
            writer.sheets['Výsledky BTT'] = ws
            
            f_bold = wb.add_format({'bold': True})
            f_hdr = wb.add_format({'bold': True, 'bottom': 1, 'bg_color': '#D9D9D9'})
            f_num = wb.add_format({'num_format': '#,##0.00'})
            f_int = wb.add_format({'num_format': '#,##0'})
            
            ws.set_column('A:A', 35)
            ws.set_column('B:G', 15)
            
            r_idx = 0
            for t in tabulky_def:
                for f_name, f_val in t["filtry"]:
                    ws.write(r_idx, 0, f_name, f_bold)
                    ws.write(r_idx, 1, f_val)
                    r_idx += 1
                if t["filtry"]: r_idx += 1
                
                df_export = t["df"]
                if not df_export.empty:
                    for c_idx, c_name in enumerate(df_export.columns):
                        ws.write(r_idx, c_idx, c_name, f_hdr)
                    r_idx += 1
                    
                    for _, r_data in df_export.iterrows():
                        for c_idx, c_name in enumerate(df_export.columns):
                            val = r_data[c_name]
                            if pd.isna(val) or val == "":
                                ws.write(r_idx, c_idx, "")
                            elif "Popisky řádků" in c_name:
                                if not str(val).startswith("   ") or "Celkový součet" in str(val):
                                    ws.write(r_idx, c_idx, str(val), f_bold)
                                else:
                                    ws.write(r_idx, c_idx, str(val))
                            elif "Poč.dokl." in c_name:
                                ws.write_number(r_idx, c_idx, val, f_int)
                            else:
                                ws.write_number(r_idx, c_idx, val, f_num)
                        r_idx += 1
                r_idx += 3

        st.divider()
        st.download_button("📥 Stáhnout opravený výkaz", data=output.getvalue(), file_name="Vysledky_BTT_hotovo.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
elif heslo:
    st.error("Nesprávné heslo.")
