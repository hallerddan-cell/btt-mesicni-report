import streamlit as st
import pandas as pd
import io
import numpy as np

st.set_page_config(page_title="BTT Měsíční Reporty", layout="wide")

# --- POMOCNÁ FUNKCE PRO TVORBU KONTINGENČNÍ TABULKY (Kompaktní formát jako v Excelu) ---
def vytvor_kontingencni_tabulku(df, row_cols, agg_dict, ma_service_fee=True):
    # Agregace dat na nejnižší úrovni
    df_grouped = df.groupby(row_cols).agg(agg_dict).reset_index()
    
    # Výpočet průměrů na nejnižší úrovni
    if "Service fee" in df_grouped.columns and ma_service_fee:
        df_grouped["Průměr z Service fee2"] = np.where(df_grouped["Poč.dokl."] > 0, df_grouped["Service fee"] / df_grouped["Poč.dokl."], 0)
    if "Provize" in df_grouped.columns:
        df_grouped["Průměr z Provize2"] = np.where(df_grouped["Poč.dokl."] > 0, df_grouped["Provize"] / df_grouped["Poč.dokl."], 0)

    vysledne_radky = []
    
    # Zpracování pro jednoúrovňovou tabulku (např. Služby, Konsolidátoři)
    if len(row_cols) == 1:
        for _, row in df_grouped.iterrows():
            zaznam = {"Popisky řádků": str(row[row_cols[0]])}
            for col in agg_dict.keys(): zaznam[f"Součet z {col}"] = row[col]
            if ma_service_fee and "Service fee" in row: zaznam["Průměr z Service fee2"] = row["Průměr z Service fee2"]
            if "Provize" in row: zaznam["Průměr z Provize2"] = row["Průměr z Provize2"]
            vysledne_radky.append(zaznam)
            
    # Zpracování pro dvouúrovňovou tabulku s mezisoučty (např. Referent -> Služba)
    elif len(row_cols) == 2:
        hlavni_skupiny = df.groupby(row_cols[0]).agg(agg_dict).reset_index()
        if ma_service_fee and "Service fee" in hlavni_skupiny.columns:
            hlavni_skupiny["Průměr z Service fee2"] = np.where(hlavni_skupiny["Poč.dokl."] > 0, hlavni_skupiny["Service fee"] / hlavni_skupiny["Poč.dokl."], 0)
        if "Provize" in hlavni_skupiny.columns:
            hlavni_skupiny["Průměr z Provize2"] = np.where(hlavni_skupiny["Poč.dokl."] > 0, hlavni_skupiny["Provize"] / hlavni_skupiny["Poč.dokl."], 0)

        for _, subtotal_row in hlavni_skupiny.iterrows():
            # Řádek mezisoučtu (Nadřazená kategorie - např. Referent)
            skupina_nazev = str(subtotal_row[row_cols[0]])
            zaznam_subtotal = {"Popisky řádků": skupina_nazev}
            for col in agg_dict.keys(): zaznam_subtotal[f"Součet z {col}"] = subtotal_row[col]
            if ma_service_fee and "Service fee" in subtotal_row: zaznam_subtotal["Průměr z Service fee2"] = subtotal_row["Průměr z Service fee2"]
            if "Provize" in subtotal_row: zaznam_subtotal["Průměr z Provize2"] = subtotal_row["Průměr z Provize2"]
            vysledne_radky.append(zaznam_subtotal)
            
            # Podřazené detaily (např. Služby vnořené pod referenta s odsazením)
            detaily = df_grouped[df_grouped[row_cols[0]] == skupina_nazev]
            for _, detail_row in detaily.iterrows():
                zaznam_detail = {"Popisky řádků": f"   {detail_row[row_cols[1]]}"} # Odsazení jako v Excelu
                for col in agg_dict.keys(): zaznam_detail[f"Součet z {col}"] = detail_row[col]
                if ma_service_fee and "Service fee" in detail_row: zaznam_detail["Průměr z Service fee2"] = detail_row["Průměr z Service fee2"]
                if "Provize" in detail_row: zaznam_detail["Průměr z Provize2"] = detail_row["Průměr z Provize2"]
                vysledne_radky.append(zaznam_detail)

    # Řádek CELKOVÝ SOUČET
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
        
        # Filtry a datové sady
        df_btt = df[df["Název org."].astype(str).str.contains("BTT", na=False)] if "Název org." in df.columns else df
        df_helpdesk = df[df["Služba"].astype(str).str.contains("Helpdesk", case=False, na=False)] if "Služba" in df.columns else df
        
        # Základní metriky pro agregaci
        base_agg = {"Poč.dokl.": "sum", "Celkem": "sum"}
        agg_with_provize = {**base_agg, "Provize": "sum"}
        agg_with_service = {**base_agg, "Service fee": "sum"}
        agg_full = {**base_agg, "Service fee": "sum", "Provize": "sum"}
        
        # --- DEFINICE VŠECH 7 TABULEK (včetně zobrazení filtrů) ---
        tabulky_def = [
            {
                "nadpis": "1. SRPEN 2026",
                "filtry": [],
                "df": vytvor_kontingencni_tabulku(df, ["Služba"], agg_with_provize, ma_service_fee=False)
            },
            {
                "nadpis": "2. BTT",
                "filtry": [("Název org.", "(Vše)"), ("Jméno referenta", "(Vše)")],
                "df": vytvor_kontingencni_tabulku(df_btt, ["Služba"], agg_full)
            },
            {
                "nadpis": "3. REFERENTI",
                "filtry": [("Název org.", "(Vše)")],
                "df": vytvor_kontingencni_tabulku(df, ["Jméno referenta", "Služba"], agg_full)
            },
            {
                "nadpis": "4. KONSOLIDÁTOŘI",
                "filtry": [],
                "df": vytvor_kontingencni_tabulku(df, ["Název org."], agg_full)
            },
            {
                "nadpis": "5. OBCHOĎÁCI",
                "filtry": [],
                "df": vytvor_kontingencni_tabulku(df, ["Obchodní zástupce"], agg_with_provize, ma_service_fee=False)
            },
            {
                "nadpis": "6. HELPDESK",
                "filtry": [],
                "df": vytvor_kontingencni_tabulku(df_helpdesk, ["Název org."], agg_with_service, ma_service_fee=False)
            },
            {
                "nadpis": "7. KLIENTI BTT",
                "filtry": [("Jméno referenta", "(Vše)")],
                "df": vytvor_kontingencni_tabulku(df_btt, ["Název org.", "Služba"], agg_full)
            }
        ]

        st.success("Data úspěšně zpracována podle originální struktury kontingenčních tabulek.")
        
        # Zobrazení na webu
        for t in tabulky_def:
            st.subheader(t["nadpis"])
            for filter_name, filter_val in t["filtry"]:
                st.caption(f"_{filter_name}: {filter_val}_")
            st.dataframe(t["df"], use_container_width=True)

        # --- EXPORT DO EXCELU (PŘESNÝ LAYOUT NA 1 LISTU) ---
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            workbook = writer.book
            worksheet = workbook.add_worksheet('Výsledky BTT')
            writer.sheets['Výsledky BTT'] = worksheet
            
            # Formáty buněk
            format_bold = workbook.add_format({'bold': True})
            format_header = workbook.add_format({'bold': True, 'bottom': 1, 'bg_color': '#D9D9D9'})
            format_num = workbook.add_format({'num_format': '#,##0.00'})
            format_int = workbook.add_format({'num_format': '#,##0'})
            
            worksheet.set_column('A:A', 35) # Rozšíření sloupce pro Popisky řádků
            worksheet.set_column('B:G', 15)
            
            row_idx = 0
            for t in tabulky_def:
                # Filtr řádky
                for filter_name, filter_val in t["filtry"]:
                    worksheet.write(row_idx, 0, filter_name, format_bold)
                    worksheet.write(row_idx, 1, filter_val)
                    row_idx += 1
                
                # Prázdný řádek před tabulkou (pokud byly filtry)
                if t["filtry"]: row_idx += 1
                
                # Zápis hlaviček sloupců
                df_export = t["df"]
                for col_idx, col_name in enumerate(df_export.columns):
                    worksheet.write(row_idx, col_idx, col_name, format_header)
                row_idx += 1
                
                # Zápis dat
                for _, row_data in df_export.iterrows():
                    for col_idx, col_name in enumerate(df_export.columns):
                        val = row_data[col_name]
                        if pd.isna(val) or val == "":
                            worksheet.write(row_idx, col_idx, "")
                        elif "Popisky řádků" in col_name:
                            # Tučné písmo pro hlavní kategorie (neobsahují odsazení) a celkový součet
                            if not str(val).startswith("   ") or "Celkový součet" in str(val):
                                worksheet.write(row_idx, col_idx, str(val), format_bold)
                            else:
                                worksheet.write(row_idx, col_idx, str(val))
                        elif "Poč.dokl." in col_name:
                            worksheet.write_number(row_idx, col_idx, val, format_int)
                        else:
                            worksheet.write_number(row_idx, col_idx, val, format_num)
                    row_idx += 1
                
                # Mezera pod tabulkou (3 řádky)
                row_idx += 3

        st.divider()
        st.download_button(
            label="📥 Stáhnout opravený výkaz (Formát Kontingenční tabulky)",
            data=output.getvalue(),
            file_name="Vysledky_BTT_presny_format.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
elif heslo:
    st.error("Nesprávné heslo.")
