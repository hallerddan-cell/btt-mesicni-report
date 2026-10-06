import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="BTT Měsíční Reporty", layout="wide")

# Funkce pro přidání řádku CELKEM s dopočtem průměrů
def pridat_celkem(df_in, group_cols):
    df_res = df_in.copy()
    
    # Výpočet průměrů pro jednotlivé řádky
    if "Poč.dokl." in df_res.columns:
        if "Service fee" in df_res.columns:
            df_res["Průměr z Service fee2"] = df_res["Service fee"] / df_res["Poč.dokl."]
        if "Provize" in df_res.columns:
            df_res["Průměr z Provize2"] = df_res["Provize"] / df_res["Poč.dokl."]

    # Výpočet součtového řádku
    celkem_dict = {}
    for col in df_res.columns:
        if col in group_cols:
            celkem_dict[col] = "CELKEM" if col == group_cols[0] else ""
        elif col in ["Poč.dokl.", "Celkem", "Service fee", "Provize"]:
            celkem_dict[col] = df_res[col].sum()
        elif col == "Průměr z Service fee2":
            celkem_dict[col] = (df_res["Service fee"].sum() / df_res["Poč.dokl."].sum()) if df_res["Poč.dokl."].sum() > 0 else 0
        elif col == "Průměr z Provize2":
            celkem_dict[col] = (df_res["Provize"].sum() / df_res["Poč.dokl."].sum()) if df_res["Poč.dokl."].sum() > 0 else 0
        else:
            celkem_dict[col] = ""

    df_celkem = pd.DataFrame([celkem_dict])
    return pd.concat([df_res, df_celkem], ignore_index=True)

# Zabezpečení
heslo = st.sidebar.text_input("Zadejte heslo", type="password")

if heslo == "TajneHeslo2026":
    st.title("📊 BTT Měsíční Výkaz – Všech 7 pohledů na jednom listu")
    
    uploaded_file = st.file_uploader("Nahrajte zdrojový Excel (.xlsx)", type=["xlsx"])
    
    if uploaded_file is not None:
        df = pd.read_excel(uploaded_file)
        
        # Kontrola základních sloupců
        req_cols = ["Služba", "Poč.dokl.", "Celkem", "Provize"]
        if all(col in df.columns for col in req_cols):
            
            # --- AGREGACE PRO VŠECH 7 TABULEK ---
            
            # 1. Celkový souhrn (SRPEN)
            t1 = df.groupby("Služba").agg({
                col: "sum" for col in ["Poč.dokl.", "Celkem", "Service fee", "Provize"] if col in df.columns
            }).reset_index()
            t1 = pridat_celkem(t1, ["Služba"])

            # 2. BTT
            df_btt = df[df["Název org."].astype(str).str.contains("BTT", case=False, na=False)] if "Název org." in df.columns else df
            t2 = df_btt.groupby("Služba").agg({
                col: "sum" for col in ["Poč.dokl.", "Celkem", "Service fee", "Provize"] if col in df.columns
            }).reset_index()
            t2 = pridat_celkem(t2, ["Služba"])

            # 3. Referenti
            group_ref = [c for c in ["Jméno referenta", "Služba"] if c in df.columns]
            if not group_ref: group_ref = ["Služba"]
            t3 = df.groupby(group_ref).agg({
                col: "sum" for col in ["Poč.dokl.", "Celkem", "Service fee", "Provize"] if col in df.columns
            }).reset_index()
            t3 = pridat_celkem(t3, group_ref)

            # 4. Konsolidátoři
            col_org = "Název org." if "Název org." in df.columns else "Služba"
            t4 = df.groupby(col_org).agg({
                col: "sum" for col in ["Poč.dokl.", "Celkem", "Service fee", "Provize"] if col in df.columns
            }).reset_index()
            t4 = pridat_celkem(t4, [col_org])

            # 5. Obchoďáci
            col_obch = "Obchodní zástupce" if "Obchodní zástupce" in df.columns else col_org
            t5 = df.groupby(col_obch).agg({
                col: "sum" for col in ["Poč.dokl.", "Celkem", "Provize"] if col in df.columns
            }).reset_index()
            t5 = pridat_celkem(t5, [col_obch])

            # 6. Helpdesk
            t6 = df.groupby(col_org).agg({
                col: "sum" for col in ["Poč.dokl.", "Celkem", "Service fee"] if col in df.columns
            }).reset_index()
            t6 = pridat_celkem(t6, [col_org])

            # 7. Klienti BTT
            group_klient = [c for c in ["Název org.", "Služba"] if c in df.columns]
            if not group_klient: group_klient = ["Služba"]
            t7 = df_btt.groupby(group_klient).agg({
                col: "sum" for col in ["Poč.dokl.", "Celkem", "Service fee", "Provize"] if col in df.columns
            }).reset_index()
            t7 = pridat_celkem(t7, group_klient)

            # --- ZOBRAZENÍ NA WEBU (ZÁLOŽKY) ---
            tabs = st.tabs(["1. Souhrn", "2. BTT", "3. Referenti", "4. Konsolidátoři", "5. Obchoďáci", "6. Helpdesk", "7. Klienti BTT"])
            sektory = [
                ("1. CELKOVÝ SOUHRN", t1), ("2. BTT", t2), ("3. REFERENTI", t3),
                ("4. KONSOLIDÁTOŘI", t4), ("5. OBCHOĎÁCI", t5), ("6. HELPDESK", t6), ("7. KLIENTI BTT", t7)
            ]
            
            for tab, (nazev, t_data) in zip(tabs, sektory):
                with tab:
                    st.subheader(nazev)
                    st.dataframe(t_data, use_container_width=True)

            # --- GENEROVÁNÍ EXCELU: VŠECHNY TABULKY NA JEDEN LIST ---
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                workbook = writer.book
                worksheet = workbook.add_worksheet('Výsledky BTT')
                writer.sheets['Výsledky BTT'] = worksheet
                
                # Formát pro nadpisy sekcí
                title_format = workbook.add_format({'bold': True, 'font_size': 14, 'font_color': '#1F4E78'})
                
                start_row = 0
                for nazev, t_data in sektory:
                    # Zápis nadpisu sekce
                    worksheet.write(start_row, 0, nazev, title_format)
                    start_row += 1
                    
                    # Zápis tabulky
                    t_data.to_excel(writer, sheet_name='Výsledky BTT', startrow=start_row, index=False)
                    start_row += len(t_data) + 3 # 3 volné řádky mezera mezi tabulkami

            st.divider()
            st.download_button(
                label="📥 Stáhnout kompletní výkaz (vše na 1 listu v Excelu)",
                data=output.getvalue(),
                file_name="Vysledky_BTT_kompletni_jednotny_list.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("Nahranému souboru chybí některé z požadovaných sloupců.")
elif heslo:
    st.sidebar.error("Nesprávné heslo.")
