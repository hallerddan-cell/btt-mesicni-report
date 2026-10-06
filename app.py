import streamlit as st
import pandas as pd
import io

# Zabezpečení jednoduchým heslem (můžete upravit)
heslo = st.sidebar.text_input("Zadejte heslo", type="password")

if heslo == "TajneHeslo2026":
    st.title("Měsíční reportovací aplikace - BTT")
    
    uploaded_file = st.file_uploader("Nahrajte zdrojová data (Excel)", type=["xlsx"])
    
    if uploaded_file is not None:
        df = pd.read_excel(uploaded_file)
        
        # Kontrola potřebných sloupců
        required_cols = ["Služba", "Poč.dokl.", "Celkem", "Provize"]
        if all(col in df.columns for col in required_cols):
            # Výpočet kontingenční tabulky
            vysledky = df.groupby("Služba").agg({
                "Poč.dokl.": "sum",
                "Celkem": "sum",
                "Provize": "sum"
            }).reset_index()
            
            # Výpočet průměru z provize
            vysledky["Průměr z Provize2"] = vysledky["Provize"] / vysledky["Poč.dokl."]
            
            st.subheader("Výsledná tabulka")
            st.dataframe(vysledky)
            
            # Generování Excelu ke stažení
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                vysledky.to_excel(writer, index=False, sheet_name='Výsledky')
            
            st.download_button(
                label="📥 Stáhnout výsledky jako Excel",
                data=output.getvalue(),
                file_name="Vysledky_BTT_zpracovano.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("Nahraný soubor neobsahuje všechny potřebné sloupce (Služba, Poč.dokl., Celkem, Provize).")
elif heslo:
    st.sidebar.error("Nesprávné heslo.")
