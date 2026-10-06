import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="BTT Měsíční Reporty", layout="wide")

# Zabezpečení heslem
heslo = st.sidebar.text_input("Zadejte heslo", type="password")

if heslo == "TajneHeslo2026":
    st.title("📊 Měsíční reportovací aplikace - BTT (Kompletní výkazy)")
    
    uploaded_file = st.file_uploader("Nahrajte zdrojová data (Excel .xlsx)", type=["xlsx"])
    
    if uploaded_file is not None:
        df = pd.read_excel(uploaded_file)
        
        required_cols = ["Služba", "Poč.dokl.", "Celkem", "Provize"]
        if all(col in df.columns for col in required_cols):
            
            # --- TABULKA 1: Seskupení dle Služeb ---
            tabulka_sluzby = df.groupby("Služba").agg({
                "Poč.dokl.": "sum",
                "Celkem": "sum",
                "Provize": "sum"
            }).reset_index()
            
            # Výpočet průměru z provize
            tabulka_sluzby["Průměr z Provize2"] = tabulka_sluzby["Provize"] / tabulka_sluzby["Poč.dokl."]
            
            # Přidání součtového řádku CELKEM
            celkem_dokl = tabulka_sluzby["Poč.dokl."].sum()
            celkem_castka = tabulka_sluzby["Celkem"].sum()
            celkem_provize = tabulka_sluzby["Provize"].sum()
            prumer_provize_celkem = celkem_provize / celkem_dokl if celkem_dokl > 0 else 0
            
            radek_celkem = pd.DataFrame([{
                "Služba": "CELKEM / SOUČET",
                "Poč.dokl.": celkem_dokl,
                "Celkem": celkem_castka,
                "Provize": celkem_provize,
                "Průměr z Provize2": prumer_provize_celkem
            }])
            
            tabulka_sluzby_s_celkem = pd.concat([tabulka_sluzby, radek_celkem], ignore_index=True)
            
            # --- TABULKA 2: Celkový souhrn (Přehledových 5 ukazatelů) ---
            souhrn_df = pd.DataFrame([{
                "Ukazatel": "Celkový počet dokladů", "Hodnota": f"{celkem_dokl:,.0f}".replace(",", " ")
            }, {
                "Ukazatel": "Celková částka (Celkem)", "Hodnota": f"{celkem_castka:,.2f} Kč".replace(",", " ")
            }, {
                "Ukazatel": "Celková Provize", "Hodnota": f"{celkem_provize:,.2f} Kč".replace(",", " ")
            }, {
                "Ukazatel": "Průměrná provize na doklad", "Hodnota": f"{prumer_provize_celkem:,.2f} Kč".replace(",", " ")
            }])

            # --- ZOBRAZENÍ NA WEBU POMOCÍ ZÁLOŽEK ---
            tab1, tab2, tab3 = st.tabs(["📌 Přehled dle Služeb", "📈 Celkový Souhrn", "📄 Zdrojová Data"])
            
            with tab1:
                st.subheader("Výsledná tabulka (Služby)")
                st.dataframe(
                    tabulka_sluzby_s_celkem.style.format({
                        "Poč.dokl.": "{:,.0f}",
                        "Celkem": "{:,.2f} Kč",
                        "Provize": "{:,.2f} Kč",
                        "Průměr z Provize2": "{:,.2f} Kč"
                    }),
                    use_container_width=True
                )

            with tab2:
                st.subheader("Celkové souhrnné ukazatele za měsíc")
                st.table(souhrn_df)
                
            with tab3:
                st.subheader("Nahraná zdrojová data")
                st.dataframe(df, use_container_width=True)

            # --- EXPORT VŠECH TABULEK DO EXCELU (Více listů) ---
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                tabulka_sluzby_s_celkem.to_excel(writer, index=False, sheet_name='Výsledky Služby')
                souhrn_df.to_excel(writer, index=False, sheet_name='Celkový Souhrn')
                df.to_excel(writer, index=False, sheet_name='Zdrojová Data')
            
            st.divider()
            st.download_button(
                label="📥 Stáhnout komplet report do Excelu (se všemi listy)",
                data=output.getvalue(),
                file_name="Vysledky_BTT_konecna_verze.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("Nahranému souboru chybí některé z požadovaných sloupců: Služba, Poč.dokl., Celkem, Provize.")
elif heslo:
    st.sidebar.error("Nesprávné heslo.")
Jak kód na GitHubu aktualizovat:
