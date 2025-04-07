
# Inkasso Automatisering - Web App backend med Streamlit

import pandas as pd
import streamlit as st
from io import BytesIO
from datetime import datetime

# --- Tilføj logo og adgangskode ---
LOGO_URL = "https://likvido.dk/wp-content/uploads/2020/10/Likvido_logo_blue.svg"
ADGANGSKODE = "likvido2025"

def vurder_faktura(dage, seneste_indbetaling):
    if pd.isna(dage):
        return "Ukendt"
    if dage > 500:
        return "OK"
    elif 300 < dage <= 500:
        return "OK" if seneste_indbetaling < pd.Timestamp.now() - pd.DateOffset(months=4) else "Afvent"
    elif 200 < dage <= 300:
        return "OK" if seneste_indbetaling < pd.Timestamp.now() - pd.DateOffset(months=4) else "Afvent"
    elif 100 < dage <= 200:
        return "OK" if seneste_indbetaling < pd.Timestamp.now() - pd.DateOffset(months=3) else "Afvent"
    elif 50 < dage <= 100:
        return "OK" if seneste_indbetaling < pd.Timestamp.now() - pd.DateOffset(days=30) else "Afvent"
    elif 0 < dage <= 50:
        return "OK" if seneste_indbetaling < pd.Timestamp.now() - pd.DateOffset(days=7) else "Afvent"
    else:
        return "Afvent"

def main():
    st.set_page_config(page_title="Likvido Inkasso Automation", layout="centered")
    st.image(LOGO_URL, width=200)
    st.title("Likvido Inkasso Automatisering")
    st.write("Upload de 3 nødvendige filer for at identificere inkassosager og oprydning.")

    # --- Adgangskontrol ---
    adgang = st.text_input("Indtast adgangskode for at fortsætte", type="password")
    if adgang != ADGANGSKODE:
        st.warning("Adgangskode påkrævet for at bruge appen.")
        st.stop()

    posteringer_file = st.file_uploader("1. Kundeindbetalinger (konto 5600 / 17110)", type=["xlsx"])
    debitor_file = st.file_uploader("2. Debitorsaldo", type=["xlsx"])
    faktura_file = st.file_uploader("3. Ubetalte fakturaer", type=["xlsx"])

    if posteringer_file and debitor_file and faktura_file:
        posteringer = pd.read_excel(posteringer_file, skiprows=5)
        posteringer = posteringer[posteringer["Type"] == "Kundeindbetaling"]
        posteringer["Dato"] = pd.to_datetime(posteringer["Dato"], errors='coerce')
        seneste = posteringer["Dato"].max()

        debitor = pd.read_excel(debitor_file, skiprows=5)
        for col in ["Efter 28 dage", "Saldo"]:
            debitor[col] = pd.to_numeric(debitor[col], errors='coerce')
        debitor_pos = debitor[(debitor["Saldo"] > 0) & (debitor["Efter 28 dage"] > 0)]

        faktura = pd.read_excel(faktura_file, skiprows=3)
        faktura["Antal dage forfalden"] = pd.to_numeric(faktura["Antal dage forfalden"], errors="coerce")
        faktura["Kundenr."] = faktura["Kundenr."].astype(str)
        debitor_pos["Nr."] = debitor_pos["Nr."].astype(str)

        merged = faktura.merge(debitor_pos, left_on="Kundenr.", right_on="Nr.", how="left")
        merged["Vurdering"] = merged["Antal dage forfalden"].apply(lambda x: vurder_faktura(x, seneste))

        klar = merged[merged["Vurdering"] == "OK"]
        opryd = debitor[debitor["Saldo"] <= 0]

        st.success(f"Seneste kundeindbetaling: {seneste.date()}")
        st.write("## Klar til inkasso")
        st.dataframe(klar)

        st.write("## Debitorer med negativ/0 saldo")
        st.dataframe(opryd)

        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            klar.to_excel(writer, sheet_name="Inkasso-kandidater", index=False)
            opryd.to_excel(writer, sheet_name="Oprydning - Saldo <= 0", index=False)
            debitor_pos.to_excel(writer, sheet_name="Debitorer +30 dage", index=False)
            posteringer.to_excel(writer, sheet_name="Kundeindbetalinger", index=False)
        output.seek(0)

        st.download_button(
            label="Download Excel Oversigt",
            data=output,
            file_name="Inkasso_oversigt.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.write("## Klar mail til bogholder")
        kunder = ", ".join(opryd['Nr.'].dropna().astype(str).unique())
        mail = f'''Emne: Oprydning af debitorer - handling påkrævet\n\nHej [bogholder],\n\nJeg er i gang med at gennemgå vores debitorer.\nSeneste bogførte indbetaling er fra {seneste.date()}.\nVil du bogføre nye indbetalinger og rydde op i følgende kunder:\n{kunder}\n\nBedste hilsner\n[Dit navn]'''
        st.text_area("Oprydningsmail", mail, height=200)

if __name__ == "__main__":
    main()
