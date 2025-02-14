import streamlit as st
import pandas as pd
import io

def daten_zusammenfuehren(fehlmengen_df, bestellungen_df):
    """Fügt Daten aus der Bestellungs-Excel-Datei in die Fehlmengen-CSV-Datei ein."""

    # Spaltennamen in bestellungen_df anpassen
    bestellungen_df = bestellungen_df.rename(columns={'Artikelnr.': 'Artikelnummer'})

    # Schleife über die Zeilen des fehlmengen_df
    for index, row in fehlmengen_df.iterrows():
        artikelnummer = row['Artikelnummer']
        if pd.isna(artikelnummer):  # Überspringe leere Zeilen
            continue

        # Filtert die Bestellungen nach Artikelnummer und Bedingungen
        passende_bestellungen = bestellungen_df[
            (bestellungen_df['Artikelnummer'] == artikelnummer) &
            (bestellungen_df['Geliefert'] == 0) &
            (bestellungen_df['Offen'] == bestellungen_df['Menge']) &  # 'Menge' statt 'Gesamtmenge'
            (bestellungen_df['Lieferdatum'] >= pd.Timestamp('today'))
        ]

        if not passende_bestellungen.empty:
            # Nimmt die erste passende Bestellung
            bestellung = passende_bestellungen.iloc

            fehlmengen_df.loc[index, 'Ist Bestellt?'] = 'Ja'
            fehlmengen_df.loc[index, 'Menge'] = bestellung['Menge']
            fehlmengen_df.loc[index, 'Lieferdatum'] = bestellung['Lieferdatum']
            fehlmengen_df.loc[index, 'Lieferant'] = bestellung['Kurzbezeichnung']  # 'Kurzbezeichnung' statt 'Lieferant'
            fehlmengen_df.loc[index, 'Bestellung'] = bestellung['Belegnr.']
        else:
            fehlmengen_df.loc[index, 'Ist Bestellt?'] = 'Nein'

    return fehlmengen_df

# Streamlit-App
st.title('Datenzusammenführung')

# Datei-Uploads
fehlmengen_file = st.file_uploader('Fehlmengen-CSV hochladen', type='csv')
bestellungen_file = st.file_uploader('Offene Bestellungen Excel hochladen', type=['xlsx'])

if fehlmengen_file and bestellungen_file:
    try:
        # Spaltennamen in Zeile 2 einlesen (0-basierter Index)
        fehlmengen_df = pd.read_csv(fehlmengen_file, sep=';', header=2)
        bestellungen_df = pd.read_excel(bestellungen_file, engine='openpyxl', header=2)

        # Daten zusammenführen
        ergebnis_df = daten_zusammenfuehren(fehlmengen_df.copy(), bestellungen_df)

        # Ergebnis anzeigen und Download-Link
        st.write('Ergebnis:')
        st.dataframe(ergebnis_df)

        # Downloadlink als Bytes erstellen
        buffer = io.BytesIO()
        ergebnis_df.to_csv(buffer, index=False, sep=';')
        st.download_button(
            label="Download",
            data=buffer,
            file_name="ergebnis.csv",
            mime="text/csv"
        )

    except Exception as e:
        st.error(f'Ein Fehler ist aufgetreten: {e}')
