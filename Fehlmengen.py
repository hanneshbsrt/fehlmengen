import streamlit as st
import pandas as pd
import openpyxl
import pytesseract
from PIL import Image
import re
import io
from google.cloud import vision
from google.oauth2 import service_account
import pandas.errors # Pandas Fehler-Modul importieren


# ... (Tesseract Pfad - optional) ...


@st.cache_data
def datei_inspektion_und_anpassung(uploaded_file, dateityp):
    """
    Liest Excel-Dateien ein und verwendet Zeile 3 als Spaltenüberschrift.
    Die ersten zwei Zeilen der Excel-Datei werden ignoriert.
    Gibt dem Benutzer die Möglichkeit, Spaltennamen anzupassen.

    Args:
        uploaded_file (streamlit.UploadedFile): Hochgeladene Datei.
        dateityp (str): Dateityp ('bestaende_excel' oder 'offene_bestellungen_excel').

    Returns:
        pandas.DataFrame: DataFrame der eingelesenen Daten oder None bei Fehler.
    """
    if uploaded_file is None:
        return None

    df = None
    fehlermeldung = None

    if dateityp == 'bestaende_excel' or dateityp == 'offene_bestellungen_excel': # Dateityp-Optionen für Excel
        try:
            df = pd.read_excel(uploaded_file, engine='openpyxl', header=2, skiprows=[0, 1]) # Header Zeile 3 (Index 2), ignoriere Zeilen 1 und 2 (Indizes 0 und 1)
            st.info(f"Excel-Datei erfolgreich gelesen (Spaltenüberschrift in Zeile 3, erste zwei Zeilen ignoriert).") # Info für Benutzer aktualisiert
        except Exception as e:
            fehlermeldung = f"Fehler beim Lesen der Excel-Datei: {e}. \n\nMögliche Ursachen: Datei ist beschädigt, falsches Format oder Spaltenüberschrift nicht in Zeile 3." # Fehlermeldung angepasst
    else:
        fehlermeldung = f"Unerwarteter Dateityp: {dateityp}.  Es werden nur Excel-Dateien für Bestände und offene Bestellungen erwartet." # Fehlermeldung für unerwarteten Dateityp

    if fehlermeldung:
        st.error(fehlermeldung)
        return None

    if df is not None:
        st.subheader(f"Vorschau der gelesenen Daten ({dateityp}, erste 5 Zeilen, **Spaltenüberschriften aus Zeile 3**):") # Dateityp und Hinweis auf Zeile 3 in Vorschau anzeigen
        st.dataframe(df.head())

        spaltennamen_neu = st.multiselect(
            f"Spaltennamen überprüfen und ggf. anpassen ({dateityp}, wähle korrekte Spalten aus, **aus Zeile 3 der Excel-Datei**):", # Dateityp und Hinweis auf Zeile 3 im Label anzeigen
            options=df.columns.tolist(),
            default=df.columns.tolist(), # Standardmäßig alle Spalten auswählen
            key=f"spaltenauswahl_{dateityp}_{uploaded_file.name}" # Eindeutiger Key für Multiselect
        )

        if spaltennamen_neu and len(spaltennamen_neu) == len(df.columns): # Sicherstellen, dass Spalten ausgewählt wurden und Anzahl stimmt
            df.columns = spaltennamen_neu # Spaltennamen im DataFrame aktualisieren
            st.success("Spaltennamen angepasst.")
            return df
        elif spaltennamen_neu:
            st.warning("Bitte wähle die korrekte Anzahl an Spaltennamen aus, die der Anzahl der Spalten in der Datei entspricht.")
            return None # DataFrame nicht zurückgeben, da Spaltenauswahl unvollständig
        else:
            st.warning("Es wurden keine Spaltennamen ausgewählt. Verwende Original-Spaltennamen.")
            return df # DataFrame mit Original-Spaltennamen zurückgeben
    else:
        return None # Fehlerfall, kein DataFrame


@st.cache_data
def artikel_stammdaten_lesen(uploaded_file):
    """Liest Artikelstammdaten aus Excel mit datei_inspektion_und_anpassung."""
    df_bestand = datei_inspektion_und_anpassung(uploaded_file, 'bestaende_excel') # Dateityp angepasst
    if df_bestand is None:
        return None

    artikel_stammdaten = {}
    for index, row in df_bestand.iterrows():
        artikelnummer = str(row['Artikel']) # Spaltenname 'Artikel'
        artikel_name = row['Kurzbezeichnung'] # Spaltenname 'Kurzbezeichnung'
        bestand_menge = str(row['Bestand']) # Spaltenname 'Bestand'
        bestand_einheit = row['ME'] # Spaltenname 'ME'
        bestand_gesamt = f"{bestand_menge} {bestand_einheit}"
        artikel_stammdaten[artikelnummer] = {
            "name": artikel_name,
            "bestand": bestand_gesamt
        }
    return artikel_stammdaten


@st.cache_data
def offene_bestellungen_lesen(uploaded_file):
    """Liest offene Bestellungen aus Excel mit datei_inspektion_und_anpassung."""
    return datei_inspektion_und_anpassung(uploaded_file, 'offene_bestellungen_excel') # Dateityp angepasst


def ist_bestellt(artikelnummer, offene_bestellungen_df):
    """... (Funktion ist_bestellt - Spaltennamen anpassen!) ..."""
    bestellungen_artikel = offene_bestellungen_df[offene_bestellungen_df['Artikelnr.'] == artikelnummer] # Spaltenname 'Artikelnr.'
    if bestellungen_artikel.empty:
        return False, None

    for index, bestellung_artikel_zeile in bestellungen_artikel.iterrows():
        belegnummer = bestellung_artikel_zeile['Belegnr.']
        bestellung_df = offene_bestellungen_df[offene_bestellungen_df['Belegnr.'] == belegnummer]
        alle_geliefert_null = (bestellung_df['Geliefert'] == 0).all()
        if alle_geliefert_null:
            return True, bestellung_df

    return False, None

def excel_tabelle_erstellen(artikelnummern, artikel_stammdaten, offene_bestellungen_df):
    """... (Funktion excel_tabelle_erstellen - Spaltennamen anpassen!) ..."""
    ausgabe_daten = []
    for artikelnummer in artikelnummern:
        stammdaten = artikel_stammdaten.get(artikelnummer)
        if stammdaten:
            name = stammdaten['name']
            bestand = stammdaten['bestand']
        else:
            name = "Artikelname nicht gefunden"
            bestand = "Bestand nicht gefunden"

        bestellt, bestellung_daten = ist_bestellt(artikelnummer, offene_bestellungen_df)
        if bestellt and bestellung_daten is not None and not bestellung_daten.empty:
            # Annahme: Erste Zeile der Bestellung enthält relevante Daten (Menge, Lieferdatum, etc.)
            bestell_zeile = bestellung_daten.iloc[0]
            menge = bestell_zeile['Menge']
            lieferdatum_roh = bestell_zeile['Lieferdatum']
            lieferdatum = pd.to_datetime(lieferdatum_roh, format='%d.%m.%Y').strftime('%d.%m.%Y') if isinstance(lieferdatum_roh, str) else lieferdatum_roh.strftime('%d.%m.%Y') if pd.notnull(lieferdatum_roh) else ""
            bearbeiter = bestell_zeile['Bearbeiter']
            belegnummer = bestell_zeile['Belegnr.']
            ist_bestellt_text = "ja"
        else:
            ist_bestellt_text = "nein"
            menge = ""
            lieferdatum = ""
            bearbeiter = ""
            belegnummer = ""

        ausgabe_daten.append([
            artikelnummer, name, bestand, ist_bestellt_text, menge, lieferdatum, bearbeiter, belegnummer
        ])

    ausgabe_df = pd.DataFrame(ausgabe_daten, columns=[
        "Art.Nr.", "Name", "Bestand", "Bestellt?", "Menge", "Lieferdatum", "Bearbeiter", "Belegnummer"
    ])
    return ausgabe_df


def artikelnummern_aus_bildern_erkennen_gcv(uploaded_files):
    """... (Funktion artikelnummern_aus_bildern_erkennen_gcv - unverändert) ..."""
    artikelnummern = []
    artikelnummer_muster = re.compile(r"A\d{5}")  # Dein Artikelnummernmuster

    credentials = service_account.Credentials.from_service_account_info(st.secrets["GOOGLE_APPLICATION_CREDENTIALS"])
    client = vision.ImageAnnotatorClient(credentials=credentials)

    for uploaded_file in uploaded_files:
        try:
            img = Image.open(uploaded_file)
            st.image(img, caption=f"Etikettenbild: {uploaded_file.name}", width=300)

            image = vision.Image(content=uploaded_file.getvalue())
            response = client.text_detection(image=image)
            erkannter_text = response.text_annotations[0].description if response.text_annotations else ""

            st.write(f"Erkannter Text (Google Cloud Vision API):\n```\n{erkannter_text}\n```")

            gefundene_artikelnummern = artikelnummer_muster.findall(erkannter_text)

            if gefundene_artikelnummern:
                beste_artikelnummer = gefundene_artikelnummern[0]
                antwort = st.radio(f"Artikelnummer in **{uploaded_file.name}** erkannt als: **{beste_artikelnummer}**. Korrekt?", ('Ja', 'Nein'), horizontal=True, key=f"radio_{uploaded_file.name}")
                if antwort == 'Ja':
                    artikelnummern.append(beste_artikelnummer)
                else:
                    manuelle_eingabe = st.text_input(f"Bitte gib die korrekte Artikelnummer für **{uploaded_file.name}** manuell ein:", key=f"manual_input_{uploaded_file.name}")
                    if manuelle_eingabe:
                        artikelnummern.append(manuelle_eingabe)
            else:
                manuelle_eingabe = st.text_input(f"Artikelnummer in **{uploaded_file.name}** konnte nicht erkannt werden. Bitte manuell eingeben:", key=f"manual_input_{uploaded_file.name}")
                if manuelle_eingabe:
                    artikelnummern.append(manuelle_eingabe)

        except Exception as e:
            st.error(f"Fehler beim Verarbeiten von {uploaded_file.name} mit Google Cloud Vision API: {e}")
            st.error(f"Fehlerdetails: {e}")
            manuelle_eingabe = st.text_input(f"Artikelnummer für **{uploaded_file.name}** manuell eingeben (Fehlerfall):", key=f"manual_input_error_{uploaded_file.name}")
            if manuelle_eingabe:
                artikelnummern.append(manuelle_eingabe)

    return artikelnummern


def main():
    st.title("Lagerbestandsautomatisierung")

    st.header("1. Etikettenbilder hochladen")
    uploaded_image_files = st.file_uploader("Etikettenbilder hochladen", type=["jpg", "jpeg", "png"], accept_multiple_files=True)

    artikelnummern_etiketten = []
    if uploaded_image_files:
        artikelnummern_etiketten = artikelnummern_aus_bildern_erkennen_gcv(uploaded_image_files)

        if artikelnummern_etiketten:
            st.success("Artikelnummernerkennung abgeschlossen (Google Cloud Vision API verwendet)!")
            st.write("Erkannte und validierte Artikelnummern von Etiketten:")
            st.write(artikelnummern_etiketten)
        else:
            st.warning("Keine Artikelnummern von den Etiketten extrahiert.")


    st.header("2. Dateien hochladen")
    bestaende_excel_file = st.file_uploader("Bestände Excel-Datei hochladen (Excel, *.xlsx, *.xls, **Spaltenüberschrift in Zeile 3**)", type=["xlsx", "xls"]) # Hinweis auf Excel-Format und Zeile 3
    offene_bestellungen_excel_file = st.file_uploader("Offene Bestellungen Excel-Datei hochladen (Excel, *.xlsx, *.xls, **Spaltenüberschrift in Zeile 3**)", type=["xlsx", "xls"]) # Hinweis auf Excel-Format und Zeile 3

    artikel_stammdaten = None # Initialisieren außerhalb der if-Bedingung
    offene_bestellungen_df = None # Initialisieren außerhalb der if-Bedingung

    if bestaende_excel_file:
        artikel_stammdaten = artikel_stammdaten_lesen(bestaende_excel_file) # Nutze bestaende_excel_file

    if offene_bestellungen_excel_file:
        offene_bestellungen_df = offene_bestellungen_lesen(offene_bestellungen_excel_file) # Nutze offene_bestellungen_excel_file


    if artikel_stammdaten and offene_bestellungen_df is not None and artikelnummern_etiketten:
        ausgabe_df = excel_tabelle_erstellen(artikelnummern_etiketten, artikel_stammdaten, offene_bestellungen_df)

        st.header("3. Ergebnis-Tabelle")
        st.dataframe(ausgabe_df)

        output_excel_file = io.BytesIO()
        with pd.ExcelWriter(output_excel_file, engine='openpyxl') as writer:
            ausgabe_df.to_excel(writer, index=False, sheet_name='Lagerbestand')
        output_excel_file.seek(0)

        st.download_button(
            label="Excel-Tabelle herunterladen",
            data=output_excel_file,
            file_name="lager_bestand_liste.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    elif artikelnummern_etiketten or bestaende_excel_file or offene_bestellungen_excel_file: # Warnungen angepasst
        if not artikelnummern_etiketten and uploaded_image_files:
            st.warning("Bitte validiere oder gib die Artikelnummern aus den Etikettenbildern ein, bevor du die Dateien hochlädst.")
        if not bestaende_excel_file:
            st.warning("Bitte lade die Bestände Excel-Datei hoch.") # Warnung für Bestände Excel
        if not offene_bestellungen_excel_file:
            st.warning("Bitte lade die Offene Bestellungen Excel-Datei hoch.") # Warnung für Offene Bestellungen Excel


if __name__ == "__main__":
    main()
