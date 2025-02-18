import streamlit as st
import pandas as pd
import openpyxl
import pytesseract
from PIL import Image
import re
import io  # Für In-Memory-Dateioperationen mit Streamlit


# Pfad zu Tesseract OCR Engine (ggf. anpassen, falls nicht im Systempfad)
# pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe' # Beispielpfad Windows


@st.cache_data  # Cache die Funktion, um wiederholtes Einlesen zu vermeiden
def artikel_stammdaten_lesen(uploaded_file):
    """
    Liest Artikelstammdaten aus einer hochgeladenen Excel-Datei.

    Args:
        uploaded_file (streamlit.UploadedFile): Hochgeladene Excel-Datei.

    Returns:
        dict: Dictionary mit Artikelstammdaten, Schlüssel ist Artikelnummer.
              Oder None bei Fehler.
    """
    if uploaded_file is None:
        return None

    try:
        df = pd.read_excel(uploaded_file)
        artikel_stammdaten = {}
        for index, row in df.iterrows():
            artikelnummer = str(row['Artikelnummer'])  # Sicherstellen, dass Artikelnummer als String behandelt wird
            artikel_name = row['Artikelname']
            bestand_menge = str(row['Bestand Menge'])  # Als String behandeln, da es mit Einheit kombiniert wird
            bestand_einheit = row['Bestand Einheit']
            bestand_gesamt = f"{bestand_menge} {bestand_einheit}"
            artikel_stammdaten[artikelnummer] = {
                "name": artikel_name,
                "bestand": bestand_gesamt
            }
        return artikel_stammdaten
    except Exception as e:
        st.error(f"Fehler beim Lesen der Artikelstammdaten-Datei: {e}")
        return None

@st.cache_data  # Cache die Funktion, um wiederholtes Einlesen zu vermeiden
def offene_bestellungen_lesen(uploaded_file):
    """
    Liest offene Bestellungen aus einer hochgeladenen CSV-Datei.

    Args:
        uploaded_file (streamlit.UploadedFile): Hochgeladene CSV-Datei.

    Returns:
        pandas.DataFrame: DataFrame mit offenen Bestellungen.
                         Oder None bei Fehler.
    """
    if uploaded_file is None:
        return None
    try:
        # Verwende io.StringIO, um direkt aus dem UploadedFile-Objekt zu lesen
        csv_string_data = io.StringIO(uploaded_file.getvalue().decode("utf-8"))
        df = pd.read_csv(csv_string_data, encoding='utf-8')
        # Konvertiere 'Geliefert' Spalte zu numerisch, Fehler werden zu NaN
        df['Geliefert'] = pd.to_numeric(df['Geliefert'], errors='coerce').fillna(0)
        return df
    except Exception as e:
        st.error(f"Fehler beim Lesen der Datei mit offenen Bestellungen: {e}")
        return None

def ist_bestellt(artikelnummer, offene_bestellungen_df):
    """
    Prüft, ob ein Artikel bestellt ist (alle Artikel einer Bestellung 'Geliefert' = 0).

    Args:
        artikelnummer (str): Artikelnummer.
        offene_bestellungen_df (pandas.DataFrame): DataFrame mit offenen Bestellungen.

    Returns:
        tuple: (bool, pandas.DataFrame or None) - True, Bestellung DataFrame wenn bestellt,
               False, None wenn nicht bestellt oder Fehler.
    """
    bestellungen_artikel = offene_bestellungen_df[offene_bestellungen_df['Artikelnummer'] == artikelnummer]
    if bestellungen_artikel.empty:
        return False, None

    for index, bestellung_artikel_zeile in bestellungen_artikel.iterrows():
        belegnummer = bestellung_artikel_zeile['Belegnr.']
        bestellung_df = offene_bestellungen_df[offene_bestellungen_df['Belegnr.'] == belegnummer]
        alle_geliefert_null = (bestellung_df['Geliefert'] == 0).all()
        if alle_geliefert_null:
            return True, bestellung_df  # Rückgabe der passenden Bestellung

    return False, None

def excel_tabelle_erstellen(artikelnummern, artikel_stammdaten, offene_bestellungen_df):
    """
    Erstellt die Pandas DataFrame für die Excel-Ausgabetabelle.

    Args:
        artikelnummern (list): Liste der Artikelnummern.
        artikel_stammdaten (dict): Dictionary mit Artikelstammdaten.
        offene_bestellungen_df (pandas.DataFrame): DataFrame mit offenen Bestellungen.

    Returns:
        pandas.DataFrame: DataFrame für die Ausgabetabelle.
    """
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
            lieferdatum = pd.to_datetime(lieferdatum_roh, format='%d.%m.%Y').strftime('%d.%m.%Y') if isinstance(lieferdatum_roh, str) else lieferdatum_roh.strftime('%d.%m.%Y') if pd.notnull(lieferdatum_roh) else ""  # Formatierung und Fehlerbehandlung
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


def artikelnummern_aus_bildern_erkennen(uploaded_files):
    """
    Erkennt Artikelnummern aus hochgeladenen Bildern mit OCR und Benutzerinteraktion.

    Args:
        uploaded_files (list): Liste von Streamlit UploadedFile-Objekten (Bilder).

    Returns:
        list: Liste der erkannten und validierten Artikelnummern.
    """
    artikelnummern = []
    # Artikelnummer beginnt mit A und hat dann fünf Zahlen (z.B. A04607)
    artikelnummer_muster = re.compile(r"A\d{5}")

    for uploaded_file in uploaded_files:
        try:
            img = Image.open(uploaded_file)
            st.image(img, caption=f"Etikettenbild: {uploaded_file.name}", width=300) # Bild in Streamlit anzeigen
            erkannter_text = pytesseract.image_to_string(img, lang='deu')
            st.write(f"Erkannter Text:\n```\n{erkannter_text}\n```") # Erkannten Text anzeigen

            gefundene_artikelnummern = artikelnummer_muster.findall(erkannter_text)

            if gefundene_artikelnummern:
                beste_artikelnummer = gefundene_artikelnummern[0]
                antwort = st.radio(f"Artikelnummer in **{uploaded_file.name}** erkannt als: **{beste_artikelnummer}**. Korrekt?", ('Ja', 'Nein'), horizontal=True, key=f"radio_{uploaded_file.name}") # Key hinzugefügt
                if antwort == 'Ja':
                    artikelnummern.append(beste_artikelnummer)
                else:
                    manuelle_eingabe = st.text_input(f"Bitte gib die korrekte Artikelnummer für **{uploaded_file.name}** manuell ein:", key=f"manual_input_{uploaded_file.name}")
                    if manuelle_eingabe: # Sicherstellen, dass der Benutzer etwas eingegeben hat
                        artikelnummern.append(manuelle_eingabe)
            else:
                manuelle_eingabe = st.text_input(f"Artikelnummer in **{uploaded_file.name}** konnte nicht erkannt werden. Bitte manuell eingeben:", key=f"manual_input_{uploaded_file.name}")
                if manuelle_eingabe: # Sicherstellen, dass der Benutzer etwas eingegeben hat
                    artikelnummern.append(manuelle_eingabe)


        except Exception as e:
            st.error(f"Fehler beim Verarbeiten von {uploaded_file.name}: {e}")
            manuelle_eingabe = st.text_input(f"Artikelnummer für **{uploaded_file.name}** manuell eingeben (Fehlerfall):", key=f"manual_input_error_{uploaded_file.name}")
            if manuelle_eingabe: # Sicherstellen, dass der Benutzer etwas eingegeben hat
                artikelnummern.append(manuelle_eingabe)

    return artikelnummern


def main():
    st.title("Lagerbestandsautomatisierung")

    st.header("1. Etikettenbilder hochladen")
    uploaded_image_files = st.file_uploader("Etikettenbilder hochladen", type=["jpg", "jpeg", "png"], accept_multiple_files=True)

    artikelnummern_etiketten = []
    if uploaded_image_files:
        artikelnummern_etiketten = artikelnummern_aus_bildern_erkennen(uploaded_image_files)

        if artikelnummern_etiketten:
            st.success("Artikelnummernerkennung abgeschlossen!")
            st.write("Erkannte und validierte Artikelnummern von Etiketten:")
            st.write(artikelnummern_etiketten)
        else:
            st.warning("Keine Artikelnummern von den Etiketten extrahiert.")


    st.header("2. Dateien hochladen")
    excel_file = st.file_uploader("Artikelstammdaten Excel-Datei hochladen", type=["xlsx", "xls"])
    csv_file = st.file_uploader("Offene Bestellungen CSV-Datei hochladen", type=["csv"])

    if excel_file and csv_file and artikelnummern_etiketten:
        artikel_stammdaten = artikel_stammdaten_lesen(excel_file)
        offene_bestellungen_df = offene_bestellungen_lesen(csv_file)

        if artikel_stammdaten and offene_bestellungen_df is not None:
            ausgabe_df = excel_tabelle_erstellen(artikelnummern_etiketten, artikel_stammdaten, offene_bestellungen_df)

            st.header("3. Ergebnis-Tabelle")
            st.dataframe(ausgabe_df)

            # Download Button für Excel
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
        else:
            st.error("Fehler beim Verarbeiten der Dateien. Bitte überprüfe die Dateien und lade sie erneut hoch.")
    elif excel_file or csv_file or artikelnummern_etiketten:
        if not artikelnummern_etiketten and uploaded_image_files:
            st.warning("Bitte validiere oder gib die Artikelnummern aus den Etikettenbildern ein, bevor du die Dateien hochlädst.")
        elif not excel_file:
            st.warning("Bitte lade die Artikelstammdaten Excel-Datei hoch.")
        elif not csv_file:
            st.warning("Bitte lade die Offene Bestellungen CSV-Datei hoch.")


if __name__ == "__main__":
    main()
