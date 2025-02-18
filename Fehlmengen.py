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
    Versucht, Dateiformat zu erkennen und liest die Datei ein (Excel oder HTML-Tabelle).
    **Excel-Dateien werden jetzt ohne Header eingelesen und Spaltennamen manuell zugewiesen (siehe artikel_stammdaten_lesen).**
    Ignoriert die ersten zwei Zeilen (falls Excel). **(Anzahl der übersprungenen Zeilen konfigurierbar)**
    Nutzt 'xlrd' Engine für .xls Dateien und 'openpyxl' für .xlsx Dateien (falls Excel).
    Erkennt und parst HTML-Tabellen in Dateien.
    Versucht zuerst UTF-16-LE Dekodierung mit Fehler-Ignorierung. Dann erweiterte Liste von Encodings (wie zuvor).
    Verbesserte HTML-Erkennung (prüft auf <html>, <!DOCTYPE html> und <TABLE>).
    Genauere Fehlermeldungen.
    **Spaltennamenanpassung erfolgt jetzt manuell im Code (für Excel). Für HTML wird Header automatisch erkannt.**

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

    if dateityp == 'bestaende_excel' or dateityp == 'offene_bestellungen_excel': # Dateityp-Optionen für Excel/HTML
        datei_inhalt_string = None
        versuchte_encodings = ['utf-16-le', 'utf-8', 'utf-8-sig', 'latin-1', 'cp1252', 'utf-16', 'utf-16-be'] # Erweiterte Liste der Encodings, UTF-16-LE zuerst

        # Versuche zuerst UTF-16-LE mit Fehler-Ignorierung
        try:
            datei_inhalt_string = uploaded_file.getvalue().decode('utf-16-le', errors='ignore') # UTF-16-LE mit Fehler-Ignorierung!
            st.info(f"Dateiinhalt erfolgreich mit Encoding 'utf-16-le' (mit Fehler-Ignorierung) gelesen.") # Info-Meldung für erfolgreiches Encoding (mit Fehler-Ignorierung)
        except Exception as e:
            st.warning(f"Encoding 'utf-16-le' (mit Fehler-Ignorierung) fehlgeschlagen. Versuche weitere Encodings...") # Warnung, wenn Encoding fehlgeschlagen (mit Fehler-Ignorierung)


        if datei_inhalt_string is None: # Falls UTF-16-LE mit Fehler-Ignorierung fehlschlägt, versuche die anderen Encodings
            for encoding in versuchte_encodings[1:]: # Starte Schleife ab dem 2. Encoding (utf-8), da utf-16-le schon versucht wurde
                try:
                    datei_inhalt_string = uploaded_file.getvalue().decode(encoding) # Dateiinhalt als String lesen, Encoding versuchen
                    st.info(f"Dateiinhalt erfolgreich mit Encoding '{encoding}' gelesen.") # Info-Meldung für erfolgreiches Encoding
                    break # Bei Erfolg aus der Schleife ausbrechen
                except UnicodeDecodeError:
                    st.warning(f"Encoding '{encoding}' fehlgeschlagen. Versuche nächstes Encoding...") # Warnung, wenn Encoding fehlschlägt
                    continue # Zum nächsten Encoding übergehen

        if datei_inhalt_string is None: # Wenn alle Encodings fehlschlagen
            fehlermeldung = f"**FATALER FEHLER: Encoding-Problem!** Fehler beim Lesen der Datei: UnicodeDecodeError.  Keines der folgenden Encodings hat funktioniert: {versuchte_encodings}. \n\nMögliche Ursachen: Datei ist **beschädigt**, **keine reine Textdatei** oder verwendet ein **völlig unbekanntes Encoding**." # Präzisere Fehlermeldung mit Liste der Encodings
            st.error(fehlermeldung)
            return None # Fehlerfall, kein DataFrame


        try: # Jetzt HTML- oder Excel-Parsing versuchen (nach erfolgreichem Encoding)
            ist_html_datei = False
            html_start_tags = ["<TABLE", "<HTML", "<!DOCTYPE html>"] # Erweiterte HTML-Erkennung:  prüfe auf <html>, <!DOCTYPE html> und <table>
            for tag in html_start_tags:
                if tag in datei_inhalt_string.upper():
                    ist_html_datei = True
                    break # Sobald ein HTML-Tag gefunden, ist es eine HTML-Datei

            if ist_html_datei: # HTML-Datei erkannt
                st.info("Datei als HTML-Tabelle erkannt. Versuche HTML-Parsing...")
                dfs = pd.read_html(datei_inhalt_string, header=0) # HTML-Tabelle(n) lesen, Header automatisch erkennen
                if dfs:
                    df = dfs[0] # Erste Tabelle aus HTML extrahieren (Annahme: es gibt nur eine Tabelle)
                    st.info(f"HTML-Tabelle erfolgreich gelesen.")
                else:
                    fehlermeldung = "**FEHLER beim HTML-Parsing:** HTML-Datei enthält keine Tabellen oder Tabellen konnten nicht geparst werden." # Genauere Fehlermeldung für HTML-Parsing-Fehler
            else: # Falls keine HTML-Datei, versuche Excel-Parsing
                engine = 'xlrd' if uploaded_file.name.lower().endswith('.xls') else 'openpyxl' # Wähle Engine basierend auf Dateiendung
                df = pd.read_excel(uploaded_file, engine=engine, header=None, skiprows=2)  # Kein Header, überspringe 2 Zeilen (oder 3, je nach Bedarf, hier 2 beibehalten)
                # df.columns = ['Artikel', 'Kurzbezeichnung', 'Bestand', 'ME', 'Geliefert', 'Offen', 'OffenBE']  # Spaltennamen werden jetzt in artikel_stammdaten_lesen zugewiesen!
                st.info(f"Datei als Excel-Datei erkannt. Erfolgreich gelesen (Engine: '{engine}', **keine Spaltenüberschriften aus Datei verwendet, manuelle Spaltenzuweisung im Code aktiv**).") # Info für Benutzer aktualisiert, Hinweis auf manuelle Spaltenzuweisung
        except Exception as e: # Allgemeine Fehler beim Parsing (HTML oder Excel)
            fehlermeldung_parsing = f"**FEHLER beim Parsen der Datei (nach erfolgreichem Encoding):** {e}. \n\nMögliche Ursachen: Datei ist beschädigt, falsches Format, HTML-Parsing-Fehler oder **unerwartetes Excel-Format**. Engine: 'xlrd'/'openpyxl' wurde verwendet (falls Excel-Parsing versucht)." # Fehlermeldung erweitert, Hinweis auf unerwartetes Excel-Format
            if fehlermeldung: # Falls es schon einen Encoding-Fehler gab, diesen beibehalten, sonst Parsing-Fehler nehmen
                fehlermeldung = fehlermeldung # Encoding-Fehler behalten
            else:
                fehlermeldung = fehlermeldung_parsing # Parsing-Fehler nehmen

            if not fehlermeldung: #  Sicherstellen, dass Fehlermeldung nicht None ist, bevor Error angezeigt wird (sollte jetzt nie None sein)
                fehlermeldung = "Unbekannter Fehler beim Lesen/Parsen der Datei. Bitte überprüfe die Datei." # Fallback-Fehlermeldung, falls alles andere fehlschlägt
            st.error(fehlermeldung)
            return None # Fehlerfall, kein DataFrame


    if fehlermeldung: #  Fehlermeldung anzeigen, falls gesetzt (Encoding- oder Parsing-Fehler)
        st.error(fehlermeldung)
        return None

    return df # DataFrame ohne Spaltennamen zurückgeben (Spaltennamenanpassung erfolgt in artikel_stammdaten_lesen/offene_bestellungen_lesen)


@st.cache_data
def artikel_stammdaten_lesen(uploaded_file):
    """Liest Artikelstammdaten aus Excel/HTML mit datei_inspektion_und_anpassung.
    **Verwendet manuelle Spaltennamen für Excel-Dateien und prüft auf erforderliche Spalten.**
    """
    df_bestand = datei_inspektion_und_anpassung(uploaded_file, 'bestaende_excel')
    if df_bestand is None:
        return None

    st.write("DataFrame Struktur (vor Spaltennamen-Zuweisung):") # **Debug-Ausgabe:** DataFrame Struktur anzeigen
    st.dataframe(df_bestand) # **Debug-Ausgabe:** DataFrame anzeigen
    if isinstance(df_bestand, pd.DataFrame):
        print(f"DataFrame shape vor Spaltennamen-Zuweisung: {df_bestand.shape}") # Shape im Backend Log ausgeben

        # Manuelle Spaltenzuweisung für Excel (unabhängig vom Header in der Datei)
        # **WICHTIG:**  Überprüfen Sie nach dem ersten Ausführen mit dieser Funktion
        # die tatsächliche Struktur des DataFrames `df_bestand` (z.B. `st.dataframe(df_bestand)` in Streamlit ausgeben).
        # Passen Sie die Spaltennamen in der nächsten Zeile **genau** an die
        # **tatsächlichen Spaltenüberschriften** der HTML-Tabelle an.
        # Die hier angegebenen Spaltennamen sind nur ein Beispiel und MÜSSEN möglicherweise angepasst werden!

        df_bestand.columns = ['Artikel', 'Kurzbezeichnung', 'Bestand', 'ME'] # **Reduzierte Spaltennamen (testweise)!**


        # Überprüfe, ob die erforderlichen Spalten vorhanden sind (NACH manueller Zuweisung!)
        required_columns = ['Artikel', 'Kurzbezeichnung', 'Bestand', 'ME']
        missing_columns = [col for col in required_columns if col not in df_bestand.columns]

        if missing_columns:
            st.error(f"**FEHLER: Fehlende Spalten nach manueller Spaltenzuweisung in Bestände-Datei:** {', '.join(missing_columns)}. \n\n**Mögliche Ursachen:** Unerwartetes Dateiformat, falsche Spaltenreihenfolge oder Anzahl an Spalten in der Datei. \n\n**Bitte überprüfe die Spaltenzuweisung im Code in der Funktion `artikel_stammdaten_lesen` und passe sie ggf. an die Datei an.**") # Erweiterte Fehlermeldung mit Hinweis auf Code-Anpassung
            return None
    else:
        st.error("Fehler beim Einlesen der Bestände-Datei. DataFrame ist nicht valide.")
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
    """Liest offene Bestellungen aus Excel/HTML mit datei_inspektion_und_anpassung.
     **Verwendet manuelle Spaltennamen für Excel-Dateien (ggf. anpassen!).**
     """
    df_offene_bestellungen = datei_inspektion_und_anpassung(uploaded_file, 'offene_bestellungen_excel')
    if df_offene_bestellungen is None:
        return None

    # Manuelle Spaltenzuweisung für Excel (unabhängig vom Header in Datei)
    df_offene_bestellungen.columns = ['Belegnr.', 'Datum', 'Kurzbezeichnung', 'Bearbeiter', 'Artikelnr.', 'Lieferdatum', 'ME', 'Menge', 'Geliefert', 'Offen', 'OffenBE'] # **Manuelle Spaltennamen zuweisen!** Spaltennamen ggf. anpassen!

    return df_offene_bestellungen # DataFrame mit manuell zugewiesenen Spaltennamen zurückgeben


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
    bestaende_excel_file = st.file_uploader("Bestände Datei hochladen (Excel *.xls, *.xlsx, oder HTML-Tabelle *.xls, **Excel: Erste 2 Zeilen ignoriert, Spaltennamen manuell im Code anpassen! HTML: Header wird automatisch erkannt.**)", type=["xls", "xlsx", "html", "htm"]) # Dateitypen und Beschreibung für HTML erweitert, Hinweis für Excel-Spaltennamen
    offene_bestellungen_excel_file = st.file_uploader("Offene Bestellungen Datei hochladen (Excel *.xls, *.xlsx, oder HTML-Tabelle *.xls, **Excel: Erste 2 Zeilen ignoriert, Spaltennamen manuell im Code anpassen! HTML: Header wird automatisch erkannt.**)", type=["xls", "xlsx", "html", "htm"]) # Dateitypen und Beschreibung für HTML erweitert, Hinweis für Excel-Spaltennamen

    artikel_stammdaten = None # Initialisieren außerhalb der if-Bedingung
    offene_bestellungen_df = None # Initialisieren außerhalb der if-Bedingung

    if bestaende_excel_file:
        artikel_stammdaten = artikel_stammdaten_lesen(bestaende_excel_file) # Nutze bestaende_excel_file
        if artikel_stammdaten is not None: # **Prüfen, ob artikel_stammdaten NICHT None ist**
            st.dataframe(artikel_stammdaten) # **DataFrame zur Kontrolle direkt in Streamlit anzeigen**


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
            st.warning("Bitte lade die Bestände Datei hoch.") # Warnung für Bestände Datei
        if not offene_bestellungen_excel_file:
            st.warning("Bitte lade die Offene Bestellungen Datei hoch.") # Warnung für Offene Bestellungen Datei


if __name__ == "__main__":
    main()
