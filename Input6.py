import math
import csv
import os
import json
import chardet
import traceback
from selenium import webdriver
from selenium.webdriver.common.keys import Keys  # Richtiger Import für Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
import time
from datetime import datetime
import sys
import threading
import pandas as pd  # Für Excel-Verarbeitung

# --------------------------------------------------------------------------------
# GLOBALE FUNKTIONEN & HILFSFUNKTIONEN
# --------------------------------------------------------------------------------

def log_error(entry, error_message):
    """
    Schreibt einen Eintrag ins Fehlerprotokoll (errorlog.csv).
    """
    with open(ERROR_LOG_FILE, mode="a", newline="", encoding="utf-8") as error_file:
        writer = csv.writer(error_file)
        if error_file.tell() == 0:
            writer.writerow(["Timestamp", "Entry", "Error Message"])
        writer.writerow([
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            str(entry),
            str(error_message)
        ])

def round_to_nearest_15(minutes):
    """Rundet Minuten auf das nächste 15-Minuten-Intervall auf."""
    return math.ceil(minutes / 15) * 15 if minutes % 15 != 0 else minutes

def convert_to_decimal_time(rounded_minutes):
    """Wandelt gerundete Minuten in Stunden im Dezimalformat um (z.B. 90 -> 1.5)."""
    return round(rounded_minutes / 60, 2)

def convert_date_format(iso_date):
    """
    Konvertiert ein Datum aus verschiedenen Formaten in 'DD/MM/YYYY'.
    Akzeptiert sowohl reine Datumsstrings im Format 'YYYY-MM-DD'
    als auch Datumsstrings mit Zeitanteil, z.B. 'YYYY-MM-DD 00:00:00'.
    Falls das Datum nicht konvertiert werden kann, wird ein Fallback-Datum verwendet.
    """
    # Wenn es bereits ein datetime-Objekt ist, formatiere es direkt:
    if isinstance(iso_date, datetime):
        return iso_date.strftime('%d/%m/%Y')
    # Entferne Zeitanteil falls vorhanden:
    if ' ' in iso_date:
        iso_date = iso_date.split(' ')[0]
    try:
        return datetime.strptime(iso_date, '%Y-%m-%d').strftime('%d/%m/%Y')
    except ValueError:
        print(f"Fehler beim Konvertieren des Datums: {iso_date} - Fallback-Datum wird verwendet.")
        return '01/01/2025'

def extract_date_from_filename(filename):
    """
    Extrahiert das Datum aus dem Dateinamen, wenn er z.B. in der Form
    "All Activities 2025-02-20.xlsx" vorliegt.
    Liefert ein datetime-Objekt zum Sortieren zurück.
    """
    try:
        base_name = os.path.splitext(filename)[0]  # z.B. "All Activities 2025-02-20"
        parts = base_name.split("All Activities")
        date_part = parts[1].strip()
        return datetime.strptime(date_part, "%Y-%m-%d")
    except Exception:
        return datetime(1900, 1, 1)

def press_tab_with_delay(actions, times, delay=1.0):
    """
    Drückt n-mal TAB mit jeweils kleinem Delay.
    Nutzt dabei ein ActionChains-Objekt.
    """
    for _ in range(times):
        actions.send_keys(Keys.TAB).perform()
        time.sleep(delay)

# --------------------------------------------------------------------------------
# GLOBALE KONSTANTEN
# --------------------------------------------------------------------------------
ERROR_LOG_FILE = "errorlog.csv"
DEFAULT_DL_NUMBER = "DL20517"
DEFAULT_SECOND_FIELD = "MAPPING nicht im JSON File"

# --------------------------------------------------------------------------------
# STOP-BEFEHL: Hintergrund-Listener (Befehl: "stop")
# --------------------------------------------------------------------------------
def stop_listener():
    global stop_flag
    while True:
        cmd = input()  # Lauscht fortlaufend auf Eingaben
        if cmd.strip().lower() == "stop":
            stop_flag = True
            print("Stop-Befehl empfangen. Das Programm wird nach dem aktuellen Schritt beendet.")
            break

# --------------------------------------------------------------------------------
# MAIN-PROCESS-FUNKTION: Initialisiert den WebDriver, verarbeitet CSV/Excel-Dateien
# --------------------------------------------------------------------------------
def process_all_files():
    global previous_dl_value, previous_second_field_value, stop_flag, entries_since_cleanup

    # Starte den Stop-Listener-Thread
    listener_thread = threading.Thread(target=stop_listener, daemon=True)
    listener_thread.start()

    # Setze Mapping-Werte und den Eintragszähler zurück
    previous_dl_value = None
    previous_second_field_value = None
    entries_since_cleanup = 0

    print("\n[INFO] Initialisiere WebDriver...")
    try:
        driver = webdriver.Chrome()
    except Exception as e:
        print("Fehler beim Starten des Chrome-WebDrivers. Stelle sicher, dass 'chromedriver' im PATH liegt.")
        traceback.print_exc()
        sys.exit(1)

    print("[INFO] Lade die Zeiterfassungs-Webseite...")
    driver.get('https://zeiterfassung-bag.msappproxy.net/bag_nav_ch_prod/WebClient/tablet.aspx?profile=SITETNTTABLET&company=A.%20Baggenstos%20%26%20Co.%20AG')
    print("Bitte logge dich im Browserfenster ein und öffne die gewünschte Seite.")
    input("Sobald du eingeloggt bist, drücke Enter, um fortzufahren...")
    time.sleep(5)
    print(f"[INFO] Aktuelle URL: {driver.current_url}")
    print("[INFO] Automatisierung wird nun fortgesetzt...")
    wait = WebDriverWait(driver, 60)
    actions = ActionChains(driver)
    select_all_key = Keys.CONTROL  # Verwende Keys.CONTROL für Windows/Linux oder Keys.COMMAND für Mac

    print("[INFO] Lade Mappings aus 'mappings.json'...")
    try:
        with open('mappings.json', 'r', encoding='utf-8') as f:
            mappings = json.load(f)
            dl_number_mapping = mappings['dl_number_mapping']
            second_field_mapping = mappings['second_field_mapping']
    except Exception as e:
        print("Fehler beim Laden der Mappings. Bitte überprüfe die Datei 'mappings.json'.")
        traceback.print_exc()
        driver.quit()
        sys.exit(1)

    print("[INFO] Suche nach Eingabedateien (CSV und Excel) im aktuellen Verzeichnis...")
    input_files = []
    for file in os.listdir(os.getcwd()):
        if file.lower().startswith("all activities") and (file.lower().endswith(".csv") or file.lower().endswith(".xlsx")):
            input_files.append(file)

    input_files_sorted = sorted(input_files, key=extract_date_from_filename)
    if not input_files_sorted:
        print("Keine passenden Dateien gefunden. Stelle sicher, dass deine Datei 'All Activities YYYY-MM-DD.csv' oder '.xlsx' heißt.")
        input("Drücke Enter, um den Browser zu schließen...")
        driver.quit()
        sys.exit(1)
    else:
        print(f"[INFO] Gefundene Dateien: {input_files_sorted}")

    # Verarbeite jede Datei einzeln
    try:
        for input_file in input_files_sorted:
            if stop_flag:
                print("Stop-Befehl erkannt. Das Programm wird beendet.")
                break

            print("\n" + "="*60)
            print(f"[INFO] Verarbeite Datei: {input_file}")
            print("="*60)
            file_path = os.path.join(os.getcwd(), input_file)

            data_entries = []
            if input_file.lower().endswith(".csv"):
                print(f"[INFO] Lese CSV-Datei: {input_file}")
                try:
                    with open(file_path, 'rb') as f:
                        file_encoding = chardet.detect(f.read())['encoding']
                except Exception as e:
                    error_msg = f"Fehler beim Lesen der Datei {input_file}: {str(e)}"
                    print(error_msg)
                    log_error(input_file, error_msg)
                    continue

                try:
                    with open(file_path, newline='', encoding=file_encoding, errors='replace') as csvfile:
                        first_line = csvfile.readline()
                        delimiter = '\t' if '\t' in first_line else ','
                        csvfile.seek(0)
                        csv_reader = csv.DictReader(csvfile, delimiter=delimiter)
                        headers = csv_reader.fieldnames
                        print(f"[INFO] Erkannte Spalten in CSV: {headers}")
                        for row in csv_reader:
                            if 'Day' not in row:
                                error_msg = f"Überspringe Zeile wegen fehlender 'Day'-Spalte: {row}"
                                print(error_msg)
                                log_error(row, error_msg)
                                continue
                            data_entries.append({
                                'Duration': row.get('Duration', '00:00:00'),
                                'Project': row.get('Project', 'Unknown'),
                                'Title': row.get('Title', ''),
                                'Day': row['Day']
                            })
                except Exception as e:
                    error_msg = f"Fehler beim Verarbeiten der CSV-Datei {input_file}: {str(e)}"
                    print(error_msg)
                    log_error(input_file, error_msg)
                    continue

            elif input_file.lower().endswith(".xlsx"):
                print(f"[INFO] Lese Excel-Datei: {input_file}")
                try:
                    df = pd.read_excel(file_path, engine='openpyxl')
                    print(f"[INFO] Excel-Datei erfolgreich geladen. Spalten: {list(df.columns)}")
                    for idx, row in df.iterrows():
                        if pd.isnull(row.get('Day')):
                            error_msg = f"Überspringe Zeile {idx} wegen fehlender 'Day'-Spalte."
                            print(error_msg)
                            log_error(f"{input_file} Zeile {idx}", error_msg)
                            continue
                        data_entries.append({
                            'Duration': row.get('Duration', '00:00:00'),
                            'Project': row.get('Project', 'Unknown'),
                            'Title': row.get('Title', ''),
                            'Day': row.get('Day')
                        })
                except Exception as e:
                    error_msg = f"Fehler beim Verarbeiten der Excel-Datei {input_file}: {str(e)}"
                    print(error_msg)
                    log_error(input_file, error_msg)
                    continue

            total_rounded_minutes = 0
            # Verarbeite jeden Eintrag in der aktuellen Datei
            for idx, entry in enumerate(data_entries):
                if stop_flag:
                    print("Stop-Befehl erkannt. Das Programm wird beendet.")
                    break

                try:
                    converted_day = convert_date_format(str(entry['Day']))
                except Exception as e:
                    converted_day = '01/01/2025'

                project_name = entry['Project'] or "Unknown"
                mapped_project = next(
                    (key for key in dl_number_mapping if key.lower() in project_name.lower()),
                    None
                )

                if mapped_project is None:
                    log_error(entry, f"Kein gültiges Mapping gefunden für Projekt '{project_name}'. Nutze Default (vorheriger Wert).")
                    dl_number = previous_dl_value if previous_dl_value is not None else DEFAULT_DL_NUMBER
                    second_field_value = previous_second_field_value if previous_second_field_value is not None else DEFAULT_SECOND_FIELD
                    title_value = "KEIN DL BELEG für den angegebenen Titel gefunden"
                else:
                    dl_number = dl_number_mapping[mapped_project]
                    second_field_value = second_field_mapping[mapped_project]
                    title_value = entry['Title']
                    previous_dl_value = dl_number
                    previous_second_field_value = second_field_value

                try:
                    h, m, s = map(int, str(entry['Duration']).split(':'))
                    total_minutes = h * 60 + m + s // 60
                    rounded_minutes = round_to_nearest_15(total_minutes)
                    total_rounded_minutes += rounded_minutes
                    decimal_duration = convert_to_decimal_time(rounded_minutes)
                except Exception as e:
                    error_msg = f"Ungültiges Dauerformat '{entry['Duration']}'. Eintrag wird übersprungen."
                    print(error_msg)
                    log_error(entry, error_msg)
                    continue

                print(f"\n[INFO] Verarbeite Eintrag {idx+1}:")
                print(f"  Projekt: {project_name}")
                print(f"  DL: {dl_number}")
                print(f"  Zweites Feld: {second_field_value}")
                print(f"  Titel: {title_value}")
                print(f"  Dauer (Dezimal): {decimal_duration} Stunden")
                print(f"  Datum: {converted_day}")

                try:
                    actions.send_keys(dl_number).perform()
                    time.sleep(1)
                    press_tab_with_delay(actions, 5, delay=1.0)
                    actions.send_keys(second_field_value).perform()
                    time.sleep(1)
                    press_tab_with_delay(actions, 2, delay=1.0)
                    actions.key_down(select_all_key).send_keys('a').key_up(select_all_key).perform()
                    actions.send_keys(Keys.BACKSPACE).perform()
                    actions.send_keys(title_value).perform()
                    time.sleep(1)
                    press_tab_with_delay(actions, 2, delay=1.0)
                    actions.send_keys(converted_day).perform()
                    time.sleep(1)
                    press_tab_with_delay(actions, 5, delay=1.0)
                    actions.send_keys(str(decimal_duration)).perform()
                    time.sleep(1)
                    press_tab_with_delay(actions, 3, delay=1.0)
                    time.sleep(2)
                    print(f"[INFO] Eintrag {idx+1} erfolgreich eingetragen.")
                except Exception as e:
                    error_msg = f"Fehler beim automatischen Eintragen (Eintrag {idx+1}): {str(e)}"
                    print(error_msg)
                    log_error(entry, error_msg)
                    continue

                entries_since_cleanup += 1
                if entries_since_cleanup >= 14:
                    print("\n[WICHTIG] Es wurden 14 Einträge vorgenommen.")
                    print("Bitte bestätige die bisherigen Einträge in der Zeiterfassung.")
                    print("Du musst 'Alle Freigeben und übertragen' durchführen, damit das Programm fortgesetzt werden kann.")
                    print("Sobald dies erledigt ist, tippe bitte 'GO'.")
                    while True:
                        confirm = input().strip().lower()
                        if confirm == "go":
                            entries_since_cleanup = 0  # Zähler zurücksetzen
                            print("[INFO] Bestätigung erhalten. Verarbeitung wird fortgesetzt...")
                            break
                        else:
                            print("Ungültige Eingabe. Bitte tippe 'GO', sobald du die Einträge bestätigt hast.")

            total_decimal_time = convert_to_decimal_time(total_rounded_minutes)
            expected_minutes = 8.5 * 60  # 8.5 Stunden = 510 Minuten
            print("\n" + "-"*60)
            if total_rounded_minutes == expected_minutes:
                print(f"[INFO] Datei '{input_file}' abgeschlossen: Gesamtdauer = {total_decimal_time} Stunden (exakt 8.5 Stunden).")
            elif total_rounded_minutes < expected_minutes:
                diff = expected_minutes - total_rounded_minutes
                print(
                    f"[INFO] Datei '{input_file}' abgeschlossen: Gesamtdauer = {total_decimal_time} Stunden. "
                    f"Es fehlen {diff // 60} Stunden und {diff % 60} Minuten, um die 8.5 Stunden-Marke zu erreichen."
                )
            else:
                diff = total_rounded_minutes - expected_minutes
                print(
                    f"[INFO] Datei '{input_file}' abgeschlossen: Gesamtdauer = {total_decimal_time} Stunden. "
                    f"Du hast {diff // 60} Stunden und {diff % 60} Minuten mehr als 8.5 Stunden erfasst."
                )
            print("-"*60)
            if stop_flag:
                break

    except Exception as e:
        print("Ein unerwarteter Fehler im Skript ist aufgetreten:")
        traceback.print_exc()
    finally:
        input("Drücke Enter, um den Browser zu schließen...")
        driver.quit()

# --------------------------------------------------------------------------------
# HAUPTSCHLEIFE: Nach Abschluss der Verarbeitung "end" oder "restart" abfragen
# --------------------------------------------------------------------------------
while True:
    stop_flag = False
    process_all_files()

    if stop_flag:
        break

    final_cmd = input("Prozess abgeschlossen. Gib 'end' ein, um zu beenden, oder 'restart', um von vorne zu beginnen: ").strip().lower()
    if final_cmd == "restart":
        print("Neustart des Prozesses...")
        continue
    elif final_cmd == "end":
        print("Programm wird beendet.")
        break
    else:
        print("Ungültiger Befehl. Programm wird beendet.")
        break

print("Programmende.")
