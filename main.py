import os
import re
import csv
import shutil
import openpyxl
import sys
import subprocess
from datetime import datetime
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from hl7apy.parser import parse_message
from hl7apy.exceptions import HL7apyException

# Pruefen, ob Pillow installiert ist (wird für Screenshots/Bilder in openpyxl benötigt)
try:
    import PIL
    HAS_PILLOW = True
except ImportError:
    HAS_PILLOW = False

# Quellordner und erlaubte Endungen für das automatische Kopieren der Testdateien
#TESTFILE_DIR = "testfile"
TESTFILE_DIR = r"C:\Users\A0047444\Downloads"
TESTFILE_EXTENSIONS = ('.hl7', '.txt', '.dat', '.opa', '.bin')

# Dateiendungen für das Einlesen der Daten bzw. Mapping-Dateien
DATA_EXTENSIONS = ('.hl7', '.dat', '.txt', '.csv', '.opa')
MAPPING_EXTENSIONS = ('.map', '.mapping')

# Maximale Anzahl aufbewahrter Sicherheitskopien von manual.xlsx
MAX_BACKUPS = 2

# Standard-Reihenfolge der HL7-Segmente für die Sortierung
STANDARD_SEGMENT_ORDER = [
    'MSH', 'EVN', 'PID', 'PD1', 'PV1', 'PV2',
    'IN1', 'IN2', 'IN3', 'NK1', 'GT1', 'AL1', 'DG1', 'PR1',
    'OBX', 'NTE', 'ORC', 'OBR', 'SPM'
]

# Bezeichnungen für Komponenten (Schlüssel ohne Segment-Index, z. B. PV1_3.1)
COMPONENT_NAMES = {
    'PV1_3.1': 'Point of Care',
    'PV1_3.2': 'Room',
    'PV1_3.3': 'Bed',
    'PV1_3.4': 'Facility',
    'PV1_3.5': 'Location Status',
    'PV1_3.6': 'Person Location Type',
    'PV1_3.7': 'Building',
}

# Felder, die bei HL7-Dateien immer angezeigt werden (auch wenn leer) und hervorgehoben sind
ALWAYS_SHOW_FIELDS = list(COMPONENT_NAMES.keys())   # PV1_3.1 ... PV1_3.7


# ==========================================
# OPA-FORMAT (EXPOPA10, Pipe-getrennt, ein Recordtyp pro Zeile)
# Aufbau je Zeile: PID | FID | Dossiernummer | Recordtyp | weitere Felder ...
# Feldnamen gemäss Schnittstellenbeschreibung "export10_d" (Stand 2025-04-01)
# ==========================================
OPA_DELIMITER = '|'
OPA_RECORD_TYPE_POS = 4          # Position (1-basiert) des Recordtyps in jeder Zeile

# Reihenfolge der Recordtypen für die Sortierung (entspricht der Spezifikation)
OPA_RECORD_ORDER = ['1', 'R', '2', '3', 'I', 'L', 'M', 'G', '!', 'T', 'P', 'S', '4']

OPA_RECORD_TITLES = {
    '1': 'Pat-Fall', 'R': 'Reservation', '2': 'Verlegungen', '3': 'Urlaube',
    'I': 'Eingriffe', 'L': 'Adressen', 'M': 'Ärzte', 'G': 'Garanten',
    '!': 'Alarme', 'T': 'Beziehungen', 'P': 'Spez. Patient', 'S': 'Spez. Fall',
    '4': 'Ende',
}

# Feldnamen je Recordtyp; Index 0 = Position 1
OPA_FIELD_NAMES = {
    '1': [
        'Patientennummer (PID)',
        'Fallnummer (FID)',
        'Dossiernummer',
        'Aufeichnungstyp',
        'Name',
        'Vorname',
        'Geburtsdatum',
        'Geschlecht Code',
        'Geschlecht Bezeichnung',
        'Zivilstand Code',
        'Zivilstand Bezeichnung',
        'Mädchenname',
        'Nationalität',
        'Nationalität Bezeichnung (Land)',
        'Heimatort',
        'Sprache Code',
        'Sprache Bezeichnung',
        'Anrede Code',
        'Anrede Bezeichnung',
        'Alte (!) AHV-Nummer',
        'Konfession Code',
        'Konfession Bezeichnung',
        'freies Feld (Zone)',
        'Statistikzone 1 Code',
        'Statistikzone 1 Bezeichnung',
        'Name des Vaters/Ehemannes',
        'Name der Mutter',
        'Passnummer',
        'Pass ausgestellt durch',
        'Anzahl Kinder',
        'Status der Aufnahme',
        'Code Mandant / Firma',
        'Eintrittstyp Code',
        'Eintrittstyp Bezeichnung',
        'Eintrittsdatum',
        'Eintrittsmodus Code',
        'Eintrittsmodus Bezeichnung',
        'Herkuntscode',
        'Herkunftsbezeichnung',
        'Eintrittsart Code',
        'Eintrittsart Bezeichnung',
        'Taxentyp Code',
        'Taxentyp Bezeichnung',
        'Fall Code',
        'Fall Bezeichnung',
        'Abteilung Code',
        'Abteilung Bezeichnung',
        'Einheit Code',
        'Einheit Bezeichnung',
        'Zimmer',
        'Bett',
        'Telephon',
        'Klasse Code',
        'Klasse Bezeichnung',
        'Tarif Code',
        'Tarif Bezeichnung',
        'Garantenart Code',
        'Garantenart Bezeichnung',
        'Eintrittszeit',
        'einweisende Instanz Code',
        'einweisende Instanz Bezeichnung',
        'Statistikzone 2 Code',
        'Statistikzone 2 Bezeichnung',
        'Statistikzone 3 Code',
        'Statistikzone 3 Bezeichnung',
        'Statistikzone 4 Code',
        'Statistikzone 4 Bezeichnung',
        'Statistikzone 5 Code',
        'Statistikzone 5 Bezeichnung',
        'Statistikzone 6 Code',
        'Statistikzone 6 Bezeichnung',
        'Diagnose Code',
        'Diagnise Bezeichnung',
        'Flag: Fall Fakturiert',
        'RESERVE',
        'Austrittsdatum',
        'Austrittsart /-modus Code',
        'Austrtittsart /-modus Bezeichnung',
        'Bestimmung Code',
        'Bestimmung Bezeichnung',
        'Austrittszeit',
        'Behandlung nach Austritt Code',
        'Behandlung nach Austritt Bezeichnung',
        'Hauptdiagnose',
        'Hauptbehandlung',
        'Todesdatum',
        'Todeszeit',
        'Flag: Schlechter Zahler',
        'RESERVE',
        'Mutationsart',
        'ehemalige PID',
        'ehemalige FID',
        'RESERVE',
        'Patienten-Nr. (PID) der Mutter',
        'Fall-Nr (FID) der Mutter',
        'Dossier-Nr (DID) der Mutter',
        'Gewicht',
        'Alias vom Zimmer',
        'Neue AHV-Nummer (N13)',
        'Name des Parter/Ehegatten',
        'Notfall Telefonnummer 1',
        'Notfall Telefonnummer 2',
        'Militärischer Grad',
        'Zimmernummer der Mutter',
        'Geburtsort (Ortschaft)',
        'Dossier-Nr vom frührenden Fall',
        'Datum vom Splitting DRG',
        'Grund für die Fallzusammenführung',
        'Patienten-Nr. (PID) des Umsystem bei Extrener Falleröffnung',
        'Fall- / Dossier-Nr. (FID/DID) des Umsystem bei Extrener Falleröffnung',
        'Inhalt Feld "Bei" Telefon Notfall 1 (aus Zusatz-Identität Patient)',
        'Inhalt Feld "Bei" Telefon Notfall 2 (aus Zusatz-Identität Patient)',
        'Fakturiere DRG / PCG',
        'Statistikzone 7 Code',
        'Statistikzone 7 Bezeichnung',
        'RESERVE',
        'RESERVE',
        'EPD Vorhanden beim Patient',
        'Flag: Patient wünscht kein EPD',
        'Datum der letzten Anfrage',
        'Zulassen Dokumentenversand EPD (B2C)',
        'Zulassen Dokumentenversand Stammgemeinschaft (B2B)',
        'Theoretischer Eintritt (ehem. Theoretischer Austritt)',
        'Voraussichtliche Dauer vom Aufenthalt in Tagen',
        'Vorgesehnes Austrittsdatum',
        'Vorgesehene Austrittszeit',
        'Flag "Geplanter Austritt bestätigt"',
    ],
    'R': [
        'Patientennummer (PID)',
        'Fallnummer (FID)',
        'Dossiernummer',
        'Aufeichnungstyp',
        'Eintrittstyp Code',
        'Eintrittstyp Bezeichnung',
        'Eintrittsart Code',
        'Eintrittsart Bezeichnung',
        'Fall Code',
        'Fall Bezeichnung',
        'Abteilung Code',
        'Abteilung Bezeichnung',
        'Pflegeeinheit Code',
        'Pflegeeinheit Bezeichnung',
        'Zimmer (vorgesehenes)',
        'Bett (vorgesehenes)',
        'Klasse Code',
        'Klasse Bezeichnung',
        'RESERVE',
        'Eintrittsdatum (vorgesehenes)',
        'Eintrittszeit (vorgesehen)',
        'Datum Vorkonsult. Anästhesie',
        'Zeit Vorkonsult. Anästhesie',
        'Datum des Eingriffs',
        'Zeit des Eingriffs',
        'Operationssaal',
        'Vorgesehene Aufenthaltsdauer',
        'Betreff Memo (Kommentar) zur Reservation',
        'Eingriffsdauer',
        'RESERVE',
        'Alias vom Zimmer aus Opale Zimmerstamm',
        'Flag: Reservation aktiv/inaktiv',
        'RESERVE',
        'RESERVE',
        'RESERVE',
        'RESERVE',
        'RESERVE',
        'RESERVE',
        'RESERVE',
        'RESERVE',
        'RESERVE',
    ],
    '2': [
        'Patientennummer (PID)',
        'Fallnummer (FID)',
        'Dossiernummer',
        'Aufeichnungstyp',
        'Sequenznummer',
        'Verlegungsdatum',
        'Verlegungszeit',
        'Fall Code',
        'Fall Bezeichnung',
        'Abteilung Code',
        'Abteilung Bezeichnung',
        'Zimmer',
        'Einheit Code',
        'Einheit Bezeichnung',
        'Bett',
        'Telephon',
        'Klasse Code',
        'Klasse Bezeichnung',
        'Tarif Code',
        'Tarif Bezeichnung',
        'Garantenart Code',
        'Garantenart Bezeichnung',
        'Taxentyp Code',
        'Taxentyp Bezeichnung',
        'Arztnummer (behandelnder)',
        'Flag Aufschiebung',
        'Flag "Separate Fakturierung"',
        'Statistikzone 6 Code',
        'Statistikzone 6 Bezeichnung',
        'Diagnose Code',
        'Diagnise Bezeichnung',
        'Verlegung annuliert',
        'RESERVE',
        'Alias vom Zimmer gemäss Opale Zimmerstamm',
        'Zeit Pflege 100%',
        'CSB Minuten',
        'Plegezeit ohne CSB',
        'RESERVE',
        'RESERVE',
        'RESERVE',
        'RESERVE',
        'RESERVE',
        'RESERVE',
        'RESERVE',
    ],
    '3': [
        'Patientennummer (PID)',
        'Fallnummer (FID)',
        'Dossiernummer',
        'Aufeichnungstyp',
        'Datum Urlaub Start',
        'Zeit Urlaub Start',
        'Datum Urlaub Ende',
        'Zeit Urlaub Ende',
        'Klasse Code',
        'Klasse Bezeichnung',
        'Austrittsmodus Code',
        'Austrittsmodus Bezeichnung',
        'Bestimmung Code',
        'Bestimmung Bezeichnung',
        'Abteilung (Bestimmung) Code',
        'Abteilung (Bestimmung) Bezeichnung',
        'Bemerkung',
        'Sequenznummer',
        'RESERVE',
        'RESERVE',
        'RESERVE',
        'RESERVE',
        'RESERVE',
        'RESERVE',
        'RESERVE',
        'RESERVE',
        'RESERVE',
        'RESERVE',
        'RESERVE',
        'RESERVE',
        'RESERVE',
    ],
    'I': [
        'Patientennummer (PID)',
        'Fallnummer (FID)',
        'Dossiernummer',
        'Aufeichnungstyp',
        'Arztnummer (behandelnder)',
        'Arztname',
        'Pflegedatum',
        'Operationsdatum',
        'Operationszeit',
        'Konsultationsdatum Prä-Anästhesie',
        'Konsultationszeit Prä-Anästhesie',
        'Operationsdauer',
        'Operationssaal',
        'Anästhesistnummer',
        'Anästhesistname',
        'Anästhesisttyp Code',
        'Anästhesisttyp Bezeichnung',
        'Operation 1',
        'Operation 2',
        'Operation 3',
        'Operation 4',
        'Operation 5',
        'Diagnose 1',
        'Diagnose 2',
        'Diagnose 3',
        'Diagnose 4',
        'Diagnose 5',
        'RESERVE',
        'RESERVE',
        'RESERVE',
        'RESERVE',
        'RESERVE',
        'RESERVE',
        'RESERVE',
        'RESERVE',
        'RESERVE',
        'RESERVE',
        'RESERVE',
    ],
    'L': [
        'Patientennummer (PID)',
        'Fallnummer (FID)',
        'Dossiernummer',
        'Aufeichnungstyp',
        'Adressentyp Code',
        'Adressentyp Bezeichnung',
        'Korrespondenzflag',
        'Fakturierungsflag',
        'Anrede Code',
        'Anrede Bezeichnung',
        'Name',
        'Zusatz',
        'Adresse',
        'Adresse 2',
        'PLZ',
        'Unterteilung PLZ (Unter-Nummer)',
        'Ort',
        'Kanton',
        'Land',
        'Telefon Privat',
        'Telefon Beruf',
        'e-Mail Adresse',
        'Beruf',
        'Arbeitgeber',
        'Arbeitsort',
        'RESERVE',
        'Opale Interne Sequenznummer',
        'Handy-Nummer',
        'Verwandschaftsgrad Code',
        'Verwandschaftsgrad Bezeichnung',
        'Länder Code gemäss BfS',
        'E-Mail Adresse Arbeitgeber',
        'Kommentar Telefon Festnetz',
        'Kommentar Telefon Geschäft',
        'Kommentar Telefon Natel',
        'Kommentar Telefon E-Mail Patient',
        'Kommentar Telefon E-Mail Arbeitgeber',
        'RESERVE',
        'RESERVE',
        'RESERVE',
        'RESERVE',
        'RESERVE',
    ],
    'M': [
        'Patientennummer (PID)',
        'Fallnummer (FID)',
        'Dossiernummer',
        'Aufeichnungstyp',
        'Arzt für Aufnahme Code',
        'Arzt für Aufnahme Bezeichnung',
        'Arztnummer',
        'Name und Vorname vom Arzt',
        'Name vom Arzt',
        'Vorname vom Arzt',
        'Anrede Code',
        'Anrede Bezeichnung',
        'Adresse',
        'RESERVE',
        'PLZ',
        'RESERVE',
        'Ort',
        'Kanton',
        'Land',
        'Telefon 1',
        'Telefon 2',
        'Telefon 3',
        'Telefon 4',
        'E-Mail Adresse',
        'Konkordats Nummer',
        'EAN / GLN Nummer',
        'RESERVE',
        'Arzttyp Code',
        'Arzttyp Bezeichnung',
        'RESERVE',
        'Opale Interne Sequenznummer',
        'Telefonnummer Zentrale',
        'Externe ID vom Arzt',
        'RESERVE',
        'RESERVE',
        'RESERVE',
        'RESERVE',
        'RESERVE',
        'RESERVE',
        'RESERVE',
        'RESERVE',
        'RESERVE',
    ],
    'G': [
        'Patientennummer (PID)',
        'Fallnummer (FID)',
        'Dossiernummer',
        'Aufeichnungstyp',
        'Priorität',
        'Unterprioritär (Sequenz der Priorität)',
        'Flag: "Garant inaktiv"',
        'Garantennummer',
        'Garantenname',
        'Zusatz',
        'Adresse',
        'Adresse 2',
        'PLZ',
        'Unterteilung PLZ (Unter-Nummer)',
        'Ort',
        'Kanton',
        'Land',
        'Policenummer',
        'Gültig-bis-Datum der Kostendeckung',
        'Klasse Code',
        'Klasse Bezeichnung',
        'Fakturierung Code',
        'Fakturierung Bezeichnung',
        'Unfallsnummer',
        'Unfallsdatum',
        'Versichertenkarten Nummer',
        'EAN Nummer vom Garant',
        'RESERVE',
        'Gültig ab Datum der Kostendeckung',
        'Telefonnummer',
        'Faxnummer',
        '1. E-Mail Adresse vom Garant',
        'Bemerkung zum Garant auf dem Aufenthalt',
        'RESERVE',
        'RESERVE',
        'RESERVE',
        'RESERVE',
        'RESERVE',
        'RESERVE',
    ],
    '!': [
        'Patientennummer (PID)',
        'Fallnummer (FID)',
        'Dossiernummer',
        'Aufeichnungstyp',
        'Alarmnummer',
        'Alarmbezeichnung',
        'Alarmbemerkung / -meldung',
        'RESERVE',
        'RESERVE',
        'RESERVE',
        'RESERVE',
        'RESERVE',
        'RESERVE',
        'RESERVE',
        'RESERVE',
        'RESERVE',
        'RESERVE',
        'RESERVE',
        'RESERVE',
    ],
    'T': [
        'Patientennummer (PID)',
        'Fallnummer (FID)',
        'Dossiernummer',
        'Aufeichnungstyp',
        'Verhältnis zu Patienten 1 Code',
        'Verhältnis zu Patienten 1 Bezeichnung',
        'Patienten-Nr du für den die Beziehung gilt',
        'Verhältnis zu Patienten 2 Code',
        'Verhältnis zu Patienten 2 Bezeichnung',
        'RESERVE',
        'RESERVE',
        'RESERVE',
        'RESERVE',
        'RESERVE',
        'RESERVE',
        'RESERVE',
        'RESERVE',
        'RESERVE',
        'RESERVE',
        'RESERVE',
    ],
    'P': [
        'Patientennummer (PID)',
        'Fallnummer (FID)',
        'Dossiernummer',
        'Aufeichnungstyp',
        'Zone 00',
        'Zone 01',
        'Zone 02',
        'Zone 03',
        'Zone 04',
        'Zone 05',
        'Zone 06',
        'Zone 07',
        'Zone 08',
        'Zone 09',
        'Zone 10',
        'Zone 11',
        'Zone 12',
        'usw.',
    ],
    'S': [
        'Patientennummer (PID)',
        'Fallnummer (FID)',
        'Dossiernummer',
        'Aufeichnungstyp',
        'Zone 00',
        'Zone 01',
        'Zone 02',
        'Zone 03',
        'Zone 04',
        'Zone 05',
        'Zone 06',
        'Zone 07',
        'Zone 08',
        'Zone 09',
        'Zone 10',
        'Zone 11',
        'Zone 12',
        'usw.',
    ],
    '4': [
        'Patientennummer (PID)',
        'Fallnummer (FID)',
        'Dossiernummer',
        'Aufeichnungstyp',
        'RESERVE',
        'Sende Datum des Datensatz EXPOPA10',
        'Sende Zeit des Datensatz EXPOPA10',
        'Herkunft / Ursprung der Daten',
        'RESERVE',
        'RESERVE',
        'RESERVE',
        'RESERVE',
        'RESERVE',
        'RESERVE',
        'RESERVE',
        'RESERVE',
    ],
}


def open_file(path):
    """Öffnet eine Datei mit dem Standardprogramm des Betriebssystems."""
    try:
        if sys.platform.startswith("win"):
            os.startfile(os.path.abspath(path))
        elif sys.platform == "darwin":
            subprocess.run(["open", path], check=False)
        else:
            subprocess.run(["xdg-open", path], check=False)
    except Exception as e:
        print(f"Datei konnte nicht automatisch geöffnet werden: {e}")


def read_text(file_path):
    """Liest eine Textdatei als str: UTF-16 (BOM), UTF-8, sonst Windows-1252/Latin-1.
    So bleiben Umlaute auch bei Dateien in ANSI-Kodierung (z. B. OPA) erhalten."""
    with open(file_path, 'rb') as f:
        data = f.read()
    if data.startswith((b'\xff\xfe', b'\xfe\xff')):
        return data.decode('utf-16', errors='ignore').strip()
    for enc in ('utf-8-sig', 'cp1252'):
        try:
            return data.decode(enc).strip()
        except UnicodeDecodeError:
            continue
    return data.decode('latin-1').strip()


def looks_like_opa(first_line):
    """True, wenn die Zeile wie ein OPA-Record (Typ 1 oder R) aussieht:
    PID|FID|Dossier|Recordtyp|... mit numerischer Fallnummer."""
    parts = first_line.split(OPA_DELIMITER)
    return (len(parts) >= 16
            and parts[OPA_RECORD_TYPE_POS - 1].strip() in ('1', 'R')
            and parts[1].strip().isdigit())


def detect_extension(file_path):
    """Erkennt anhand des Inhalts, ob eine Datei hl7, opa, dat oder txt war. None, wenn unklar."""
    with open(file_path, 'rb') as f:
        data = f.read(4096)

    if data.startswith((b'\xff\xfe', b'\xfe\xff')):
        text = data.decode('utf-16', errors='ignore')
    else:
        text = data.decode('utf-8-sig', errors='ignore')

    # BOM, Leerraum und MLLP-Startzeichen (\x0b) vorne entfernen
    text = text.lstrip('\x0b\ufeff \t\r\n')

    if text.startswith('MSH|'):
        return '.hl7'

    first_line = next((l for l in text.splitlines() if l.strip()), '')
    if looks_like_opa(first_line):
        return '.opa'
    pipes, semis = first_line.count('|'), first_line.count(';')
    if pipes == 0 and semis == 0:
        return None
    return '.dat' if pipes > semis else '.txt'


def copy_testfiles(target_dir, source_dir=TESTFILE_DIR):
    """Kopiert die neueste hl7/txt/dat/opa/bin-Datei aus source_dir nach target_dir.
    .bin-Dateien werden anhand des Inhalts mit der richtigen Endung kopiert.
    Gibt die Anzahl kopierter Dateien zurück (None bei Fehler)."""
    if not os.path.isdir(source_dir):
        print(f"FEHLER: Quellordner '{source_dir}' wurde nicht gefunden.")
        return None

    candidates = [
        os.path.join(source_dir, f) for f in os.listdir(source_dir)
        if f.lower().endswith(TESTFILE_EXTENSIONS)
        and os.path.isfile(os.path.join(source_dir, f))
    ]
    if not candidates:
        print(f"FEHLER: Keine .hl7/.txt/.dat/.opa/.bin-Dateien in '{source_dir}' gefunden.")
        return None

    # Neueste Datei anhand des Änderungsdatums
    src = max(candidates, key=os.path.getmtime)
    f = os.path.basename(src)
    stem, ext = os.path.splitext(f)

    if ext.lower() == '.bin':
        new_ext = detect_extension(src)
        if new_ext is None:
            print(f"FEHLER: Dateityp der neuesten Datei '{f}' konnte nicht erkannt werden.")
            return None
        target_name = stem + new_ext
        print(f"'{f}' als '{target_name}' erkannt.")
    else:
        target_name = f

    shutil.copy2(src, os.path.join(target_dir, target_name))
    print(f"Neueste Datei '{f}' aus '{source_dir}' nach '{target_dir}' kopiert.")
    return 1


def clean_sheet_title(title, max_len=31):
    """Reinigt den Namen für ein Excel-Tabellenblatt."""
    cleaned = re.sub(r'[\\/*?:\[\]]', '_', str(title))
    return cleaned[:max_len]


def get_field_sort_key(field_id):
    """Sortier-Schlüssel für logische HL7-Reihenfolge (inkl. Komponenten)."""
    base = field_id.split('_')[0]
    last = field_id.split('_')[-1]          # z. B. "3" oder "3.1"
    parts = last.split('.')

    field_num = int(parts[0]) if parts[0].isdigit() else 999
    comp_num = int(parts[1]) if len(parts) > 1 and parts[1].isdigit() else 0

    if '[' in base:
        seg_name = base.split('[')[0]
        seg_idx = int(base.split('[')[1].replace(']', ''))
    else:
        seg_name = base
        seg_idx = 1

    if seg_name in STANDARD_SEGMENT_ORDER:
        seg_priority = STANDARD_SEGMENT_ORDER.index(seg_name)
    else:
        seg_priority = 99

    return (seg_priority, seg_idx, field_num, comp_num)


def get_opa_sort_key(field_id):
    """Sortier-Schlüssel für OPA-Felder: Recordtyp (Spezifikation), Wiederholung, Position.
    Feld-IDs: '1_5', 'R_9', 'L_11', 'L[2]_11' (2. Adress-Record)."""
    m = re.match(r'^(.+?)(?:\[(\d+)\])?_(\d+)$', field_id)
    if not m:
        return (99, 0, 0)
    rtype, rep_idx, pos = m.group(1), int(m.group(2) or 1), int(m.group(3))
    order = OPA_RECORD_ORDER.index(rtype) if rtype in OPA_RECORD_ORDER else 98
    return (order, rep_idx, pos)


# ==========================================
# PARSER: MAPPING-DATEI (.map / .mapping)
# ==========================================

def extract_mapping_descriptions(file_path):
    """Liest eine Mapping-Datei (.map / .mapping) ein."""
    field_mapping = {}
    meta_info = {
        'BESCH_BESONDERHEITEN': '',
        'BESCH_LOGIK_ORCHESTRA': '',
        'BESCH_LOGIK_UMSYSTEM': '',
        'BESCH_ANPASSUNG': ''
    }

    if not file_path or not os.path.exists(file_path):
        return field_mapping, meta_info

    with open(file_path, 'r', encoding='utf-8-sig', errors='ignore') as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith('#') or line.startswith('//'):
                continue

            if '=' in line:
                key, val = line.split('=', 1)
            elif ':' in line:
                key, val = line.split(':', 1)
            else:
                continue

            field_id = key.strip()
            description = val.strip()

            if field_id in meta_info:
                meta_info[field_id] = description
            elif field_id:
                field_mapping[field_id] = description

    return field_mapping, meta_info


# ==========================================
# PARSER: HL7, OPA & TEXTDATEIEN (.dat, .txt, .csv)
# ==========================================

def extract_hl7_fields(file_path):
    """Liest eine HL7-Datei ein und liefert ein Dict der Felder zurück."""
    with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
        raw_content = f.read().strip()

    formatted_content = raw_content.replace('\r\n', '\r').replace('\n', '\r')
    hl7_message = parse_message(formatted_content, find_groups=False)

    fields_data = {}
    segment_counters = {}

    for element in hl7_message.children:
        seg_name = element.name.upper()

        segment_counters[seg_name] = segment_counters.get(seg_name, 0) + 1
        seg_index = segment_counters[seg_name]
        seg_display_name = f"{seg_name} ({seg_index})" if seg_index > 1 else seg_name

        for field in element.children:
            # Komponenten mit ^ (bzw. Subkomponenten mit &) korrekt ausgeben
            field_val = field.to_er7()
            if not field_val:
                continue

            raw_num = field.name.split('_')[-1]
            if seg_index > 1:
                field_id = f"{seg_name}[{seg_index}]_{raw_num}"
            else:
                field_id = f"{seg_name}_{raw_num}"

            if field_id in fields_data:
                # Wiederholung (~) desselben Feldes: anhängen statt überschreiben
                fields_data[field_id]['value'] += "~" + field_val
            else:
                long_name = field.long_name if field.long_name else field_id
                fields_data[field_id] = {
                    'segment': seg_display_name,
                    'long_name': long_name,
                    'value': field_val
                }

            # Komponenten (^) als eigene Zeilen: PV1_3.1, PV1_3.2, ...
            if '^' in field_val:
                for comp_no, comp_val in enumerate(field_val.split('^'), start=1):
                    if not comp_val:
                        continue  # leere Komponenten nicht anzeigen
                    comp_id = f"{field_id}.{comp_no}"
                    if comp_id in fields_data:
                        # Wiederholung (~): anhängen
                        fields_data[comp_id]['value'] += "~" + comp_val
                    else:
                        # Segment-Index entfernen, damit PV1[2]_3.1 ebenfalls PV1_3.1 findet
                        lookup_key = re.sub(r'\[\d+\]', '', comp_id)
                        base_name = fields_data[field_id]['long_name']
                        fields_data[comp_id] = {
                            'segment': seg_display_name,
                            'long_name': COMPONENT_NAMES.get(lookup_key, f"{base_name}.{comp_no}"),
                            'value': comp_val
                        }

    return fields_data


def extract_opa_fields(file_path):
    """Liest eine OPA-Datei (EXPOPA10) ein. Jede Zeile ist ein Record, der Typ steht an
    Position 4. Feld-IDs: <Typ>_<Position>, bei wiederholten Records <Typ>[n]_<Position>
    (z. B. L_11, L[2]_11). Leere Felder werden wie bei HL7 nicht aufgenommen.
    Die Identitätsfelder 1-3 (PID/FID/Dossier) erscheinen nur einmal (Typ 1), das Feld 4
    (Aufzeichnungstyp) gar nicht, da es im Segment/Typ bereits steht."""
    fields_data = {}
    counters = {}

    for line in read_text(file_path).splitlines():
        if not line.strip():
            continue
        parts = line.split(OPA_DELIMITER)
        if len(parts) < OPA_RECORD_TYPE_POS:
            continue

        rtype = parts[OPA_RECORD_TYPE_POS - 1].strip()
        counters[rtype] = counters.get(rtype, 0) + 1
        rec_idx = counters[rtype]

        title = OPA_RECORD_TITLES.get(rtype, 'Unbekannt')
        seg_display = f"{rtype} {title}" + (f" ({rec_idx})" if rec_idx > 1 else "")
        names = OPA_FIELD_NAMES.get(rtype, [])

        for pos, val in enumerate(parts, start=1):
            val = val.strip()
            if not val:
                continue
            if pos == OPA_RECORD_TYPE_POS:
                continue
            if pos < OPA_RECORD_TYPE_POS and not (rtype == '1' and rec_idx == 1):
                continue

            field_id = f"{rtype}[{rec_idx}]_{pos}" if rec_idx > 1 else f"{rtype}_{pos}"
            fields_data[field_id] = {
                'segment': seg_display,
                'long_name': names[pos - 1] if pos <= len(names) else f"Feld {pos}",
                'value': val
            }

    return fields_data


def extract_text_fields(file_path):
    """Liest strukturierte Textdateien (.dat, .txt, .csv) ein."""
    fields_data = {}
    ext = os.path.splitext(file_path)[1].lower()

    if ext == '.dat':
        delimiter = '|'
    elif ext == '.txt':
        delimiter = ';'
    else:
        delimiter = ','

    with open(file_path, 'r', encoding='utf-8-sig', errors='ignore') as f:
        content = f.read().strip()

    if not content:
        return fields_data

    lines = [line for line in content.splitlines() if line.strip()]
    reader = list(csv.reader(lines, delimiter=delimiter))

    if not reader:
        return fields_data

    type_label = ext.replace('.', '').upper()
    first_row = reader[0]

    for idx, val in enumerate(first_row, start=1):
        field_id = f"Feld {idx}"
        fields_data[field_id] = {
            'segment': type_label,
            'long_name': field_id,
            'value': val.strip()
        }

    return fields_data


def read_raw_file_content(file_path):
    """Liest den exakten Rohinhalt der Datei für die Fußzeile aus."""
    try:
        return read_text(file_path)
    except Exception:
        return ""


# ==========================================
# EXCEL BEFÜLLUNG (DYNAMISCHE SHEETS)
# ==========================================

def write_data_sheet(ws, file_paths, mapping_file=None):
    """Befüllt ein Excel-Tabellenblatt mit Metadaten, Dateidaten und Rohcode."""
    parsed_files = []
    all_fields_set = set()
    field_order_preserve = []
    has_hl7 = False
    has_opa = False

    field_descriptions, meta_info = extract_mapping_descriptions(mapping_file) if mapping_file else ({}, {})

    # 1. Dateien verarbeiten
    for fp in file_paths:
        ext = os.path.splitext(fp)[1].lower()
        try:
            if ext == '.hl7':
                p_data = extract_hl7_fields(fp)
                has_hl7 = True
            elif ext == '.opa':
                p_data = extract_opa_fields(fp)
                has_opa = True
            else:
                p_data = extract_text_fields(fp)

            raw_content = read_raw_file_content(fp)
            parsed_files.append((fp, p_data, raw_content))

            for k in p_data.keys():
                if k not in all_fields_set:
                    all_fields_set.add(k)
                    field_order_preserve.append(k)
        except Exception as e:
            print(f"Fehler beim Einlesen von '{fp}': {e}")

    if not parsed_files:
        return

    # Pflichtfelder immer aufnehmen, auch wenn sie in keiner Datei befüllt sind
    if has_hl7:
        all_fields_set.update(ALWAYS_SHOW_FIELDS)

    # 2. Feld-Sortierung
    if has_hl7:
        all_fields_ordered = sorted(list(all_fields_set), key=get_field_sort_key)
    elif has_opa:
        all_fields_ordered = sorted(list(all_fields_set), key=get_opa_sort_key)
    else:
        def text_sort_key(f_name):
            num = re.findall(r'\d+', f_name)
            return int(num[0]) if num else 999

        all_fields_ordered = sorted(list(all_fields_set), key=text_sort_key)

    # 3. Formate & Stile
    font_file_header  = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    font_raw_header   = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    font_seg_col      = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    font_id_col       = Font(name="Calibri", size=10, bold=True, color="1F4E79")
    font_name_col     = Font(name="Calibri", size=10, bold=True, color="1F4E79")
    font_remark_col   = Font(name="Calibri", size=10, italic=True, color="333333")
    font_data         = Font(name="Calibri", size=10)
    font_raw_code     = Font(name="Consolas", size=9, color="1F2937")

    font_meta_label   = Font(name="Calibri", size=10, bold=True, color="1F4E79")
    font_meta_val     = Font(name="Calibri", size=10, italic=True, color="1F2937")
    fill_meta_label   = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    fill_meta_val     = PatternFill(start_color="F2F4F7", end_color="F2F4F7", fill_type="solid")

    fill_file_header  = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    fill_raw_header   = PatternFill(start_color="333333", end_color="333333", fill_type="solid")
    fill_seg_col      = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
    fill_id_col       = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    fill_name_col     = PatternFill(start_color="B4C6E7", end_color="B4C6E7", fill_type="solid")
    fill_remark_col   = PatternFill(start_color="E9EEF4", end_color="E9EEF4", fill_type="solid")
    fill_data_even    = PatternFill(start_color="F9FAFB", end_color="F9FAFB", fill_type="solid")
    fill_data_odd     = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    fill_raw_cell     = PatternFill(start_color="F2F4F7", end_color="F2F4F7", fill_type="solid")
    fill_always       = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )

    total_cols = len(parsed_files) + 4

    # 4. METADATEN-BLOCK ERSTELLEN
    meta_keys = [
        ("Besonderheiten", meta_info.get('BESCH_BESONDERHEITEN', '')),
        ("Logik Orchestra", meta_info.get('BESCH_LOGIK_ORCHESTRA', '')),
        ("Logik Umsystem", meta_info.get('BESCH_LOGIK_UMSYSTEM', '')),
        ("Anpassung", meta_info.get('BESCH_ANPASSUNG', ''))
    ]
    # Nur Einträge anzeigen, die in der map-Datei nicht auskommentiert (und nicht leer) sind
    meta_keys = [(label, val) for label, val in meta_keys if val]

    for idx, (label, val) in enumerate(meta_keys, start=1):
        ws.merge_cells(start_row=idx, start_column=1, end_row=idx, end_column=3)
        c_lbl = ws.cell(row=idx, column=1, value=label)
        c_lbl.font = font_meta_label
        c_lbl.fill = fill_meta_label
        c_lbl.alignment = Alignment(horizontal="left", vertical="center")

        for c_i in range(1, 4):
            ws.cell(row=idx, column=c_i).fill = fill_meta_label
            ws.cell(row=idx, column=c_i).border = thin_border

        ws.merge_cells(start_row=idx, start_column=4, end_row=idx, end_column=total_cols)
        c_val = ws.cell(row=idx, column=4, value=val)
        c_val.font = font_meta_val
        c_val.fill = fill_meta_val
        c_val.alignment = Alignment(horizontal="left", vertical="center")

        for c_i in range(4, total_cols + 1):
            ws.cell(row=idx, column=c_i).fill = fill_meta_val
            ws.cell(row=idx, column=c_i).border = thin_border

    # 5. TABELLEN-HEADER ERSTELLEN
    header_row = len(meta_keys) + 2 if meta_keys else 1

    ws.cell(row=header_row, column=1, value="Segment/Typ").font = font_file_header
    ws.cell(row=header_row, column=1).fill = fill_file_header

    ws.cell(row=header_row, column=2, value="Feld-ID").font = font_file_header
    ws.cell(row=header_row, column=2).fill = fill_file_header

    ws.cell(row=header_row, column=3, value="Feldbezeichnung").font = font_file_header
    ws.cell(row=header_row, column=3).fill = fill_file_header

    ws.cell(row=header_row, column=4, value="Bemerkung").font = font_file_header
    ws.cell(row=header_row, column=4).fill = fill_file_header

    for c in range(1, 5):
        ws.cell(row=header_row, column=c).alignment = Alignment(horizontal="center", vertical="center")

    for file_idx, (fp, _, _) in enumerate(parsed_files, start=5):
        filename = os.path.basename(fp)
        c = ws.cell(row=header_row, column=file_idx, value=filename)
        c.font = font_file_header
        c.fill = fill_file_header
        c.alignment = Alignment(horizontal="center", vertical="center")

    # 6. DATENZEILEN BEFÜLLEN
    start_data_row = header_row + 1
    last_data_row = start_data_row

    for row_offset, f_id in enumerate(all_fields_ordered):
        row_idx = start_data_row + row_offset
        last_data_row = row_idx

        seg_name, long_name = "", ""
        for _, p_data, _ in parsed_files:
            if f_id in p_data:
                seg_name = p_data[f_id]['segment']
                long_name = p_data[f_id]['long_name']
                break

        # Pflichtfeld ohne Wert in allen Dateien: Segment und Bezeichnung selbst setzen
        is_always = has_hl7 and f_id in ALWAYS_SHOW_FIELDS
        if is_always and not seg_name:
            seg_name = f_id.split('_')[0]                 # "PV1"
            long_name = COMPONENT_NAMES.get(f_id, f_id)

        remark = field_descriptions.get(f_id, "")
        row_fill = fill_data_even if row_idx % 2 == 0 else fill_data_odd
        if is_always:
            row_fill = fill_always

        c_seg = ws.cell(row=row_idx, column=1, value=seg_name)
        c_seg.font, c_seg.fill = font_seg_col, fill_seg_col
        c_seg.alignment = Alignment(horizontal="center", vertical="center")

        c_id = ws.cell(row=row_idx, column=2, value=f_id)
        c_id.font, c_id.fill = font_id_col, fill_id_col
        c_id.alignment = Alignment(horizontal="left", vertical="center")

        c_name = ws.cell(row=row_idx, column=3, value=long_name)
        c_name.font, c_name.fill = font_name_col, fill_name_col
        c_name.alignment = Alignment(horizontal="left", vertical="center")

        c_remark = ws.cell(row=row_idx, column=4, value=remark)
        c_remark.font, c_remark.fill = font_remark_col, (fill_always if is_always else fill_remark_col)
        c_remark.alignment = Alignment(horizontal="left", vertical="center")

        for file_idx, (_, p_data, _) in enumerate(parsed_files, start=5):
            val = p_data.get(f_id, {}).get('value', '')
            cell = ws.cell(row=row_idx, column=file_idx, value=val)
            cell.font = font_data
            cell.alignment = Alignment(horizontal="left", vertical="center")
            cell.border = thin_border
            cell.fill = row_fill

    for r in range(header_row, last_data_row + 1):
        for c in range(1, total_cols + 1):
            ws.cell(row=r, column=c).border = thin_border

    # 7. UNTERER BEREICH: VOLLSTÄNDIGER HL7- / RAW-CODE
    raw_row = last_data_row + 3

    ws.merge_cells(start_row=raw_row, start_column=1, end_row=raw_row, end_column=4)
    c_label = ws.cell(row=raw_row, column=1, value="Vollständige HL7- / Dateinachricht")
    c_label.font = font_raw_header
    c_label.fill = fill_raw_header
    c_label.alignment = Alignment(horizontal="center", vertical="center")

    for col_i in range(1, 5):
        cell_m = ws.cell(row=raw_row, column=col_i)
        cell_m.fill = fill_raw_header
        cell_m.border = thin_border

    for file_idx, (_, _, raw_content) in enumerate(parsed_files, start=5):
        cell_raw = ws.cell(row=raw_row, column=file_idx, value=raw_content)
        cell_raw.font = font_raw_code
        cell_raw.fill = fill_raw_cell
        cell_raw.border = thin_border
        cell_raw.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)

    max_lines = max([len(raw.splitlines()) for _, _, raw in parsed_files] + [1])
    ws.row_dimensions[raw_row].height = max(30, min(max_lines * 15, 300))

    # 8. Spaltenbreiten anpassen
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in list(col)[header_row - 1:last_data_row]:
            val_str = str(cell.value or '')
            max_len = max(max_len, len(val_str))
        ws.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 50)

    ws.freeze_panes = f'E{header_row + 1}'


# ==========================================
# SICHERUNG MANUELLER BLÄTTER IN manual.xlsx
# ==========================================

def backup_static_file(static_excel_path):
    """Legt im Ordner von static_excel_path eine Sicherheitskopie mit Zeitstempel an
    (z. B. manual_backup_20261006_073300.xlsx). Gibt den Pfad der Kopie zurück,
    None wenn keine Datei zu sichern war. Wirft eine Exception, wenn die Kopie scheitert."""
    if not os.path.exists(static_excel_path):
        return None

    folder = os.path.dirname(static_excel_path) or "."
    stem, ext = os.path.splitext(os.path.basename(static_excel_path))
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = os.path.join(folder, f"{stem}_backup_{stamp}{ext}")

    shutil.copy2(static_excel_path, backup_path)
    print(f"Sicherheitskopie erstellt: '{backup_path}'")

    # Alte Sicherungen aufräumen: nur die neuesten MAX_BACKUPS behalten
    pattern = re.compile(
        rf"^{re.escape(stem)}_backup_\d{{8}}_\d{{6}}{re.escape(ext)}$", re.IGNORECASE
    )
    backups = sorted(f for f in os.listdir(folder) if pattern.match(f))  # Zeitstempel -> sortierbar
    for old in backups[:-MAX_BACKUPS]:
        try:
            os.remove(os.path.join(folder, old))
            print(f"Alte Sicherheitskopie gelöscht: '{old}'")
        except OSError as e:
            print(f"WARNUNG: '{old}' konnte nicht gelöscht werden: {e}")

    return backup_path


def sync_manual_sheets(output_excel_path, static_excel_path, import_dir):
    """Sichert alle Blätter der bestehenden Ausgabedatei, die NICHT aus data/ erzeugt
    wurden, in die statische Vorlage (static_excel_path wird überschrieben).

    Schutzmechanismen:
    - Ohne Pillow wird manual.xlsx NICHT überschrieben (sonst gingen Bilder dauerhaft verloren).
    - Vor dem Überschreiben wird immer zuerst eine Sicherheitskopie angelegt.
    - Gespeichert wird erst in eine temporäre Datei und dann ersetzt, damit manual.xlsx
      bei einem Fehler nicht halb geschrieben zurückbleibt.
    Gibt True zurück, wenn manual.xlsx neu geschrieben wurde."""
    if not os.path.exists(output_excel_path):
        return False

    if not HAS_PILLOW:
        print(f"WARNUNG: Pillow fehlt - '{static_excel_path}' wird NICHT überschrieben, "
              f"damit keine Bilder verloren gehen.")
        return False

    # Namen der Blätter, die der Code aus data/ erzeugt
    generated = set()
    if os.path.isdir(import_dir):
        for root, dirs, files in os.walk(import_dir):
            if root == import_dir:
                continue
            if any(f.lower().endswith(DATA_EXTENSIONS) for f in files):
                generated.add(clean_sheet_title(os.path.basename(root)))

    try:
        wb = openpyxl.load_workbook(output_excel_path)
    except Exception as e:
        print(f"WARNUNG: '{output_excel_path}' konnte nicht gelesen werden, "
              f"manual.xlsx bleibt unverändert: {e}")
        return False

    for name in list(wb.sheetnames):
        if name in generated:
            wb.remove(wb[name])

    if not wb.sheetnames:
        print("Keine manuellen Blätter in der Ausgabedatei gefunden, manual.xlsx bleibt unverändert.")
        return False

    # Aktives Blatt zurücksetzen (das bisherige könnte entfernt worden sein)
    wb.active = 0
    for i, sheet in enumerate(wb.worksheets):
        sheet.sheet_view.tabSelected = (i == 0)

    os.makedirs(os.path.dirname(static_excel_path) or ".", exist_ok=True)

    # ZUERST: Sicherheitskopie der bestehenden manual.xlsx
    try:
        backup_static_file(static_excel_path)
    except Exception as e:
        print(f"FEHLER: Sicherheitskopie von '{static_excel_path}' fehlgeschlagen: {e}")
        print("manual.xlsx wird aus Sicherheitsgründen NICHT überschrieben.")
        return False

    tmp_path = static_excel_path + ".tmp"
    try:
        wb.save(tmp_path)
        os.replace(tmp_path, static_excel_path)
        print(f"{len(wb.sheetnames)} manuelle(s) Blatt/Blätter in '{static_excel_path}' gesichert.")
        return True
    except PermissionError:
        print(f"FEHLER: '{static_excel_path}' ist gesperrt (in Excel geöffnet?). Abbruch.")
        raise
    finally:
        if os.path.exists(tmp_path):
            try:
                os.remove(tmp_path)
            except OSError:
                pass


# ==========================================
# HAUPTFUNKTION / STEUERUNG
# ==========================================

def run_import(import_dir="data", output_excel_path="vergleich_transponiert.xlsx",
               static_excel_path=os.path.join("static_files", "manual.xlsx"),
               target_folder=None):
    """Lädt die statische Basis-Arbeitsmappe (inkl. Grafiken/Screenshots) und erweitert sie um die dynamischen Daten-Sheets.

    Optional: target_folder = Name eines Unterordners in import_dir. Dann werden die
    hl7/txt/dat/opa-Dateien aus 'testfile' dorthin kopiert und nach dem Speichern wird
    das gleichnamige Sheet direkt angezeigt.
    """

    if not HAS_PILLOW:
        print("\n" + "!"*80)
        print("WARNUNG: Das Python-Paket 'Pillow' (PIL) ist NICHT installiert!")
        print("Ohne Pillow entfernt openpyxl automatisch alle Screenshots/Bilder aus der Excel-Datei!")
        print("manual.xlsx wird in diesem Lauf NICHT überschrieben.")
        print("Bitte installiere es bei Bedarf mit:  pip install pillow")
        print("!"*80 + "\n")

    # Manuelle Blätter aus der letzten Ausgabedatei zurück in manual.xlsx sichern
    # (mit Sicherheitskopie, und nur wenn Pillow vorhanden ist)
    sync_manual_sheets(output_excel_path, static_excel_path, import_dir)

    # Optional: Testdateien in einen bestimmten Unterordner kopieren
    if target_folder:
        target_dir = os.path.join(import_dir, target_folder)
        if not os.path.isdir(target_dir):
            print(f"FEHLER: Unterordner '{target_folder}' existiert nicht in '{import_dir}'.")
            return
        if copy_testfiles(target_dir) is None:
            return

    created_sheets = {}

    # 1. Basis-Arbeitsmappe laden:
    if os.path.exists(static_excel_path):
        wb = openpyxl.load_workbook(static_excel_path)
        static_count = len(wb.sheetnames)
        print(f"Statische Vorlage '{static_excel_path}' mit {static_count} Blatt/Blättern geladen.")
    else:
        wb = openpyxl.Workbook()
        default_sheet = wb.active
        static_count = 0

    # 2. Dynamische Blätter aus data/ verarbeiten
    folders_processed = 0

    if os.path.exists(import_dir):
        for root, dirs, files in os.walk(import_dir):
            if root == import_dir:
                continue

            matching_files = [
                os.path.join(root, f) for f in files
                if f.lower().endswith(DATA_EXTENSIONS)
            ]

            matching_files.sort(key=lambda p: os.path.basename(p).swapcase()) #nach Dateiname sortieren

            if not matching_files:
                continue

            mapping_file = None
            for f in files:
                if f.lower().endswith(MAPPING_EXTENSIONS):
                    mapping_file = os.path.join(root, f)
                    break

            folder_name = os.path.basename(root)
            sheet_title = clean_sheet_title(folder_name)

            map_info = f" [mit Mapping: {os.path.basename(mapping_file)}]" if mapping_file else ""
            print(f"Verarbeite Ordner '{folder_name}' ({len(matching_files)} Datei(en)){map_info}...")

            ws = wb.create_sheet(title=sheet_title)
            created_sheets[folder_name] = ws
            write_data_sheet(ws, matching_files, mapping_file=mapping_file)
            folders_processed += 1

    # Falls keine manual.xlsx existierte, aber Ordner verarbeitet wurden, das Standard-Sheet löschen
    if static_count == 0 and folders_processed > 0 and 'default_sheet' in locals():
        wb.remove(default_sheet)

    if static_count > 0 or folders_processed > 0:
        # Gewünschtes Sheet als aktives Blatt setzen, damit Excel es direkt anzeigt
        if target_folder:
            target_ws = created_sheets.get(target_folder)
            if target_ws is None:
                print(f"WARNUNG: Für '{target_folder}' wurde kein Sheet erzeugt.")
            else:
                wb.active = wb.index(target_ws)
                for sheet in wb.worksheets:
                    sheet.sheet_view.tabSelected = (sheet is target_ws)

        wb.save(output_excel_path)
        print(f"\nErfolgreich gespeichert in '{output_excel_path}' ({static_count} statische(s) Sheet(s), {folders_processed} dynamische(s) Sheet(s)).")
        open_file(output_excel_path)
    else:
        print(f"Keine Dateien in '{import_dir}' oder '{static_excel_path}' gefunden.")


if __name__ == "__main__":
    # Optional: Unterordnername als Argument, z. B.  python hl7_vergleich.py MeinOrdner
    folder = sys.argv[1] if len(sys.argv) > 1 else None
    run_import("data", "vergleich_transponiert.xlsx", target_folder=folder)