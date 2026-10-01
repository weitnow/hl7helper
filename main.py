import os
import re
import csv
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from hl7apy.parser import parse_message
from hl7apy.exceptions import HL7apyException

# Standard-Reihenfolge der HL7-Segmente für die Sortierung
STANDARD_SEGMENT_ORDER = [
    'MSH', 'EVN', 'PID', 'PD1', 'PV1', 'PV2', 
    'IN1', 'IN2', 'IN3', 'NK1', 'GT1', 'AL1', 'DG1', 'PR1', 
    'OBX', 'NTE', 'ORC', 'OBR', 'SPM'
]

def clean_sheet_title(title, max_len=31):
    """Reinigt den Namen für ein Excel-Tabellenblatt."""
    cleaned = re.sub(r'[\\/*?:\[\]]', '_', title)
    return cleaned[:max_len]


def get_field_sort_key(field_id):
    """Sortier-Schlüssel für logische HL7-Reihenfolge."""
    base = field_id.split('_')[0]
    field_num = int(field_id.split('_')[-1]) if field_id.split('_')[-1].isdigit() else 999
    
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

    return (seg_priority, seg_idx, field_num)


# ==========================================
# PARSER: MAPPING-DATEI (.map / .mapping)
# ==========================================

def extract_mapping_descriptions(file_path):
    """Liest eine Mapping-Datei (.map / .mapping) ein.
    Erwartet Zeilen im Format: Feld_ID=Beschreibung oder Feld_ID:Beschreibung
    """
    mapping = {}
    if not file_path or not os.path.exists(file_path):
        return mapping

    with open(file_path, 'r', encoding='utf-8-sig', errors='ignore') as f:
        for line in f:
            line = line.strip()
            # Kommentare und leere Zeilen überspringen
            if not line or line.startswith('#') or line.startswith('//'):
                continue
            
            # Trennung nach '=' oder ':'
            if '=' in line:
                key, val = line.split('=', 1)
            elif ':' in line:
                key, val = line.split(':', 1)
            else:
                continue

            field_id = key.strip()
            description = val.strip()
            if field_id:
                mapping[field_id] = description

    return mapping


# ==========================================
# PARSER: HL7 & TEXTDATEIEN (.dat, .txt, .csv)
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
            # Falls das Feld Wiederholungen (Repetitions mit ~) hat, diese verknüpfen:
            if hasattr(field, 'children') and len(field.children) > 0 and hasattr(field.children[0], 'value'):
                # Wiederholende Instanzen zusammenfügen
                val_list = [str(rep.value) for rep in field.children if rep.value]
                field_val = "~".join(val_list)
            else:
                field_val = str(field.value) if field.value else ""

            if field_val:
                raw_num = field.name.split('_')[-1]
                if seg_index > 1:
                    field_id = f"{seg_name}[{seg_index}]_{raw_num}"
                else:
                    field_id = f"{seg_name}_{raw_num}"

                long_name = field.long_name if field.long_name else field_id

                fields_data[field_id] = {
                    'segment': seg_display_name,
                    'long_name': long_name,
                    'value': field_val
                }

    return fields_data


def extract_text_fields(file_path):
    """Liest strukturierte Textdateien (.dat, .txt, .csv) ein.
    Verwendet | für .dat und ; für .txt.
    Erzeugt fortlaufende Felder: Feld 1, Feld 2, Feld 3 ...
    """
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




# ==========================================
# EXCEL BEFÜLLUNG
# ==========================================

def write_data_sheet(ws, file_paths, mapping_file=None):
    """Befüllt ein Excel-Tabellenblatt mit verarbeiteten Dateidaten, HL7-Feldbezeichnungen und Bemerkungen aus Mapping."""
    parsed_files = []
    all_fields_set = set()
    field_order_preserve = []
    has_hl7 = False

    # Mapping-Datei einlesen (falls vorhanden)
    field_descriptions = extract_mapping_descriptions(mapping_file) if mapping_file else {}

    # 1. Dateien verarbeiten
    for fp in file_paths:
        ext = os.path.splitext(fp)[1].lower()
        try:
            if ext == '.hl7':
                p_data = extract_hl7_fields(fp)
                has_hl7 = True
            else:
                p_data = extract_text_fields(fp)

            parsed_files.append((fp, p_data))
            for k in p_data.keys():
                if k not in all_fields_set:
                    all_fields_set.add(k)
                    field_order_preserve.append(k)
        except Exception as e:
            print(f"Fehler beim Einlesen von '{fp}': {e}")

    if not parsed_files:
        return

    # 2. Feld-Sortierung
    if has_hl7:
        all_fields_ordered = sorted(list(all_fields_set), key=get_field_sort_key)
    else:
        def text_sort_key(f_name):
            num = re.findall(r'\d+', f_name)
            return int(num[0]) if num else 999

        all_fields_ordered = sorted(list(all_fields_set), key=text_sort_key)

    # 3. Formate & Stile
    font_file_header  = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    font_seg_col      = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    font_id_col       = Font(name="Calibri", size=10, bold=True, color="1F4E79")
    font_name_col     = Font(name="Calibri", size=10, bold=True, color="1F4E79")
    font_remark_col   = Font(name="Calibri", size=10, italic=True, color="333333")
    font_data         = Font(name="Calibri", size=10)

    fill_file_header  = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    fill_seg_col      = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
    fill_id_col       = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    fill_name_col     = PatternFill(start_color="B4C6E7", end_color="B4C6E7", fill_type="solid")
    fill_remark_col   = PatternFill(start_color="E9EEF4", end_color="E9EEF4", fill_type="solid")
    fill_data_even    = PatternFill(start_color="F9FAFB", end_color="F9FAFB", fill_type="solid")
    fill_data_odd     = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )

    # 4. Header-Zeile erstellen
    ws.cell(row=1, column=1, value="Segment/Typ").font = font_file_header
    ws.cell(row=1, column=1).fill = fill_file_header

    ws.cell(row=1, column=2, value="Feld-ID").font = font_file_header
    ws.cell(row=1, column=2).fill = fill_file_header

    ws.cell(row=1, column=3, value="Feldbezeichnung").font = font_file_header
    ws.cell(row=1, column=3).fill = fill_file_header

    ws.cell(row=1, column=4, value="Bemerkung").font = font_file_header
    ws.cell(row=1, column=4).fill = fill_file_header

    for c in range(1, 5):
        ws.cell(row=1, column=c).alignment = Alignment(horizontal="center", vertical="center")

    # Dateinamen ab Spalte E (Spalte 5) eintragen
    for file_idx, (fp, _) in enumerate(parsed_files, start=5):
        filename = os.path.basename(fp)
        c = ws.cell(row=1, column=file_idx, value=filename)
        c.font = font_file_header
        c.fill = fill_file_header
        c.alignment = Alignment(horizontal="center", vertical="center")

    # 5. Datenzeilen befüllen
    for row_idx, f_id in enumerate(all_fields_ordered, start=2):
        seg_name, long_name = "", ""
        for _, p_data in parsed_files:
            if f_id in p_data:
                seg_name = p_data[f_id]['segment']
                long_name = p_data[f_id]['long_name']
                break

        # Bemerkung aus der Mapping-Datei auslesen (falls vorhanden)
        remark = field_descriptions.get(f_id, "")

        row_fill = fill_data_even if row_idx % 2 == 0 else fill_data_odd

        # Spalte A: Segment / Typ
        c_seg = ws.cell(row=row_idx, column=1, value=seg_name)
        c_seg.font, c_seg.fill = font_seg_col, fill_seg_col
        c_seg.alignment = Alignment(horizontal="center", vertical="center")

        # Spalte B: Feld-ID
        c_id = ws.cell(row=row_idx, column=2, value=f_id)
        c_id.font, c_id.fill = font_id_col, fill_id_col
        c_id.alignment = Alignment(horizontal="left", vertical="center")

        # Spalte C: Feldbezeichnung (originale HL7- / Datei-Bezeichnung)
        c_name = ws.cell(row=row_idx, column=3, value=long_name)
        c_name.font, c_name.fill = font_name_col, fill_name_col
        c_name.alignment = Alignment(horizontal="left", vertical="center")

        # Spalte D: Bemerkung (aus map-Datei)
        c_remark = ws.cell(row=row_idx, column=4, value=remark)
        c_remark.font, c_remark.fill = font_remark_col, fill_remark_col
        c_remark.alignment = Alignment(horizontal="left", vertical="center")

        # Ab Spalte E: Feldwerte je Datei eintragen
        for file_idx, (_, p_data) in enumerate(parsed_files, start=5):
            val = p_data.get(f_id, {}).get('value', '')
            cell = ws.cell(row=row_idx, column=file_idx, value=val)
            cell.font = font_data
            cell.alignment = Alignment(horizontal="left", vertical="center")
            cell.border = thin_border
            cell.fill = row_fill

    # Ränder setzen
    for r in range(1, len(all_fields_ordered) + 2):
        for c in range(1, len(parsed_files) + 5):
            ws.cell(row=r, column=c).border = thin_border

    # 6. Spaltenbreiten anpassen
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            val_str = str(cell.value or '')
            max_len = max(max_len, len(val_str))
        ws.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 50)

    # Fixieren der ersten 4 Spalten (Segment, Feld-ID, Feldbezeichnung, Bemerkung) und der ersten Zeile
    ws.freeze_panes = 'E2'


# ==========================================
# HAUPTFUNKTION / STEUERUNG
# ==========================================

def run_import(import_dir="data", output_excel_path="vergleich_transponiert.xlsx"):
    """Liest alle Unterordner aus 'data' aus."""
    if not os.path.exists(import_dir):
        print(f"Ordner '{import_dir}' wurde nicht gefunden.")
        return

    data_extensions = ('.hl7', '.dat', '.txt', '.csv')
    mapping_extensions = ('.map', '.mapping')

    wb = openpyxl.Workbook()
    default_sheet = wb.active
    folders_processed = 0

    for root, dirs, files in os.walk(import_dir):
        if root == import_dir:
            continue

        # Daten-Dateien filtern
        matching_files = [
            os.path.join(root, f) for f in files 
            if f.lower().endswith(data_extensions)
        ]

        if not matching_files:
            continue

        # Nach optionaler Mapping-Datei im selben Ordner suchen
        mapping_file = None
        for f in files:
            if f.lower().endswith(mapping_extensions):
                mapping_file = os.path.join(root, f)
                break

        folder_name = os.path.basename(root)
        sheet_title = clean_sheet_title(folder_name)

        map_info = f" [mit Mapping: {os.path.basename(mapping_file)}]" if mapping_file else ""
        print(f"Verarbeite Ordner '{folder_name}' ({len(matching_files)} Datei(en)){map_info}...")

        ws = wb.create_sheet(title=sheet_title)
        write_data_sheet(ws, matching_files, mapping_file=mapping_file)
        folders_processed += 1

    if folders_processed > 0:
        wb.remove(default_sheet)
        wb.save(output_excel_path)
        print(f"\nErfolgreich gespeichert in '{output_excel_path}' ({folders_processed} Sheet(s) erstellt).")
    else:
        print(f"Keine Unterordner mit unterstützten Dateien in '{import_dir}' gefunden.")


if __name__ == "__main__":
    run_import("data", "vergleich_transponiert.xlsx")