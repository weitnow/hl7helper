import os
import re
import csv
import openpyxl
import sys
import subprocess
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

# Standard-Reihenfolge der HL7-Segmente für die Sortierung
STANDARD_SEGMENT_ORDER = [
    'MSH', 'EVN', 'PID', 'PD1', 'PV1', 'PV2', 
    'IN1', 'IN2', 'IN3', 'NK1', 'GT1', 'AL1', 'DG1', 'PR1', 
    'OBX', 'NTE', 'ORC', 'OBR', 'SPM'
]

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


def clean_sheet_title(title, max_len=31):
    """Reinigt den Namen für ein Excel-Tabellenblatt."""
    cleaned = re.sub(r'[\\/*?:\[\]]', '_', str(title))
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
    """Liest eine Mapping-Datei (.map / .mapping) ein."""
    field_mapping = {}
    meta_info = {
        'BESCH_BESONDERHEITEN': '',
        'BESCH_LOGIK_ORCHESTRA': '',
        'BESCH_LOGIK_UMSYSTEM': ''
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
        with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
            return f.read().strip()
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

    field_descriptions, meta_info = extract_mapping_descriptions(mapping_file) if mapping_file else ({}, {})

    # 1. Dateien verarbeiten
    for fp in file_paths:
        ext = os.path.splitext(fp)[1].lower()
        try:
            if ext == '.hl7':
                p_data = extract_hl7_fields(fp)
                has_hl7 = True
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
        ("Logik Umsystem", meta_info.get('BESCH_LOGIK_UMSYSTEM', ''))
    ]

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
    header_row = 5

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
    start_data_row = 6
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

        remark = field_descriptions.get(f_id, "")
        row_fill = fill_data_even if row_idx % 2 == 0 else fill_data_odd

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
        c_remark.font, c_remark.fill = font_remark_col, fill_remark_col
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

    ws.freeze_panes = 'E6'


# ==========================================
# HAUPTFUNKTION / STEUERUNG
# ==========================================

def run_import(import_dir="data", output_excel_path="vergleich_transponiert.xlsx", static_excel_path=os.path.join("static_files", "manual.xlsx")):
    """Lädt die statische Basis-Arbeitsmappe (inkl. Grafiken/Screenshots) und erweitert sie um die dynamischen Daten-Sheets."""
    
    if not HAS_PILLOW:
        print("\n" + "!"*80)
        print("WARNUNG: Das Python-Paket 'Pillow' (PIL) ist NICHT installiert!")
        print("Ohne Pillow entfernt openpyxl automatisch alle Screenshots/Bilder aus der Excel-Datei!")
        print("Bitte installiere es bei Bedarf mit:  pip install pillow")
        print("!"*80 + "\n")

    data_extensions = ('.hl7', '.dat', '.txt', '.csv')
    mapping_extensions = ('.map', '.mapping')

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
                if f.lower().endswith(data_extensions)
            ]

            if not matching_files:
                continue

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

    # Falls keine manual.xlsx existierte, aber Ordner verarbeitet wurden, das Standard-Sheet löschen
    if static_count == 0 and folders_processed > 0 and 'default_sheet' in locals():
        wb.remove(default_sheet)

    if static_count > 0 or folders_processed > 0:
        wb.save(output_excel_path)
        print(f"\nErfolgreich gespeichert in '{output_excel_path}' ({static_count} statische(s) Sheet(s), {folders_processed} dynamische(s) Sheet(s)).")
        open_file(output_excel_path)
    else:
        print(f"Keine Dateien in '{import_dir}' oder '{static_excel_path}' gefunden.")


if __name__ == "__main__":
    run_import("data", "vergleich_transponiert.xlsx")