import io
import os
import re
import xml.etree.ElementTree as ET
from datetime import date, datetime, time

import pandas as pd
import streamlit as st

SS_NS = "urn:schemas-microsoft-com:office:spreadsheet"
NS = {"ss": SS_NS}
SS = f"{{{SS_NS}}}"

# Configuración inicial de la página
st.set_page_config(page_title="Control de llegadas", page_icon="✅", layout="wide")

# Estilos en modo oscuro (Azul Oscuro / Slate)
st.markdown("""
    
""", unsafe_allow_html=True)


def get_avatar_path(name, person_id):
    """Busca la foto del archivo local en el repositorio actual."""
    name_str = str(name).lower()

    if "ashlin" in name_str or "matutes" in name_str:
        filename = "Ashlin.jpeg"
    elif "deivyd" in name_str or "deyvid" in name_str or "fonseca" in name_str:
        filename = "Deyvid.jpeg"
    elif "nelson" in name_str or "lopez" in name_str or "lópez" in name_str:
        filename = "Nelson.jpeg"
    elif "yolanda" in name_str or "isabella" in name_str or "hernandez" in name_str or "hernández" in name_str:
        filename = "Isabella.jpeg"
    else:
        return f"https://api.dicebear.com/7.x/avataaars/svg?seed={person_id}"

    # Verificar si el archivo existe físicamente en el repositorio
    if os.path.exists(filename):
        return filename
    return f"https://api.dicebear.com/7.x/avataaars/svg?seed={person_id}"


def clean(value):
    return str(value).strip() if value is not None else ""


def parse_excel_xml_2003(file_bytes: bytes):
    root = ET.fromstring(file_bytes)
    sheets = {}
    for ws in root.findall("ss:Worksheet", NS):
        name = ws.attrib.get(SS + "Name", "Hoja")
        table = ws.find("ss:Table", NS)
        out_rows = []
        if table is None:
            sheets[name] = out_rows
            continue

        for row in table.findall("ss:Row", NS):
            values = {}
            col = 1
            for cell in row.findall("ss:Cell", NS):
                idx = cell.attrib.get(SS + "Index")
                if idx:
                    col = int(idx)
                data = cell.find("ss:Data", NS)
                values[col] = data.text if data is not None and data.text is not None else ""
                col += 1
            max_col = max(values.keys(), default=0)
            out_rows.append([values.get(i, "") for i in range(1, max_col + 1)])
        sheets[name] = out_rows
    return sheets


def xml_report_to_attendance(file_bytes: bytes):
    sheets = parse_excel_xml_2003(file_bytes)
    if "Detalles" not in sheets:
        raise ValueError("No se encontró la hoja 'Detalles' esperada en el reporte.")

    rows = sheets["Detalles"]
    period_start, period_end = None, None
    period_regex = re.compile(r"(\d{4})\.(\d{2})\.(\d{2})~(\d{4})\.(\d{2})\.(\d{2})")

    for row in rows[:10]:
        for value in row:
            m = period_regex.search(clean(value))
            if m:
                period_start = date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
                period_end = date(int(m.group(4)), int(m.group(5)), int(m.group(6)))
                break
        if period_start:
            break

    if not period_start:
        raise ValueError("No se pudo identificar el periodo del reporte.")

    people, records = [], []
    def at(row, idx): return clean(row[idx]) if idx < len(row) else ""

    i = 0
    while i < len(rows):
        row = rows[i]
        if at(row, 0) == "ID:" and at(row, 2) == "Nombre:":
            person_id, name = at(row, 1), at(row, 3)
            sector = at(row, 5) if at(row, 4) == "Sector:" else ""
            people.append({"ID": person_id, "Nombre": name, "Sector": sector})

            marks_row = rows[i + 1] if i + 1 < len(rows) else []
            days_in_period = (period_end - period_start).days + 1

            for day_offset in range(days_in_period):
                d = period_start + pd.Timedelta(days=day_offset)
                cell = clean(marks_row[day_offset]) if day_offset < len(marks_row) else ""
                times = re.findall(r"\b([01]?\d|2[0-3]):([0-5]\d)\b", cell)
                normalized = [f"{int(hh):02d}:{mm}" for hh, mm in times]
                
                norm_unique = []
                for t in normalized:
                    if t not in norm_unique: norm_unique.append(t)

                arrival = norm_unique[0] if norm_unique else ""
                departure = norm_unique[-1] if len(norm_unique) >= 2 else ""
                records.append({
                    "Fecha": pd.Timestamp(d).date(),
                    "ID": person_id,
                    "Nombre": name,
                    "Sector": sector,
                    "Llegada": arrival,
                    "Salida": departure,
                    "Marcaciones": " · ".join(norm_unique),
                })
            i += 2
        else:
            i += 1

    people_df = pd.DataFrame(people).drop_duplicates(subset=["ID"], keep="first")
    attendance_df = pd.DataFrame(records)
    return people_df, attendance_df, period_start, period_end


def generic_excel_to_attendance(file_bytes: bytes, filename: str):
    bio = io.BytesIO(file_bytes)
    ext = filename.lower().rsplit(".", 1)[-1]
    engine = "openpyxl" if ext == "xlsx" else "xlrd"
    xls = pd.ExcelFile(bio, engine=engine)

    candidates = []
    for sheet in xls.sheet_names:
        df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet, engine=engine)
        cols = {str(c).strip().lower(): c for c in df.columns}
        score = sum(1 for col in cols if any(k in col for k in ["nombre", "name", "id", "fecha", "hora"]))
        candidates.append((score, sheet, df))

    _, sheet, df = max(candidates, key=lambda x: x[0])
    lower = {str(c).strip().lower(): c for c in df.columns}

    def find_col(words):
        for low, orig in lower.items():
            if any(w in low for w in words): return orig
        return None

    id_col = find_col(["id", "codigo", "código"])
    name_col = find_col(["nombre", "name", "empleado"])
    sector_col = find_col(["sector", "grupo", "curso"])
    date_col = find_col(["fecha", "date"])
    time_col = find_col(["hora", "time", "entrada", "llegada"])

    if name_col is None:
        raise ValueError(f"La hoja '{sheet}' no tiene columna de nombres.")

    work = pd.DataFrame()
    work["ID"] = df[id_col].astype(str) if id_col else range(1, len(df) + 1)
    work["Nombre"] = df[name_col].astype(str)
    work["Sector"] = df[sector_col].astype(str) if sector_col else ""
    work["Fecha"] = pd.to_datetime(df[date_col], errors="coerce").dt.date if date_col else date.today()
    
    if time_col:
        def fmt_time(v):
            if pd.isna(v): return ""
            if isinstance(v, (datetime, pd.Timestamp, time)): return v.strftime("%H:%M")
            m = re.search(r"\b([01]?\d|2[0-3]):([0-5]\d)\b", str(v))
            return f"{int(m.group(1)):02d}:{m.group(2)}" if m else ""
        work["Llegada"] = df[time_col].apply(fmt_time)
    else:
        work["Llegada"] = ""

    work["Salida"] = ""
    work["Marcaciones"] = work["Llegada"]
    people = work[["ID", "Nombre", "Sector"]].drop_duplicates(subset=["ID"])
    dates = work["Fecha"].dropna()
    start = min(dates) if len(dates) else date.today()
    end = max(dates) if len(dates) else date.today()
    return people, work, start, end


def load_report(uploaded_file):
    data = uploaded_file.getvalue()
    head = data[:300].lstrip()
    if head.startswith(b" 0, "Fecha"]
default_date = max(with_marks) if not with_marks.empty else period_start

c1, c2, c3 = st.columns([1.2, 1, 1])
with c1:
    selected_date = st.date_input("Fecha a consultar", value=default_date, min_value=period_start, max_value=period_end)
with c2:
    expected_time = st.time_input("Hora esperada de llegada", value=time(6, 20))
with c3:
    tolerance = st.number_input("Tolerancia (minutos)", min_value=0, max_value=180, value=0, step=5)

view = attendance[attendance["Fecha"] == selected_date].copy()
view = people.merge(view, on=["ID", "Nombre", "Sector"], how="left")
view["Fecha"] = selected_date
view["Llegada"] = view["Llegada"].fillna("")
view["Salida"] = view["Salida"].fillna("")
view["Marcaciones"] = view["Marcaciones"].fillna("")
view["Estado"] = view["Llegada"].apply(lambda x: "Llegó" if clean(x) else "Sin marcación")

limit_minutes = expected_time.hour * 60 + expected_time.minute + int(tolerance)

def punctuality(arrival):
    if not arrival: return "Pendiente"
    h, m = map(int, arrival.split(":"))
    return "A tiempo" if h * 60 + m <= limit_minutes else "Tarde"

view["Puntualidad"] = view["Llegada"].apply(punctuality)

# Asignar la ruta local de la foto existente en el repo
view["Foto"] = view.apply(lambda row: get_avatar_path(row["Nombre"], row["ID"]), axis=1)

arrived = int((view["Estado"] == "Llegó").sum())
pending = int((view["Estado"] == "Sin marcación").sum())
late = int((view["Puntualidad"] == "Tarde").sum())
on_time = int((view["Puntualidad"] == "A tiempo").sum())
rate = (arrived / len(view) * 100) if len(view) else 0

m1, m2, m3, m4 = st.columns(4)
m1.metric("Total Personas", len(view))
m2.metric("Llegaron", arrived)
m3.metric("Sin marcación", pending)
m4.metric("% Asistencia", f"{rate:.1f}%")

st.write("---")

st.subheader("📋 Detalle de la lista")
export_cols = ["Foto", "Fecha", "ID", "Nombre", "Sector", "Llegada", "Salida", "Puntualidad", "Estado", "Marcaciones"]

st.dataframe(
    view[export_cols],
    column_config={
        "Foto": st.column_config.ImageColumn("Foto / Avatar", help="Foto cargada en el repositorio"),
    },
    hide_index=True,
)

st.subheader("📌 Listas rápidas")
tab1, tab2, tab3 = st.tabs(["✅ Llegaron", "❌ No han llegado", "⏰ Tardanzas"])

with tab1:
    arrived_df = view[view["Estado"] == "Llegó"][["Foto", "ID", "Nombre", "Sector", "Llegada", "Puntualidad"]]
    st.dataframe(arrived_df, column_config={"Foto": st.column_config.ImageColumn("Foto")}, hide_index=True)

with tab2:
    pending_df = view[view["Estado"] == "Sin marcación"][["Foto", "ID", "Nombre", "Sector", "Puntualidad"]]
    st.dataframe(pending_df, column_config={"Foto": st.column_config.ImageColumn("Foto")}, hide_index=True)

with tab3:
    late_df = view[view["Puntualidad"] == "Tarde"][["Foto", "ID", "Nombre", "Sector", "Llegada"]]
    st.dataframe(late_df, column_config={"Foto": st.column_config.ImageColumn("Foto")}, hide_index=True)

export = view[export_cols].copy()
file_bytes, mime_type, file_ext = to_excel_bytes(export)

st.download_button(
    "Descargar resultado en Excel",
    data=file_bytes,
    file_name=f"asistencia_{selected_date.isoformat()}.{file_ext}",
    mime=mime_type,
)