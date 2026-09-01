import io
import re
import xml.etree.ElementTree as ET
from datetime import date, datetime, time

import pandas as pd
import streamlit as st

SS_NS = "urn:schemas-microsoft-com:office:spreadsheet"
NS = {"ss": SS_NS}
SS = f"{{{SS_NS}}}"

st.set_page_config(page_title="Control de llegadas", page_icon="✅", layout="wide")


def clean(value):
    if value is None:
        return ""
    return str(value).strip()


def parse_excel_xml_2003(file_bytes: bytes):
    """Lee reportes XML Spreadsheet 2003 aunque tengan extensión .xls."""
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
    period_start = None
    period_end = None

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
        raise ValueError("No pude identificar el periodo del reporte.")

    people = []
    records = []

    def at(row, idx):
        return clean(row[idx]) if idx < len(row) else ""

    i = 0
    while i < len(rows):
        row = rows[i]
        if at(row, 0) == "ID:" and at(row, 2) == "Nombre:":
            person_id = at(row, 1)
            name = at(row, 3)
            sector = at(row, 5) if at(row, 4) == "Sector:" else ""
            people.append({"ID": person_id, "Nombre": name, "Sector": sector})

            marks_row = rows[i + 1] if i + 1 < len(rows) else []
            days_in_period = (period_end - period_start).days + 1

            for day_offset in range(days_in_period):
                d = period_start + pd.Timedelta(days=day_offset)
                cell = clean(marks_row[day_offset]) if day_offset < len(marks_row) else ""

                times = re.findall(r"\b([01]?\d|2[0-3]):([0-5]\d)\b", cell)
                normalized = []
                for hh, mm in times:
                    t = f"{int(hh):02d}:{mm}"
                    if t not in normalized:
                        normalized.append(t)

                arrival = normalized[0] if normalized else ""
                departure = normalized[-1] if len(normalized) >= 2 else ""
                records.append({
                    "Fecha": pd.Timestamp(d).date(),
                    "ID": person_id,
                    "Nombre": name,
                    "Sector": sector,
                    "Llegada": arrival,
                    "Salida": departure,
                    "Marcaciones": " · ".join(normalized),
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
        score = sum(1 for col in cols if any(k in col for k in ["nombre", "name", "id", "fecha", "hora", "time"]))
        candidates.append((score, sheet, df))

    _, sheet, df = max(candidates, key=lambda x: x[0])
    lower = {str(c).strip().lower(): c for c in df.columns}

    def find_col(words):
        for low, orig in lower.items():
            if any(w in low for w in words):
                return orig
        return None

    id_col = find_col(["id", "codigo", "código", "documento"])
    name_col = find_col(["nombre", "name", "empleado", "persona"])
    sector_col = find_col(["sector", "grupo", "curso", "departamento"])
    date_col = find_col(["fecha", "date"])
    time_col = find_col(["hora", "time", "entrada", "llegada"])

    if name_col is None:
        raise ValueError(f"La hoja '{sheet}' no tiene una columna de nombre reconocible.")

    work = pd.DataFrame()
    work["ID"] = df[id_col].astype(str) if id_col is not None else range(1, len(df) + 1)
    work["Nombre"] = df[name_col].astype(str)
    work["Sector"] = df[sector_col].astype(str) if sector_col is not None else ""
    work["Fecha"] = pd.to_datetime(df[date_col], errors="coerce").dt.date if date_col is not None else date.today()

    if time_col is not None:
        def fmt_time(v):
            if pd.isna(v):
                return ""
            if isinstance(v, (datetime, pd.Timestamp)):
                return v.strftime("%H:%M")
            if isinstance(v, time):
                return v.strftime("%H:%M")
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
    if head.startswith(b"<?xml") and b"Excel.Sheet" in data[:2000]:
        return xml_report_to_attendance(data)
    return generic_excel_to_attendance(data, uploaded_file.name)


def to_excel_bytes(df):
    output = io.BytesIO()
    try:
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Asistencia")
        return output.getvalue(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "xlsx"
    except Exception:
        df.to_csv(output, index=False, encoding="utf-8-sig")
        return output.getvalue(), "text/csv", "csv"


st.title("Control de llegadas")
st.caption("Carga el reporte XLS/XLSX del equipo de asistencia y consulta quién llegó, a qué hora y quién no tiene marcación.")

uploaded = st.file_uploader("Cargar reporte de asistencia", type=["xls", "xlsx"])

if uploaded is None:
    st.info("Carga un archivo para comenzar. La aplicación reconoce el formato XML de Excel 2003 usado por el reporte 09-Resumen.xls.")
    st.stop()

try:
    people, attendance, period_start, period_end = load_report(uploaded)
except Exception as exc:
    st.error(f"No pude procesar el archivo: {exc}")
    st.stop()

if people.empty:
    st.warning("El archivo se pudo abrir, pero no encontré personas.")
    st.stop()

with_marks = attendance.loc[attendance["Llegada"].astype(str).str.len() > 0, "Fecha"]
default_date = max(with_marks) if not with_marks.empty else period_start
if default_date < period_start or default_date > period_end:
    default_date = period_start

st.success(f"Archivo reconocido: {len(people)} personas · Periodo {period_start.strftime('%d/%m/%Y')} a {period_end.strftime('%d/%m/%Y')}")

c1, c2, c3 = st.columns([1.2, 1, 1])
with c1:
    selected_date = st.date_input(
        "Fecha a consultar",
        value=default_date,
        min_value=period_start,
        max_value=period_end,
    )
with c2:
    expected_time = st.time_input("Hora esperada de llegada", value=time(6, 20))
with c3:
    tolerance = st.number_input("Tolerancia (minutos)", min_value=0, max_value=180, value=0, step=5)

# Construcción de la vista diaria
view = attendance[attendance["Fecha"] == selected_date].copy()
view = people.merge(view, on=["ID", "Nombre", "Sector"], how="left")
view["Fecha"] = selected_date
view["Llegada"] = view["Llegada"].fillna("")
view["Salida"] = view["Salida"].fillna("")
view["Marcaciones"] = view["Marcaciones"].fillna("")
view["Estado"] = view["Llegada"].apply(lambda x: "Llegó" if clean(x) else "Sin marcación")

limit_minutes = expected_time.hour * 60 + expected_time.minute + int(tolerance)

def punctuality(arrival):
    if not arrival:
        return "Sin marcación"
    h, m = map(int, arrival.split(":"))
    return "A tiempo" if h * 60 + m <= limit_minutes else "Tarde"

view["Puntualidad"] = view["Llegada"].apply(punctuality)

arrived = int((view["Estado"] == "Llegó").sum())
pending = int((view["Estado"] == "Sin marcación").sum())
late = int((view["Puntualidad"] == "Tarde").sum())
on_time = int((view["Puntualidad"] == "A tiempo").sum())
rate = (arrived / len(view) * 100) if len(view) else 0

m1, m2, m3, m4 = st.columns(4)
m1.metric("Total Personas", len(view))
m2.metric("Llegaron", arrived)
m3.metric("Sin marcación (No llegaron)", pending)
m4.metric("% Asistencia", f"{rate:.1f}%")

if arrived:
    st.caption(f"De los que llegaron: {on_time} a tiempo y {late} con tardanza (Límite: {expected_time.strftime('%H:%M')} + {int(tolerance)} min).")

# =========================
# GRÁFICAS DEL DÍA SELECCIONADO
# =========================
st.markdown("---")
st.subheader(f"📊 Estado de asistencia del día ({selected_date.strftime('%d/%m/%Y')})")

g1, g2 = st.columns(2)

with g1:
    st.markdown("**Resumen general del día**")
    status_df = pd.DataFrame(
        {
            "Cantidad": [on_time, late, pending]
        },
        index=["Llegó A tiempo", "Llegó Tarde", "Sin marcación (No llegó)"]
    )
    st.bar_chart(status_df)

with g2:
    st.markdown("**Asistencia por Sector / Grupo**")
    if "Sector" in view.columns and not view["Sector"].dropna().empty:
        sector_chart = (
            view.groupby(["Sector", "Estado"])
            .size()
            .unstack(fill_value=0)
        )
        st.bar_chart(sector_chart)
    else:
        st.info("No hay datos de sectores/grupos para graficar.")

st.markdown("---")

# Filtros para la tabla de abajo
st.subheader("📋 Detalle de la lista")
f1, f2, f3 = st.columns([1.2, 1, 1])
with f1:
    search = st.text_input("Buscar persona", placeholder="Nombre o ID")
with f2:
    sectors = sorted([x for x in view["Sector"].dropna().astype(str).unique() if x and x != "nan"])
    sector_filter = st.multiselect("Sector / grupo", sectors, default=sectors)
with f3:
    status_filter = st.multiselect("Estado", ["Llegó", "Sin marcación"], default=["Llegó", "Sin marcación"])

filtered = view.copy()
if search.strip():
    q = search.strip().lower()
    filtered = filtered[
        filtered["Nombre"].astype(str).str.lower().str.contains(q, na=False)
        | filtered["ID"].astype(str).str.lower().str.contains(q, na=False)
    ]
if sectors:
    filtered = filtered[filtered["Sector"].astype(str).isin(sector_filter)]
filtered = filtered[filtered["Estado"].isin(status_filter)]

show_cols = ["ID", "Nombre", "Sector", "Llegada", "Salida", "Puntualidad", "Estado", "Marcaciones"]
st.dataframe(filtered[show_cols], hide_index=True)

st.subheader("📌 Listas rápidas")
tab1, tab2, tab3 = st.tabs(["✅ Llegaron", "❌ No han llegado", "⏰ Tardanzas"])
with tab1:
    arrived_df = view[view["Estado"] == "Llegó"][["ID", "Nombre", "Sector", "Llegada", "Puntualidad"]]
    st.dataframe(arrived_df, hide_index=True)
with tab2:
    pending_df = view[view["Estado"] == "Sin marcación"][["ID", "Nombre", "Sector"]]
    st.dataframe(pending_df, hide_index=True)
with tab3:
    late_df = view[view["Puntualidad"] == "Tarde"][["ID", "Nombre", "Sector", "Llegada"]]
    st.dataframe(late_df, hide_index=True)

export = view[["Fecha", "ID", "Nombre", "Sector", "Llegada", "Salida", "Puntualidad", "Estado", "Marcaciones"]].copy()

file_bytes, mime_type, file_ext = to_excel_bytes(export)

st.download_button(
    "Descargar resultado",
    data=file_bytes,
    file_name=f"asistencia_{selected_date.isoformat()}.{file_ext}",
    mime=mime_type,
)