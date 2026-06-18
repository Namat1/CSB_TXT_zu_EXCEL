from __future__ import annotations

import io
import re
import traceback
from copy import copy
from typing import Dict, Tuple, List

import pandas as pd
import streamlit as st
from openpyxl.styles import Alignment, Font, PatternFill


# ------------------------------------------------------------
# Grundeinstellungen
# ------------------------------------------------------------

st.set_page_config(
    page_title="CSB Textdatei Auswertung",
    page_icon="🚚",
    layout="wide",
)

st.markdown(
    """
    <style>
        .block-container {
            max-width: 100% !important;
            padding-top: 1.2rem;
            padding-left: 1.6rem;
            padding-right: 1.6rem;
        }
        div[data-testid="stMetric"] {
            background: #111827;
            border: 1px solid #263244;
            border-radius: 14px;
            padding: 14px 16px;
        }
        div[data-testid="stMetric"] label {
            color: #cbd5e1 !important;
        }
        div[data-testid="stMetric"] div {
            color: #ffffff !important;
        }
        .small-info {
            color: #9ca3af;
            font-size: 0.92rem;
        }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("🚚 CSB Textdatei Auswertung mit Diagnose")
st.caption("Nur CSB Textdatei hochladen, Touren und Kunden prüfen und als Excel-Datei exportieren.")
st.info(
    "CSB mit 103/F8 für die ganzen Wochentage von 1001-1886 bis 6001-6886 generieren "
    "und als Textdatei exportieren. Danach die Textdatei hier hochladen."
)

WOCHENTAG_MAP = {
    "montag": "Mo",
    "dienstag": "Die",
    "mittwoch": "Mitt",
    "donnerstag": "Don",
    "freitag": "Fr",
    "samstag": "Sam",
    "sonntag": "So",
}

TAG_REIHENFOLGE = ["Mo", "Die", "Mitt", "Don", "Fr", "Sam", "So", ""]


# ------------------------------------------------------------
# Hilfsfunktionen
# ------------------------------------------------------------

def clean_text(value) -> str:
    if value is None:
        return ""
    value = str(value)
    value = value.replace("\x0c", " ")
    value = value.replace("\xa0", " ")
    value = value.replace("\x81", " ")
    if re.fullmatch(r"\d+\.0", value.strip()):
        value = value.strip()[:-2]
    value = re.sub(r"\s+", " ", value)
    return value.strip(" \t\r\n.;")


def norm_num(value) -> str:
    value = clean_text(value)
    if value == "":
        return ""
    value = value.replace(",", ".")
    if re.fullmatch(r"\d+\.0", value):
        value = value[:-2]
    if re.fullmatch(r"\d+", value):
        return str(int(value))
    return value


def norm_tour(value) -> str:
    return norm_num(value)


def normalize_day(value: str) -> str:
    v = clean_text(value).lower().replace(".", "")
    if v in WOCHENTAG_MAP:
        return WOCHENTAG_MAP[v]

    aliases = {
        "mo": "Mo",
        "mon": "Mo",
        "monday": "Mo",
        "die": "Die",
        "di": "Die",
        "dienst": "Die",
        "tuesday": "Die",
        "mitt": "Mitt",
        "mi": "Mitt",
        "mittw": "Mitt",
        "wednesday": "Mitt",
        "don": "Don",
        "do": "Don",
        "thursday": "Don",
        "fr": "Fr",
        "frei": "Fr",
        "friday": "Fr",
        "sam": "Sam",
        "sa": "Sam",
        "saturday": "Sam",
        "so": "So",
        "son": "So",
        "sunday": "So",
    }
    return aliases.get(v, clean_text(value))


def liefetag_aus_tour(tour: str) -> str:
    tour = norm_tour(tour)
    if not tour:
        return ""
    return {
        "1": "Mo",
        "2": "Die",
        "3": "Mitt",
        "4": "Don",
        "5": "Fr",
        "6": "Sam",
        "7": "So",
    }.get(tour[0], "")


def decode_txt_bytes(data: bytes) -> str:
    for enc in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
        try:
            return data.decode(enc)
        except UnicodeDecodeError:
            continue
    return data.decode("cp1252", errors="replace")


# ------------------------------------------------------------
# CSB Textdatei lesen
# ------------------------------------------------------------

def analyze_customer_line(line: str):
    """
    Gibt zurück:
    - ("OK", tuple_mit_kundendaten)
    - ("GRUND", None), wenn die Zeile nicht als Kunde erkannt wurde

    Bewusst etwas toleranter als die erste Version:
    Die Punkt-Spalten am Ende werden nicht mehr zwingend verlangt.
    """
    raw = line.rstrip("\r\n").replace("\xa0", " ")

    if not raw.strip():
        return "Leerzeile", None

    # Kopf-, Trenn- und Summenzeilen ignorieren
    if re.search(r"\b(Tour|Wochentag|Fahrer|Anzahl Kunden|LKW|Datum|Seite)\b", raw, re.IGNORECASE):
        return "Kopf-/Summenzeile", None

    # Kundenzeile muss im Regelfall eine 3- bis 6-stellige CSB Nummer am Anfang haben.
    # Unterstützt:
    #          10502 Kunde ...
    #     1    13822 Kunde ...
    # Außerdem toleranter, falls weniger führende Leerzeichen vorhanden sind.
    match_start = re.match(r"^\s*(?:(\d{1,3})\s+)?(\d{3,6})\s+", raw)
    if not match_start:
        # Nur kundenähnliche Zeilen in die Diagnose aufnehmen.
        if re.search(r"\d{3,6}", raw) and re.search(r"[A-Za-zÄÖÜäöüß]", raw):
            return "Keine passende CSB Nummer am Zeilenanfang", None
        return "Keine Kundenzeile", None

    ladereihenfolge_aus_textdatei = match_start.group(1) or ""
    csb = match_start.group(2)

    plz_matches = list(re.finditer(r"\b\d{5}\b", raw))
    if not plz_matches:
        return "Keine fünfstellige Postleitzahl gefunden", None

    plz_match = plz_matches[-1]
    plz = plz_match.group(0)

    ort_raw = raw[plz_match.end():]
    ort_raw = re.sub(r"(?:\s+\.){2,}.*$", "", ort_raw)
    ort = clean_text(ort_raw)

    mid = raw[match_start.end():plz_match.start()].rstrip()

    if not clean_text(mid):
        return "Kein Kundenname / keine Straße zwischen CSB und Postleitzahl", None

    # CSB Festbreite: Name ungefähr 21 Zeichen, danach Straße.
    kunde = clean_text(mid[:21])
    strasse = clean_text(mid[21:])

    return "OK", (ladereihenfolge_aus_textdatei, csb, kunde, strasse, plz, ort)


def extract_customer_line(line: str):
    status, customer = analyze_customer_line(line)
    if status == "OK":
        return customer
    return None


def parse_csb_ladeplan(text: str) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    current_tour = ""
    current_wochentag_raw = ""
    current_liefertag = ""
    current_tour_text = ""
    position = 0

    kunden_rows = []
    nicht_erkannte_rows: List[dict] = []
    tour_meta: Dict[str, dict] = {}

    tour_re = re.compile(r"^\s*Tour\s+(\d{3,6})\b(.*?)(?:LKW:|$)", re.IGNORECASE)
    day_re = re.compile(r"^\s*Wochentag\s+(.+?)(?:Fahrer:|$)", re.IGNORECASE)
    count_re = re.compile(r"^\s*(\d+)\s+Anzahl Kunden\b", re.IGNORECASE)

    for raw_line in text.splitlines():
        line = raw_line.rstrip("\n\r")

        day_match = day_re.search(line)
        if day_match:
            current_wochentag_raw = clean_text(day_match.group(1))
            current_liefertag = normalize_day(current_wochentag_raw)

        tour_match = tour_re.search(line)
        if tour_match:
            current_tour = norm_tour(tour_match.group(1))
            current_tour_text = clean_text(tour_match.group(2))
            position = 0

            if not current_liefertag and current_tour:
                current_liefertag = liefetag_aus_tour(current_tour)

            tour_meta[current_tour] = {
                "Tour": current_tour,
                "Liefertag": current_liefertag,
                "Wochentag aus Textdatei": current_wochentag_raw,
                "Tour Text": current_tour_text,
                "Erwartete Kunden": None,
            }
            continue

        count_match = count_re.search(line)
        if count_match and current_tour:
            tour_meta.setdefault(
                current_tour,
                {
                    "Tour": current_tour,
                    "Liefertag": current_liefertag,
                    "Wochentag aus Textdatei": current_wochentag_raw,
                    "Tour Text": current_tour_text,
                    "Erwartete Kunden": None,
                },
            )
            tour_meta[current_tour]["Erwartete Kunden"] = int(count_match.group(1))
            continue

        status, customer = analyze_customer_line(line)
        if status == "OK" and customer and current_tour:
            position += 1
            ladereihenfolge_aus_textdatei, csb, kunde, strasse, plz, ort = customer

            kunden_rows.append(
                {
                    "Tour": current_tour,
                    "Liefertag": current_liefertag,
                    "Wochentag aus Textdatei": current_wochentag_raw,
                    "Ladereihenfolge aus Textdatei": ladereihenfolge_aus_textdatei,
                    "Position im Tourblock": position,
                    "CSB": norm_num(csb),
                    "Kunde": kunde,
                    "Straße": strasse,
                    "Postleitzahl": norm_num(plz),
                    "Ort": ort,
                    "Tour Text": current_tour_text,
                }
            )
        elif current_tour and status not in ("Leerzeile", "Kopf-/Summenzeile", "Keine Kundenzeile"):
            nicht_erkannte_rows.append(
                {
                    "Tour": current_tour,
                    "Liefertag": current_liefertag,
                    "Grund": status,
                    "Originalzeile": line,
                }
            )

    kunden_df = pd.DataFrame(kunden_rows)

    if kunden_df.empty:
        empty_touren = pd.DataFrame(
            columns=[
                "Tour",
                "Liefertag",
                "Wochentag aus Textdatei",
                "Tour Text",
                "Erwartete Kunden",
                "Erkannte Kunden",
                "Differenz",
                "Status",
            ]
        )
        empty_pruefung = pd.DataFrame(
            columns=["Tour", "Liefertag", "Erwartete Kunden", "Erkannte Kunden", "Differenz", "Status"]
        )
        return kunden_df, empty_touren, empty_pruefung, pd.DataFrame(nicht_erkannte_rows)

    erkannte = (
        kunden_df.groupby("Tour", as_index=False)
        .size()
        .rename(columns={"size": "Erkannte Kunden"})
    )

    touren_df = pd.DataFrame(tour_meta.values()).merge(erkannte, on="Tour", how="outer")
    touren_df["Erwartete Kunden"] = pd.to_numeric(touren_df["Erwartete Kunden"], errors="coerce")
    touren_df["Erkannte Kunden"] = (
        pd.to_numeric(touren_df["Erkannte Kunden"], errors="coerce").fillna(0).astype(int)
    )
    touren_df["Differenz"] = touren_df["Erkannte Kunden"] - touren_df["Erwartete Kunden"]

    def status(row):
        if pd.isna(row["Erwartete Kunden"]):
            return "Keine Sollzahl gefunden"
        if row["Differenz"] == 0:
            return "OK"
        return "Abweichung"

    touren_df["Status"] = touren_df.apply(status, axis=1)

    kunden_df = kunden_df.sort_values(["Tour", "Position im Tourblock"], kind="stable").reset_index(drop=True)
    touren_df = touren_df.sort_values("Tour", kind="stable").reset_index(drop=True)
    pruefung_df = touren_df[
        ["Tour", "Liefertag", "Erwartete Kunden", "Erkannte Kunden", "Differenz", "Status"]
    ].copy()

    nicht_erkannte_df = pd.DataFrame(nicht_erkannte_rows)

    return kunden_df, touren_df, pruefung_df, nicht_erkannte_df


# ------------------------------------------------------------
# Zusatzauswertungen
# ------------------------------------------------------------

def build_tagesuebersicht(kunden_df: pd.DataFrame, touren_df: pd.DataFrame) -> pd.DataFrame:
    if kunden_df.empty:
        return pd.DataFrame(columns=["Liefertag", "Touren", "Kunden"])

    kunden_je_tag = (
        kunden_df.groupby("Liefertag", dropna=False)
        .agg(Touren=("Tour", "nunique"), Kunden=("CSB", "count"))
        .reset_index()
    )

    if not touren_df.empty and "Erwartete Kunden" in touren_df.columns:
        erwartet = (
            touren_df.groupby("Liefertag", dropna=False)["Erwartete Kunden"]
            .sum(min_count=1)
            .reset_index()
            .rename(columns={"Erwartete Kunden": "Erwartete Kunden"})
        )
        kunden_je_tag = kunden_je_tag.merge(erwartet, on="Liefertag", how="left")
        kunden_je_tag["Differenz"] = kunden_je_tag["Kunden"] - kunden_je_tag["Erwartete Kunden"]

    sort_map = {tag: i for i, tag in enumerate(TAG_REIHENFOLGE)}
    kunden_je_tag["_sort"] = kunden_je_tag["Liefertag"].map(sort_map).fillna(99)
    kunden_je_tag = kunden_je_tag.sort_values("_sort", kind="stable").drop(columns="_sort")
    return kunden_je_tag.reset_index(drop=True)


def build_doppelte_kunden(kunden_df: pd.DataFrame) -> pd.DataFrame:
    if kunden_df.empty:
        return pd.DataFrame(columns=kunden_df.columns)

    counts = kunden_df.groupby(["CSB", "Liefertag"], as_index=False).size().rename(columns={"size": "Anzahl"})
    mehrfach = counts[counts["Anzahl"] > 1]
    if mehrfach.empty:
        return pd.DataFrame(columns=list(kunden_df.columns) + ["Anzahl"])

    return kunden_df.merge(mehrfach, on=["CSB", "Liefertag"], how="inner").sort_values(
        ["CSB", "Liefertag", "Tour", "Position im Tourblock"], kind="stable"
    )


def build_excel_export(
    kunden_df: pd.DataFrame,
    touren_df: pd.DataFrame,
    pruefung_df: pd.DataFrame,
    tagesuebersicht_df: pd.DataFrame,
    doppelte_kunden_df: pd.DataFrame,
    nicht_erkannte_df: pd.DataFrame,
) -> bytes:
    output = io.BytesIO()

    auffaellige_touren_df = pruefung_df[pruefung_df["Status"] != "OK"].copy() if not pruefung_df.empty else pruefung_df

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        kunden_df.to_excel(writer, sheet_name="Kunden aus Textdatei", index=False)
        touren_df.to_excel(writer, sheet_name="Touren Gesamt", index=False)
        pruefung_df.to_excel(writer, sheet_name="Touren Prüfung", index=False)
        auffaellige_touren_df.to_excel(writer, sheet_name="Auffällige Touren", index=False)
        tagesuebersicht_df.to_excel(writer, sheet_name="Tagesübersicht", index=False)
        doppelte_kunden_df.to_excel(writer, sheet_name="Doppelte Kunden", index=False)
        nicht_erkannte_df.to_excel(writer, sheet_name="Nicht erkannte Zeilen", index=False)

        workbook = writer.book
        header_fill = PatternFill(fill_type="solid", fgColor="1F2937")
        header_font = Font(bold=True, color="FFFFFF")
        header_alignment = Alignment(horizontal="center", vertical="center")

        for worksheet in workbook.worksheets:
            worksheet.freeze_panes = "A2"
            worksheet.auto_filter.ref = worksheet.dimensions

            for cell in worksheet[1]:
                cell.font = copy(header_font)
                cell.fill = copy(header_fill)
                cell.alignment = copy(header_alignment)

            for column in worksheet.columns:
                letter = column[0].column_letter
                max_len = 0
                for cell in column:
                    value = "" if cell.value is None else str(cell.value)
                    max_len = max(max_len, len(value))
                worksheet.column_dimensions[letter].width = min(max(max_len + 2, 10), 55)

    return output.getvalue()


def show_dataframe(df: pd.DataFrame, height: int = 560):
    st.dataframe(df, use_container_width=True, hide_index=True, height=height)


# ------------------------------------------------------------
# Oberfläche
# ------------------------------------------------------------

uploaded_txt = st.file_uploader("CSB Textdatei hochladen", type=["txt"])

if uploaded_txt is None:
    st.info("Bitte eine CSB Textdatei hochladen.")
    st.stop()

try:
    with st.spinner("Textdatei wird gelesen und ausgewertet..."):
        txt_text = decode_txt_bytes(uploaded_txt.getvalue())
        csb_kunden_df, csb_touren_df, csb_pruefung_df, nicht_erkannte_df = parse_csb_ladeplan(txt_text)
        tagesuebersicht_df = build_tagesuebersicht(csb_kunden_df, csb_touren_df)
        doppelte_kunden_df = build_doppelte_kunden(csb_kunden_df)

    if csb_kunden_df.empty:
        st.error("In der CSB Textdatei wurden keine Kunden erkannt.")
        st.stop()

    anzahl_touren = csb_kunden_df["Tour"].nunique()
    anzahl_kunden = len(csb_kunden_df)
    auffaellige_touren = int((csb_pruefung_df["Status"] != "OK").sum()) if not csb_pruefung_df.empty else 0
    ohne_sollzahl = int((csb_pruefung_df["Status"] == "Keine Sollzahl gefunden").sum()) if not csb_pruefung_df.empty else 0
    doppelte_anzahl = len(doppelte_kunden_df)

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Kunden erkannt", f"{anzahl_kunden:,}".replace(",", "."))
    c2.metric("Touren erkannt", f"{anzahl_touren:,}".replace(",", "."))
    c3.metric("Auffällige Touren", f"{auffaellige_touren:,}".replace(",", "."))
    c4.metric("Ohne Sollzahl", f"{ohne_sollzahl:,}".replace(",", "."))
    c5.metric("Doppelte Kunden", f"{doppelte_anzahl:,}".replace(",", "."))

    if auffaellige_touren == 0:
        st.success("Die erkannte Kundenzahl passt bei allen Touren zur Sollzahl aus der Textdatei.")
    else:
        st.warning("Es gibt Touren mit abweichender Kundenzahl oder ohne gefundene Sollzahl.")

    if not nicht_erkannte_df.empty:
        st.info(f"Es wurden {len(nicht_erkannte_df)} kundenähnliche Zeilen nicht übernommen. Details stehen im Reiter „Nicht erkannte Zeilen“ und im Excel-Export.")

    excel_bytes = build_excel_export(
        csb_kunden_df,
        csb_touren_df,
        csb_pruefung_df,
        tagesuebersicht_df,
        doppelte_kunden_df,
        nicht_erkannte_df,
    )

    st.download_button(
        "Excel-Auswertung herunterladen",
        data=excel_bytes,
        file_name="CSB_Textdatei_Auswertung.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

    st.divider()

    filter_col1, filter_col2, filter_col3 = st.columns([1, 1, 2])

    with filter_col1:
        vorhandene_tage = [tag for tag in TAG_REIHENFOLGE if tag in set(csb_kunden_df["Liefertag"].fillna(""))]
        tag_filter = st.multiselect("Liefertag filtern", vorhandene_tage, default=vorhandene_tage)

    with filter_col2:
        tour_suche = st.text_input("Tour suchen", value="")

    with filter_col3:
        suchtext = st.text_input("Kunde, CSB, Straße oder Ort suchen", value="")

    gefiltert_df = csb_kunden_df.copy()

    if tag_filter:
        gefiltert_df = gefiltert_df[gefiltert_df["Liefertag"].isin(tag_filter)]

    if tour_suche.strip():
        gefiltert_df = gefiltert_df[gefiltert_df["Tour"].astype(str).str.contains(tour_suche.strip(), case=False, na=False)]

    if suchtext.strip():
        pattern = re.escape(suchtext.strip())
        suchspalten = ["CSB", "Kunde", "Straße", "Postleitzahl", "Ort", "Tour Text"]
        mask = False
        for spalte in suchspalten:
            mask = mask | gefiltert_df[spalte].astype(str).str.contains(pattern, case=False, na=False)
        gefiltert_df = gefiltert_df[mask]

    tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(
        [
            "Kunden aus Textdatei",
            "Touren Prüfung",
            "Auffällige Touren",
            "Tagesübersicht",
            "Doppelte Kunden",
            "Nicht erkannte Zeilen",
        ]
    )

    with tab1:
        st.caption("Gefilterte Kundenliste aus der CSB Textdatei.")
        show_dataframe(gefiltert_df)

    with tab2:
        st.caption("Prüfung der erwarteten Kundenzahl gegen die erkannte Kundenzahl je Tour.")
        show_dataframe(csb_pruefung_df)

    with tab3:
        auffaellige_touren_df = csb_pruefung_df[csb_pruefung_df["Status"] != "OK"].copy()
        show_dataframe(auffaellige_touren_df)

    with tab4:
        st.caption("Zusammenfassung je Liefertag.")
        show_dataframe(tagesuebersicht_df, height=360)

    with tab5:
        if doppelte_kunden_df.empty:
            st.success("Keine doppelten Kunden je Liefertag gefunden.")
        else:
            show_dataframe(doppelte_kunden_df)

    with tab6:
        st.caption("Diese Zeilen sahen kundenähnlich aus, konnten aber nicht sauber als Kunde gelesen werden.")
        if nicht_erkannte_df.empty:
            st.success("Keine kundenähnlichen Zeilen übersprungen.")
        else:
            show_dataframe(nicht_erkannte_df)

except Exception:
    st.error("Fehler beim Verarbeiten der Textdatei.")
    st.code(traceback.format_exc())
