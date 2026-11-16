import streamlit as st
import sqlite3
import pandas as pd
import requests
import json
import io

from datetime import datetime
from docx import Document
from docx.shared import Pt


# =====================================================
# SOLARAI 1.0
# PV-Planung, Projektverwaltung und Offerten
# =====================================================

st.set_page_config(
    page_title="SolarAI",
    page_icon="☀️",
    layout="wide",
    initial_sidebar_state="expanded"
)


# =====================================================
# DESIGN
# =====================================================

st.markdown("""
<style>

.stApp {
    background-color: #080D18;
    color: #E8F0FF;
}

section[data-testid="stSidebar"] {
    background-color: #101827;
}

h1, h2, h3 {
    color: #EAF2FF;
}

.stButton > button {
    background: linear-gradient(90deg, #0066FF, #00BFFF);
    color: white;
    border: none;
    border-radius: 10px;
    padding: 10px 20px;
    font-weight: bold;
}

.stButton > button:hover {
    background: #0088FF;
    color: white;
}

div[data-testid="stMetric"] {
    background: #121E31;
    border: 1px solid #203B60;
    border-radius: 12px;
    padding: 18px;
}

div[data-testid="stForm"] {
    background: #101A2B;
    border: 1px solid #243B5A;
    border-radius: 12px;
    padding: 20px;
}

</style>
""", unsafe_allow_html=True)


# =====================================================
# DATENBANK
# =====================================================

DB_NAME = "solarai.db"


def get_connection():
    return sqlite3.connect(DB_NAME)


def init_database():

    conn = get_connection()
    cursor = conn.cursor()

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS projects (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            customer TEXT,
            address TEXT,
            building_type TEXT,
            roof_type TEXT,
            module_count INTEGER,
            module_power REAL,
            inverter TEXT,
            battery REAL,
            created_at TEXT
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS products (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            category TEXT,
            unit TEXT,
            purchase_price REAL,
            selling_price REAL
        )
    """)

    conn.commit()
    conn.close()


init_database()


# =====================================================
# DATENBANK FUNKTIONEN
# =====================================================

def save_project(
    name,
    customer,
    address,
    building_type,
    roof_type,
    module_count,
    module_power,
    inverter,
    battery
):

    conn = get_connection()

    conn.execute("""
        INSERT INTO projects (
            name,
            customer,
            address,
            building_type,
            roof_type,
            module_count,
            module_power,
            inverter,
            battery,
            created_at
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        name,
        customer,
        address,
        building_type,
        roof_type,
        module_count,
        module_power,
        inverter,
        battery,
        datetime.now().strftime("%d.%m.%Y")
    ))

    conn.commit()
    conn.close()


def load_projects():

    conn = get_connection()

    df = pd.read_sql_query(
        "SELECT * FROM projects ORDER BY id DESC",
        conn
    )

    conn.close()

    return df


def save_product(
    name,
    category,
    unit,
    purchase_price,
    selling_price
):

    conn = get_connection()

    conn.execute("""
        INSERT INTO products (
            name,
            category,
            unit,
            purchase_price,
            selling_price
        )
        VALUES (?, ?, ?, ?, ?)
    """, (
        name,
        category,
        unit,
        purchase_price,
        selling_price
    ))

    conn.commit()
    conn.close()


def load_products():

    conn = get_connection()

    df = pd.read_sql_query(
        "SELECT * FROM products ORDER BY id DESC",
        conn
    )

    conn.close()

    return df


# =====================================================
# BERECHNUNGEN
# =====================================================

def calculate_pv_power(module_count, module_power):

    return module_count * module_power / 1000


def calculate_annual_yield(power_kwp, specific_yield):

    return power_kwp * specific_yield


def calculate_quote(material_items, labor_items, margin, vat):

    material_cost = sum(
        item["Menge"] * item["Einzelpreis"]
        for item in material_items
    )

    labor_cost = sum(
        item["Stunden"] * item["Stundensatz"]
        for item in labor_items
    )

    subtotal = material_cost + labor_cost

    margin_amount = subtotal * margin / 100

    net_total = subtotal + margin_amount

    vat_amount = net_total * vat / 100

    gross_total = net_total + vat_amount

    return {
        "material": material_cost,
        "labor": labor_cost,
        "subtotal": subtotal,
        "margin": margin_amount,
        "net": net_total,
        "vat": vat_amount,
        "gross": gross_total
    }


# =====================================================
# WORD OFFERTE
# =====================================================

def create_word_quote(
    project_name,
    customer,
    address,
    material_items,
    labor_items,
    totals,
    vat_rate
):

    doc = Document()

    title = doc.add_heading("SOLARAI", 0)
    title.style = "Title"

    doc.add_heading("Offerte Photovoltaikanlage", level=1)

    doc.add_paragraph(
        f"Datum: {datetime.now().strftime('%d.%m.%Y')}"
    )

    doc.add_paragraph(f"Projekt: {project_name}")
    doc.add_paragraph(f"Kunde: {customer}")
    doc.add_paragraph(f"Adresse: {address}")

    doc.add_heading("Material", level=2)

    table = doc.add_table(rows=1, cols=4)
    table.style = "Table Grid"

    headers = [
        "Position",
        "Menge",
        "Einzelpreis CHF",
        "Total CHF"
    ]

    for i, header in enumerate(headers):
        table.rows[0].cells[i].text = header

    for item in material_items:

        row = table.add_row().cells

        row[0].text = item["Position"]
        row[1].text = str(item["Menge"])
        row[2].text = f'{item["Einzelpreis"]:.2f}'
        row[3].text = (
            f'{item["Menge"] * item["Einzelpreis"]:.2f}'
        )

    doc.add_heading("Arbeitsleistungen", level=2)

    table = doc.add_table(rows=1, cols=4)
    table.style = "Table Grid"

    headers = [
        "Arbeit",
        "Stunden",
        "CHF/h",
        "Total CHF"
    ]

    for i, header in enumerate(headers):
        table.rows[0].cells[i].text = header

    for item in labor_items:

        row = table.add_row().cells

        row[0].text = item["Arbeit"]
        row[1].text = str(item["Stunden"])
        row[2].text = f'{item["Stundensatz"]:.2f}'
        row[3].text = (
            f'{item["Stunden"] * item["Stundensatz"]:.2f}'
        )

    doc.add_heading("Zusammenfassung", level=2)

    doc.add_paragraph(
        f'Material: CHF {totals["material"]:,.2f}'
    )

    doc.add_paragraph(
        f'Arbeit: CHF {totals["labor"]:,.2f}'
    )

    doc.add_paragraph(
        f'Zuschlag: CHF {totals["margin"]:,.2f}'
    )

    doc.add_paragraph(
        f'Netto: CHF {totals["net"]:,.2f}'
    )

    doc.add_paragraph(
        f'MWST ({vat_rate:.1f}%): CHF {totals["vat"]:,.2f}'
    )

    doc.add_heading(
        f'Gesamttotal: CHF {totals["gross"]:,.2f}',
        level=2
    )

    doc.add_paragraph(
        "Diese Offerte ist ein Entwurf und muss vor "
        "der Abgabe fachlich und kaufmännisch geprüft werden."
    )

    buffer = io.BytesIO()

    doc.save(buffer)

    buffer.seek(0)

    return buffer


# =====================================================
# SIDEBAR
# =====================================================

st.sidebar.title("☀️ SolarAI")

st.sidebar.caption("PV Intelligence Platform")

page = st.sidebar.radio(
    "Navigation",
    [
        "🏠 Dashboard",
        "📂 Projekte",
        "📐 PV-Planung",
        "💰 Offerten",
        "📦 Produkte",
        "🤖 KI-Assistent",
        "⚙️ Einstellungen"
    ]
)

st.sidebar.divider()

st.sidebar.success("SolarAI 1.0")

st.sidebar.caption("Lokale Entwicklungsumgebung")


# =====================================================
# DASHBOARD
# =====================================================

if page == "🏠 Dashboard":

    st.title("☀️ SolarAI Dashboard")

    st.write(
        "Willkommen bei deiner persönlichen "
        "PV-Planungsplattform."
    )

    projects = load_projects()
    products = load_products()

    total_projects = len(projects)

    total_power = 0

    if not projects.empty:

        total_power = (
            projects["module_count"] *
            projects["module_power"] / 1000
        ).sum()

    col1, col2, col3 = st.columns(3)

    col1.metric(
        "PV-Projekte",
        total_projects
    )

    col2.metric(
        "Geplante Leistung",
        f"{total_power:.2f} kWp"
    )

    col3.metric(
        "Produkte",
        len(products)
    )

    st.divider()

    st.subheader("Letzte Projekte")

    if projects.empty:

        st.info("Noch keine Projekte vorhanden.")

    else:

        st.dataframe(
            projects[
                [
                    "name",
                    "customer",
                    "module_count",
                    "module_power",
                    "created_at"
                ]
            ],
            use_container_width=True,
            hide_index=True
        )


# =====================================================
# PROJEKTE
# =====================================================

elif page == "📂 Projekte":

    st.title("📂 Projektverwaltung")

    tab1, tab2 = st.tabs([
        "Neues Projekt",
        "Projektübersicht"
    ])

    with tab1:

        with st.form("new_project"):

            name = st.text_input("Projektname")

            customer = st.text_input("Kunde")

            address = st.text_input("Projektadresse")

            building_type = st.selectbox(
                "Gebäudetyp",
                [
                    "Einfamilienhaus",
                    "Mehrfamilienhaus",
                    "Gewerbe",
                    "Industrie",
                    "Andere"
                ]
            )

            roof_type = st.selectbox(
                "Dachtyp",
                [
                    "Ziegeldach",
                    "Flachdach",
                    "Blechdach",
                    "Fassade",
                    "Andere"
                ]
            )

            module_count = st.number_input(
                "Anzahl Module",
                min_value=0,
                value=30
            )

            module_power = st.number_input(
                "Modulleistung Wp",
                min_value=0.0,
                value=450.0
            )

            inverter = st.text_input(
                "Wechselrichter"
            )

            battery = st.number_input(
                "Batteriespeicher kWh",
                min_value=0.0,
                value=0.0
            )

            submitted = st.form_submit_button(
                "Projekt speichern"
            )

            if submitted:

                if not name.strip():

                    st.error(
                        "Bitte einen Projektnamen eingeben."
                    )

                else:

                    save_project(
                        name,
                        customer,
                        address,
                        building_type,
                        roof_type,
                        module_count,
                        module_power,
                        inverter,
                        battery
                    )

                    st.success(
                        "Projekt erfolgreich gespeichert!"
                    )

    with tab2:

        projects = load_projects()

        if projects.empty:

            st.info("Keine Projekte vorhanden.")

        else:

            st.dataframe(
                projects,
                use_container_width=True,
                hide_index=True
            )


# =====================================================
# PV PLANUNG
# =====================================================

elif page == "📐 PV-Planung":

    st.title("📐 PV-Planung")

    st.write("Berechne die wichtigsten Anlagendaten.")

    col1, col2 = st.columns(2)

    with col1:

        module_count = st.number_input(
            "Anzahl PV-Module",
            min_value=1,
            value=30
        )

        module_power = st.number_input(
            "Modulleistung (Wp)",
            min_value=1.0,
            value=450.0
        )

    with col2:

        specific_yield = st.number_input(
            "Spezifischer Jahresertrag (kWh/kWp)",
            min_value=0.0,
            value=1000.0
        )

        battery = st.number_input(
            "Batteriespeicher (kWh)",
            min_value=0.0,
            value=0.0
        )

    power = calculate_pv_power(
        module_count,
        module_power
    )

    annual_yield = calculate_annual_yield(
        power,
        specific_yield
    )

    st.divider()

    col1, col2, col3 = st.columns(3)

    col1.metric(
        "Anlagenleistung",
        f"{power:.2f} kWp"
    )

    col2.metric(
        "Geschätzter Jahresertrag",
        f"{annual_yield:,.0f} kWh"
    )

    col3.metric(
        "Batteriespeicher",
        f"{battery:.1f} kWh"
    )

    st.info(
        "Der Jahresertrag ist eine vereinfachte "
        "Schätzung auf Basis des eingegebenen "
        "spezifischen Ertrags. Standort, Ausrichtung, "
        "Verschattung und Verluste müssen separat "
        "berücksichtigt werden."
    )


# =====================================================
# OFFERTEN
# =====================================================

elif page == "💰 Offerten":

    st.title("💰 Offertengenerator")

    st.write(
        "Erstelle eine professionelle PV-Offerte "
        "mit Material- und Arbeitskosten."
    )

    projects = load_projects()

    if projects.empty:

        st.warning(
            "Bitte zuerst ein Projekt erstellen."
        )

    else:

        project_options = {
            f'{row["id"]} - {row["name"]}': row
            for _, row in projects.iterrows()
        }

        selected = st.selectbox(
            "Projekt auswählen",
            list(project_options.keys())
        )

        project = project_options[selected]

        st.subheader("Materialpositionen")

        products = load_products()

        default_material = pd.DataFrame([
            {
                "Position": "PV-Modul",
                "Menge": int(project["module_count"]),
                "Einzelpreis": 0.0
            },
            {
                "Position": "Wechselrichter",
                "Menge": 1,
                "Einzelpreis": 0.0
            },
            {
                "Position": "Montagesystem",
                "Menge": 1,
                "Einzelpreis": 0.0
            }
        ])

        material_df = st.data_editor(
            default_material,
            num_rows="dynamic",
            use_container_width=True,
            key=f'material_{project["id"]}'
        )

        st.subheader("Arbeitsleistungen")

        default_labor = pd.DataFrame([
            {
                "Arbeit": "PV-Montage",
                "Stunden": 0.0,
                "Stundensatz": 85.0
            },
            {
                "Arbeit": "Elektroinstallation",
                "Stunden": 0.0,
                "Stundensatz": 95.0
            },
            {
                "Arbeit": "Inbetriebnahme",
                "Stunden": 0.0,
                "Stundensatz": 95.0
            }
        ])

        labor_df = st.data_editor(
            default_labor,
            num_rows="dynamic",
            use_container_width=True,
            key=f'labor_{project["id"]}'
        )

        col1, col2 = st.columns(2)

        with col1:

            margin = st.number_input(
                "Zuschlag (%)",
                min_value=0.0,
                value=0.0
            )

        with col2:

            vat = st.number_input(
                "MWST (%)",
                min_value=0.0,
                value=8.1
            )

        if st.button("Offerte berechnen"):

            try:

                clean_material = material_df.dropna(
                    subset=["Position", "Menge", "Einzelpreis"]
                ).copy()

                clean_labor = labor_df.dropna(
                    subset=["Arbeit", "Stunden", "Stundensatz"]
                ).copy()

                clean_material["Menge"] = pd.to_numeric(
                    clean_material["Menge"]
                )

                clean_material["Einzelpreis"] = pd.to_numeric(
                    clean_material["Einzelpreis"]
                )

                clean_labor["Stunden"] = pd.to_numeric(
                    clean_labor["Stunden"]
                )

                clean_labor["Stundensatz"] = pd.to_numeric(
                    clean_labor["Stundensatz"]
                )

                if (
                    (clean_material["Menge"] < 0).any()
                    or (clean_material["Einzelpreis"] < 0).any()
                    or (clean_labor["Stunden"] < 0).any()
                    or (clean_labor["Stundensatz"] < 0).any()
                ):

                    st.error(
                        "Mengen, Stunden und Preise "
                        "dürfen nicht negativ sein."
                    )

                else:

                    material_items = clean_material.to_dict(
                        "records"
                    )

                    labor_items = clean_labor.to_dict(
                        "records"
                    )

                    totals = calculate_quote(
                        material_items,
                        labor_items,
                        margin,
                        vat
                    )

                    st.session_state["quote"] = {
                        "project_id": int(project["id"]),
                        "material": material_items,
                        "labor": labor_items,
                        "totals": totals,
                        "vat": vat
                    }

            except (ValueError, TypeError) as error:

                st.error(
                    f"Bitte die Eingaben überprüfen: {error}"
                )

        quote = st.session_state.get("quote")

        if quote and quote["project_id"] == int(project["id"]):

            totals = quote["totals"]

            st.divider()

            st.subheader("Offertenzusammenfassung")

            col1, col2, col3 = st.columns(3)

            col1.metric(
                "Material",
                f'CHF {totals["material"]:,.2f}'
            )

            col2.metric(
                "Arbeit",
                f'CHF {totals["labor"]:,.2f}'
            )

            col3.metric(
                "Gesamttotal inkl. MWST",
                f'CHF {totals["gross"]:,.2f}'
            )

            word_file = create_word_quote(
                project["name"],
                project["customer"],
                project["address"],
                quote["material"],
                quote["labor"],
                totals,
                quote["vat"]
            )

            st.download_button(
                "📄 Word-Offerte herunterladen",
                data=word_file,
                file_name=f'SolarAI_Offerte_{project["id"]}.docx',
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "wordprocessingml.document"
                )
            )

            st.warning(
                "Vor dem Versand alle Preise, Mengen, "
                "Leistungen und Vertragsbedingungen prüfen."
            )


# =====================================================
# PRODUKTDATENBANK
# =====================================================

elif page == "📦 Produkte":

    st.title("📦 Produktdatenbank")

    st.write(
        "Verwalte PV-Module, Wechselrichter, "
        "Batteriespeicher und weiteres Material."
    )

    with st.form("new_product"):

        name = st.text_input("Produktname")

        category = st.selectbox(
            "Kategorie",
            [
                "PV-Modul",
                "Wechselrichter",
                "Batteriespeicher",
                "Montagesystem",
                "Elektromaterial",
                "Sonstiges"
            ]
        )

        unit = st.selectbox(
            "Einheit",
            [
                "Stück",
                "Meter",
                "Set",
                "Pauschale"
            ]
        )

        purchase_price = st.number_input(
            "Einkaufspreis CHF",
            min_value=0.0,
            value=0.0
        )

        selling_price = st.number_input(
            "Verkaufspreis CHF",
            min_value=0.0,
            value=0.0
        )

        submitted = st.form_submit_button(
            "Produkt speichern"
        )

        if submitted:

            if not name.strip():

                st.error(
                    "Bitte einen Produktnamen eingeben."
                )

            else:

                save_product(
                    name,
                    category,
                    unit,
                    purchase_price,
                    selling_price
                )

                st.success(
                    "Produkt erfolgreich gespeichert!"
                )

    st.divider()

    st.subheader("Gespeicherte Produkte")

    products = load_products()

    if products.empty:

        st.info("Noch keine Produkte vorhanden.")

    else:

        st.dataframe(
            products,
            use_container_width=True,
            hide_index=True
        )


# =====================================================
# KI ASSISTENT
# =====================================================

elif page == "🤖 KI-Assistent":

    st.title("🤖 SolarAI Assistent")

    st.write(
        "Dein persönlicher KI-Assistent "
        "für Photovoltaikprojekte."
    )

    st.info(
        "Für den KI-Chat muss Ollama installiert "
        "sein und ein kompatibles KI-Modell laufen."
    )

    model = st.text_input(
        "Ollama-Modell",
        value="qwen2.5:3b"
    )

    if "messages" not in st.session_state:

        st.session_state.messages = []

    for message in st.session_state.messages:

        with st.chat_message(message["role"]):

            st.markdown(message["content"])

    prompt = st.chat_input(
        "Stelle SolarAI eine Frage..."
    )

    if prompt:

        st.session_state.messages.append({
            "role": "user",
            "content": prompt
        })

        with st.chat_message("user"):

            st.markdown(prompt)

        system_prompt = """
Du bist SolarAI, ein professioneller Assistent
für Photovoltaikplanung in der Schweiz.

Du unterstützt bei:
- Photovoltaikplanung
- Technischen Fragen
- Materiallisten
- Offertenvorbereitung
- Batteriespeichern
- Wechselrichtern
- Projektabwicklung

Wichtige Regeln:
- Erfinde keine technischen Daten.
- Erfinde keine Produktpreise.
- Kennzeichne Annahmen ausdrücklich.
- Weise auf fehlende Projektdaten hin.
- Gib keine ungeprüften sicherheitskritischen
  Auslegungen als freigegebene Planung aus.
- Behaupte nicht, aktuelle Schweizer Normen,
  Vorschriften oder Förderungen geprüft zu haben.
- Antworte verständlich und auf Deutsch.
"""

        messages = [
            {
                "role": "system",
                "content": system_prompt
            }
        ]

        messages.extend(
            st.session_state.messages
        )

        try:

            response = requests.post(
                "http://localhost:11434/api/chat",
                json={
                    "model": model,
                    "messages": messages,
                    "stream": False
                },
                timeout=180
            )

            response.raise_for_status()

            answer = response.json()["message"]["content"]

        except requests.RequestException:

            answer = (
                "Die lokale KI ist nicht erreichbar. "
                "Bitte prüfe, ob Ollama installiert ist, "
                "läuft und das ausgewählte Modell "
                "heruntergeladen wurde."
            )

        except (KeyError, ValueError):

            answer = (
                "Die KI hat keine gültige Antwort geliefert."
            )

        st.session_state.messages.append({
            "role": "assistant",
            "content": answer
        })

        with st.chat_message("assistant"):

            st.markdown(answer)


# =====================================================
# EINSTELLUNGEN
# =====================================================

elif page == "⚙️ Einstellungen":

    st.title("⚙️ Einstellungen")

    st.subheader("SolarAI")

    st.write("Version: 1.0")

    st.write("Betriebsmodus: Lokal")

    st.write("Datenbank: SQLite")

    st.write("KI-Schnittstelle: Ollama")

    st.divider()

    st.subheader("Hinweise")

    st.info(
        "Diese Version ist ein lokaler Prototyp. "
        "Sie besitzt noch keine Benutzerverwaltung, "
        "keine automatische Datensicherung und "
        "keine abgesicherte Online-Bereitstellung."
    )

    st.warning(
        "Bitte zunächst nur Testdaten verwenden. "
        "Für echte Firmen- und Kundendaten müssen "
        "Datenschutz, Zugriffsrechte und "
        "Datensicherung geklärt werden."
    )


# =====================================================
# FOOTER
# =====================================================

st.sidebar.divider()

st.sidebar.caption(
    "☀️ SolarAI | PV Intelligence Platform"
)