
import streamlit as st
import pandas as pd
import requests
import io

from datetime import datetime
from docx import Document
from supabase import create_client


# =====================================================
# SOLARAI 2.0
# PV-Planung, Cloud-Projekte, Produkte und Offerten
# =====================================================

st.set_page_config(
    page_title="SolarAI",
    page_icon="☀️",
    layout="wide",
    initial_sidebar_state="expanded",
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
# SUPABASE
# =====================================================

def create_supabase_client():
    """
    Jeder Streamlit-Benutzersitzung gehört ein eigener Client.
    Kein global geteilter, angemeldeter Client.
    """
    return create_client(
        st.secrets["SUPABASE_URL"],
        st.secrets["SUPABASE_KEY"],
    )


def get_client():
    if "supabase_client" not in st.session_state:
        st.session_state.supabase_client = (
            create_supabase_client()
        )

    return st.session_state.supabase_client


def clear_login():
    st.session_state.pop("supabase_client", None)
    st.session_state.pop("auth_session", None)
    st.session_state.pop("user_id", None)
    st.session_state.pop("user_email", None)
    st.session_state.pop("quote", None)
    st.session_state.pop("messages", None)


def login_user(email, password):
    client = get_client()

    result = client.auth.sign_in_with_password({
        "email": email,
        "password": password,
    })

    if not result.session or not result.user:
        raise ValueError("Anmeldung fehlgeschlagen.")

    st.session_state.auth_session = result.session
    st.session_state.user_id = result.user.id
    st.session_state.user_email = result.user.email


def restore_login():
    """
    Prüft die gespeicherte Sitzung und erneuert
    das Zugriffstoken, falls erforderlich.
    """
    session = st.session_state.get("auth_session")

    if session is None:
        return False

    try:
        client = get_client()

        result = client.auth.set_session(
            session.access_token,
            session.refresh_token,
        )

        if not result.session or not result.user:
            clear_login()
            return False

        st.session_state.auth_session = result.session
        st.session_state.user_id = result.user.id
        st.session_state.user_email = result.user.email

        return True

    except Exception:
        clear_login()
        return False


def logout_user():
    try:
        get_client().auth.sign_out()
    except Exception:
        pass

    clear_login()
    st.rerun()


# =====================================================
# LOGIN-BILDSCHIRM
# =====================================================

if not restore_login():

    st.title("☀️ SolarAI")
    st.subheader("Anmelden")
    st.write(
        "Melde dich mit deinem SolarAI-Konto an."
    )

    with st.form("login_form"):
        email = st.text_input("E-Mail-Adresse")
        password = st.text_input(
            "Passwort",
            type="password",
        )

        submitted = st.form_submit_button(
            "🔐 Anmelden"
        )

    if submitted:
        if not email.strip() or not password:
            st.error(
                "Bitte E-Mail-Adresse und Passwort eingeben."
            )
        else:
            try:
                login_user(
                    email.strip(),
                    password,
                )
                st.rerun()

            except Exception:
                st.error(
                    "Anmeldung fehlgeschlagen. "
                    "Bitte Zugangsdaten prüfen."
                )

    st.info(
        "Dein Konto wird in Supabase unter "
        "Authentication → Users verwaltet."
    )

    st.stop()


# =====================================================
# CLOUD-DATENBANK
# =====================================================

PROJECT_COLUMNS = [
    "id",
    "name",
    "customer",
    "address",
    "building_type",
    "roof_type",
    "module_count",
    "module_power",
    "inverter",
    "battery_kwh",
    "created_at",
]

PRODUCT_COLUMNS = [
    "id",
    "name",
    "category",
    "unit",
    "purchase_price",
    "selling_price",
]


def save_project(
    name,
    customer,
    address,
    building_type,
    roof_type,
    module_count,
    module_power,
    inverter,
    battery,
):
    user_id = st.session_state.user_id

    data = {
        "user_id": user_id,
        "name": name.strip(),
        "customer": customer,
        "address": address,
        "building_type": building_type,
        "roof_type": roof_type,
        "module_count": int(module_count),
        "module_power": float(module_power),
        "module_power_wp": float(module_power),
        "inverter": inverter,
        "battery_kwh": float(battery),
        "battery": str(battery),
    }

    get_client().table("projects").insert(
        data
    ).execute()


def load_projects():
    result = (
        get_client()
        .table("projects")
        .select("*")
        .eq("user_id", st.session_state.user_id)
        .order("created_at", desc=True)
        .execute()
    )

    df = pd.DataFrame(result.data or [])

    for column in PROJECT_COLUMNS:
        if column not in df.columns:
            df[column] = None

    for column in [
        "module_count",
        "module_power",
        "battery_kwh",
    ]:
        df[column] = pd.to_numeric(
            df[column],
            errors="coerce",
        ).fillna(0)

    return df


def delete_project(project_id):
    (
        get_client()
        .table("projects")
        .delete()
        .eq("id", str(project_id))
        .eq("user_id", st.session_state.user_id)
        .execute()
    )


def save_product(
    name,
    category,
    unit,
    purchase_price,
    selling_price,
):
    data = {
        "user_id": st.session_state.user_id,
        "name": name.strip(),
        "category": category,
        "unit": unit,
        "purchase_price": float(purchase_price),
        "selling_price": float(selling_price),
    }

    get_client().table("products").insert(
        data
    ).execute()


def load_products():
    result = (
        get_client()
        .table("products")
        .select("*")
        .eq("user_id", st.session_state.user_id)
        .order("created_at", desc=True)
        .execute()
    )

    df = pd.DataFrame(result.data or [])

    for column in PRODUCT_COLUMNS:
        if column not in df.columns:
            df[column] = None

    return df


def delete_product(product_id):
    (
        get_client()
        .table("products")
        .delete()
        .eq("id", str(product_id))
        .eq("user_id", st.session_state.user_id)
        .execute()
    )


# =====================================================
# BERECHNUNGEN
# =====================================================

def calculate_pv_power(module_count, module_power):
    return module_count * module_power / 1000


def calculate_annual_yield(
    power_kwp,
    specific_yield,
):
    return power_kwp * specific_yield


def calculate_quote(
    material_items,
    labor_items,
    margin,
    vat,
):
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
        "gross": gross_total,
    }


# =====================================================
# WORD-OFFERTE
# =====================================================

def create_word_quote(
    project_name,
    customer,
    address,
    material_items,
    labor_items,
    totals,
    vat_rate,
):
    doc = Document()

    doc.add_heading("SOLARAI", 0)
    doc.add_heading(
        "Offerte Photovoltaikanlage",
        level=1,
    )

    doc.add_paragraph(
        f"Datum: {datetime.now():%d.%m.%Y}"
    )
    doc.add_paragraph(
        f"Projekt: {project_name}"
    )
    doc.add_paragraph(
        f"Kunde: {customer}"
    )
    doc.add_paragraph(
        f"Adresse: {address}"
    )

    doc.add_heading("Material", level=2)

    table = doc.add_table(
        rows=1,
        cols=4,
    )
    table.style = "Table Grid"

    headers = [
        "Position",
        "Menge",
        "Einzelpreis CHF",
        "Total CHF",
    ]

    for i, header in enumerate(headers):
        table.rows[0].cells[i].text = header

    for item in material_items:
        row = table.add_row().cells

        row[0].text = str(item["Position"])
        row[1].text = str(item["Menge"])
        row[2].text = (
            f'{item["Einzelpreis"]:.2f}'
        )
        row[3].text = (
            f'{item["Menge"] * item["Einzelpreis"]:.2f}'
        )

    doc.add_heading(
        "Arbeitsleistungen",
        level=2,
    )

    table = doc.add_table(
        rows=1,
        cols=4,
    )
    table.style = "Table Grid"

    headers = [
        "Arbeit",
        "Stunden",
        "CHF/h",
        "Total CHF",
    ]

    for i, header in enumerate(headers):
        table.rows[0].cells[i].text = header

    for item in labor_items:
        row = table.add_row().cells

        row[0].text = str(item["Arbeit"])
        row[1].text = str(item["Stunden"])
        row[2].text = (
            f'{item["Stundensatz"]:.2f}'
        )
        row[3].text = (
            f'{item["Stunden"] * item["Stundensatz"]:.2f}'
        )

    doc.add_heading(
        "Zusammenfassung",
        level=2,
    )

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
        f'MWST ({vat_rate:.1f}%): '
        f'CHF {totals["vat"]:,.2f}'
    )

    doc.add_heading(
        f'Gesamttotal: CHF {totals["gross"]:,.2f}',
        level=2,
    )

    doc.add_paragraph(
        "Diese Offerte ist ein Entwurf und muss "
        "vor der Abgabe fachlich und kaufmännisch "
        "geprüft werden."
    )

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)

    return buffer


# =====================================================
# SEITENLEISTE
# =====================================================

st.sidebar.title("☀️ SolarAI")
st.sidebar.caption("PV Intelligence Platform")

st.sidebar.success("Angemeldet")
st.sidebar.caption(
    st.session_state.user_email or ""
)

if st.sidebar.button("🚪 Abmelden"):
    logout_user()

st.sidebar.divider()

page = st.sidebar.radio(
    "Navigation",
    [
        "🏠 Dashboard",
        "📂 Projekte",
        "📐 PV-Planung",
        "💰 Offerten",
        "📦 Produkte",
        "🤖 KI-Assistent",
        "⚙️ Einstellungen",
    ],
)

st.sidebar.divider()
st.sidebar.success("SolarAI 2.0")
st.sidebar.caption("Supabase Cloud")


# =====================================================
# DASHBOARD
# =====================================================

if page == "🏠 Dashboard":

    st.title("☀️ SolarAI Dashboard")
    st.write(
        "Willkommen bei deiner persönlichen "
        "PV-Planungsplattform."
    )

    try:
        projects = load_projects()
        products = load_products()

        total_projects = len(projects)

        total_power = 0.0

        if not projects.empty:
            total_power = (
                projects["module_count"]
                * projects["module_power"]
                / 1000
            ).sum()

        col1, col2, col3 = st.columns(3)

        col1.metric(
            "PV-Projekte",
            total_projects,
        )
        col2.metric(
            "Geplante Leistung",
            f"{total_power:.2f} kWp",
        )
        col3.metric(
            "Produkte",
            len(products),
        )

        st.divider()
        st.subheader("Letzte Projekte")

        if projects.empty:
            st.info(
                "Noch keine Projekte vorhanden."
            )
        else:
            st.dataframe(
                projects[
                    [
                        "name",
                        "customer",
                        "module_count",
                        "module_power",
                        "created_at",
                    ]
                ],
                use_container_width=True,
                hide_index=True,
            )

    except Exception as error:
        st.error(
            "Cloud-Daten konnten nicht geladen werden."
        )
        st.exception(error)


# =====================================================
# PROJEKTE
# =====================================================

elif page == "📂 Projekte":

    st.title("📂 Projektverwaltung")

    tab1, tab2 = st.tabs(
        [
            "Neues Projekt",
            "Projektübersicht",
        ]
    )

    with tab1:

        with st.form("new_project"):

            name = st.text_input("Projektname")
            customer = st.text_input("Kunde")
            address = st.text_input(
                "Projektadresse"
            )

            building_type = st.selectbox(
                "Gebäudetyp",
                [
                    "Einfamilienhaus",
                    "Mehrfamilienhaus",
                    "Gewerbe",
                    "Industrie",
                    "Andere",
                ],
            )

            roof_type = st.selectbox(
                "Dachtyp",
                [
                    "Ziegeldach",
                    "Flachdach",
                    "Blechdach",
                    "Fassade",
                    "Andere",
                ],
            )

            module_count = st.number_input(
                "Anzahl Module",
                min_value=0,
                value=30,
            )

            module_power = st.number_input(
                "Modulleistung Wp",
                min_value=0.0,
                value=450.0,
            )

            inverter = st.text_input(
                "Wechselrichter"
            )

            battery = st.number_input(
                "Batteriespeicher kWh",
                min_value=0.0,
                value=0.0,
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
                try:
                    save_project(
                        name,
                        customer,
                        address,
                        building_type,
                        roof_type,
                        module_count,
                        module_power,
                        inverter,
                        battery,
                    )
                    st.success(
                        "Projekt in Supabase gespeichert!"
                    )

                except Exception as error:
                    st.error(
                        "Projekt konnte nicht gespeichert "
                        "werden."
                    )
                    st.exception(error)

    with tab2:

        try:
            projects = load_projects()

            if projects.empty:
                st.info(
                    "Keine Projekte vorhanden."
                )

            else:
                st.dataframe(
                    projects[
                        [
                            "name",
                            "customer",
                            "address",
                            "building_type",
                            "roof_type",
                            "module_count",
                            "module_power",
                            "inverter",
                            "battery_kwh",
                            "created_at",
                        ]
                    ],
                    use_container_width=True,
                    hide_index=True,
                )

                options = {
                    f'{row["name"]} | {str(row["id"])[:8]}':
                    str(row["id"])
                    for _, row in projects.iterrows()
                }

                selected = st.selectbox(
                    "Projekt zum Löschen auswählen",
                    list(options.keys()),
                )

                confirm = st.checkbox(
                    "Ich möchte dieses Projekt "
                    "wirklich löschen."
                )

                if st.button(
                    "🗑️ Projekt löschen",
                    disabled=not confirm,
                ):
                    delete_project(
                        options[selected]
                    )
                    st.success(
                        "Projekt gelöscht."
                    )
                    st.rerun()

        except Exception as error:
            st.error(
                "Projektübersicht konnte nicht "
                "geladen werden."
            )
            st.exception(error)


# =====================================================
# PV-PLANUNG
# =====================================================

elif page == "📐 PV-Planung":

    st.title("📐 PV-Planung")
    st.write(
        "Berechne die wichtigsten Anlagendaten."
    )

    col1, col2 = st.columns(2)

    with col1:
        module_count = st.number_input(
            "Anzahl PV-Module",
            min_value=1,
            value=30,
        )

        module_power = st.number_input(
            "Modulleistung (Wp)",
            min_value=1.0,
            value=450.0,
        )

    with col2:
        specific_yield = st.number_input(
            "Spezifischer Jahresertrag (kWh/kWp)",
            min_value=0.0,
            value=1000.0,
        )

        battery = st.number_input(
            "Batteriespeicher (kWh)",
            min_value=0.0,
            value=0.0,
        )

    power = calculate_pv_power(
        module_count,
        module_power,
    )

    annual_yield = calculate_annual_yield(
        power,
        specific_yield,
    )

    st.divider()

    col1, col2, col3 = st.columns(3)

    col1.metric(
        "Anlagenleistung",
        f"{power:.2f} kWp",
    )

    col2.metric(
        "Geschätzter Jahresertrag",
        f"{annual_yield:,.0f} kWh",
    )

    col3.metric(
        "Batteriespeicher",
        f"{battery:.1f} kWh",
    )

    st.info(
        "Der Jahresertrag ist eine vereinfachte "
        "Schätzung. Standort, Ausrichtung, "
        "Verschattung und Verluste müssen "
        "separat berücksichtigt werden."
    )


# =====================================================
# OFFERTEN
# =====================================================

elif page == "💰 Offerten":

    st.title("💰 Offertengenerator")
    st.write(
        "Erstelle eine PV-Offerte mit Material- "
        "und Arbeitskosten."
    )

    try:
        projects = load_projects()
    except Exception as error:
        st.error(
            "Projekte konnten nicht geladen werden."
        )
        st.exception(error)
        st.stop()

    if projects.empty:
        st.warning(
            "Bitte zuerst ein Projekt erstellen."
        )

    else:
        project_options = {
            f'{row["name"]} | {str(row["id"])[:8]}':
            row
            for _, row in projects.iterrows()
        }

        selected = st.selectbox(
            "Projekt auswählen",
            list(project_options.keys()),
        )

        project = project_options[selected]
        project_id = str(project["id"])

        st.subheader("Materialpositionen")

        default_material = pd.DataFrame(
            [
                {
                    "Position": "PV-Modul",
                    "Menge": int(
                        project["module_count"]
                    ),
                    "Einzelpreis": 0.0,
                },
                {
                    "Position": "Wechselrichter",
                    "Menge": 1,
                    "Einzelpreis": 0.0,
                },
                {
                    "Position": "Montagesystem",
                    "Menge": 1,
                    "Einzelpreis": 0.0,
                },
            ]
        )

        material_df = st.data_editor(
            default_material,
            num_rows="dynamic",
            use_container_width=True,
            key=f"material_{project_id}",
        )

        st.subheader("Arbeitsleistungen")

        default_labor = pd.DataFrame(
            [
                {
                    "Arbeit": "PV-Montage",
                    "Stunden": 0.0,
                    "Stundensatz": 85.0,
                },
                {
                    "Arbeit": "Elektroinstallation",
                    "Stunden": 0.0,
                    "Stundensatz": 95.0,
                },
                {
                    "Arbeit": "Inbetriebnahme",
                    "Stunden": 0.0,
                    "Stundensatz": 95.0,
                },
            ]
        )

        labor_df = st.data_editor(
            default_labor,
            num_rows="dynamic",
            use_container_width=True,
            key=f"labor_{project_id}",
        )

        col1, col2 = st.columns(2)

        with col1:
            margin = st.number_input(
                "Zuschlag (%)",
                min_value=0.0,
                value=0.0,
            )

        with col2:
            vat = st.number_input(
                "MWST (%)",
                min_value=0.0,
                value=8.1,
            )

        if st.button("Offerte berechnen"):

            try:
                clean_material = (
                    material_df.dropna(
                        subset=[
                            "Position",
                            "Menge",
                            "Einzelpreis",
                        ]
                    ).copy()
                )

                clean_labor = (
                    labor_df.dropna(
                        subset=[
                            "Arbeit",
                            "Stunden",
                            "Stundensatz",
                        ]
                    ).copy()
                )

                for column in [
                    "Menge",
                    "Einzelpreis",
                ]:
                    clean_material[column] = (
                        pd.to_numeric(
                            clean_material[column],
                            errors="raise",
                        )
                    )

                for column in [
                    "Stunden",
                    "Stundensatz",
                ]:
                    clean_labor[column] = (
                        pd.to_numeric(
                            clean_labor[column],
                            errors="raise",
                        )
                    )

                negative_values = (
                    (clean_material["Menge"] < 0).any()
                    or (
                        clean_material["Einzelpreis"] < 0
                    ).any()
                    or (
                        clean_labor["Stunden"] < 0
                    ).any()
                    or (
                        clean_labor["Stundensatz"] < 0
                    ).any()
                )

                if negative_values:
                    st.error(
                        "Mengen, Stunden und Preise "
                        "dürfen nicht negativ sein."
                    )

                else:
                    material_items = (
                        clean_material.to_dict("records")
                    )

                    labor_items = (
                        clean_labor.to_dict("records")
                    )

                    totals = calculate_quote(
                        material_items,
                        labor_items,
                        margin,
                        vat,
                    )

                    st.session_state["quote"] = {
                        "project_id": project_id,
                        "material": material_items,
                        "labor": labor_items,
                        "totals": totals,
                        "vat": vat,
                    }

            except (ValueError, TypeError) as error:
                st.error(
                    "Bitte Eingaben überprüfen: "
                    f"{error}"
                )

        quote = st.session_state.get("quote")

        if (
            quote
            and quote["project_id"] == project_id
        ):
            totals = quote["totals"]

            st.divider()
            st.subheader(
                "Offertenzusammenfassung"
            )

            col1, col2, col3 = st.columns(3)

            col1.metric(
                "Material",
                f'CHF {totals["material"]:,.2f}',
            )

            col2.metric(
                "Arbeit",
                f'CHF {totals["labor"]:,.2f}',
            )

            col3.metric(
                "Gesamttotal inkl. MWST",
                f'CHF {totals["gross"]:,.2f}',
            )

            word_file = create_word_quote(
                project["name"],
                project["customer"] or "",
                project["address"] or "",
                quote["material"],
                quote["labor"],
                totals,
                quote["vat"],
            )

            st.download_button(
                "📄 Word-Offerte herunterladen",
                data=word_file,
                file_name=(
                    f"SolarAI_Offerte_"
                    f"{project_id[:8]}.docx"
                ),
                mime=(
                    "application/vnd.openxmlformats-"
                    "officedocument.wordprocessingml."
                    "document"
                ),
            )

            st.warning(
                "Vor dem Versand alle Preise, Mengen, "
                "Leistungen und Vertragsbedingungen "
                "prüfen."
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
                "Sonstiges",
            ],
        )

        unit = st.selectbox(
            "Einheit",
            [
                "Stück",
                "Meter",
                "Set",
                "Pauschale",
            ],
        )

        purchase_price = st.number_input(
            "Einkaufspreis CHF",
            min_value=0.0,
            value=0.0,
        )

        selling_price = st.number_input(
            "Verkaufspreis CHF",
            min_value=0.0,
            value=0.0,
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
            try:
                save_product(
                    name,
                    category,
                    unit,
                    purchase_price,
                    selling_price,
                )

                st.success(
                    "Produkt in Supabase gespeichert!"
                )

            except Exception as error:
                st.error(
                    "Produkt konnte nicht gespeichert "
                    "werden."
                )
                st.exception(error)

    st.divider()
    st.subheader("Gespeicherte Produkte")

    try:
        products = load_products()

        if products.empty:
            st.info(
                "Noch keine Produkte vorhanden."
            )

        else:
            st.dataframe(
                products[PRODUCT_COLUMNS],
                use_container_width=True,
                hide_index=True,
            )

            product_options = {
                f'{row["name"]} | {str(row["id"])[:8]}':
                str(row["id"])
                for _, row in products.iterrows()
            }

            selected_product = st.selectbox(
                "Produkt zum Löschen auswählen",
                list(product_options.keys()),
            )

            confirm_product = st.checkbox(
                "Ich möchte dieses Produkt "
                "wirklich löschen."
            )

            if st.button(
                "🗑️ Produkt löschen",
                disabled=not confirm_product,
            ):
                delete_product(
                    product_options[selected_product]
                )
                st.success(
                    "Produkt gelöscht."
                )
                st.rerun()

    except Exception as error:
        st.error(
            "Produkte konnten nicht geladen werden."
        )
        st.exception(error)


# =====================================================
# KI-ASSISTENT
# =====================================================

elif page == "🤖 KI-Assistent":

    st.title("🤖 SolarAI Assistent")

    st.write(
        "Dein persönlicher KI-Assistent "
        "für Photovoltaikprojekte."
    )

    st.info(
        "Der KI-Chat verwendet weiterhin "
        "Ollama auf deinem eigenen Computer. "
        "Auf Streamlit Community Cloud ist "
        "dieser lokale Dienst normalerweise "
        "nicht erreichbar."
    )

    model = st.text_input(
        "Ollama-Modell",
        value="qwen2.5:3b",
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
            "content": prompt,
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

        messages = [{
            "role": "system",
            "content": system_prompt,
        }]

        messages.extend(
            st.session_state.messages
        )

        try:
            response = requests.post(
                "http://localhost:11434/api/chat",
                json={
                    "model": model,
                    "messages": messages,
                    "stream": False,
                },
                timeout=180,
            )

            response.raise_for_status()

            answer = (
                response.json()["message"]["content"]
            )

        except requests.RequestException:
            answer = (
                "Die lokale KI ist nicht erreichbar. "
                "Bitte prüfen, ob Ollama läuft und "
                "das ausgewählte Modell installiert ist."
            )

        except (KeyError, ValueError):
            answer = (
                "Die KI hat keine gültige Antwort "
                "geliefert."
            )

        st.session_state.messages.append({
            "role": "assistant",
            "content": answer,
        })

        with st.chat_message("assistant"):
            st.markdown(answer)


# =====================================================
# EINSTELLUNGEN
# =====================================================

elif page == "⚙️ Einstellungen":

    st.title("⚙️ Einstellungen")

    st.subheader("SolarAI")

    st.write("Version: 2.0")
    st.write("Betriebsmodus: Cloud")
    st.write("Datenbank: Supabase PostgreSQL")
    st.write("Anmeldung: Supabase Auth")
    st.write("KI-Schnittstelle: Ollama (lokal)")

    st.divider()

    st.subheader("Benutzerkonto")

    st.write(
        "Angemeldet als:",
        st.session_state.user_email,
    )

    if st.button("🚪 Jetzt abmelden"):
        logout_user()

    st.divider()

    st.subheader("Hinweise")

    st.info(
        "Projekte und Produkte werden in "
        "Supabase gespeichert. Die Offerten "
        "werden momentan im Browser berechnet "
        "und als Word-Datei heruntergeladen."
    )

    st.warning(
        "Diese Version ist weiterhin ein "
        "Prototyp. Vor dem Einsatz mit echten "
        "Kundendaten müssen Datenschutz, "
        "Zugriffsrechte, Backups und die "
        "Online-Bereitstellung geprüft werden."
    )


# =====================================================
# FOOTER
# =====================================================

st.sidebar.divider()
st.sidebar.caption(
    "☀️ SolarAI | PV Intelligence Platform"
)
