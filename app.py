# -*- coding: utf-8 -*-
"""Interfață pentru completarea contractului de înstrăinare-dobândire auto."""

from __future__ import annotations

from datetime import date

import streamlit as st

import contract_logic as logic

st.set_page_config(
    page_title="Contract auto",
    page_icon=":material/description:",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
    .block-container {padding-top: 1.4rem; padding-bottom: 3rem; max-width: 1180px;}
    h1 {font-weight: 650; letter-spacing: -0.03em; margin-bottom: 0.15rem;}
    div[data-testid="stFileUploader"] {background: rgba(255,255,255,0.55); border-radius: 10px; padding: 0.15rem 0.4rem 0.3rem;}
    </style>
    """,
    unsafe_allow_html=True,
)

MANUAL_FIELDS = {"pret_cifre", "pret_litere", "data_contract", "loc_contract"}


def init_state() -> None:
    if "gemini_api_key" not in st.session_state:
        try:
            st.session_state["gemini_api_key"] = st.secrets.get("GEMINI_API_KEY", "")
        except Exception:
            st.session_state["gemini_api_key"] = ""
    if "data_contract" not in st.session_state:
        st.session_state["data_contract"] = date.today().strftime("%d.%m.%Y")
    for key in logic.FIELDS:
        st.session_state.setdefault(key, "")
    st.session_state.setdefault("contract_bytes", None)
    st.session_state.setdefault("contract_name", "Contract.docx")
    st.session_state.setdefault("contract_error", "")
    st.session_state.setdefault("price_error", "")
    st.session_state.setdefault("extract_note", "")


def text_field(key: str, label: str, placeholder: str = "") -> None:
    st.text_input(label, key=key, placeholder=placeholder)


def warn_id(key: str) -> None:
    value = str(st.session_state.get(key, "")).replace(" ", "")
    if not value or value.upper().startswith("RO"):
        return
    if value.isdigit() and len(value) != 13:
        st.caption("Dacă este CNP, trebuie să aibă 13 cifre. Pentru CUI, verifică numărul din actul firmei.")


def person_section(prefix: str, title: str) -> None:
    number_key = f"{prefix}_reprezentant_ci_nr" if prefix == "vanzator" else f"{prefix}_reprezentant_ci_numar"
    st.subheader(title)
    text_field(f"{prefix}_nume", "Nume / denumire firmă", "Popescu Andrei sau SC Exemplu SRL")
    warn_id(f"{prefix}_cnp")
    c1, c2, c3 = st.columns(3)
    with c1:
        text_field(f"{prefix}_ci_serie", "Serie act", "KX")
    with c2:
        text_field(f"{prefix}_ci_nr", "Număr act", "123456")
    with c3:
        text_field(f"{prefix}_cnp", "CNP / CUI", "1234567890123")
    c1, c2, c3 = st.columns(3)
    with c1:
        text_field(f"{prefix}_judet", "Județ", "Cluj")
    with c2:
        text_field(f"{prefix}_localitate", "Localitate", "Cluj-Napoca")
    with c3:
        text_field(f"{prefix}_sat_sector", "Sat / sector", "Sector 1")
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        text_field(f"{prefix}_strada", "Strada", "Florilor")
    with c2:
        text_field(f"{prefix}_nr", "Număr", "10A")
    with c3:
        text_field(f"{prefix}_bloc", "Bloc", "B2")
    with c4:
        text_field(f"{prefix}_scara", "Scară", "A")
    c1, c2 = st.columns(2)
    with c1:
        text_field(f"{prefix}_etaj", "Etaj", "3")
    with c2:
        text_field(f"{prefix}_apartament", "Apartament", "12")
    st.markdown("**Reprezentant** — se completează doar dacă vinde sau cumpără o firmă ori un împuternicit")
    text_field(f"{prefix}_reprezentant_nume", "Nume reprezentant")
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        text_field(f"{prefix}_reprezentant_ci_serie", "Serie CI")
    with c2:
        text_field(number_key, "Număr CI")
    with c3:
        text_field(f"{prefix}_reprezentant_cnp", "CNP reprezentant")
    with c4:
        text_field(
            f"{prefix}_reprezentant_calitate_administrator_imputernicit",
            "Calitate",
            "administrator sau împuternicit",
        )
    warn_id(f"{prefix}_reprezentant_cnp")


def collect_documents(slots: list[tuple[str, list]]) -> list[tuple[str, bytes, str]]:
    documents = []
    for label, files in slots:
        for index, uploaded in enumerate(files or [], start=1):
            try:
                data, mime = logic.prepare_upload(uploaded.name, uploaded.getvalue(), uploaded.type or "")
            except Exception as exc:
                raise ValueError(f"Nu pot citi „{uploaded.name}” ({label}): {exc}") from exc
            documents.append((f"{label} — fișier {index}: {uploaded.name}", data, mime))
    return documents


def on_fill_price() -> None:
    st.session_state["price_error"] = ""
    try:
        st.session_state["pret_litere"] = logic.pret_in_litere(str(st.session_state.get("pret_cifre", "")))
    except ValueError as exc:
        st.session_state["price_error"] = str(exc)


def on_generate() -> None:
    st.session_state["contract_error"] = ""
    if not str(st.session_state.get("pret_litere", "")).strip() and str(st.session_state.get("pret_cifre", "")).strip():
        try:
            st.session_state["pret_litere"] = logic.pret_in_litere(str(st.session_state["pret_cifre"]))
        except ValueError as exc:
            st.session_state["contract_error"] = str(exc)
            st.session_state["contract_bytes"] = None
            return
    values = {key: st.session_state.get(key, "") for key in logic.FIELDS}
    try:
        st.session_state["contract_bytes"] = logic.render_contract(values)
        st.session_state["contract_name"] = logic.contract_filename(values)
    except Exception as exc:
        st.session_state["contract_bytes"] = None
        st.session_state["contract_error"] = f"Nu am putut genera contractul: {exc}"

def main() -> None:
    init_state()

    # Protecție cu parolă (activă doar dacă este setată APP_PASSWORD în secrets, ex. online)
    try:
        app_pass = st.secrets.get("APP_PASSWORD", "")
    except Exception:
        app_pass = ""
    if app_pass and not st.session_state.get("authenticated"):
        st.title("🔒 Acces restricționat")
        entered = st.text_input("Introdu parola de acces:", type="password")
        if st.button("Intră în aplicație", type="primary"):
            if entered == app_pass:
                st.session_state["authenticated"] = True
                st.rerun()
            else:
                st.error("Parolă incorectă.")
        st.stop()

    try:
        logic.repair_template()
    except FileNotFoundError as exc:
        st.error(str(exc))
        st.stop()

    with st.sidebar:
        st.header("Gemini")
        st.text_input(
            "Cheie Google Gemini API",
            type="password",
            key="gemini_api_key",
            help="Cheia rămâne în sesiunea din browser și nu este salvată pe disc.",
        )
        st.caption(f"Model folosit la extragere: {logic.MODEL_NAME}")
        st.caption(
            "Pozele pot fi JPG, PNG sau WEBP. Actele de mai multe pagini pot fi PDF "
            "sau mai multe poze în aceeași zonă."
        )

    st.title("Contract de înstrăinare-dobândire")
    st.caption(
        "Completează automat contractul pentru un mijloc de transport. "
        "Extrage datele din acte, verifică-le, apoi descarcă documentul Word."
    )

    st.subheader("Acte")
    left, right = st.columns(2)
    with left:
        vanzator_ci = st.file_uploader(
            "Buletin Vânzător (CI vechi sau CEI nou)",
            type=["jpg", "jpeg", "png", "webp", "pdf"],
            accept_multiple_files=True,
            key="up_vanzator_ci",
        )
        vanzator_dom = st.file_uploader(
            "Dovadă domiciliu / Extras RoCEIReader Vânzător (opțional)",
            type=["jpg", "jpeg", "png", "webp", "pdf"],
            accept_multiple_files=True,
            key="up_vanzator_dom",
        )
        vanzator_firma = st.file_uploader(
            "Acte Firmă / Procură Vânzător (opțional)",
            type=["jpg", "jpeg", "png", "webp", "pdf"],
            accept_multiple_files=True,
            key="up_vanzator_firma",
        )
        civ = st.file_uploader(
            "Carte de Identitate a Vehiculului (CIV)",
            type=["jpg", "jpeg", "png", "webp", "pdf"],
            accept_multiple_files=True,
            key="up_civ",
        )
    with right:
        cumparator_ci = st.file_uploader(
            "Buletin Cumpărător (CI vechi sau CEI nou)",
            type=["jpg", "jpeg", "png", "webp", "pdf"],
            accept_multiple_files=True,
            key="up_cumparator_ci",
        )
        cumparator_dom = st.file_uploader(
            "Dovadă domiciliu / Extras RoCEIReader Cumpărător (opțional)",
            type=["jpg", "jpeg", "png", "webp", "pdf"],
            accept_multiple_files=True,
            key="up_cumparator_dom",
        )
        cumparator_firma = st.file_uploader(
            "Acte Firmă / Procură Cumpărător (opțional)",
            type=["jpg", "jpeg", "png", "webp", "pdf"],
            accept_multiple_files=True,
            key="up_cumparator_firma",
        )
        talon = st.file_uploader(
            "Certificat de Înmatriculare (Talon)",
            type=["jpg", "jpeg", "png", "webp", "pdf"],
            accept_multiple_files=True,
            key="up_talon",
        )

    slots = [
        ("Buletin vânzător", vanzator_ci),
        ("Dovadă domiciliu vânzător", vanzator_dom),
        ("Acte firmă sau procură vânzător", vanzator_firma),
        ("Buletin cumpărător", cumparator_ci),
        ("Dovadă domiciliu cumpărător", cumparator_dom),
        ("Acte firmă sau procură cumpărător", cumparator_firma),
        ("Carte de identitate a vehiculului", civ),
        ("Certificat de înmatriculare", talon),
    ]
    loaded = [label for label, files in slots if files]
    if loaded:
        st.caption("Pregătite pentru extragere: " + " · ".join(loaded))

    if st.button("Extrage Datele cu AI", type="primary", disabled=not loaded, use_container_width=True):
        try:
            documents = collect_documents(slots)
            with st.spinner("Se citesc actele și se extrag datele..."):
                extracted = logic.extract_with_gemini(st.session_state.get("gemini_api_key", ""), documents)
            for key, value in extracted.items():
                if key in MANUAL_FIELDS and not value:
                    continue
                st.session_state[key] = value
            st.session_state["contract_bytes"] = None
            filled = sum(1 for key, value in extracted.items() if value and key not in MANUAL_FIELDS)
            st.session_state["extract_note"] = f"Am completat {filled} câmpuri. Verifică fiecare secțiune înainte să generezi contractul."
        except Exception as exc:
            st.session_state["extract_note"] = ""
            st.error(str(exc))

    if st.session_state.get("extract_note"):
        st.success(st.session_state["extract_note"])

    st.divider()
    with st.container(border=True):
        person_section("vanzator", "Vânzător")
    st.write("")
    with st.container(border=True):
        person_section("cumparator", "Cumpărător")
    st.write("")
    with st.container(border=True):
        st.subheader("Mașină")
        c1, c2, c3 = st.columns(3)
        with c1:
            text_field("auto_marca", "Marcă", "DACIA")
        with c2:
            text_field("auto_tipul", "Tip", "LOGAN")
        with c3:
            text_field("auto_numar_inmatriculare", "Număr înmatriculare", "B 123 ABC")
        c1, c2, c3 = st.columns(3)
        with c1:
            text_field("auto_serie_sasiu", "Serie șasiu (VIN)")
        with c2:
            text_field("auto_serie_motor", "Serie motor")
        with c3:
            text_field("auto_capacitate_cilindrica", "Capacitate cilindrică (cmc)", "1598")
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            text_field("auto_nr_carte_identitate", "Serie CIV")
        with c2:
            text_field("auto_an_fabricatie", "An fabricație", "2018")
        with c3:
            text_field("auto_norma_euro", "Normă Euro", "EURO 6")
        with c4:
            text_field("auto_data_itp", "Data expirării ITP", "25.11.2026")
    st.write("")
    with st.container(border=True):
        st.subheader("Preț")
        c1, c2 = st.columns(2)
        with c1:
            text_field("pret_cifre", "Preț în cifre", "15000")
        with c2:
            text_field("pret_litere", "Preț în litere", "Cincisprezece mii de lei")
        st.button("Completează prețul în litere", on_click=on_fill_price)
        if st.session_state.get("price_error"):
            st.caption(st.session_state["price_error"])
    st.write("")
    with st.container(border=True):
        st.subheader("Data și locul contractului")
        c1, c2 = st.columns(2)
        with c1:
            text_field("data_contract", "Data contractului", "25.11.2026")
        with c2:
            text_field("loc_contract", "Locul contractului", "Cluj-Napoca")

    st.write("")
    st.button("Generează Contractul Word", type="primary", on_click=on_generate, use_container_width=True)
    if st.session_state.get("contract_error"):
        st.error(st.session_state["contract_error"])
    if st.session_state.get("contract_bytes"):
        st.download_button(
            "Descarcă Contractul Completat",
            data=st.session_state["contract_bytes"],
            file_name=st.session_state.get("contract_name") or "Contract.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary",
            use_container_width=True,
        )
        st.caption("Deschide fișierul și mai verifică o dată datele înainte de semnare. Câmpurile goale rămân goale.")


if __name__ == "__main__":
    main()
