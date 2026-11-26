# -*- coding: utf-8 -*-
"""Extragere date (Gemini) și completare contract Word pentru înstrăinare auto."""

from __future__ import annotations

import html
import io
import json
import re
import zipfile
from decimal import Decimal, InvalidOperation
from pathlib import Path

from docxtpl import DocxTemplate
from jinja2 import Environment, Undefined
from PIL import Image, ImageOps

TEMPLATE_NAME = "sablon.docx"
MODEL_NAME = "gemini-3.8-flash"

FIELDS = [
    "vanzator_nume",
    "vanzator_judet",
    "vanzator_localitate",
    "vanzator_sat_sector",
    "vanzator_strada",
    "vanzator_nr",
    "vanzator_bloc",
    "vanzator_scara",
    "vanzator_etaj",
    "vanzator_apartament",
    "vanzator_ci_serie",
    "vanzator_ci_nr",
    "vanzator_cnp",
    "vanzator_reprezentant_nume",
    "vanzator_reprezentant_ci_serie",
    "vanzator_reprezentant_ci_nr",
    "vanzator_reprezentant_cnp",
    "vanzator_reprezentant_calitate_administrator_imputernicit",
    "cumparator_nume",
    "cumparator_judet",
    "cumparator_localitate",
    "cumparator_sat_sector",
    "cumparator_strada",
    "cumparator_nr",
    "cumparator_bloc",
    "cumparator_scara",
    "cumparator_etaj",
    "cumparator_apartament",
    "cumparator_ci_serie",
    "cumparator_ci_nr",
    "cumparator_cnp",
    "cumparator_reprezentant_nume",
    "cumparator_reprezentant_ci_serie",
    "cumparator_reprezentant_ci_numar",
    "cumparator_reprezentant_cnp",
    "cumparator_reprezentant_calitate_administrator_imputernicit",
    "auto_marca",
    "auto_tipul",
    "auto_serie_sasiu",
    "auto_serie_motor",
    "auto_capacitate_cilindrica",
    "auto_numar_inmatriculare",
    "auto_nr_carte_identitate",
    "auto_an_fabricatie",
    "auto_norma_euro",
    "auto_data_itp",
    "pret_cifre",
    "pret_litere",
    "data_contract",
    "loc_contract",
]

_W_T = re.compile(r"(<w:t(?:\s[^>]*)?>)(.*?)(</w:t>)", re.DOTALL)
_TAG = re.compile(r"\{\{.*?\}\}", re.DOTALL)

SYSTEM_INSTRUCTION = (
    "Ești un extractor de date pentru contracte de înstrăinare-dobândire a unui "
    "mijloc de transport din România. Citești buletine, cărți electronice de "
    "identitate, extrase de domiciliu, acte de firmă, procuri, CIV și talon. "
    "Răspunzi exclusiv cu un obiect JSON valid, cu exact cheile cerute. "
    "Toate valorile sunt string-uri. Nu inventa date care nu apar în documente. "
    "Câmpul gol este \"\"."
)


class SilentUndefined(Undefined):
    def __str__(self) -> str:
        return ""

    def __html__(self) -> str:
        return ""

    def __iter__(self):
        return iter(())

    def __bool__(self) -> bool:
        return False


def template_path(base_dir: Path | None = None) -> Path:
    root = base_dir or Path(__file__).resolve().parent
    return root / TEMPLATE_NAME


def _xml_escape(text: str) -> str:
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _merge_split_tags(texts: list[str]) -> list[str]:
    full = "".join(texts)
    if "{{" not in full:
        return texts
    owners: list[int] = []
    for index, piece in enumerate(texts):
        owners.extend([index] * len(piece))
    for match in _TAG.finditer(full):
        first = owners[match.start()]
        for pos in range(match.start(), match.end()):
            owners[pos] = first
    rebuilt = [""] * len(texts)
    for pos, char in enumerate(full):
        rebuilt[owners[pos]] += char
    return rebuilt


def _patch_text_nodes(texts: list[str]) -> list[str]:
    texts = _merge_split_tags(texts)
    patched: list[str] = []
    for index, text in enumerate(texts):
        text = text.replace("vanzator_sat/sector", "vanzator_sat_sector")
        text = text.replace("loc_contarct", "loc_contract")
        patched.append(text)
    for index, text in enumerate(patched):
        if text == "\u00a0\u00a0\u00a0\u00a0" and index > 0 and patched[index - 1] == ", et. ":
            window = "".join(patched[max(0, index - 20) : index])
            if "cumparator_scara" in window and "cumparator_etaj" not in window:
                patched[index] = "{{cumparator_etaj}}"
            elif "vanzator_scara" in window and "vanzator_etaj" not in window:
                patched[index] = "{{vanzator_etaj}}"
        if text == " euro" and "auto_norma_euro" not in text:
            patched[index] = " euro {{auto_norma_euro}}, ITP {{auto_data_itp}}"
    return patched


def repair_template(path: Path | None = None) -> bool:
    """Unește etichetele {{...}} tăiate de Word și corectează numele incompatibile cu Jinja."""
    path = path or template_path()
    if not path.exists():
        raise FileNotFoundError(f"Nu găsesc șablonul {path.name} în {path.parent}.")

    with zipfile.ZipFile(path, "r") as archive:
        xml = archive.read("word/document.xml").decode("utf-8")
        others = {name: archive.read(name) for name in archive.namelist() if name != "word/document.xml"}

    matches = list(_W_T.finditer(xml))
    texts = [html.unescape(match.group(2)) for match in matches]
    patched = _patch_text_nodes(texts)
    if patched == texts:
        return False

    pieces: list[str] = []
    cursor = 0
    for match, new_text in zip(matches, patched):
        pieces.append(xml[cursor : match.start(2)])
        pieces.append(_xml_escape(new_text))
        cursor = match.end(2)
    pieces.append(xml[cursor:])
    new_xml = "".join(pieces).encode("utf-8")

    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for name, payload in others.items():
            archive.writestr(name, payload)
        archive.writestr("word/document.xml", new_xml)
    path.write_bytes(buffer.getvalue())
    return True


def blank_context() -> dict[str, str]:
    return {key: "" for key in FIELDS}


def clean_value(key: str, value) -> str:
    if value is None or isinstance(value, bool):
        return ""
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    text = str(value).strip()
    if text.lower() in {"none", "null", "n/a", "na", "-", "—", "undefined"}:
        return ""
    text = re.sub(r"\s+", " ", text)
    if key.endswith("cnp") or key.endswith("_cui"):
        text = text.replace(" ", "")
    if key in {"auto_serie_sasiu", "auto_serie_motor"}:
        text = re.sub(r"[^0-9A-Za-z]", "", text).upper()
    elif key in {
        "auto_numar_inmatriculare",
        "auto_nr_carte_identitate",
        "vanzator_ci_serie",
        "cumparator_ci_serie",
        "vanzator_reprezentant_ci_serie",
        "cumparator_reprezentant_ci_serie",
    }:
        text = text.upper()
    if key == "auto_capacitate_cilindrica":
        digits = re.sub(r"[^\d]", "", text)
        if 2 <= len(digits) <= 5:
            text = digits
    return text


def context_from_model(payload: dict) -> dict[str, str]:
    if not isinstance(payload, dict):
        raise ValueError("Răspunsul modelului nu este un obiect JSON.")
    cleaned = blank_context()
    for key in FIELDS:
        if key in payload:
            cleaned[key] = clean_value(key, payload.get(key))
    return cleaned


def _below_100(number: int, feminine: bool) -> str:
    units_m = ["", "unu", "doi", "trei", "patru", "cinci", "șase", "șapte", "opt", "nouă"]
    units_f = ["", "una", "două", "trei", "patru", "cinci", "șase", "șapte", "opt", "nouă"]
    teens_m = [
        "zece",
        "unsprezece",
        "doisprezece",
        "treisprezece",
        "paisprezece",
        "cincisprezece",
        "șaisprezece",
        "șaptesprezece",
        "optsprezece",
        "nouăsprezece",
    ]
    teens_f = [
        "zece",
        "unsprezece",
        "douăsprezece",
        "treisprezece",
        "paisprezece",
        "cincisprezece",
        "șaisprezece",
        "șaptesprezece",
        "optsprezece",
        "nouăsprezece",
    ]
    tens = [
        "",
        "",
        "douăzeci",
        "treizeci",
        "patruzeci",
        "cincizeci",
        "șaizeci",
        "șaptezeci",
        "optzeci",
        "nouăzeci",
    ]
    units = units_f if feminine else units_m
    teens = teens_f if feminine else teens_m
    if number < 10:
        return units[number]
    if number < 20:
        return teens[number - 10]
    ten, unit = divmod(number, 10)
    if unit == 0:
        return tens[ten]
    return f"{tens[ten]} și {units[unit]}"


def _below_1000(number: int, feminine: bool) -> str:
    if number < 100:
        return _below_100(number, feminine)
    hundreds, rest = divmod(number, 100)
    if hundreds == 1:
        head = "o sută"
    elif hundreds == 2:
        head = "două sute"
    else:
        head = f"{_below_100(hundreds, True)} sute"
    if rest == 0:
        return head
    return f"{head} {_below_100(rest, feminine)}"


def uses_de(number: int) -> bool:
    if number < 20:
        return False
    last = number % 100
    return last == 0 or last >= 20


def number_to_words(number: int, feminine: bool = False) -> str:
    if number == 0:
        return "zero"
    if number < 0:
        raise ValueError("Numărul nu poate fi negativ.")
    if number >= 1_000_000_000:
        raise ValueError("Suma este prea mare pentru a fi scrisă în litere.")

    parts: list[str] = []
    millions, rest = divmod(number, 1_000_000)
    thousands, units = divmod(rest, 1000)

    if millions:
        if millions == 1:
            parts.append("un milion")
        else:
            link = " de" if uses_de(millions) else ""
            parts.append(f"{_below_1000(millions, True)}{link} milioane")
    if thousands:
        if thousands == 1:
            parts.append("o mie")
        else:
            link = " de" if uses_de(thousands) else ""
            parts.append(f"{_below_1000(thousands, True)}{link} mii")
    if units:
        parts.append(_below_1000(units, feminine))
    return " ".join(parts)


def parse_amount(raw: str) -> tuple[int, int]:
    text = raw.strip().lower().replace("\u00a0", "").replace(" ", "")
    text = re.sub(r"(lei|ron)$", "", text)
    if not text:
        raise ValueError("Introdu prețul în cifre.")

    if "," in text and "." in text:
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")
    elif "," in text:
        left, right = text.split(",", 1)
        if right.isdigit() and len(right) in {1, 2}:
            text = left.replace(".", "") + "." + right
        elif re.fullmatch(r"\d{1,3}(,\d{3})+", text):
            text = text.replace(",", "")
        else:
            raise ValueError("Prețul nu este un număr valid.")
    elif re.fullmatch(r"\d{1,3}(\.\d{3})+", text):
        text = text.replace(".", "")

    try:
        amount = Decimal(text)
    except InvalidOperation as exc:
        raise ValueError("Prețul nu este un număr valid.") from exc
    if amount < 0:
        raise ValueError("Prețul nu poate fi negativ.")
    amount = amount.quantize(Decimal("0.01"))
    lei = int(amount)
    bani = int((amount - Decimal(lei)) * 100)
    return lei, bani


def _money_part(amount: int, singular: str, plural: str) -> str:
    if amount == 1:
        return f"un {singular}"
    link = " de" if uses_de(amount) else ""
    return f"{number_to_words(amount, feminine=False)}{link} {plural}"


def pret_in_litere(raw: str) -> str:
    lei, bani = parse_amount(raw)
    if lei == 0 and bani == 0:
        words = "zero lei"
    else:
        words = _money_part(lei, "leu", "lei") if lei or not bani else "zero lei"
        if bani:
            words = f"{words} și {_money_part(bani, 'ban', 'bani')}"
    return words[0].upper() + words[1:]


def prepare_upload(filename: str, data: bytes, mime: str) -> tuple[bytes, str]:
    suffix = Path(filename or "").suffix.lower()
    mime = (mime or "").lower()
    if suffix == ".pdf" or mime == "application/pdf":
        return data, "application/pdf"
    if suffix not in {".jpg", ".jpeg", ".png", ".webp"} and not mime.startswith("image/"):
        raise ValueError(f"Format neacceptat: {filename}. Folosește JPG, PNG, WEBP sau PDF.")

    image = Image.open(io.BytesIO(data))
    image = ImageOps.exif_transpose(image)
    if image.mode in {"RGBA", "LA"} or (image.mode == "P" and "transparency" in image.info):
        image = image.convert("RGBA")
        background = Image.new("RGB", image.size, (255, 255, 255))
        background.paste(image, mask=image.getchannel("A"))
        image = background
    elif image.mode != "RGB":
        image = image.convert("RGB")

    max_side = 2200
    width, height = image.size
    if max(width, height) > max_side:
        scale = max_side / float(max(width, height))
        image = image.resize((max(1, int(width * scale)), max(1, int(height * scale))), Image.Resampling.LANCZOS)

    output = io.BytesIO()
    image.save(output, format="JPEG", quality=88, optimize=True)
    return output.getvalue(), "image/jpeg"


def _parse_json(raw: str) -> dict:
    text = (raw or "").strip()
    if text.startswith("```"):
        text = re.sub(r"^```(?:json)?\s*", "", text)
        text = re.sub(r"\s*```$", "", text)
    try:
        payload = json.loads(text)
    except json.JSONDecodeError:
        match = re.search(r"\{.*\}", text, re.DOTALL)
        if not match:
            raise ValueError("Modelul nu a returnat JSON valid.") from None
        payload = json.loads(match.group(0))
    if isinstance(payload, list):
        if len(payload) == 1 and isinstance(payload[0], dict):
            payload = payload[0]
        else:
            raise ValueError("Modelul a returnat o listă, nu un obiect JSON.")
    return payload


def build_prompt() -> str:
    schema = {key: "" for key in FIELDS}
    keys = ", ".join(FIELDS)
    return f"""
Extrage datele pentru un contract de înstrăinare-dobândire a unui mijloc de transport (România, model 2026).
Primești documente etichetate. Unele pot lipsi. Nu copia datele vânzătorului la cumpărător sau invers.

Reguli pentru persoană fizică (fără act de firmă și fără procură):
- Completează numele, adresa, seria și numărul actului de identitate și CNP-ul.
- Lasă goale toate câmpurile „reprezentant” și câmpul de calitate.
- Numele este numele și prenumele, în ordinea din act, cu diacritice dacă se văd.

Reguli pentru adresă:
- Dacă există dovadă de domiciliu / extras RoCEIReader, adresa completă se ia DIN ACEL DOCUMENT (are prioritate față de buletin).
- Dacă nu există dovadă separată, adresa se ia de pe buletinul clasic (CI vechi), dacă este tipărită acolo.
- Cartea electronică de identitate (CEI) nouă nu are adresa tipărită. Fără extras RoCEIReader, lasă câmpurile de adresă goale. Nu inventa adresa.
- Județul: doar numele județului, fără „Județul” (exemplu: Cluj, Ilfov, București).
- Localitatea: doar numele (exemplu: Cluj-Napoca, Voluntari, București), fără „municipiul/orașul/comuna”.
- Sat/sector: satul sau, în București, „Sector 1” … „Sector 6”. Gol dacă nu există.
- Strada: fără prefixul „Str.”. Dacă este bulevard, cale, alee sau șosea, păstrează abrevierea (Bd., Cal., Al., Șos.) plus numele.
- Număr, bloc, scară, etaj, apartament: doar valoarea, fără „nr.”, „bl.”, „sc.”, „et.”, „ap.”. Etajul poate fi „Parter”.

Reguli pentru actul de identitate:
- Seria și numărul se iau din buletin / CEI, nu din extrasul de adresă.
- Dacă există serie și număr separat, pune seria în câmpul de serie (de obicei 2 litere) și numărul în câmpul de număr.
- Dacă documentul are un singur identificator, lasă seria goală și pune identificatorul la număr.

Reguli pentru firmă sau procură:
- Dacă documentele arată că partea este o firmă: numele firmei în câmpul nume, CUI/CIF (cum apare, de exemplu RO12345678) în câmpul cnp, sediul firmei în câmpurile de adresă. Seria și numărul de buletin ale firmei rămân goale.
- Datele persoanei fizice (nume, serie, număr, CNP) merg în câmpurile de reprezentant.
- Calitatea reprezentantului: „administrator” dacă actele arată administrator, „împuternicit” dacă există procură / împuternicire. Nu completa alte formulări.
- Dacă este procură pentru o persoană fizică (fără firmă): proprietarul din procură este partea (nume, CNP, adresă), iar împuternicitul este reprezentantul, cu calitatea „împuternicit”. Nu inventa CNP-ul sau adresa dacă nu apar în acte.
- Dacă nu există firmă și nu există procură, reprezentantul rămâne gol.

Reguli pentru vehicul, din acte românești (CIV / talon) SAU acte străine (Germania - Zulassungsbescheinigung Teil I/II, Polonia - Dowód Rejestracyjny / Karta Pojazdu, Austria - Zulassungsschein / Typenschein, Cehia - Osvědčení o registraci, Olanda - Kentekenbewijs, Franța - Certificat d'immatriculation, Italia - Carta di Circolazione, Anglia/UK - V5C Logbook etc.):
- Folosește codurile europene harmonizate când citești acte străine (inclusiv UK V5C):
  * auto_marca: rubrica D.1 (Make / Marke / Marka / Merk / Marque / Marca) — exemplu: DACIA, VOLKSWAGEN, BMW.
  * auto_tipul: rubrica D.2 / D.3 (Type / Variant / Model / Handelsbezeichnung / Typ / Handelsbenaming).
  * auto_serie_sasiu: rubrica E (VIN / Vehicle Identification Number / Fahrzeug-Identifizierungsnummer / Numer identyfikacyjny).
  * auto_serie_motor: rubrica P.5 (Engine number / Motornummer) dacă apare; altfel "".
  * auto_capacitate_cilindrica: rubrica P.1 (Cylinder capacity / Hubraum / Pojemność / Cilinderinhoud) — doar cifre în cmc, fără „cmc”.
  * auto_numar_inmatriculare: rubrica A (Registration number / Kennzeichen / Numer rejestracyjny / Kenteken / Targa) sau numărul străin din act.
  * auto_nr_carte_identitate: seria CIV românesc (dacă există CIV România); pentru acte străine fără CIV românesc lasă "".
  * auto_an_fabricatie: anul fabricației (4 cifre) dacă este menționat expres (ex. Polonia „Rok produkcji”, UK „Year of manufacture”), altfel anul din rubrica B (Data primei înmatriculări / First registration / Erstzulassung).
  * auto_norma_euro: rubrica V.9 sau clasa de emisii (ex. EURO 5, EURO 6, Euro 6d) tradusă în format „EURO X”; dacă nu apare, lasă "".
  * auto_data_itp: data expirării ITP/TUV/APK/MOT dacă apare în acte, format ZZ.LL.AAAA; altfel "".
- Dacă o valoare lipsește de pe acte, las-o "".

Prețul, data și locul NU apar pe acte. Lasă pret_cifre, pret_litere, data_contract și loc_contract goale.

Cheile obligatorii, toate string:
{keys}

Exemplu de formă (valorile reale vin din documente):
{json.dumps(schema, ensure_ascii=False)}
""".strip()


def extract_with_gemini(api_key: str, documents: list[tuple[str, bytes, str]]) -> dict[str, str]:
    if not api_key or not api_key.strip():
        raise ValueError("Lipește cheia Google Gemini API în bara din stânga.")
    if not documents:
        raise ValueError("Încarcă cel puțin un document înainte de extragere.")

    import time
    from google import genai
    from google.genai import types

    parts = [types.Part.from_text(text=build_prompt())]
    for label, data, mime in documents:
        parts.append(types.Part.from_text(text=f"DOCUMENT: {label}"))
        parts.append(types.Part.from_bytes(data=data, mime_type=mime))

    client = genai.Client(api_key=api_key.strip())
    config_base = dict(
        temperature=0,
        response_mime_type="application/json",
        max_output_tokens=8192,
        system_instruction=SYSTEM_INSTRUCTION,
    )
    configs = [types.GenerateContentConfig(**config_base)]

    candidate_models = [
        "gemini-3.6-flash",
        "gemini-3.5-flash",
        "gemini-3.1-flash-lite",
        "gemini-3.7-flash",
        "gemini-3.8-flash",
    ]

    response = None
    last_error = None
    for model_id in candidate_models:
        for attempt in range(1):
            for config in configs:
                try:
                    response = client.models.generate_content(
                        model=model_id,
                        contents=parts,
                        config=config,
                    )
                    break
                except Exception as exc:
                    last_error = exc
                    err_str = str(exc)
                    if "thinking" in err_str.lower() or "400" in err_str or "INVALID_ARGUMENT" in err_str:
                        continue
                    if any(code in err_str for code in ("503", "UNAVAILABLE", "429", "404", "NOT_FOUND")):
                        break
                    raise
            if response is not None:
                break
        if response is not None:
            break

    if response is None:
        raise RuntimeError(f"Serverele Google sunt temporar aglomerate. Mai încearcă o dată. Detaliu: {last_error}")

    raw = getattr(response, "text", None) or ""
    if not raw.strip():
        feedback = getattr(response, "prompt_feedback", None)
        raise RuntimeError(
            "Gemini nu a returnat date. "
            f"Verifică cheia API și imaginile încărcate. Detaliu: {feedback or 'răspuns gol'}"
        )
    return context_from_model(_parse_json(raw))



def render_contract(values: dict, template: Path | None = None) -> bytes:
    path = template or template_path()
    repair_template(path)
    context = blank_context()
    for key, value in values.items():
        if key in context:
            cleaned = clean_value(key, value)
            context[key] = f"  {cleaned}  " if cleaned else ""

    puncte_reprezentant = {
        "vanzator_reprezentant_nume": "." * 35,
        "vanzator_reprezentant_act": "." * 28,
        "vanzator_reprezentant_ci_serie": "." * 6,
        "vanzator_reprezentant_ci_nr": "." * 14,
        "vanzator_reprezentant_cnp": "." * 22,
        "vanzator_reprezentant_calitate_administrator_imputernicit": "." * 25,
        "cumparator_reprezentant_nume": "." * 35,
        "cumparator_reprezentant_act": "." * 28,
        "cumparator_reprezentant_ci_serie": "." * 6,
        "cumparator_reprezentant_ci_nr": "." * 14,
        "cumparator_reprezentant_ci_numar": "." * 14,
        "cumparator_reprezentant_cnp": "." * 22,
        "cumparator_reprezentant_calitate_administrator_imputernicit": "." * 25,
    }
    for k, dots in puncte_reprezentant.items():
        if k in context and not context[k]:
            context[k] = dots

    document = DocxTemplate(str(path))
    environment = Environment(undefined=SilentUndefined, autoescape=True)
    document.render(context, environment, autoescape=True)
    buffer = io.BytesIO()
    document.save(buffer)
    payload = buffer.getvalue()

    with zipfile.ZipFile(io.BytesIO(payload)) as archive:
        xml = archive.read("word/document.xml").decode("utf-8")
    if ">None<" in xml or "None" in re.findall(r"\{\{[^}]+\}\}", xml):
        raise RuntimeError("Șablonul a lăsat valori necompletate.")
    leftovers = re.findall(r"\{\{[^{}]+\}\}", xml)
    if leftovers:
        raise RuntimeError("Au rămas etichete neînlocuite: " + ", ".join(leftovers))
    return payload


def contract_filename(values: dict) -> str:
    plate = clean_value("auto_numar_inmatriculare", values.get("auto_numar_inmatriculare", ""))
    plate = re.sub(r"[^A-Z0-9]+", "-", plate).strip("-") or "AUTO"
    day = clean_value("data_contract", values.get("data_contract", ""))
    day = re.sub(r"[^0-9]+", ".", day).strip(".") or "fara-data"
    return f"Contract_{plate}_{day}.docx"
