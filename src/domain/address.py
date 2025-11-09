# src/domain/address.py
import re

ARG_PROVINCES = [
    "Buenos Aires",
    "CABA",
    "Ciudad Autónoma de Buenos Aires",
    "Catamarca",
    "Chaco",
    "Chubut",
    "Córdoba",
    "Corrientes",
    "Entre Ríos",
    "Formosa",
    "Jujuy",
    "La Pampa",
    "La Rioja",
    "Mendoza",
    "Misiones",
    "Neuquén",
    "Río Negro",
    "Salta",
    "San Juan",
    "San Luis",
    "Santa Cruz",
    "Santa Fe",
    "Santiago del Estero",
    "Tierra del Fuego",
    "Tierra del Fuego, Antártida e Islas del Atlántico Sur",
    "Tucumán",
]

_prov_alt = sorted(ARG_PROVINCES, key=len, reverse=True)
_prov_pattern = r"(" + "|".join(re.escape(p) for p in _prov_alt) + r")\s*$"
_PROV_REGEX = re.compile(_prov_pattern, flags=re.IGNORECASE)

def _smart_strip(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip())

def _extract_province(s: str):
    m = _PROV_REGEX.search(s or "")
    if not m:
        return s, None
    prov = _smart_strip(m.group(1))
    rest = _smart_strip(s[:m.start()])
    return rest, prov

def _split_domicilio_localidad(rest: str):
    text = _smart_strip(rest)
    if "," in text:
        left, right = text.rsplit(",", 1)
        return _smart_strip(left), _smart_strip(right)

    tokens = text.split(" ")
    if len(tokens) <= 2:
        return "", text

    last_num_idx = None
    for i, tok in enumerate(tokens):
        if re.fullmatch(r"\d+[A-Za-z]?", tok):
            last_num_idx = i
    if last_num_idx is not None and last_num_idx < len(tokens) - 1:
        domicilio = " ".join(tokens[: last_num_idx + 1])
        localidad = " ".join(tokens[last_num_idx + 1 :])
        return _smart_strip(domicilio), _smart_strip(localidad)

    for k in (3, 2):
        if len(tokens) > k:
            domicilio = " ".join(tokens[:-k])
            localidad = " ".join(tokens[-k:])
            if re.search(r"\b(\d+|Av\.?|Avenida|Calle|Ruta|RN|RP)\b", domicilio, flags=re.IGNORECASE):
                return _smart_strip(domicilio), _smart_strip(localidad)

    domicilio = " ".join(tokens[:-1])
    localidad = tokens[-1]
    return _smart_strip(domicilio), _smart_strip(localidad)

def parse_domicilio_fiscal(line: str):
    """
    Recibe la línea completa de 'Domicilio (Fiscal)' y devuelve:
    (domicilio_solo, localidad, provincia)

    - NO modifica el campo original; que quede igual es decisión del caller.
    - Si no se detecta provincia, devuelve (original, "", "").
    """
    s = _smart_strip(line)
    rest, prov = _extract_province(s)
    if prov is None:
        return s, "", ""
    domicilio, localidad = _split_domicilio_localidad(rest)
    return domicilio, localidad, prov
