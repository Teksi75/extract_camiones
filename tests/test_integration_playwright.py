from src.portal.scraper import only_digits, split_domicilio


def test_only_digits_extrae_numeros():
    assert only_digits("VPE 0012345") == "0012345"


def test_split_domicilio_multilinea():
    dom, loc, prov = split_domicilio("Ruta 7 km 35\nLuján de Cuyo\nMendoza")
    assert dom == "Ruta 7 km 35"
    assert loc == "Luján de Cuyo"
    assert prov == "Mendoza"
