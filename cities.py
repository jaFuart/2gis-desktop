"""Offline catalogue of public 2GIS map projects."""

import json
import re
from pathlib import Path
from urllib.parse import urlsplit

ASSETS = Path(__file__).resolve().parent / "assets"
with (ASSETS / "cities.json").open(encoding="utf-8") as source:
    CITIES = json.load(source)

LABELS = {code: name for name, code in reversed(list(CITIES.items()))}


def search_key(value):
    return " ".join(value.casefold().replace("ё", "е").replace("‑", "-").split())


ALIASES = {search_key(name): code for name, code in CITIES.items()}
ALIASES.update({"питер": "spb", "спб": "spb", "екб": "ekaterinburg", "нижний": "n_novgorod"})


def city_code(value):
    name = search_key(value.strip())
    if name in ALIASES:
        return ALIASES[name]
    if name.startswith(("https://2gis.ru/", "http://2gis.ru/", "https://www.2gis.ru/",
                        "http://www.2gis.ru/", "2gis.ru/")):
        link = urlsplit(name if "://" in name else f"https://{name}")
        name = link.path.strip("/").split("/", 1)[0]
    if not re.fullmatch(r"[a-z0-9][a-z0-9_-]*", name, flags=re.ASCII):
        raise ValueError("Выберите город из списка или вставьте ссылку на его карту в 2ГИС.")
    return name
