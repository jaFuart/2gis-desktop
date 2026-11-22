"""Collect public 2GIS organization cards into a text-only Excel workbook or CSV."""

import argparse
import csv
import json
import math
import os
import re
import sys
import time
from collections import Counter
from contextlib import closing
from itertools import chain, islice
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.parse import parse_qsl, quote, unquote, urlencode, urlsplit
from urllib.request import ProxyHandler, build_opener
from xml.sax.saxutils import escape
from zipfile import ZIP_DEFLATED, ZipFile


BASE = "https://2gis.ru"
API = "https://catalog.api.2gis.com/3.0/items"
DAYS = dict(zip(
    ("Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"),
    ("понедельник", "вторник", "среда", "четверг", "пятница", "суббота", "воскресенье"),
))
HOURS = {day: f"Часы работы {name}" for day, name in DAYS.items()}
GROUND_NAME = "Название ближайшей наземной остановки"
GROUND_DISTANCE = "Расстояние до ближайшей наземной остановки"
GROUPS = {
    "Способы оплаты": ("способы оплаты", "способы быстрой оплаты", "оплата"),
    "Опции ресторана/кафе": ("ресторан / кафе", "ресторан/кафе", "ресторан и кафе"),
    "Опции меню": ("меню", "кухня", "ассортимент", "общественное питание", "кондитерские изделия"),
    "Опции напитков": ("напитки",),
    "Сервисные опции": ("сервис", "услуги", "кейтеринг", "организация и проведение праздников", "банкетные залы"),
    "Опции справочника": ("в справочнике",),
    "Дополнительные опции": ("дополнительно",),
    "Доступная среда": ("доступная среда", "инфраструктура для детских колясок"),
    "Опции доставки": ("доставка", "доставка готовых блюд"),
}
COLUMNS = (
    "Название места", "Название организации", "Краткое описание", *HOURS.values(),
    "Рейтинг", "Адрес", "Сайт", "Количество парковочных мест",
    "Количество ближайших станций метро", "Название ближайшей станции метро",
    "Расстояние до ближайшей станции метро в метрах", GROUND_NAME,
    GROUND_DISTANCE, "Электронные адреса", "Категории",
    "Номера телефонов", "Средний чек", *GROUPS,
)
FIELDS = "name_ex,org,summary,reviews,full_address_name,rubrics,schedule,attribute_groups,links"
FIRM = re.compile(r"^/([a-z0-9_-]+)/firm/(\d+)(?:/|$)")
PHONE = re.compile(r"\+?\d[\d\s().‐‑–-]{8,}\d")
INTERVAL = re.compile(r"\d{1,2}:\d{2}\s*[-–—−]\s*\d{1,2}:\d{2}")
CHECK = re.compile(r"(?:средний\s+)?чек\s*(?:от\s+)?([\d ]+)\s*(?:₽|руб)", re.I)
PAGE_STATE = """
    const state = typeof initialState === 'object' ? initialState?.data : null;
    const profiles = state?.entity?.profile ?? {};
"""


class ExportCancelled(Exception):
    pass


def progress_event(args, event, **data):
    if getattr(args, "on_progress", None):
        args.on_progress(event, data)


def cancel_checkpoint(args):
    if getattr(args, "is_cancelled", None) and args.is_cancelled():
        raise ExportCancelled


def clean(value):
    return " ".join(str(value if value is not None else "").replace("\u200b", "").split())


def joined(values):
    return ", ".join(dict.fromkeys(text for value in values if (text := clean(value))))


def phones(values):
    numbers = {}
    for value in values:
        text = clean(value)
        if "..." in text or "…" in text:
            continue
        for number in PHONE.findall(text):
            digits = re.sub(r"\D", "", number)
            if 10 <= len(digits) <= 15:
                numbers.setdefault(digits, number)
    return joined(numbers.values())


def website_contact(value):
    value = clean(value)
    link = urlsplit(value)
    if link.hostname == "link.2gis.ru":
        value = unquote(link.query)
        link = urlsplit(value)
    if link.scheme in ("http", "https") and link.hostname:
        query = [(key, val) for key, val in parse_qsl(link.query, keep_blank_values=True)
                 if not key.casefold().startswith("utm_")]
        return link._replace(query=urlencode(query)).geturl()
    if re.fullmatch(r"(?:[\w-]+\.)+[\w-]+(?:/[^ ]*)?", value):
        return value
    return ""


def set_group(row, label, values):
    values = [clean(value) for value in values if clean(value)]
    for value in values:
        if (match := CHECK.search(value)) and not row.get("Средний чек"):
            row["Средний чек"] = match[1].strip().replace(" ", "")
    for column, aliases in GROUPS.items():
        if clean(label).casefold() in aliases and values:
            row[column] = joined(filter(None, (row.get(column), joined(values))))


def record_row(item, *, include_phone=True):
    org = item.get("org") or {}
    name = item.get("name_ex") or {}
    reviews = item.get("reviews") or {}
    contacts = {}
    for group in item.get("contact_groups") or []:
        for contact in group.get("contacts") or []:
            kind = contact.get("type")
            value = contact.get("value") or contact.get("text") or contact.get("url")
            if kind in {"website", "site"}:
                value = website_contact(contact.get("url")) or website_contact(value)
            if value:
                contacts.setdefault(kind, []).append(value)
    address_data = item.get("address") or {}
    street_parts = (component for component in address_data.get("components") or []
                    if component.get("type") == "street_number")
    street_address = joined(clean(f"{part.get('street') or ''}, {part.get('number') or ''}".strip(", "))
                            for part in street_parts)
    short_address = (item.get("address_name") or street_address or address_data.get("landmark_name")
                     or address_data.get("building_name"))
    city = next((division.get("name") for division in item.get("adm_div") or []
                 if division.get("type") in {"city", "settlement"} and division.get("name")), None)
    address = joined((short_address, city, address_data.get("postcode"))) if short_address else ""
    rubrics = joined(r.get("name") for r in item.get("rubrics") or [])
    row = {
        "Название места": name.get("primary") or item.get("name"),
        "Название организации": name.get("legal_name") or org.get("name"),
        "Краткое описание": (item.get("summary") or {}).get("text") or item.get("description"),
        "Рейтинг": reviews.get("general_rating") or reviews.get("rating"),
        "Адрес": item.get("full_address_name") or address,
        "Категории": rubrics,
        "Опции справочника": rubrics,
        "Сайт": joined(contacts.get("website", []) + contacts.get("site", [])),
        "Электронные адреса": joined(contacts.get("email", [])),
        "Номера телефонов": phones(contacts.get("phone", [])) if include_phone else "",
    }
    if item.get("point"):
        row["_point"] = item["point"]
    schedule = item.get("schedule") or {}
    if any(day in schedule for day in DAYS) or schedule.get("is_24x7"):
        for day, column in HOURS.items():
            intervals = (schedule.get(day) or {}).get("working_hours") or []
            times = joined(f"{s['from']}–{s['to']}" for s in intervals if s.get("from") and s.get("to"))
            row[column] = "Круглосуточно" if schedule.get("is_24x7") else times
            if day not in schedule and not schedule.get("is_24x7"):
                row[column] = "Выходной"
    for group in item.get("attribute_groups") or []:
        set_group(row, group.get("name"), (a.get("name") or a.get("value") for a in group.get("attributes") or []))
    links = item.get("links") or {}
    capacities = [p["capacity"] for p in links.get("parking") or [] if clean(p.get("capacity")).isdigit()]
    if capacities:
        row["Количество парковочных мест"] = sum(map(int, capacities))
    stations = links.get("nearest_stations") or []
    metro = links.get("nearest_metro")
    if metro is None and stations:
        metro = [station for station in stations if "metro" in (station.get("route_types") or [])]
    if metro is not None:
        row["Количество ближайших станций метро"] = len({s["id"] for s in metro if s.get("id")})
        nearest = min((s for s in metro if s.get("distance") is not None), key=lambda s: s["distance"], default=None)
        if nearest:
            row["Расстояние до ближайшей станции метро в метрах"] = nearest["distance"]
            row["Название ближайшей станции метро"] = next(
                (s.get("name") for s in stations if s.get("id") == nearest.get("id")), nearest.get("name"),
            )
    ground = min(
        (s for s in (*stations, *(links.get("nearest_platforms") or []))
         if set(s.get("route_types") or [])
         & {"bus", "trolleybus", "tram", "shuttle_bus"} and s.get("distance") is not None),
        key=lambda s: s["distance"], default=None,
    )
    if ground:
        row[GROUND_NAME] = ground.get("name")
        row[GROUND_DISTANCE] = ground["distance"]
    return row


def fill_missing(target, source):
    for column, value in source.items():
        if not clean(target.get(column)) and clean(value):
            target[column] = value


def api_items(args, key):
    opener = build_opener(ProxyHandler({}))
    fields = FIELDS + (",contact_groups" if args.api_contacts else "")
    if args.ground_stops:
        fields += ",point"
    size = min(args.limit or 50, 50)
    page = 1
    while args.pages is None or page <= args.pages:
        cancel_checkpoint(args)
        params = {
            "q": args.query, "type": "branch", "region_id": args.region_id,
            "locale": "ru_RU", "page": page, "page_size": size, "key": key,
            "fields": ",".join(f"items.{field}" for field in fields.split(",")),
        }
        try:
            with opener.open(f"{API}?{urlencode(params)}", timeout=20) as response:
                data = json.load(response)
        except HTTPError as exc:
            raise RuntimeError(f"Places API вернул HTTP {exc.code}. Проверьте ключ, права и лимиты.") from None
        except (URLError, ValueError, TimeoutError):
            raise RuntimeError("Не удалось получить JSON от Places API.") from None
        code = data.get("meta", {}).get("code")
        if code not in (200, 404):
            raise RuntimeError(f"Places API вернул код {code}. Проверьте ключ, права и параметры.")
        items = data.get("result", {}).get("items") or []
        progress_event(args, "page", page=page, total=data.get("result", {}).get("total"), found=len(items))
        yield from items
        if len(items) < size:
            return
        total = data.get("result", {}).get("total")
        if isinstance(total, int) and page * size >= total:
            return
        page += 1
        time.sleep(1)


def page_snapshot(driver):
    return driver.execute_script("""
        return {
          text: document.body.innerText,
          name: document.querySelector('h1')?.innerText,
          description: document.querySelector('[itemprop="description"]')?.textContent,
          socialDescription: document.querySelector('meta[property="og:description"]')?.content,
          metas: Array.from(document.querySelectorAll('[itemprop][content]'), n =>
            [n.getAttribute('itemprop'), n.getAttribute('content')]),
          links: Array.from(document.querySelectorAll('a[href]'), a => [a.href, a.innerText]),
          structured: Array.from(document.querySelectorAll('script[type="application/ld+json"]'),
            n => n.textContent),
        };
    """)


def contact_rack(driver, url):
    path = urlsplit(url).path.rstrip("/")
    return driver.execute_script(r"""
        const path = arguments[0];
        if (location.pathname.replace(/\/$/, '') !== path) return null;
        const isContactTab = link => link.origin === location.origin
            && link.pathname.replace(/\/$/, '') === path
            && link.textContent.trim() === 'Контакты';
        const heading = document.querySelector('h1');
        for (let panel = heading?.parentElement; panel; panel = panel.parentElement) {
            if (!Array.from(panel.querySelectorAll('a[href]')).some(isContactTab)) continue;
            const racks = Array.from(panel.querySelectorAll('[data-rack="true"]'));
            const contactRacks = racks.filter(rack => rack.querySelector(
                'a[href^="tel:"], a[href^="mailto:"], a[href*="/geo/"]'));
            return contactRacks.length === 1 ? contactRacks[0] : null;
        }
        return null;
    """, path)


def scoped_contacts(driver, url, by):
    rack = contact_rack(driver, url)
    if rack is None:
        return []
    from selenium.common.exceptions import WebDriverException

    buttons = rack.find_elements(by.XPATH, ".//*[self::button or @role='button'][contains(., 'Показать телефон')]")
    for button in buttons:
        if button.is_displayed():
            try:
                button.click()
            except WebDriverException:
                pass
    return driver.execute_script("""
        return Array.from(arguments[0].querySelectorAll('a[href]'), a => [a.href, a.innerText]);
    """, contact_rack(driver, url) or rack)


def contact_fields(links):
    websites = []
    telephone = []
    emails = []
    for href, label in links:
        if href.startswith("tel:") and re.match(r"^\+?\d", clean(label)):
            telephone.append(unquote(href[4:]))
        elif href.startswith("mailto:"):
            emails.append(unquote(href[7:].split("?")[0]))
        elif urlsplit(href).hostname == "link.2gis.ru" and re.fullmatch(
            r"(?:[\w-]+\.)+[\w-]+(?:/[^ ]*)?", clean(label)
        ):
            websites.append(website_contact(href) or clean(label))
    return {"Номера телефонов": phones(telephone), "Электронные адреса": joined(emails), "Сайт": joined(websites)}


def structured_card(snapshot):
    def nodes(value):
        if isinstance(value, list):
            for item in value:
                yield from nodes(item)
        elif isinstance(value, dict):
            yield value
            yield from nodes(value.get("@graph"))
            yield from nodes(value.get("mainEntity"))

    name = clean(snapshot["name"]).casefold()
    for script in snapshot["structured"]:
        try:
            document = json.loads(script)
        except ValueError:
            continue
        for node in nodes(document):
            if clean(node.get("name")).casefold() == name and any(
                k in node for k in ("address", "telephone", "aggregateRating", "openingHoursSpecification")
            ):
                return node
    return {}


def lines(text):
    return [value for part in re.split(r"[\r\n\u200b]+", text) if (value := clean(part))]


def visible_card(main, info):
    data = structured_card(main)
    meta = dict(main["metas"])
    contacts = lines(main["text"])
    details = lines(info["text"])
    address = data.get("address") or {}
    if isinstance(address, dict):
        address = joined(address.get(k) for k in ("streetAddress", "addressLocality", "postalCode"))
    if not address:
        address = joined(text for href, text in main["links"]
                         if urlsplit(href).hostname == "2gis.ru" and "/geo/" in urlsplit(href).path
                         and re.search(r"\d", text) and len(clean(text)) < 140)
    rating = (data.get("aggregateRating") or {}).get("ratingValue") or meta.get("ratingValue")
    if not rating and (match := re.search(r"\bоценка\s+([1-5](?:[.,]\d{1,2})?)\b",
                                         main.get("socialDescription") or "", re.I)):
        rating = match[1]
    if not rating:
        rating = next((value for value, after in zip(contacts, contacts[1:])
                       if re.fullmatch(r"[1-5](?:[.,]\d{1,2})?", value)
                       and re.match(r"[\d\s]+\s+(?:оцен|отзыв)", after, re.I)), None)
    row = {
        "Название места": main["name"],
        "Название организации": data.get("legalName") or meta.get("legalName"),
        "Краткое описание": data.get("description") or main["description"],
        "Адрес": address, "Рейтинг": rating,
        "Категории": joined(text for href, text in info["links"]
                            if urlsplit(href).hostname == "2gis.ru"
                            and re.search(r"/rubricId/\d+", urlsplit(href).path, re.I)),
    }
    day_map = {day.casefold(): column for day, column in zip(DAYS.values(), HOURS.values())}
    for index, value in enumerate(contacts):
        day_match = re.match(r"^(" + "|".join(day_map) + r")\b\s*[:—–-]?\s*(.*)$", value, re.I)
        if day_match:
            next_lines = [day_match[2], *contacts[index + 1:index + 3]]
            for candidate in next_lines:
                if times := INTERVAL.findall(candidate):
                    row[day_map[day_match[1].casefold()]] = joined(times)
                    break
                if candidate.casefold() in {"выходной", "круглосуточно", "закрыто"}:
                    row[day_map[day_match[1].casefold()]] = candidate
                    break
                if re.match(r"^(?:" + "|".join(day_map) + r")\b", candidate, re.I):
                    break
    spec = data.get("openingHoursSpecification") or []
    for entry in spec if isinstance(spec, list) else [spec]:
        days = entry.get("dayOfWeek") or []
        if isinstance(days, str):
            days = [days]
        for day in days:
            column = HOURS.get(day.rstrip("/").rsplit("/", 1)[-1][:3].title())
            if column and entry.get("opens") and entry.get("closes"):
                row[column] = joined((row.get(column), f"{entry['opens']}–{entry['closes']}"))
    sections = {alias for aliases in GROUPS.values() for alias in aliases}
    sections.update(("парковка", "транспорт", "актуальность данных", "также люди ищут", "время работы"))
    for index, value in enumerate(details):
        section = value.casefold()
        if section not in sections:
            continue
        body = []
        for line in details[index + 1:]:
            if line.casefold() in sections or line.startswith("Нашли ошибку"):
                break
            body.append(line)
        set_group(row, value, body)
        if section == "парковка":
            if match := re.search(r"\b([\d ]+)\s+мест\b", " ".join(body), re.I):
                row["Количество парковочных мест"] = match[1].strip().replace(" ", "")
    return row


def check_page(driver):
    text = driver.find_element("tag name", "body").text[:2000].casefold()
    if "captcha" in driver.title.casefold() or "что вы не робот" in text:
        raise RuntimeError("2ГИС показал CAPTCHA на текущем подключении.")
    signals = ("403 forbidden", "access denied", "429 too many requests",
               "err_name_not_resolved", "err_connection", "доступ ограничен")
    if any(signal in text for signal in signals):
        raise RuntimeError("В браузере 2ГИС недоступен или запрос ограничен.")


def browser_card(driver, open_page, url, by, wait):
    open_page(url)
    check_page(driver)
    wait.until(lambda d: any(heading.text.strip() for heading in d.find_elements(by.CSS_SELECTOR, "h1")))
    identifier = FIRM.match(urlsplit(url).path)[2]
    record = driver.execute_script(PAGE_STATE + "return profiles[arguments[0]]?.data ?? null;", identifier)
    row = record_row(record) if record else {}
    contacts = contact_fields(scoped_contacts(driver, url, by))
    if contacts["Номера телефонов"]:
        row["Номера телефонов"] = contacts["Номера телефонов"]
    fill_missing(row, contacts)
    if not any(clean(row.get(column)) for column in HOURS.values()):
        selector = (
            "//*[self::button or @role='button'][contains(., 'Откроется') or "
            "contains(., 'Закроется') or contains(., 'График работы') or contains(., 'Время работы')]"
        )
        buttons = [element for element in driver.find_elements(by.XPATH, selector) if element.is_displayed()]
        if buttons:
            from selenium.common.exceptions import WebDriverException
            from selenium.webdriver.support.ui import WebDriverWait
            try:
                buttons[0].click()
                WebDriverWait(driver, 4).until(
                    lambda d: "понедельник" in d.find_element(by.TAG_NAME, "body").text.casefold())
            except WebDriverException:
                print(f"Часть полей не раскрылась: {url}", file=sys.stderr)
    main = page_snapshot(driver)
    info_href = next((href for href, label in main["links"]
                      if clean(label) == "Инфо" and urlsplit(href).path.endswith("/tab/info")), None)
    info = {"text": "", "links": []}
    needs_info = not record or not record.get("attribute_groups")
    if needs_info and info_href and urlsplit(info_href).hostname == "2gis.ru":
        open_page(info_href)
        check_page(driver)
        wait.until(lambda d: any(heading.text.strip() for heading in d.find_elements(by.CSS_SELECTOR, "h1")))
        info = page_snapshot(driver)
    fill_missing(row, visible_card(main, info))
    return row


def collect_browser(args, candidates):
    try:
        from seleniumbase import SB
        from selenium.common.exceptions import TimeoutException, WebDriverException
        from selenium.webdriver.common.by import By
        from selenium.webdriver.support.ui import WebDriverWait
    except ImportError:
        raise RuntimeError("Установите SeleniumBase командой python -m pip install seleniumbase.") from None

    cancel_checkpoint(args)
    progress_event(args, "connecting")
    with SB(browser="chrome", uc=args.uc, headless=not args.uc and not args.headed,
            headless2=args.uc and not args.headed, locale_code="ru",
            chromium_arg="--no-proxy-server", driver_version="browser") as sb:
        driver = sb.driver
        driver.set_page_load_timeout(25)
        wait = WebDriverWait(driver, 12)

        def open_page(url):
            cancel_checkpoint(args)
            if args.uc:
                sb.uc_open_with_reconnect(url, reconnect_time=4)
            else:
                sb.open(url)
            cancel_checkpoint(args)

        if candidates is None:
            candidates = islice(browser_search(driver, open_page, args, By, wait), args.limit)
        skipped = 0
        for url, row in candidates:
            cancel_checkpoint(args)
            if url and any(not clean(row.get(column)) for column in COLUMNS):
                progress_event(args, "card", name=row.get("Название места") or "", url=url)
                try:
                    web_row = browser_card(driver, open_page, url, By, wait)
                except (TimeoutException, WebDriverException):
                    print(f"Карточка не загрузилась: {url}", file=sys.stderr)
                    skipped += 1
                    web_row = {}
                if args.source == "browser":
                    fill_missing(web_row, row)
                    row = web_row
                else:
                    fill_missing(row, web_row)
                time.sleep(1)
            if clean(row.get("Название места")):
                yield row
            cancel_checkpoint(args)
        if skipped:
            raise RuntimeError(f"Не удалось проверить {skipped} карточек в браузере.")


def search_page_number(url, search):
    link = urlsplit(url)
    base = urlsplit(search)
    if link.hostname != base.hostname:
        return None
    path = unquote(link.path).rstrip("/")
    original = unquote(base.path).rstrip("/")
    if path == original:
        return 1
    match = re.fullmatch(re.escape(original) + r"/page/(\d+)", path)
    return int(match[1]) if match else None


def browser_search(driver, open_page, args, by, wait):
    search = f"{BASE}/{args.city}/search/{quote(args.query, safe='')}"
    url = search
    seen = set()
    fingerprints = set()
    expected_page = 1
    while url:
        progress_event(args, "searching", page=expected_page)
        open_page(url)
        check_page(driver)
        actual_page = search_page_number(driver.current_url, search)
        if actual_page != expected_page:
            raise RuntimeError(f"Выдача перенаправила со страницы {expected_page} на другой адрес.")
        wait.until(lambda d: d.find_elements(by.CSS_SELECTOR, f'a[href*="/{args.city}/firm/"]')
                   or "ничего не найдено" in d.find_element(by.TAG_NAME, "body").text.casefold())
        snapshot = driver.execute_script(PAGE_STATE + """
            const search = state?.search;
            const number = String(arguments[0]);
            const matching = Object.entries(search?.profile ?? {}).find(
                ([key, value]) => String(value.data?.currentPage) === number
                    && Array.isArray(search?.pagination?.[key]?.[number]?.data));
            const ids = matching ? search.pagination[matching[0]][number].data : [];
            const profile = matching?.[1]?.data;
            return {
                links: Array.from(document.links, a => a.href),
                ids,
                records: Object.fromEntries(ids.map(id => [id, profiles[id]?.data ?? null])),
                descriptions: Object.fromEntries(ids.map(id =>
                    [id, state?.searchContext?.[matching[0] + '_' + id]?.ads?.text ?? null])),
                total: profile?.total ?? null,
                pages: profile?.pages ?? null,
                more: profile?.hasPagesToLoad ?? null,
            };
        """, expected_page)
        page_ids = {}
        next_url = None
        for href in snapshot["links"]:
            link = urlsplit(href)
            match = FIRM.match(link.path)
            if link.hostname == "2gis.ru" and match and match[1] == args.city:
                page_ids.setdefault(match[2], f"{BASE}/{match[1]}/firm/{match[2]}")
            if search_page_number(href, search) == expected_page + 1:
                next_url = href
        if snapshot["ids"]:
            page_ids = {identifier: f"{BASE}/{args.city}/firm/{identifier}" for identifier in snapshot["ids"]
                        if isinstance(identifier, str) and identifier.isdigit()
                        and (identifier in page_ids or (snapshot["records"].get(identifier) or {}).get("type") == "branch")}
        if not next_url and snapshot["more"] and isinstance(snapshot["pages"], int) and expected_page < snapshot["pages"]:
            next_url = f"{search}/page/{expected_page + 1}"
        fingerprint = frozenset(page_ids)
        if fingerprint and fingerprint in fingerprints:
            raise RuntimeError(f"Страница {expected_page} повторяет ранее полученную выдачу.")
        fingerprints.add(fingerprint)
        new_ids = [identifier for identifier in page_ids if identifier not in seen]
        seen.update(new_ids)
        total = snapshot["total"]
        page_count = snapshot["pages"]
        page_label = f"{expected_page}/{page_count}" if isinstance(page_count, int) else str(expected_page)
        print(f"Страница {page_label}: карточек {len(page_ids)}, новых {len(new_ids)}.", file=sys.stderr)
        progress_event(args, "page", page=expected_page, pages=page_count, found=len(new_ids), total=total)
        for identifier in new_ids:
            record = snapshot["records"].get(identifier)
            row = record_row(record, include_phone=False) if record else {}
            if not clean(row.get("Краткое описание")):
                description = snapshot["descriptions"].get(identifier)
                if description:
                    row["Краткое описание"] = description
            yield page_ids[identifier], row
        if args.pages is not None and expected_page >= args.pages:
            return
        if not next_url:
            print(f"Конец доступной выдачи. Уникальных карточек: {len(seen)}.", file=sys.stderr)
            if isinstance(total, int) and len(seen) < total:
                raise RuntimeError(f"2ГИС указал {total} результатов, но выдал {len(seen)} уникальных карточек.")
        url = next_url
        expected_page += 1


def coordinates(lat, lon):
    try:
        lat, lon = float(lat), float(lon)
    except (TypeError, ValueError):
        return None
    if -90 <= lat <= 90 and -180 <= lon <= 180:
        return lat, lon
    return None


def read_ground_stops(path):
    with path.open(encoding="utf-8-sig", newline="") as file:
        reader = csv.DictReader(file)
        required = {"stop_name", "stop_lat", "stop_lon"}
        if not required.issubset(reader.fieldnames or []):
            raise ValueError("В справочнике остановок нужны поля stop_name, stop_lat и stop_lon.")
        stops = set()
        for row in reader:
            name = clean(row["stop_name"])
            point = coordinates(row["stop_lat"], row["stop_lon"])
            if name and point:
                stops.add((name, point))
    if not stops:
        raise ValueError("В справочнике нет именованных остановок с координатами.")
    return stops


def distance_in_metres(origin, target):
    lat1, lon1 = map(math.radians, origin)
    lat2, lon2 = map(math.radians, target)
    haversine = math.sin((lat2 - lat1) / 2) ** 2
    haversine += math.cos(lat1) * math.cos(lat2) * math.sin((lon2 - lon1) / 2) ** 2
    return 12_742_000 * math.asin(math.sqrt(min(1, haversine)))


def fill_ground_stops(rows, stops):
    filled = 0
    for row in rows:
        if clean(row.get(GROUND_NAME)) or clean(row.get(GROUND_DISTANCE)):
            continue
        point = row.get("_point") or {}
        origin = coordinates(point.get("lat"), point.get("lon"))
        if origin is None:
            continue
        candidates = ((distance_in_metres(origin, target), name) for name, target in stops
                      if abs(target[0] - origin[0]) <= 0.01)
        distance, name = min(candidates, default=(math.inf, ""))
        if distance <= 1000:
            row[GROUND_NAME] = name
            row[GROUND_DISTANCE] = round(distance)
            filled += 1
    return filled


def write_csv(path, rows, incomplete=False):
    values = [[clean(row.get(column)) for column in COLUMNS] for row in rows]
    with path.open("x", newline="", encoding="utf-8-sig") as file:
        writer = csv.writer(file, delimiter=";")
        writer.writerow(COLUMNS)
        writer.writerows(values)
    report_export(path, values, incomplete)


def report_export(path, values, incomplete):
    counts = Counter(column for row in values for column, value in zip(COLUMNS, row) if value)
    status = "Частичная выгрузка" if incomplete else "Выгрузка"
    print(f"{status}: {len(values)} карточек. Файл: {path}")
    for column in COLUMNS:
        print(f"{column}: {counts[column]}/{len(values)}")


def excel_text(value):
    value = re.sub(r"_(?=x[0-9a-fA-F]{4}_)", "_x005F_", value)
    value = re.sub(r"[\x00-\x08\x0b-\x0d\x0e-\x1f\ufffe\uffff]",
                   lambda match: f"_x{ord(match[0]):04X}_", value)
    return escape(value)


def excel_column(index):
    name = ""
    while index:
        index, digit = divmod(index - 1, 26)
        name = chr(65 + digit) + name
    return name


def write_xlsx(path, columns, values):
    spreadsheet = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    office = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    package = "http://schemas.openxmlformats.org/package/2006/relationships"
    content = "http://schemas.openxmlformats.org/package/2006/content-types"
    excel_type = "application/vnd.openxmlformats-officedocument.spreadsheetml"
    parts = {
        "[Content_Types].xml": (
            f'<Types xmlns="{content}">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            f'<Override PartName="/xl/workbook.xml" ContentType="{excel_type}.sheet.main+xml"/>'
            f'<Override PartName="/xl/worksheets/sheet1.xml" ContentType="{excel_type}.worksheet+xml"/>'
            f'<Override PartName="/xl/styles.xml" ContentType="{excel_type}.styles+xml"/></Types>'
        ),
        "_rels/.rels": (
            f'<Relationships xmlns="{package}">'
            f'<Relationship Id="rId1" Type="{office}/officeDocument" Target="xl/workbook.xml"/>'
            '</Relationships>'
        ),
        "xl/workbook.xml": (
            f'<workbook xmlns="{spreadsheet}" xmlns:r="{office}">'
            '<sheets><sheet name="Карточки" sheetId="1" r:id="rId1"/></sheets></workbook>'
        ),
        "xl/_rels/workbook.xml.rels": (
            f'<Relationships xmlns="{package}">'
            f'<Relationship Id="rId1" Type="{office}/worksheet" Target="worksheets/sheet1.xml"/>'
            f'<Relationship Id="rId2" Type="{office}/styles" Target="styles.xml"/></Relationships>'
        ),
        "xl/styles.xml": (
            f'<styleSheet xmlns="{spreadsheet}">'
            '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>'
            '<fills count="2"><fill><patternFill patternType="none"/></fill>'
            '<fill><patternFill patternType="gray125"/></fill></fills>'
            '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
            '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
            '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
            '<xf numFmtId="49" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
            '</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
            '</styleSheet>'
        ),
    }
    with path.open("xb") as file, ZipFile(file, "w", compression=ZIP_DEFLATED) as archive:
        for name, xml in parts.items():
            archive.writestr(name, xml.encode("utf-8"))
        with archive.open("xl/worksheets/sheet1.xml", "w") as sheet:
            sheet.write(f'<worksheet xmlns="{spreadsheet}"><cols>'.encode("utf-8"))
            for index in range(1, len(columns) + 1):
                longest = max(len(row[index - 1]) for row in chain((columns,), values))
                width = min(55, max(12, longest + 2))
                sheet.write(f'<col min="{index}" max="{index}" width="{width}" customWidth="1" style="1"/>'.encode("utf-8"))
            sheet.write(b"</cols><sheetData>")
            names = [excel_column(index) for index in range(1, len(columns) + 1)]
            for index, row in enumerate(chain((columns,), values), 1):
                cells = "".join(
                    f'<c r="{name}{index}" s="1" t="inlineStr">'
                    f'<is><t xml:space="preserve">{excel_text(value)}</t></is></c>'
                    for name, value in zip(names, row)
                )
                sheet.write(f'<row r="{index}">{cells}</row>'.encode("utf-8"))
            sheet.write(b"</sheetData></worksheet>")


def main(argv=None, *, on_progress=None, is_cancelled=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--source", choices=("browser", "api", "hybrid"), default="browser")
    parser.add_argument("--city", default="moscow", help="Код города в URL 2ГИС")
    parser.add_argument("--region-id", type=int, default=32, help="ID проекта Places API (32 = Москва)")
    parser.add_argument("--query", default="Кафе")
    parser.add_argument("--limit", type=int, help="Максимум карточек. По умолчанию до конца доступной выдачи")
    parser.add_argument("--pages", type=int, help="Максимум страниц")
    parser.add_argument("--output", type=Path, default=Path("2gis-cards.xlsx"), help="Путь к файлу .xlsx или .csv")
    parser.add_argument("--headed", action="store_true", help="Показывать окно Chrome")
    parser.add_argument("--uc", action="store_true", help="Запускать Chrome через SeleniumBase UC")
    parser.add_argument("--api-contacts", action="store_true", help="Запросить доступные ключу контакты API")
    parser.add_argument("--ground-stops", type=Path, help="Файл stops.txt с координатами наземных остановок")
    args = parser.parse_args(argv)
    args.on_progress = on_progress
    args.is_cancelled = is_cancelled
    if (not re.fullmatch(r"[a-z0-9_-]+", args.city) or not args.query.strip() or args.region_id < 1
            or any(value is not None and value < 1 for value in (args.limit, args.pages))):
        parser.error("Укажите город, запрос и положительные значения лимитов и ID региона.")
    if args.output.exists():
        parser.error(f"Файл уже существует: {args.output}")
    extension = args.output.suffix.casefold()
    if extension not in {".xlsx", ".csv"}:
        parser.error("Укажите выходной файл с расширением .xlsx или .csv.")
    stops = None
    if args.ground_stops:
        try:
            stops = read_ground_stops(args.ground_stops)
        except (OSError, ValueError) as error:
            parser.error(str(error))
    key = os.getenv("TWOGIS_API_KEY")
    if args.source != "browser" and not key:
        parser.error("Для выбранного режима задайте переменную TWOGIS_API_KEY.")
    rows = []
    api_candidates = {}
    incomplete = False
    cancelled = False
    try:
        if args.source == "browser":
            with closing(collect_browser(args, None)) as candidates:
                for row in candidates:
                    rows.append(row)
                    progress_event(args, "row", count=len(rows), row=row)
        else:
            for item in api_items(args, key):
                cancel_checkpoint(args)
                identifier = str(item.get("id", "")).split("_")[0]
                if not identifier.isdigit() or identifier in api_candidates:
                    continue
                api_candidates[identifier] = record_row(item)
                progress_event(args, "row", count=len(api_candidates), row=api_candidates[identifier])
                if len(api_candidates) == args.limit:
                    break
            if args.source == "hybrid" and api_candidates:
                cards = ((f"{BASE}/{args.city}/firm/{identifier}", row) for identifier, row in api_candidates.items())
                try:
                    with closing(collect_browser(args, cards)) as candidates:
                        for _ in candidates:
                            pass
                except ExportCancelled:
                    raise
                except Exception as error:
                    incomplete = True
                    print(f"Данные API будут сохранены. Браузерный сбор прерван: {error}", file=sys.stderr)
    except ExportCancelled:
        cancelled = True
        incomplete = True
    except Exception as error:
        incomplete = True
        print(f"Сбор прерван: {type(error).__name__}: {str(error).splitlines()[0] if str(error) else ''}",
              file=sys.stderr)
        progress_event(args, "warning", message=str(error).splitlines()[0] if str(error) else type(error).__name__)
    rows.extend(api_candidates.values())
    if not rows:
        if cancelled:
            return 130
        print("Карточки не найдены, файл не создан.", file=sys.stderr)
        return 2 if on_progress and not incomplete else 1
    progress_event(args, "saving", count=len(rows))
    if stops:
        print(f"Наземные остановки из справочника: {fill_ground_stops(rows, stops)}.", file=sys.stderr)
    if extension == ".xlsx":
        values = [[clean(row.get(column)) for column in COLUMNS] for row in rows]
        write_xlsx(args.output, COLUMNS, values)
        report_export(args.output, values, incomplete)
    else:
        write_csv(args.output, rows, incomplete)
    return 130 if cancelled else int(incomplete)


if __name__ == "__main__":
    sys.exit(main())
