"""Small, polite adapters for Moroccan job boards not supported by JobSpy."""

from __future__ import annotations

import json
import re
import time
import unicodedata
from datetime import datetime, timedelta, timezone
from urllib.parse import parse_qs, urlencode, urljoin, urlsplit, urlunsplit

import pandas as pd
import requests
from bs4 import BeautifulSoup


EMPLOI_MA = "emploi_ma"
MAROC_ANNONCES = "marocannonces"
REKRUTE = "rekrute"
MAROC_EMPLOI = "marocemploi"
LOCAL_SOURCES = {EMPLOI_MA, MAROC_ANNONCES, REKRUTE, MAROC_EMPLOI}

_EMPLOI_ROOT = "https://www.emploi.ma"
_MAROC_ROOT = "https://www.marocannonces.com"
_REKRUTE_ROOT = "https://www.rekrute.com"
_MAROC_EMPLOI_ROOT = "https://marocemploi.net"
_SOURCE_LISTING_URLS = {
    MAROC_ANNONCES: "https://www.marocannonces.com/emploi",
    REKRUTE: f"{_REKRUTE_ROOT}/fr/offres-emploi-maroc.html",
    MAROC_EMPLOI: f"{_MAROC_EMPLOI_ROOT}/offre/",
}
_PAGE_SIZES = {
    MAROC_ANNONCES: 20,
    REKRUTE: 10,
    MAROC_EMPLOI: 20,
}
_REQUEST_INTERVAL_SECONDS = 2.0
_REQUEST_TIMEOUT_SECONDS = 20
_LAST_REQUEST_BY_HOST = {}
_FORM_PAGE_CACHE = {}
_PAGE_CACHE = {}
_HTTP_SESSION = requests.Session()

_CITIES = (
    "Casablanca", "Rabat", "Tanger", "Tangier", "Marrakech", "Fès",
    "Fez", "Agadir", "Kénitra", "Kenitra", "Meknès", "Meknes",
    "Oujda", "Tétouan", "Tetouan", "El Jadida", "Safi", "Salé",
    "Sale", "Mohammedia", "Béni Mellal", "Beni Mellal", "Nador",
    "Settat", "Khouribga", "Laâyoune", "Laayoune", "Dakhla", "Temara",
    "Bouskoura", "Berrechid", "Nouaceur", "Mediouna", "Tout le Maroc",
)

def scrape_local_jobs(source, keyword, city, results_wanted=20, hours_old=None):
    """Return public job listings in the same dataframe shape as JobSpy."""

    if source == EMPLOI_MA:
        return _scrape_emploi_ma(
            keyword, city, int(results_wanted), hours_old
        )
    if source == MAROC_ANNONCES:
        return _scrape_marocannonces(
            keyword, city, int(results_wanted), hours_old
        )
    if source in (REKRUTE, MAROC_EMPLOI):
        return _scrape_form_source(
            source, keyword, city, int(results_wanted), hours_old
        )
    raise ValueError(f"Unknown local job source: {source}")


def _scrape_emploi_ma(keyword, city, limit, hours_old):
    keyword_slug = _slugify(keyword)
    keyword_url = f"{_EMPLOI_ROOT}/recherche-jobs-maroc/{keyword_slug}"
    keyword_html = _get_html(keyword_url)
    records = _parse_listing_page(
        keyword_html,
        source="Emploi.ma",
        base_url=keyword_url,
        requested_city=city,
        keyword=keyword,
        server_keyword=True,
        server_city=False,
    )

    # Emploi.ma also exposes city result pages. Search those when the keyword
    # results do not fill the requested limit, then match the keyword locally.
    if len(records) < limit:
        city_url = f"{_EMPLOI_ROOT}/recherche-jobs-maroc/{_slugify(city)}"
        try:
            city_html = _get_html(city_url)
        except requests.HTTPError as error:
            if not error.response or error.response.status_code != 404:
                raise
        else:
            records.extend(_parse_listing_page(
                city_html,
                source="Emploi.ma",
                base_url=city_url,
                requested_city=city,
                keyword=keyword,
                server_keyword=False,
                server_city=True,
            ))

    return _filter_and_limit(records, limit, hours_old)


def _scrape_marocannonces(keyword, city, limit, hours_old):
    return _scrape_form_source(
        MAROC_ANNONCES, keyword, city, limit, hours_old
    )


def _scrape_form_source(source, keyword, city, limit, hours_old):
    listing_url = _SOURCE_LISTING_URLS[source]
    now = time.monotonic()
    cached = _FORM_PAGE_CACHE.get(listing_url)
    if cached is None or now - cached[0] > 1800:
        form_html = _get_html(listing_url)
        _FORM_PAGE_CACHE[listing_url] = (now, form_html)
    else:
        form_html = cached[1]

    request_spec = _discover_search_form(
        form_html, keyword, city, fallback_action=listing_url
    )
    if request_spec is None:
        page_url = listing_url
        search_html = form_html
        server_keyword = False
        server_city = False
        page_request = None
    else:
        method, action, values, server_keyword, server_city = request_spec
        search_html, page_url = _submit_form(method, action, values)
        page_request = (method, action, values)

    records = _parse_listing_page(
        search_html,
        source=_source_display_name(source),
        base_url=page_url,
        requested_city=city,
        keyword=keyword,
        server_keyword=server_keyword,
        server_city=server_city,
    )

    # Follow a small number of result pages when a query yields fewer rows
    # than requested. Each request is rate limited by host.
    page_size = _PAGE_SIZES[source]
    max_pages = max(1, min(5, (limit + page_size - 1) // page_size))
    if not server_keyword or not server_city:
        max_pages = min(5, max(max_pages, 3))
    current_html = search_html
    seen_page_urls = {page_url}
    for _ in range(1, max_pages):
        if len(records) >= limit:
            break
        next_url = _next_page_url(current_html, page_url)
        if next_url is None or next_url in seen_page_urls:
            break
        seen_page_urls.add(next_url)

        if page_request is not None and page_request[0] == "post":
            method, action, values = page_request
            page_number = _page_number(next_url)
            values = dict(values)
            values["pge"] = str(page_number)
            current_html, page_url = _submit_form(method, action, values)
        else:
            current_html = _get_html(next_url)
            page_url = next_url
        records.extend(
            _parse_listing_page(
                current_html,
                source=_source_display_name(source),
                base_url=page_url,
                requested_city=city,
                keyword=keyword,
                server_keyword=server_keyword,
                server_city=server_city,
            )
        )

    return _filter_and_limit(records, limit, hours_old)


def _source_display_name(source):
    return {
        MAROC_ANNONCES: "MarocAnnonces",
        REKRUTE: "ReKrute.com",
        MAROC_EMPLOI: "MarocEmploi.net",
    }[source]


def _discover_search_form(html, keyword, city, fallback_action):
    soup = BeautifulSoup(html, "html.parser")
    for form in soup.find_all("form"):
        fields = {}
        keyword_field = None
        city_field = None
        city_field_type = None
        text_fields = []

        for element in form.find_all(["input", "select"]):
            name = element.get("name")
            if not name:
                continue
            element_type = (element.get("type") or "").lower()
            if element.name == "input":
                if element_type in ("submit", "button", "image", "file"):
                    continue
                if element_type in ("checkbox", "radio") and not element.has_attr("checked"):
                    continue
                fields[name] = element.get("value", "")
                label = _field_label(element)
                is_text = element_type in ("", "text", "search")
                has_keyword_label = any(word in label for word in (
                    "mot", "keyword", "search", "query", "recherche", "key",
                    "poste", "metier", "emploi",
                ))
                has_city_label = any(word in label for word in (
                    "ville", "city", "location", "region", "localite",
                ))
                if is_text:
                    text_fields.append((name, has_city_label, has_keyword_label))
                if is_text and has_city_label and city_field is None:
                    city_field = name
                    city_field_type = "input"
                if is_text and has_keyword_label and keyword_field is None:
                    keyword_field = name
            else:
                options = element.find_all("option")
                selected_value = ""
                label = _field_label(element)
                if any(word in label for word in (
                    "ville", "city", "location", "region", "localite",
                )):
                    city_field = name
                    city_field_type = "select"
                    requested = _normalize(city)
                    for option in options:
                        option_label = _normalize(option.get_text(" ", strip=True))
                        option_value = _normalize(option.get("value", ""))
                        if requested and (
                            requested == option_label or requested == option_value
                        ):
                            selected_value = option.get("value", "")
                            break
                if not selected_value:
                    selected = element.find("option", selected=True)
                    selected_value = selected.get("value", "") if selected else ""
                fields[name] = selected_value

        if keyword_field is None:
            fallback = next(
                (field for field in text_fields if not field[1]),
                text_fields[0] if text_fields else None,
            )
            if fallback:
                keyword_field = fallback[0]

        if keyword_field:
            fields[keyword_field] = keyword
            server_city = False
            if city_field and city_field != keyword_field:
                if city_field_type == "input":
                    fields[city_field] = city
                    server_city = True
                else:
                    city_select = form.find("select", attrs={"name": city_field})
                    if city_select:
                        requested = _normalize(city)
                        for option in city_select.find_all("option"):
                            label = _normalize(option.get_text(" ", strip=True))
                            value = _normalize(option.get("value", ""))
                            if requested and (requested == label or requested == value):
                                fields[city_field] = option.get("value", "")
                                server_city = True
                                break
            action = form.get("action") or fallback_action
            action = urljoin(fallback_action, action)
            return (
                (form.get("method") or "get").lower(),
                action,
                fields,
                True,
                server_city,
            )
    return None


def _field_label(element):
    parts = [
        str(element.get(attr, ""))
        for attr in ("name", "id", "placeholder", "aria-label")
    ]
    element_id = element.get("id")
    if element_id:
        for label in element.find_parent().find_all("label", attrs={"for": element_id}):
            parts.append(label.get_text(" ", strip=True))
    return _normalize(" ".join(parts))


def _submit_form(method, action, values):
    if method == "post":
        response = _request(action, method="post", data=values)
    else:
        response = _request(action, params=values)
    return response.text, response.url


def _parse_listing_page(
    html,
    source,
    base_url,
    requested_city,
    keyword,
    server_keyword,
    server_city,
):
    soup = BeautifulSoup(html, "html.parser")
    records = []
    seen = set()

    for item in _job_postings_from_jsonld(soup, source, base_url):
        normalized_blob = item.get("_search_text", "")
        if not server_keyword and not _keyword_matches(keyword, normalized_blob):
            continue
        if not server_city and not _city_matches(
            requested_city, item.get("location"), normalized_blob
        ):
            continue
        key = item.get("job_url") or item.get("title")
        if key and key not in seen:
            seen.add(key)
            records.append(item)

    for anchor in soup.find_all("a", href=True):
        href = anchor.get("href", "").strip()
        if source == "ReKrute.com" and "matching4K" in href:
            continue
        if not _is_job_link(href, source):
            continue

        card = _find_card_container(anchor)
        card_text = card.get_text(" ", strip=True)
        title = _extract_title(anchor, card)
        if not title or len(title) < 3:
            continue
        if _is_navigation_title(title):
            continue

        normalized_blob = _normalize(f"{title} {card_text}")
        if not server_keyword and not _keyword_matches(keyword, normalized_blob):
            continue

        location = _extract_location(card_text)
        if not server_city and not _city_matches(requested_city, location, card_text):
            continue

        job_url = _absolute_job_url(href, source, base_url)
        key = job_url or _normalize(title)
        if key in seen:
            continue
        seen.add(key)

        records.append({
            "site": source,
            "title": title,
            "company": _extract_company(card),
            "location": location,
            "date_posted": _extract_date(card),
            "job_type": _extract_job_type(card_text),
            "min_amount": None,
            "max_amount": None,
            "currency": "MAD" if _has_salary(card_text) else None,
            "job_url": job_url,
            "_search_text": normalized_blob,
        })

    return records


def _job_postings_from_jsonld(soup, source, base_url):
    records = []

    def walk(value):
        if isinstance(value, list):
            for child in value:
                walk(child)
            return
        if not isinstance(value, dict):
            return

        types = value.get("@type", [])
        if isinstance(types, str):
            types = [types]
        if "JobPosting" in types:
            company = value.get("hiringOrganization") or {}
            if isinstance(company, dict):
                company = company.get("name")
            location = value.get("jobLocation") or {}
            if isinstance(location, list):
                location = location[0] if location else {}
            if isinstance(location, dict):
                address = location.get("address") or {}
                if isinstance(address, dict):
                    location = ", ".join(
                        part for part in (
                            address.get("addressLocality"),
                            address.get("addressRegion"),
                            address.get("addressCountry"),
                        ) if part
                    )
                else:
                    location = str(location.get("name", ""))
            records.append({
                "site": source,
                "title": value.get("title"),
                "company": company,
                "location": location,
                "date_posted": value.get("datePosted"),
                "job_type": value.get("employmentType"),
                "min_amount": None,
                "max_amount": None,
                "currency": None,
                "job_url": value.get("url") or value.get("@id") or base_url,
                "_search_text": _normalize(json.dumps(value, ensure_ascii=False)),
            })
        for child in value.values():
            if isinstance(child, (dict, list)):
                walk(child)

    for script in soup.find_all("script", type="application/ld+json"):
        try:
            walk(json.loads(script.string or script.get_text()))
        except (TypeError, ValueError):
            continue
    return records


def _is_job_link(href, source):
    path = urlsplit(href).path.lower()
    if source == "MarocAnnonces":
        return bool(re.search(r"/annonce/\d+", path))
    if source == "ReKrute.com":
        return "/offre-emploi-" in path and path.endswith(".html")
    if source == "MarocEmploi.net":
        return bool(re.search(r"/offre/[^/]+/?$", path))
    return (
        "/offre-emploi" in path
        or "/emploi/" in path
        or "/job/" in path
        or "/node/" in path
    )


def _find_card_container(anchor):
    best = anchor
    for parent in anchor.parents:
        if parent.name in ("body", "html"):
            break
        text = parent.get_text(" ", strip=True)
        if len(text) > 2200:
            break
        classes = " ".join(parent.get("class", []))
        marker = f"{parent.name} {classes} {parent.get('id', '')}".lower()
        if (
            parent.name in ("article", "li")
            or re.search(r"job|emploi|annonce|offer|listing|views-row|card|result", marker)
        ) and len(text) >= 15:
            best = parent
            break
        if len(text) >= 15:
            best = parent
    return best


def _extract_title(anchor, card):
    heading = card.find(re.compile(r"^h[1-6]$"))
    if heading:
        text = heading.get_text(" ", strip=True)
        if text:
            return _trim_title(text)
    title_node = card.find(attrs={"class": re.compile(r"title|titre", re.I)})
    if title_node:
        text = title_node.get_text(" ", strip=True)
        if text:
            return _trim_title(text)
    return _trim_title(anchor.get_text(" ", strip=True))


def _trim_title(value):
    value = re.sub(r"\s+", " ", value).strip(" -|\t\n")
    value = re.split(
        r"\b(?:niveau d[’']etudes|niveau d[’']exp[eé]rience|salaire\s*:|"
        r"contrat propos[eé]\s*:|r[eé]gion de\s*:|comp[eé]tences cl[eé]s\s*:)",
        value,
        maxsplit=1,
        flags=re.I,
    )[0]
    return value.strip(" -|\t\n")


def _extract_company(card):
    for node in card.find_all(True):
        marker = " ".join(node.get("class", [])) + " " + str(node.get("itemprop", ""))
        if re.search(r"company|employer|recruiter|entreprise|recruteur", marker, re.I):
            text = node.get_text(" ", strip=True)
            if text and len(text) < 180:
                return text
    return None


def _extract_location(text):
    normalized = _normalize(text)
    for city in sorted(_CITIES, key=len, reverse=True):
        candidate = _normalize(city)
        if re.search(rf"(?<!\w){re.escape(candidate)}(?!\w)", normalized):
            return city
    return None


def _extract_date(card):
    time_node = card.find("time")
    if time_node:
        value = time_node.get("datetime") or time_node.get_text(" ", strip=True)
        parsed = _parse_posted_date(value)
        if parsed:
            return parsed
    return _parse_posted_date(card.get_text(" ", strip=True))


def _parse_posted_date(value):
    if not value:
        return None
    text = str(value).strip()
    normalized = _normalize(text)
    now = datetime.now(timezone.utc)

    if "aujourd" in normalized or re.search(r"\btoday\b", normalized):
        return now.isoformat()
    if re.search(r"\bhier\b|\byesterday\b", normalized):
        return (now - timedelta(days=1)).isoformat()

    relative = re.search(
        r"(?:il y a\s*)?(\d+)\s*(minutes?|mins?|mn|heures?|hours?|h|jours?|days?|j|semaines?|weeks?|mois)",
        normalized,
    )
    if relative:
        count = int(relative.group(1))
        unit = relative.group(2)
        if unit.startswith(("minute", "min", "mn")):
            delta = timedelta(minutes=count)
        elif unit.startswith(("heure", "hour", "h")):
            delta = timedelta(hours=count)
        elif unit.startswith(("jour", "day", "j")):
            delta = timedelta(days=count)
        elif unit.startswith(("semaine", "week")):
            delta = timedelta(weeks=count)
        else:
            delta = timedelta(days=30 * count)
        return (now - delta).isoformat()

    date_match = re.search(r"\b(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})\b", text)
    if date_match:
        day, month, year = map(int, date_match.groups())
        if year < 100:
            year += 2000
        try:
            return datetime(year, month, day, tzinfo=timezone.utc).isoformat()
        except ValueError:
            return None

    try:
        parsed = datetime.fromisoformat(text.replace("Z", "+00:00"))
        if parsed.tzinfo is None:
            parsed = parsed.replace(tzinfo=timezone.utc)
        return parsed.astimezone(timezone.utc).isoformat()
    except ValueError:
        return None


def _extract_job_type(text):
    match = re.search(
        r"(?:contrat propos[eé]\s*:\s*)([^|.]{2,80})",
        text,
        flags=re.I,
    )
    return match.group(1).strip() if match else None


def _has_salary(text):
    return bool(re.search(r"\b(?:salaire|dh|dhs|mad)\b", text, re.I))


def _keyword_matches(keyword, searchable):
    query = _normalize(keyword)
    if not query:
        return True
    tokens = [token for token in query.split() if len(token) > 2]
    return query in searchable or not tokens or any(token in searchable for token in tokens)


def _city_matches(city, location, searchable):
    requested = _normalize(city)
    if not requested:
        return True
    return requested in _normalize(location or "") or requested in searchable


def _filter_and_limit(records, limit, hours_old):
    cutoff = None
    if hours_old is not None:
        cutoff = datetime.now(timezone.utc) - timedelta(hours=hours_old)

    filtered = []
    seen = set()
    for record in records:
        title = record.get("title")
        if not title:
            continue
        date_value = record.get("date_posted")
        if cutoff and date_value:
            parsed = pd.to_datetime(date_value, utc=True, errors="coerce")
            if not pd.isna(parsed) and parsed.to_pydatetime() < cutoff:
                continue
        key = record.get("job_url") or _normalize(title)
        if key in seen:
            continue
        seen.add(key)
        filtered.append({
            key: value for key, value in record.items()
            if not key.startswith("_")
        })
        if len(filtered) >= limit:
            break
    return pd.DataFrame(filtered)


def _next_page_url(html, current_url):
    soup = BeautifulSoup(html, "html.parser")
    candidates = []
    current_page = _page_number(current_url)
    for anchor in soup.find_all("a", href=True):
        href = anchor.get("href", "")
        query = parse_qs(urlsplit(href).query)
        page_values = query.get("pge") or query.get("page")
        path_page = re.search(r"/(?:page|pages?)/(\d+)(?:/|$)", urlsplit(href).path)
        label = _normalize(anchor.get_text(" ", strip=True))
        rel = anchor.get("rel", [])
        if isinstance(rel, str):
            rel = rel.split()
        if page_values:
            try:
                page_num = int(page_values[0])
            except ValueError:
                continue
            if page_num > current_page:
                candidates.append((page_num, _absolute_url(href, current_url)))
        elif path_page and int(path_page.group(1)) > current_page:
            candidates.append((int(path_page.group(1)), _absolute_url(href, current_url)))
        elif "next" in rel or label in ("suivant", "next", ">", "»"):
            candidates.append((current_page + 1, _absolute_url(href, current_url)))

    for select in soup.find_all("select"):
        marker = _normalize(
            f"{select.get('name', '')} {select.get('id', '')} "
            f"{select.get('aria-label', '')}"
        )
        if not any(word in marker for word in ("page", "pge")):
            continue
        for option in select.find_all("option"):
            value = (option.get("value") or "").strip()
            if value.isdigit() and int(value) > current_page:
                parts = urlsplit(current_url)
                query = parse_qs(parts.query)
                query["page"] = [value]
                next_url = urlunsplit((
                    parts.scheme,
                    parts.netloc,
                    parts.path,
                    urlencode(query, doseq=True),
                    parts.fragment,
                ))
                candidates.append((int(value), next_url))
                break
    if not candidates:
        return None
    return sorted(candidates, key=lambda candidate: candidate[0])[0][1]


def _page_number(url):
    values = parse_qs(urlsplit(url).query).get("pge") or parse_qs(urlsplit(url).query).get("page")
    if values:
        try:
            return int(values[0])
        except ValueError:
            pass
    return 1


def _absolute_job_url(href, source, base_url):
    if href.startswith("http://") or href.startswith("https://") or href.startswith("//"):
        result = urljoin(base_url, href)
    elif source == "MarocAnnonces":
        if href.startswith("/"):
            path = href
        else:
            path = "/" + href.lstrip("/")
        duplicate = "/categorie/309/categorie/309/"
        path = path.replace(duplicate, "/categorie/309/")
        result = _MAROC_ROOT + path
    else:
        result = _absolute_url(href, base_url)
    result = result.replace(
        "/categorie/309/categorie/309/", "/categorie/309/"
    )
    return result.split("#", 1)[0].rstrip("?")


def _absolute_url(href, base_url):
    result = urljoin(base_url, href)
    return result.replace(
        "/categorie/309/categorie/309/", "/categorie/309/"
    )


def _get_html(url, params=None):
    if params is None:
        cached = _PAGE_CACHE.get(url)
        if cached is not None and time.monotonic() - cached[0] < 300:
            return cached[1]
    html = _request(url, params=params).text
    if params is None:
        _PAGE_CACHE[url] = (time.monotonic(), html)
    return html


def _request(url, method="get", params=None, data=None):
    host = urlsplit(url).netloc
    elapsed = time.monotonic() - _LAST_REQUEST_BY_HOST.get(host, 0)
    if elapsed < _REQUEST_INTERVAL_SECONDS:
        time.sleep(_REQUEST_INTERVAL_SECONDS - elapsed)
    _LAST_REQUEST_BY_HOST[host] = time.monotonic()

    response = _HTTP_SESSION.request(
        method,
        url,
        params=params,
        data=data,
        timeout=_REQUEST_TIMEOUT_SECONDS,
        headers={
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/131.0.0.0 Safari/537.36"
            ),
            "Accept": (
                "text/html,application/xhtml+xml,application/xml;q=0.9,"
                "*/*;q=0.8"
            ),
            "Accept-Language": "fr-MA,fr;q=0.9,en;q=0.8",
        },
    )
    _LAST_REQUEST_BY_HOST[host] = time.monotonic()
    response.raise_for_status()
    return response


def _slugify(value):
    value = unicodedata.normalize("NFKD", str(value))
    value = "".join(char for char in value if not unicodedata.combining(char))
    return re.sub(r"[^a-z0-9]+", "-", value.lower()).strip("-")


def _normalize(value):
    value = unicodedata.normalize("NFKD", str(value or "").lower())
    value = "".join(char for char in value if not unicodedata.combining(char))
    return re.sub(r"\s+", " ", value).strip()


def _is_navigation_title(title):
    return _normalize(title) in {
        "emploi", "offres emploi", "offre emploi", "suivant", "precedent",
        "page suivante", "voir plus", "emploi maroc",
    }
