import asyncio
import logging
from dataclasses import dataclass, field
from typing import Optional
from urllib.parse import urljoin

import httpx
import yaml
from bs4 import BeautifulSoup

logger = logging.getLogger(__name__)


@dataclass
class ScrapedRecord:
    title: str = ""
    authors: str = ""
    record_url: str = ""
    fulltext_url: str = ""
    format_type: str = ""


@dataclass
class TargetResult:
    target_id: str
    target_name: str
    records: list[ScrapedRecord] = field(default_factory=list)
    total_count: int = 0
    error: Optional[str] = None


@dataclass
class TargetConfig:
    id: str
    name: str
    description: str
    url_template: str
    selectors: dict


@dataclass
class AppConfig:
    base_url: str
    max_results_per_target: int
    results_per_page: int
    targets: list[TargetConfig]


def load_config(path: str = "config/targets.yaml") -> AppConfig:
    with open(path, "r", encoding="utf-8") as f:
        raw = yaml.safe_load(f)

    targets = [
        TargetConfig(
            id=t["id"],
            name=t["name"],
            description=t.get("description", ""),
            url_template=t["url_template"],
            selectors=t["selectors"],
        )
        for t in raw["targets"]
    ]

    return AppConfig(
        base_url=raw["base_url"],
        max_results_per_target=raw.get("max_results_per_target", 50),
        results_per_page=raw.get("results_per_page", 20),
        targets=targets,
    )


def _parse_record(record_html, selectors: dict, base_url: str) -> ScrapedRecord | None:
    try:
        soup = record_html if isinstance(record_html, BeautifulSoup) else BeautifulSoup(str(record_html), "html.parser")

        title_el = soup.select_one(selectors["title"])
        title = title_el.get_text(strip=True) if title_el else ""

        author_els = soup.select(selectors["author"])
        authors = "; ".join(a.get_text(strip=True) for a in author_els) if author_els else ""

        record_url_el = soup.select_one(selectors["record_url"])
        record_url = urljoin(base_url, record_url_el.get("href", "")) if record_url_el else ""

        fulltext_el = soup.select_one(selectors["fulltext_url"])
        fulltext_url = fulltext_el.get("href", "") if fulltext_el else ""

        format_els = soup.select(selectors["format"])
        format_type = "; ".join(f.get_text(strip=True) for f in format_els) if format_els else ""

        if not title:
            return None

        return ScrapedRecord(
            title=title,
            authors=authors,
            record_url=record_url,
            fulltext_url=fulltext_url,
            format_type=format_type,
        )
    except Exception:
        logger.exception("Error parsing record")
        return None


async def _fetch_page(
    client: httpx.AsyncClient,
    url: str,
    timeout: int = 30,
) -> str | None:
    try:
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36",
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
            "Accept-Language": "es-ES,es;q=0.9,en;q=0.8",
        }
        response = await client.get(url, headers=headers, timeout=timeout, follow_redirects=True)
        response.raise_for_status()
        return response.text
    except httpx.HTTPStatusError as e:
        logger.error("HTTP error %s for %s: %s", e.response.status_code, url, e)
        return None
    except Exception:
        logger.exception("Error fetching %s", url)
        return None


async def scrape_target(
    config: AppConfig,
    target: TargetConfig,
    query: str,
    timeout: int = 30,
) -> TargetResult:
    result = TargetResult(target_id=target.id, target_name=target.name)

    pages_needed = (config.max_results_per_target + config.results_per_page - 1) // config.results_per_page

    async with httpx.AsyncClient() as client:
        for page in range(1, pages_needed + 1):
            url = urljoin(config.base_url, target.url_template.format(query=query, page=page))
            logger.info("Scraping %s page %s: %s", target.name, page, url)

            html = await _fetch_page(client, url, timeout)

            if html is None:
                if page == 1:
                    result.error = f"No se pudo conectar a {target.name}"
                break

            soup = BeautifulSoup(html, "html.parser")
            record_elements = soup.select(target.selectors["container"])

            if not record_elements:
                if page == 1:
                    result.error = f"No se encontraron resultados en {target.name}"
                break

            for record_el in record_elements:
                if len(result.records) >= config.max_results_per_target:
                    break
                parsed = _parse_record(record_el, target.selectors, config.base_url)
                if parsed:
                    result.records.append(parsed)

            if len(record_elements) < config.results_per_page:
                break

            await asyncio.sleep(0.3)

    return result


async def scrape_all_targets(
    config: AppConfig,
    query: str,
    timeout: int = 30,
) -> list[TargetResult]:
    tasks = [scrape_target(config, target, query, timeout) for target in config.targets]
    results = await asyncio.gather(*tasks)
    return list(results)
