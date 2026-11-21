#!/usr/bin/env python3
"""PartsManager: John Deere, бренды ALLBEST и PM PRO, предложения пронумерованы."""

import argparse
import asyncio
import json
import os
import random
import re
import signal
import sys
import traceback
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from playwright.async_api import TimeoutError as PlaywrightTimeout
from playwright.async_api import async_playwright

CHROME = "/opt/pw-browsers/chromium_headless_shell-1248/chrome-headless-shell-linux64/chrome-headless-shell"
HOME = "https://partsmanager.ru/"
STEP = 100
COLUMNS = ["Артикул", "Номер", "Бренд", "Код_детали", "Описание", "Наличие", "Цена", "URL"]
KNOWN = {
    "ALLBEST": "https://partsmanager.ru/mycatalog/john-deere-allbest/john-deere-allbest?start={}",
}
LOG_FILE: Path | None = None
CONTINUE: Path | None = None
PAUSE: Path | None = None


def log(message: str) -> None:
    sys.stdout.buffer.write((message + "\n").encode("utf-8", "replace"))
    sys.stdout.buffer.flush()
    if LOG_FILE is None:
        return
    try:
        with LOG_FILE.open("a", encoding="utf-8") as handle:
            handle.write(f"{datetime.now():%H:%M:%S} {message}\n")
    except OSError:
        pass


def log_error(message: str) -> None:
    log(f"ОШИБКА {message}")
    if LOG_FILE is None:
        return
    trace = traceback.format_exc().rstrip()
    if trace and trace != "NoneType: None":
        try:
            with LOG_FILE.open("a", encoding="utf-8") as handle:
                handle.write(trace + "\n")
        except OSError:
            pass


def parse_price(value):
    if isinstance(value, (int, float)):
        return float(value)
    if not value:
        return None
    cleaned = re.sub(r"[^\d.,]", "", str(value).replace(" ", "").replace("\xa0", "")).replace(",", ".")
    try:
        return float(cleaned) if cleaned else None
    except ValueError:
        return None


def brand_key(value: str) -> str:
    text = re.sub(r"[^a-z0-9]+", "", (value or "").lower())
    if "allbest" in text:
        return "ALLBEST"
    if "pmpro" in text or text == "pm":
        return "PMPRO"
    return ""


async def is_challenge(page) -> bool:
    try:
        return bool(await page.evaluate(
            """() => {
                const text = (document.body && document.body.innerText || "").toLowerCase();
                if (text.includes("access restricted") || text.includes("security check") || text.includes("проверка безопасности")) return true;
                const selectors = ["#captcha", ".captcha", "img[src*='captcha']", "input[name*='captcha']", "iframe[src*='recaptcha']", "iframe[src*='hcaptcha']", "iframe[src*='turnstile']"];
                return selectors.some((selector) => Array.from(document.querySelectorAll(selector)).some((node) => {
                    const style = getComputedStyle(node);
                    const box = node.getBoundingClientRect();
                    return style.display !== "none" && style.visibility !== "hidden" && box.width > 20 && box.height > 20;
                }));
            }"""
        ))
    except Exception:
        return False


async def wait_user(page, stop, reason: str) -> bool:
    log(f"ПАУЗА {reason} Реши это в окне браузера и нажми «Продолжить».")
    try:
        await page.bring_to_front()
    except Exception:
        pass
    while not stop.is_set():
        if CONTINUE is not None and CONTINUE.exists():
            CONTINUE.unlink(missing_ok=True)
            log("Продолжаю")
            return True
        await asyncio.sleep(0.4)
    return False


async def honor_pause(page, stop) -> bool:
    if PAUSE is None or not PAUSE.exists():
        return not stop.is_set()
    PAUSE.unlink(missing_ok=True)
    return await wait_user(page, stop, "Пауза по кнопке.")


async def guard(page, stop, reason: str) -> bool:
    if await is_challenge(page):
        return await wait_user(page, stop, reason)
    return not stop.is_set()


async def discover_catalogs(page) -> list[dict]:
    found = dict(KNOWN)
    links = await page.evaluate(
        """() => Array.from(document.querySelectorAll("a[href]")).map((link) => ({
            href: link.href,
            text: (link.textContent || "").replace(/\\s+/g, " ").trim()
        }))"""
    )
    for link in links or []:
        href = (link.get("href") or "").split("?")[0].rstrip("/")
        text = (link.get("text") or "").lower()
        low = href.lower()
        if "john" not in low and "john" not in text and "deere" not in low and "дир" not in text:
            if "allbest" not in low and "pmpro" not in low and "pm-pro" not in low:
                continue
        key = brand_key(href + " " + text)
        if not key or "mycatalog" not in low and "/parts/" not in low:
            continue
        if "mycatalog" in low:
            found[key] = href + "?start={}"
    catalogs = [{"name": name, "url": url} for name, url in found.items()]
    if "PMPRO" not in found:
        log("Каталог PM PRO по ссылке не нашёл. Если он в другом адресе — скажи его.")
    return catalogs


async def catalog_items(page) -> list[dict]:
    raw = await page.evaluate(
        """() => Array.from(document.querySelectorAll("li.item")).map((card) => {
            const value = (selector) => {
                const node = card.querySelector(selector);
                return node ? (node.value || node.textContent || "").trim() : "";
            };
            const link = card.querySelector(".priceButton a, .articleDesc a[href*='/parts/'], a[href*='/parts/']");
            return {
                sku: value("input[name='number']") || value(".articleDesc a"),
                brand: value("input[name='brand']"),
                name: value(".articleDesc a, .articleDesc"),
                href: link ? link.href || link.getAttribute("href") || "" : ""
            };
        })"""
    )
    items = []
    for card in raw or []:
        href = card.get("href") or ""
        if href.startswith("/"):
            href = "https://partsmanager.ru" + href
        brand = brand_key(card.get("brand") or "")
        if not card.get("sku") or not href or not brand:
            continue
        items.append({"sku": card["sku"], "brand": brand, "name": re.sub(r"\s+", " ", card.get("name") or "").strip(), "url": href})
    return items


async def offers_from_page(page, article: str, url: str) -> list[dict]:
    raw = await page.evaluate(
        """() => {
            const rows = Array.from(document.querySelectorAll("tr.resultTr2, table tr")).filter((row) =>
                row.querySelector(".resultBrand, .resultPartCode, td")
            );
            return rows.map((row) => {
                const text = (selector) => {
                    const node = row.querySelector(selector);
                    return node ? node.innerText.trim() : "";
                };
                const cells = Array.from(row.querySelectorAll("td")).map((cell) => cell.innerText.trim());
                return {
                    brand: text(".resultBrand .favorite-brand, .resultBrand") || cells[1] || "",
                    code: text(".resultPartCode a, .resultPartCode") || cells[2] || "",
                    name: text(".resultDescription") || cells[3] || "",
                    stock: text(".resultAvailability span, .resultAvailability") || cells[4] || "",
                    price: row.getAttribute("data-output-price") || text(".resultPrice") || cells[cells.length - 2] || ""
                };
            }).filter((row) => row.brand || row.code || row.price);
        }"""
    )
    offers = []
    for number, row in enumerate(raw or [], start=1):
        code = re.sub(r"\s+", " ", row.get("code") or "").strip()
        brand = re.sub(r"\s+", " ", row.get("brand") or "").strip()
        if not code and not brand:
            continue
        offers.append({
            "Артикул": article,
            "Номер": number,
            "Бренд": brand,
            "Код_детали": code,
            "Описание": re.sub(r"\s+", " ", row.get("name") or "").strip(),
            "Наличие": re.sub(r"\s+", " ", row.get("stock") or "").strip(),
            "Цена": parse_price(row.get("price")),
            "URL": url,
        })
    return offers


def write_xlsx(path: Path, by_brand: dict[str, list[dict]]) -> None:
    book = Workbook()
    summary = book.active
    summary.title = "Сводка"
    header_fill = PatternFill("solid", fgColor="1E40AF")
    header_font = Font(bold=True, color="FFFFFF")
    group_fill = PatternFill("solid", fgColor="F2F6FC")
    used = {"Сводка"}
    summary.append(["Бренд", "Строк", "Артикулов"])
    for cell in summary[1]:
        cell.fill, cell.font = header_fill, header_font
    for name in ("ALLBEST", "PMPRO"):
        rows = by_brand.get(name, [])
        if not rows:
            continue
        title = name if name not in used else f"{name} 2"
        used.add(title)
        sheet = book.create_sheet(title)
        sheet.append(COLUMNS)
        for cell in sheet[1]:
            cell.fill, cell.font = header_fill, header_font
            cell.alignment = Alignment(horizontal="center", wrap_text=True)
        previous = None
        for row in rows:
            sheet.append([row.get(column) for column in COLUMNS])
            if row.get("Артикул") != previous:
                previous = row.get("Артикул")
            if row.get("Номер", 1) > 1:
                for cell in sheet[sheet.max_row]:
                    cell.fill = group_fill
        for column in range(1, len(COLUMNS) + 1):
            sheet.column_dimensions[get_column_letter(column)].width = 28
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = sheet.dimensions
        summary.append([name, len(rows), len({row.get("Артикул") for row in rows})])
    for column in range(1, 4):
        summary.column_dimensions[get_column_letter(column)].width = 22
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(".tmp.xlsx")
    book.save(tmp)
    tmp.replace(path)


async def scrape_catalog(page, catalog: dict, limit: int | None, seen: set[str], stop) -> list[dict]:
    rows = []
    for page_num in range(limit or 500):
        if stop.is_set() or not await honor_pause(page, stop):
            break
        url = catalog["url"].format(page_num * STEP)
        log(f"{catalog['name']} страница {page_num + 1}")
        try:
            await page.goto(url, timeout=90000, wait_until="domcontentloaded")
        except Exception as error:
            log_error(f"Страница не открылась: {error}")
            break
        await asyncio.sleep(random.uniform(1.0, 1.6))
        if not await guard(page, stop, f"Капча в каталоге {catalog['name']}."):
            break
        try:
            await page.wait_for_selector("li.item", timeout=12000)
        except PlaywrightTimeout:
            log(f"{catalog['name']}: карточек больше нет")
            break
        items = await catalog_items(page)
        items = [item for item in items if item["brand"] == catalog["name"] or catalog["name"] == "ALLBEST"]
        log(f"{catalog['name']}: карточек {len(items)}")
        if not items:
            break
        for item in items:
            if stop.is_set() or not await honor_pause(page, stop):
                return rows
            if item["url"] in seen:
                continue
            log(f"{item['brand']} {item['sku']}")
            try:
                await page.goto(item["url"], timeout=90000, wait_until="domcontentloaded")
                await asyncio.sleep(random.uniform(0.8, 1.4))
            except Exception as error:
                log(f"Карточка не открылась: {error}")
                continue
            if not await guard(page, stop, f"Капча на {item['sku']}."):
                return rows
            offers = await offers_from_page(page, item["sku"], item["url"])
            if not offers:
                offers = [{
                    "Артикул": item["sku"], "Номер": 1, "Бренд": item["brand"], "Код_детали": item["sku"],
                    "Описание": item["name"], "Наличие": "", "Цена": None, "URL": item["url"],
                }]
            rows.extend(offers)
            seen.add(item["url"])
            log(f"предложений {len(offers)}")
        await asyncio.sleep(random.uniform(0.8, 1.4))
    return rows


async def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--pages", type=int, default=0)
    parser.add_argument("--out", required=True)
    parser.add_argument("--show", action="store_true")
    args = parser.parse_args()
    limit = args.pages if args.pages > 0 else None
    out = Path(args.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    global LOG_FILE, CONTINUE, PAUSE
    LOG_FILE = out.with_name("partsmanager.log")
    CONTINUE = out.with_name("continue.flag")
    PAUSE = out.with_name("pause.flag")
    for flag in (CONTINUE, PAUSE):
        if flag.exists():
            flag.unlink()
    progress_path = out.with_name("progress.json")
    saved = json.loads(progress_path.read_text(encoding="utf-8")) if progress_path.is_file() else {}
    by_brand = saved.get("rows", {})
    seen = set(saved.get("seen", []))

    stop = asyncio.Event()
    loop = asyncio.get_running_loop()
    try:
        loop.add_signal_handler(signal.SIGINT, stop.set)
        loop.add_signal_handler(signal.SIGTERM, stop.set)
    except NotImplementedError:
        signal.signal(signal.SIGINT, lambda *_: loop.call_soon_threadsafe(stop.set))

    browsers = Path("/opt/pw-browsers")
    if browsers.is_dir() and "PLAYWRIGHT_BROWSERS_PATH" not in os.environ:
        os.environ["PLAYWRIGHT_BROWSERS_PATH"] = str(browsers)
    chrome = os.environ.get("GNG_CHROME") or CHROME
    launch = {"headless": not args.show, "slow_mo": 150 if args.show else 0, "args": ["--disable-blink-features=AutomationControlled", "--no-sandbox"]}
    if not args.show and chrome and Path(chrome).is_file():
        launch["executable_path"] = chrome

    async with async_playwright() as pw:
        browser = await pw.chromium.launch(**launch)
        context = await browser.new_context(locale="ru-RU", viewport={"width": 1400, "height": 900})
        page = await context.new_page()
        log("Открываю PartsManager")
        try:
            await page.goto(HOME, timeout=90000, wait_until="domcontentloaded")
            await page.bring_to_front()
        except Exception as error:
            log(f"Сайт открылся с ошибкой: {error}")
        if not await wait_user(page, stop, "Пауза в начале. Войди в аккаунт и реши капчу, если она есть."):
            await browser.close()
            return
        catalogs = await discover_catalogs(page)
        log("Каталоги: " + ", ".join(item["name"] for item in catalogs))
        for catalog in catalogs:
            if stop.is_set():
                break
            rows = await scrape_catalog(page, catalog, limit, seen, stop)
            by_brand.setdefault(catalog["name"], []).extend(rows)
            progress_path.write_text(json.dumps({"rows": by_brand, "seen": list(seen)}, ensure_ascii=False), encoding="utf-8")
            write_xlsx(out, by_brand)
            log(f"ROWS {sum(len(items) for items in by_brand.values())}")
        await browser.close()
    write_xlsx(out, by_brand)
    log(f"ROWS {sum(len(items) for items in by_brand.values())}")
    log("DONE")


if __name__ == "__main__":
    try:
        asyncio.run(main())
    except Exception as error:
        log_error(f"Парсер остановился: {error}")
        sys.exit(1)
