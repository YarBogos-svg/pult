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
import urllib.parse
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
KNOWN = [
    {"name": "ALLBEST", "url": "https://partsmanager.ru/mycatalog/john-deere-allbest/john-deere-allbest?start={}"},
    {"name": "PMPRO", "url": "https://partsmanager.ru/mycatalog/john-deere-pmpro/john-deere-pmpro?start={}"},
    {"name": "PMPRO", "url": "https://partsmanager.ru/mycatalog/pmpro/pmpro?start={}"},
    {"name": "PMPRO", "url": "https://partsmanager.ru/mycatalog/pm-pro/pm-pro?start={}"},
    {"name": "PMPRO", "url": "https://partsmanager.ru/mycatalog/john-deere-pm-pro/john-deere-pm-pro?start={}"},
]
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
    text = urllib.parse.unquote((value or "").lower())
    text = re.sub(r"[^a-z0-9]+", "", text)
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


async def open_url(page, url: str, stop, what: str) -> bool:
    for attempt in range(3):
        if stop.is_set():
            return False
        try:
            await page.goto(url, timeout=90000, wait_until="domcontentloaded")
            return True
        except Exception as error:
            log(f"{what} не открылась, попытка {attempt + 1}: {str(error)[:80]}")
            await asyncio.sleep(2 + attempt)
    log(f"Пропускаю {what}: сайт не ответил")
    return False


async def try_login(page, login: str, password: str, stop) -> None:
    try:
        log("Открываю PartsManager")
        await page.goto(HOME, timeout=90000, wait_until="domcontentloaded")
        await page.bring_to_front()
        await asyncio.sleep(1.5)
        if await is_challenge(page):
            if not await wait_user(page, stop, "Сайт просит проверку перед входом."):
                return
        if login and password:
            for selector in ("text=Вход", "text=Войти", "a[href*='login']"):
                locator = page.locator(selector).first
                if await locator.count() and await locator.is_visible():
                    await locator.click()
                    await page.wait_for_timeout(800)
                    break
            secret = page.locator("input[type='password']").first
            if await secret.count():
                form = page.locator("form").filter(has=page.locator("input[type='password']")).first
                scope = form if await form.count() else page
                user = scope.locator("input[type='email'], input[name*='mail'], input[name*='login'], input[type='text']").first
                if await user.count():
                    await user.fill(login)
                    await secret.fill(password)
                    submit = scope.locator("button[type='submit'], button:has-text('Войти')").first
                    if await submit.count():
                        await submit.click()
                    else:
                        await secret.press("Enter")
                    await page.wait_for_timeout(2500)
        if stop.is_set():
            return
        if await is_challenge(page):
            if not await wait_user(page, stop, "После входа сайт просит проверку."):
                return
        if not await wait_user(page, stop, "Пауза после входа. Войди в PartsManager, если ещё не вошёл."):
            return
    except Exception as error:
        if stop.is_set() or "has been closed" in str(error):
            stop.set()
            return
        log_error(f"Вход не удался: {error}")
        await wait_user(page, stop, "Вход не прошёл. Войди вручную.")


async def open_offers(page) -> None:
    for text in ("Посмотреть другие предложения", "Показать ещё", "Показать все", "Показать больше"):
        button = page.locator(f"text={text}").first
        try:
            if await button.count() and await button.is_visible():
                await button.click()
                await page.wait_for_timeout(1200)
                log(f"Нажал «{text}»")
                return
        except Exception:
            continue


async def guard(page, stop, reason: str) -> bool:
    if await is_challenge(page):
        return await wait_user(page, stop, reason)
    return not stop.is_set()


async def discover_catalogs(page) -> list[dict]:
    found = list(KNOWN)
    await page.goto(HOME, timeout=90000, wait_until="domcontentloaded")
    await asyncio.sleep(1.5)
    for label in ("JOHN DEERE", "PM PRO", "PMPRO", "ALLBEST"):
        locator = page.locator(f"text={label}").first
        try:
            if await locator.count() and await locator.is_visible():
                await locator.hover()
                await asyncio.sleep(0.4)
        except Exception:
            pass
    links = await page.evaluate(
        """() => Array.from(document.querySelectorAll("a[href]")).map((link) => ({
            href: link.href,
            text: (link.textContent || "").replace(/\\s+/g, " ").trim()
        }))"""
    )
    for link in links or []:
        href = (link.get("href") or "").split("?")[0].rstrip("/")
        text = link.get("text") or ""
        key = brand_key(href + " " + text)
        low = href.lower()
        if not key or "mycatalog" not in low:
            continue
        found.append({"name": key, "url": href + "?start={}"})
        log(f"Нашёл каталог {key}: {href}")
    unique, seen = [], set()
    for item in found:
        if item["url"] in seen:
            continue
        seen.add(item["url"])
        unique.append(item)
    return unique


async def catalog_items(page, catalog_name: str) -> list[dict]:
    raw = await page.evaluate(
        """() => {
            const cards = Array.from(document.querySelectorAll("li.item, .catalog-item, .product-item, .product-card, [class*='product'], [class*='item']"));
            const buttons = Array.from(document.querySelectorAll("a")).filter((link) =>
                /цены и аналоги|посмотреть/i.test(link.textContent || "")
            );
            const sources = buttons.length ? buttons.map((button) => button.closest("li, .product-card, [class*='product'], [class*='item']") || button.parentElement) : cards;
            return sources.filter(Boolean).map((card) => {
                const value = (selector) => {
                    const node = card.querySelector(selector);
                    return node ? (node.value || node.textContent || "").trim() : "";
                };
                const link = card.querySelector("a[href*='/parts/']") || card.querySelector("a");
                const article = Array.from(card.querySelectorAll("a")).map((node) => node.textContent.trim()).find((text) => /[A-Z0-9]{4,}/.test(text)) || "";
                return {
                    sku: value("input[name='number']") || article,
                    brand: value("input[name='brand']"),
                    name: value(".articleDesc a, .articleDesc") || card.innerText.split("\\n").slice(0, 3).join(" "),
                    text: (card.innerText || "").replace(/\\s+/g, " ").trim(),
                    href: link ? link.href || link.getAttribute("href") || "" : ""
                };
            });
        }"""
    )
    items = []
    seen = set()
    for card in raw or []:
        href = card.get("href") or ""
        if href.startswith("/"):
            href = "https://partsmanager.ru" + href
        if not href or href in seen:
            continue
        if "/parts/" not in href and "mycatalog" in href:
            continue
        seen.add(href)
        brand = brand_key(" ".join([card.get("brand") or "", href, card.get("name") or "", card.get("text") or ""])) or catalog_name
        sku = card.get("sku") or href.rstrip("/").split("/")[-1]
        items.append({"sku": sku, "brand": brand, "name": re.sub(r"\s+", " ", card.get("name") or "").strip(), "url": href})
    return items


async def offers_from_page(page, article: str, url: str) -> list[dict]:
    raw = await page.evaluate(
        """() => {
            const main = {
                brand: (document.querySelector("h1, .product-title, .article-title") || {}).innerText || "",
                code: "",
                name: "",
                stock: "",
                price: ""
            };
            const header = document.querySelector(".product-info, .article-info, .product-card") || document.body;
            const stockEl = header.querySelector("[class*='availability'], [class*='stock'], [class*='nalich']");
            const priceEl = header.querySelector("[class*='price']");
            if (stockEl) main.stock = stockEl.innerText.trim();
            if (priceEl) main.price = priceEl.innerText.trim();
            const rows = Array.from(document.querySelectorAll("tr.resultTr2, table tr")).filter((row) =>
                row.querySelector(".resultBrand, .resultPartCode, td")
            );
            const table = rows.map((row) => {
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
            if (!table.length && main.brand) table.unshift(main);
            return table;
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
        if not await open_url(page, url, stop, f"{catalog['name']} страница {page_num + 1}"):
            break
        await asyncio.sleep(random.uniform(1.0, 1.6))
        if not await guard(page, stop, f"Капча в каталоге {catalog['name']}."):
            break
        try:
            await page.wait_for_selector("a:has-text('Цены и аналоги'), a[href*='/parts/']", timeout=12000)
        except PlaywrightTimeout:
            log(f"{catalog['name']}: карточек больше нет")
            break
        buttons = page.locator("a:has-text('Цены и аналоги')")
        count = await buttons.count()
        log(f"{catalog['name']}: кнопок «Цены и аналоги» {count}")
        if count == 0:
            items = await catalog_items(page, catalog["name"])
            log(f"{catalog['name']}: карточек по ссылкам {len(items)}")
            if not items:
                break
            for item in items:
                if stop.is_set() or not await honor_pause(page, stop):
                    return rows
                if item["url"] in seen:
                    continue
                log(f"Открываю карточку {item['sku']}")
                try:
                    await page.goto(item["url"], timeout=90000, wait_until="domcontentloaded")
                    await asyncio.sleep(random.uniform(0.8, 1.4))
                except Exception as error:
                    log(f"Карточка не открылась: {error}")
                    continue
                if not await guard(page, stop, f"Капча на {item['sku']}."):
                    return rows
                await open_offers(page)
                offers = await offers_from_page(page, item["sku"], item["url"])
                if not offers:
                    offers = [{
                        "Артикул": item["sku"], "Номер": 1, "Бренд": item["brand"], "Код_детали": item["sku"],
                        "Описание": item["name"], "Наличие": "", "Цена": None, "URL": item["url"],
                    }]
                rows.extend(offers)
                seen.add(item["url"])
                log(f"предложений {len(offers)}")
            continue
        for index in range(count):
            if stop.is_set() or not await honor_pause(page, stop):
                return rows
            button = page.locator("a:has-text('Цены и аналоги')").nth(index)
            try:
                href = await button.get_attribute("href") or ""
                if href.startswith("/"):
                    href = "https://partsmanager.ru" + href
                sku = href.rstrip("/").split("/")[-1] or f"card-{index + 1}"
                if href and href in seen:
                    continue
                log(f"Открываю карточку {sku}")
                if href:
                    if not await open_url(page, href, stop, f"карточка {sku}"):
                        continue
                else:
                    await button.click()
                    await page.wait_for_timeout(1200)
                await asyncio.sleep(random.uniform(0.8, 1.4))
            except Exception as error:
                log(f"Карточка не открылась: {error}")
                continue
            if not await guard(page, stop, f"Капча на {sku}."):
                return rows
            await open_offers(page)
            offers = await offers_from_page(page, sku, page.url)
            if not offers:
                offers = [{
                    "Артикул": sku, "Номер": 1, "Бренд": catalog["name"], "Код_детали": sku,
                    "Описание": "", "Наличие": "", "Цена": None, "URL": page.url,
                }]
            rows.extend(offers)
            if href:
                seen.add(href)
            log(f"предложений {len(offers)}")
            await open_url(page, url, stop, "возврат в каталог")
            await asyncio.sleep(0.6)
        await asyncio.sleep(random.uniform(0.4, 0.8))
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
    if progress_path.is_file():
        progress_path.unlink()
        log("Старый прогресс сброшен, собираю заново")
    by_brand = {}
    seen = set()

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
        await try_login(page, os.environ.get("PULT_LOGIN", "").strip(), os.environ.get("PULT_PASSWORD", ""), stop)
        if stop.is_set():
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
