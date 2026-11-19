#!/usr/bin/env python3
"""PartsManager: один вход, каталоги ALLBEST и PM PRO на отдельных листах."""

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
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from playwright.async_api import TimeoutError as PlaywrightTimeout
from playwright.async_api import async_playwright

CHROME = "/opt/pw-browsers/chromium_headless_shell-1248/chrome-headless-shell-linux64/chrome-headless-shell"
HOME = "https://partsmanager.ru/"
STEP = 100
COLUMNS = [
    "Артикул_каталога", "Название_каталога", "Бренд_каталога",
    "Срок_каталога", "Наличие_каталога", "Цена_каталога", "Комментарий_каталога", "URL_каталога",
    "Бренд", "Артикул_предложения", "Название_предложения",
    "Цена", "Наличие", "Цена_текст", "Наличие_текст",
]
KNOWN = [
    {"name": "ALLBEST", "url": "https://partsmanager.ru/mycatalog/john-deere-allbest/john-deere-allbest?start={}"},
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


def sheet_name(name: str, used: set[str]) -> str:
    clean = re.sub(r"[\[\]:*?/\\]", " ", name).strip()[:31] or "Каталог"
    base, number = clean, 2
    while clean in used:
        suffix = f" {number}"
        clean = (base[: 31 - len(suffix)] + suffix).strip()
        number += 1
    used.add(clean)
    return clean


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


async def page_text(page) -> str:
    try:
        return (await page.content()).lower()
    except Exception:
        return ""


async def is_challenge(page) -> bool:
    html = await page_text(page)
    return any(marker in html for marker in ("captcha", "recaptcha", "капча", "я не робот", "проверка безопасности", "security check", "cloudflare"))


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


async def try_login(page, login: str, password: str, stop) -> None:
    try:
        log("Открываю PartsManager")
        await page.goto(HOME, timeout=90000, wait_until="domcontentloaded")
        await page.bring_to_front()
        await asyncio.sleep(1.5)
        if await is_challenge(page):
            if not await wait_user(page, stop, "Сайт просит проверку перед входом."):
                return
        html = await page_text(page)
        if "выход" in html or "личный кабинет" in html:
            log("Уже авторизован")
        elif login and password:
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
        html = await page_text(page)
        if "выход" in html or "личный кабинет" in html:
            log("Вход выполнен")
        elif not await wait_user(page, stop, "Войди в PartsManager, если ещё не вошёл."):
            return
        if stop.is_set():
            return
        await wait_user(page, stop, "Пауза после входа в кабинет.")
    except Exception as error:
        if stop.is_set() or "has been closed" in str(error):
            stop.set()
            return
        log_error(f"Вход не удался: {error}")
        await wait_user(page, stop, "Вход не прошёл. Войди вручную.")


async def discover_catalogs(page) -> list[dict]:
    found = {item["url"]: item for item in KNOWN}
    links = await page.evaluate(
        """() => Array.from(document.querySelectorAll("a[href*='mycatalog']")).map((link) => ({
            href: link.href,
            text: (link.textContent || "").replace(/\\s+/g, " ").trim()
        }))"""
    )
    for link in links or []:
        href = (link.get("href") or "").split("?")[0].rstrip("/")
        text = (link.get("text") or "").lower()
        if not href:
            continue
        low = href.lower()
        if "allbest" in low:
            name = "ALLBEST"
        elif "pmpro" in low or "pm-pro" in low or "pm_pro" in low or "pm pro" in text:
            name = "PMPRO"
        else:
            continue
        found[href + "?start={}"] = {"name": name, "url": href + "?start={}"}
    catalogs = list(found.values())
    if not any(item["name"] == "PMPRO" for item in catalogs):
        log("Ссылку PM PRO на странице не нашёл. Если каталог другой — скажи адрес, добавлю.")
    return catalogs


async def catalog_cards(page) -> list[dict]:
    raw = await page.evaluate(
        """() => Array.from(document.querySelectorAll("li.item")).map((card) => {
            const value = (selector) => {
                const node = card.querySelector(selector);
                return node ? (node.value || node.textContent || "").trim() : "";
            };
            const link = card.querySelector(".priceButton a, a[href*='/parts/']");
            return {
                sku: value("input[name='number']"),
                brand: value("input[name='brand']"),
                name: value(".articleDesc a, .articleDesc"),
                href: link ? link.getAttribute("href") || "" : "",
                text: (card.innerText || "").replace(/\\s+/g, " ").trim()
            };
        })"""
    )
    items = []
    for card in raw or []:
        href = card.get("href") or ""
        if href.startswith("/"):
            href = "https://partsmanager.ru" + href
        if not card.get("sku") or not href:
            continue
        text = card.get("text") or ""
        items.append({
            "Артикул_каталога": card.get("sku") or "",
            "Название_каталога": re.sub(r"\s+", " ", card.get("name") or "").strip(),
            "Бренд_каталога": card.get("brand") or "",
            "Срок_каталога": "",
            "Наличие_каталога": "",
            "Цена_каталога": parse_price(text),
            "Комментарий_каталога": "",
            "URL_каталога": href,
        })
    return items


async def offers_from_page(page) -> list[dict]:
    raw = await page.evaluate(
        """() => Array.from(document.querySelectorAll("tr.resultTr2")).map((row) => ({
            price: row.getAttribute("data-output-price") || "",
            stock: row.getAttribute("data-availability") || "",
            brand: (row.querySelector(".resultBrand .favorite-brand") || {}).innerText || "",
            sku: (row.querySelector(".resultPartCode a") || {}).innerText || "",
            name: (row.querySelector(".resultDescription") || {}).innerText || "",
            stockText: (row.querySelector(".resultAvailability span") || {}).innerText || "",
            priceText: (row.querySelector(".resultPrice") || {}).innerText || ""
        }))"""
    )
    offers = []
    for row in raw or []:
        offers.append({
            "Бренд": re.sub(r"\s+", " ", row.get("brand") or "").strip(),
            "Артикул_предложения": re.sub(r"\s+", " ", row.get("sku") or "").strip(),
            "Название_предложения": re.sub(r"\s+", " ", row.get("name") or "").strip(),
            "Цена": parse_price(row.get("price")),
            "Наличие": int(row["stock"]) if str(row.get("stock") or "").isdigit() else None,
            "Цена_текст": re.sub(r"\s+", " ", row.get("priceText") or "").strip(),
            "Наличие_текст": re.sub(r"\s+", " ", row.get("stockText") or "").strip(),
        })
    return offers


def write_xlsx(path: Path, by_catalog: dict[str, list[dict]]) -> None:
    book = Workbook()
    summary = book.active
    summary.title = "Сводка"
    header_fill = PatternFill("solid", fgColor="1E40AF")
    header_font = Font(bold=True, color="FFFFFF")
    colors = ["FFFFFF", "F2F6FC", "FFF9E6", "F3FBEF", "FDF2F7", "F2F2F2"]
    fill_min = PatternFill("solid", fgColor="C6EFCE")
    fill_max = PatternFill("solid", fgColor="FFC7CE")
    thin = Border(
        left=Side(style="thin", color="D0D0D0"),
        right=Side(style="thin", color="D0D0D0"),
        top=Side(style="thin", color="D0D0D0"),
        bottom=Side(style="thin", color="D0D0D0"),
    )
    used = {"Сводка"}
    summary.append(["Каталог", "Строк", "Артикулов"])
    for cell in summary[1]:
        cell.fill = header_fill
        cell.font = header_font
    for name, rows in by_catalog.items():
        sheet = book.create_sheet(sheet_name(name, used))
        sheet.append(COLUMNS)
        for cell in sheet[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", wrap_text=True)
        sheet.row_dimensions[1].height = 32
        previous, color = None, -1
        for row in rows:
            sheet.append([row.get(column) for column in COLUMNS])
            key = row.get("Артикул_каталога")
            if key != previous:
                color = (color + 1) % len(colors)
                previous = key
            fill = PatternFill("solid", fgColor=colors[color])
            for cell in sheet[sheet.max_row]:
                cell.fill = fill
                cell.border = thin
                cell.alignment = Alignment(vertical="top", wrap_text=True)
        price_col = COLUMNS.index("Цена") + 1
        sku_col = COLUMNS.index("Артикул_каталога") + 1
        lows, highs = {}, {}
        for index in range(2, sheet.max_row + 1):
            sku = sheet.cell(index, sku_col).value
            price = sheet.cell(index, price_col).value
            if isinstance(price, (int, float)):
                sheet.cell(index, price_col).number_format = '#,##0.00" ₽"'
                lows[sku] = min(lows.get(sku, price), price)
                highs[sku] = max(highs.get(sku, price), price)
        for index in range(2, sheet.max_row + 1):
            sku = sheet.cell(index, sku_col).value
            price = sheet.cell(index, price_col).value
            if not isinstance(price, (int, float)) or lows.get(sku) == highs.get(sku):
                continue
            if price == lows.get(sku):
                sheet.cell(index, price_col).fill = fill_min
            elif price == highs.get(sku):
                sheet.cell(index, price_col).fill = fill_max
        for column in range(1, len(COLUMNS) + 1):
            sheet.column_dimensions[get_column_letter(column)].width = 22
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = sheet.dimensions
        summary.append([name, len(rows), len({row.get("Артикул_каталога") for row in rows})])
    summary.append(["Итого", sum(len(rows) for rows in by_catalog.values()), ""])
    for column in range(1, 4):
        summary.column_dimensions[get_column_letter(column)].width = 24
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
        await asyncio.sleep(random.uniform(1.2, 2.0))
        if await is_challenge(page):
            if not await wait_user(page, stop, f"Капча в {catalog['name']}."):
                break
            await page.goto(url, timeout=90000, wait_until="domcontentloaded")
        try:
            await page.wait_for_selector("li.item", timeout=10000)
        except PlaywrightTimeout:
            log(f"{catalog['name']}: карточек больше нет")
            break
        cards = await catalog_cards(page)
        log(f"{catalog['name']}: карточек {len(cards)}")
        if not cards:
            break
        for card in cards:
            if stop.is_set() or not await honor_pause(page, stop):
                return rows
            if card["URL_каталога"] in seen:
                continue
            log(f"{catalog['name']} {card['Артикул_каталога']}")
            try:
                await page.goto(card["URL_каталога"], timeout=90000, wait_until="domcontentloaded")
                await asyncio.sleep(random.uniform(0.8, 1.4))
            except Exception as error:
                log(f"Карточка не открылась: {error}")
                continue
            if await is_challenge(page):
                if not await wait_user(page, stop, f"Капча на {card['Артикул_каталога']}."):
                    return rows
            offers = await offers_from_page(page)
            if not offers:
                offers = [{}]
            for offer in offers:
                item = dict(card)
                item.update(offer)
                rows.append(item)
            seen.add(card["URL_каталога"])
            log(f"предложений {len(offers)}")
        await asyncio.sleep(random.uniform(1.0, 1.8))
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
    by_catalog = saved.get("rows", {})
    seen = set(saved.get("seen", []))
    if by_catalog:
        log(f"Продолжаю: каталогов {len(by_catalog)}")

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
    launch = {"headless": not args.show, "slow_mo": 200 if args.show else 0, "args": ["--disable-blink-features=AutomationControlled", "--no-sandbox"]}
    if not args.show and chrome and Path(chrome).is_file():
        launch["executable_path"] = chrome
    if args.show:
        log("Браузер откроется отдельным окном")

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
            if by_catalog.get(catalog["name"]):
                log(f"{catalog['name']} уже есть, дополняю")
            rows = await scrape_catalog(page, catalog, limit, seen, stop)
            by_catalog.setdefault(catalog["name"], []).extend(rows)
            progress_path.write_text(json.dumps({"rows": by_catalog, "seen": list(seen)}, ensure_ascii=False), encoding="utf-8")
            write_xlsx(out, by_catalog)
            log(f"ROWS {sum(len(items) for items in by_catalog.values())}")
        await browser.close()
    write_xlsx(out, by_catalog)
    log(f"ROWS {sum(len(items) for items in by_catalog.values())}")
    log("DONE")


if __name__ == "__main__":
    try:
        asyncio.run(main())
    except Exception as error:
        log_error(f"Парсер остановился: {error}")
        sys.exit(1)
