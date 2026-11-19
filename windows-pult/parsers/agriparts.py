#!/usr/bin/env python3
"""AgriParts: ищет артикулы из Книга3.xlsx на agriparts.su."""

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

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from playwright.async_api import async_playwright

CHROME = "/opt/pw-browsers/chromium_headless_shell-1248/chrome-headless-shell-linux64/chrome-headless-shell"
HOME = "https://agriparts.su"
COLUMNS = [
    "Артикул_запроса", "Статус", "Найденный_оригинал", "Найденный_аналог",
    "Бренд", "Название", "Цена", "Наличие", "Акция", "Все_аналоги", "Все_оригиналы",
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


def load_articles(path: Path) -> list[str]:
    book = load_workbook(path, read_only=True, data_only=True)
    sheet = book.active
    seen, articles = set(), []
    for row in sheet.iter_rows(values_only=True):
        value = "" if row[0] is None else str(row[0]).strip()
        if not value or value.lower() in {"nan", "артикул", "артикул_запроса"}:
            continue
        if value not in seen:
            seen.add(value)
            articles.append(value)
    return articles


async def page_text(page) -> str:
    try:
        return (await page.content()).lower()
    except Exception:
        return ""


async def is_challenge(page) -> bool:
    html = await page_text(page)
    return any(marker in html for marker in ("captcha", "recaptcha", "капча", "я не робот", "security check"))


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
        log("Открываю AgriParts")
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
        if not await wait_user(page, stop, "Пауза после входа. Войди в AgriParts, если ещё не вошёл."):
            return
    except Exception as error:
        if stop.is_set() or "has been closed" in str(error):
            stop.set()
            return
        log_error(f"Вход не удался: {error}")
        await wait_user(page, stop, "Вход не прошёл. Войди вручную.")


async def search_article(page, article: str) -> dict:
    data = await page.evaluate(
        """() => Array.from(document.querySelectorAll(".r-product-card")).map((card) => {
            const text = (selector) => {
                const node = card.querySelector(selector);
                return node ? node.innerText.trim() : "";
            };
            return {
                type: text(".r-product-card__badges .r-badge span"),
                brand: text(".r-product-card__brand .r-badge span"),
                article: text(".r-product-card__article"),
                title: text(".r-product-card__title"),
                price: text(".r-product-card__price .price span").replace(/\\u00a0/g, " "),
                stock: text(".r-product-card__availability button").split("\\n")[0],
                sale: card.querySelector(".r-product-card__sale") ? "Да" : ""
            };
        })"""
    )
    found = []
    analog = original = ""
    for card in data or []:
        if not card.get("article"):
            continue
        item = {
            "Тип": card.get("type") or "",
            "Бренд": card.get("brand") or "",
            "Артикул": card.get("article") or "",
            "Название": card.get("title") or "",
            "Цена": card.get("price") or "",
            "Наличие": card.get("stock") or "",
            "Акция": card.get("sale") or "",
        }
        found.append(item)
        kind = item["Тип"].lower()
        if "аналог" in kind and not analog:
            analog = item["Артикул"]
        elif "оригинал" in kind and not original:
            original = item["Артикул"]
    return {"query": article, "found": found, "analog_article": analog, "original_article": original, "found_count": len(found)}


def in_stock(offer: dict) -> bool:
    stock = (offer.get("Наличие") or "").strip().lower()
    if not stock or "под заказ" in stock:
        return False
    match = re.search(r":\s*(\d+)", stock)
    return not (match and int(match.group(1)) == 0)


def best_offer(offers: list[dict]):
    available = [offer for offer in offers if in_stock(offer)]
    return available[0] if available else (offers[0] if offers else None)


def format_offers(offers: list[dict]) -> str:
    parts = []
    for offer in offers:
        parts.append(f"{offer.get('Артикул', '')} ({offer.get('Бренд', '')}) — {offer.get('Цена', '')} / {offer.get('Наличие', '')}")
    return "; ".join(parts)


def to_row(item: dict) -> dict:
    found = item.get("found", [])
    analogs = [offer for offer in found if "аналог" in (offer.get("Тип") or "").lower()]
    originals = [offer for offer in found if "оригинал" in (offer.get("Тип") or "").lower()]
    analog = best_offer(analogs)
    original = best_offer(originals)
    if analog and in_stock(analog):
        main, status = analog, "В наличии (аналог)"
    elif original and in_stock(original):
        main, status = original, "В наличии (оригинал)"
    elif analog:
        main, status = analog, "Нет в наличии"
    elif original:
        main, status = original, "Нет в наличии"
    else:
        main, status = None, "Не найдено"
    return {
        "Артикул_запроса": item["query"],
        "Статус": status,
        "Найденный_оригинал": originals[0]["Артикул"] if originals else "",
        "Найденный_аналог": analogs[0]["Артикул"] if analogs else "",
        "Бренд": (main or {}).get("Бренд", ""),
        "Название": (main or {}).get("Название", ""),
        "Цена": (main or {}).get("Цена", ""),
        "Наличие": (main or {}).get("Наличие", ""),
        "Акция": (main or {}).get("Акция", ""),
        "Все_аналоги": format_offers(analogs),
        "Все_оригиналы": format_offers(originals),
    }


def write_xlsx(path: Path, items: list[dict]) -> None:
    book = Workbook()
    sheet = book.active
    sheet.title = "AGRIPARTS"
    header_fill = PatternFill("solid", fgColor="1E40AF")
    header_font = Font(bold=True, color="FFFFFF")
    fills = {
        "В наличии": PatternFill("solid", fgColor="C6EFCE"),
        "Нет в наличии": PatternFill("solid", fgColor="FFC7CE"),
        "Не найдено": PatternFill("solid", fgColor="E7E6E6"),
    }
    thin = Border(
        left=Side(style="thin", color="D0D0D0"),
        right=Side(style="thin", color="D0D0D0"),
        top=Side(style="thin", color="D0D0D0"),
        bottom=Side(style="thin", color="D0D0D0"),
    )
    sheet.append(COLUMNS)
    for cell in sheet[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
    sheet.row_dimensions[1].height = 32
    for item in items:
        row = to_row(item)
        sheet.append([row.get(column) for column in COLUMNS])
        fill = next((value for key, value in fills.items() if key in row["Статус"]), None)
        for cell in sheet[sheet.max_row]:
            cell.border = thin
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            if fill:
                cell.fill = fill
    for column in range(1, len(COLUMNS) + 1):
        sheet.column_dimensions[get_column_letter(column)].width = 60 if COLUMNS[column - 1].startswith("Все_") else 24
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(".tmp.xlsx")
    book.save(tmp)
    tmp.replace(path)


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
    LOG_FILE = out.with_name("agriparts.log")
    CONTINUE = out.with_name("continue.flag")
    PAUSE = out.with_name("pause.flag")
    for flag in (CONTINUE, PAUSE):
        if flag.exists():
            flag.unlink()
    source = Path(__file__).with_name("articles.xlsx")
    if not source.is_file():
        source = Path(__file__).with_name("Книга3.xlsx")
    if not source.is_file():
        log_error("Нет файла articles.xlsx рядом с парсером.")
        return
    articles = load_articles(source)
    if limit:
        articles = articles[:limit]
    progress_path = out.with_name("progress.json")
    saved = json.loads(progress_path.read_text(encoding="utf-8")) if progress_path.is_file() else {}
    results = saved.get("results", [])
    done = {item["query"] for item in results}
    log(f"Артикулов: {len(articles)}. Уже проверено: {len(done)}")

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
    if args.show:
        log("Браузер откроется отдельным окном")

    async with async_playwright() as pw:
        browser = await pw.chromium.launch(**launch)
        context = await browser.new_context(locale="ru-RU", viewport={"width": 1400, "height": 900})
        page = await context.new_page()
        await try_login(page, os.environ.get("PULT_LOGIN", "").strip(), os.environ.get("PULT_PASSWORD", ""), stop)
        for index, article in enumerate(articles, start=1):
            if stop.is_set() or not await honor_pause(page, stop):
                break
            if article in done:
                continue
            log(f"[{index}/{len(articles)}] {article}")
            try:
                await page.goto(f"{HOME}/search/?q={article}", timeout=90000, wait_until="domcontentloaded")
                await asyncio.sleep(random.uniform(1.2, 1.8))
                if await is_challenge(page):
                    if not await wait_user(page, stop, f"Капча на {article}."):
                        break
                parsed = await search_article(page, article)
            except Exception as error:
                log(f"Ошибка {article}: {error}")
                parsed = {"query": article, "found": [], "analog_article": "", "original_article": "", "found_count": 0}
            results.append(parsed)
            done.add(article)
            status = to_row(parsed)["Статус"]
            log(f"{status}, карточек {parsed['found_count']}")
            if index % 10 == 0:
                progress_path.write_text(json.dumps({"results": results}, ensure_ascii=False), encoding="utf-8")
                write_xlsx(out, results)
                log(f"ROWS {len(results)}")
            await asyncio.sleep(random.uniform(1.4, 2.2))
        await browser.close()
    write_xlsx(out, results)
    progress_path.write_text(json.dumps({"results": results}, ensure_ascii=False), encoding="utf-8")
    log(f"ROWS {len(results)}")
    log("DONE")


if __name__ == "__main__":
    try:
        asyncio.run(main())
    except Exception as error:
        log_error(f"Парсер остановился: {error}")
        sys.exit(1)
