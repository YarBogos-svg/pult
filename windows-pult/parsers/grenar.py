#!/usr/bin/env python3
"""Парсер каталога John Deere на grenar.ru с точным количеством через API."""

import argparse
import asyncio
import json
import os
import random
import re
import signal
import sys
import traceback
import urllib.request
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from playwright.async_api import TimeoutError as PlaywrightTimeout
from playwright.async_api import async_playwright

CHROME = "/opt/pw-browsers/chromium_headless_shell-1248/chrome-headless-shell-linux64/chrome-headless-shell"
BASE = "https://grenar.ru/brands/john-deere/"
API_PRODUCT = "https://grenar.ru/api/v1/catalog/products/{}"
HEADERS = ["Артикул", "Название", "Бренд", "Цена", "Цена_старая", "Скидка", "Количество", "Отгрузка"]
LOG_FILE: Path | None = None
CONTINUE: Path | None = None
PAUSE: Path | None = None
CONCURRENCY = 12


def log(message: str) -> None:
    sys.stdout.buffer.write((message + "\n").encode("utf-8", "replace"))
    sys.stdout.buffer.flush()
    if LOG_FILE is None:
        return
    try:
        stamp = datetime.now().strftime("%H:%M:%S")
        with LOG_FILE.open("a", encoding="utf-8") as handle:
            handle.write(f"{stamp} {message}\n")
    except OSError:
        pass


def log_error(message: str) -> None:
    log(f"ОШИБКА {message}")
    if LOG_FILE is None:
        return
    trace = traceback.format_exc().rstrip()
    if not trace or trace == "NoneType: None":
        return
    try:
        with LOG_FILE.open("a", encoding="utf-8") as handle:
            handle.write(trace + "\n")
    except OSError:
        pass


def parse_price(value: str):
    if not value:
        return None
    cleaned = re.sub(r"[^\d.,]", "", value.replace(" ", "").replace("\xa0", "")).replace(",", ".")
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
    if not html:
        return False
    markers = (
        "captcha",
        "recaptcha",
        "hcaptcha",
        "я не робот",
        "подтвердите, что вы не робот",
        "проверка безопасности",
        "капча",
    )
    return any(marker in html for marker in markers)


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
        log("Открываю Grenar")
        await page.goto(BASE, timeout=90000, wait_until="domcontentloaded")
        await page.bring_to_front()
        await asyncio.sleep(1.5)
        if await is_challenge(page):
            if not await wait_user(page, stop, "Сайт просит проверку перед входом."):
                return
        if login and password:
            for selector in ("text=Войти", "text=Вход", "a[href*='login']", "button:has-text('Войти')"):
                locator = page.locator(selector).first
                if await locator.count() and await locator.is_visible():
                    await locator.click()
                    await page.wait_for_timeout(800)
                    break
            secret = page.locator("input[type='password']").first
            if await secret.count():
                form = page.locator("form").filter(has=page.locator("input[type='password']")).first
                scope = form if await form.count() else page
                user = scope.locator(
                    "input[type='email'], input[name*='login'], input[name*='mail'], input[name*='user'], input[type='text']"
                ).first
                if await user.count():
                    await user.fill(login)
                    await secret.fill(password)
                    submit = scope.locator(
                        "button[type='submit'], input[type='submit'], button:has-text('Войти'), button:has-text('Вход')"
                    ).first
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
        if "выйти" in html or "выход" in html or "logout" in html or "мой профиль" in html:
            log("Вход выполнен")
        else:
            if not await wait_user(page, stop, "Войди в аккаунт Grenar, если ещё не вошёл."):
                return
        if stop.is_set():
            return
        await wait_user(page, stop, "Пауза после входа в кабинет.")
    except Exception as error:
        if stop.is_set() or "has been closed" in str(error):
            log("Остановка по запросу")
            stop.set()
            return
        log_error(f"Вход не удался: {error}")
        await wait_user(page, stop, "Вход не прошёл. Войди вручную.")


def write_xlsx(path: Path, rows: list[dict]) -> None:
    book = Workbook()
    sheet = book.active
    sheet.title = "John Deere"
    sheet.append(HEADERS)
    fill = PatternFill("solid", fgColor="1B5E20")
    font = Font(bold=True, color="FFFFFF")
    for cell in sheet[1]:
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(horizontal="center")
    for row in rows:
        sheet.append([row.get(name, "") for name in HEADERS])
    price_col = HEADERS.index("Цена") + 1
    old_col = HEADERS.index("Цена_старая") + 1
    qty_col = HEADERS.index("Количество") + 1
    for index in range(2, sheet.max_row + 1):
        price = sheet.cell(index, price_col)
        old = sheet.cell(index, old_col)
        qty = sheet.cell(index, qty_col)
        if isinstance(price.value, float):
            price.number_format = '#,##0.00" ₽"'
        if isinstance(old.value, float):
            old.number_format = '#,##0.00" ₽"'
        if isinstance(qty.value, int):
            qty.number_format = "#,##0"
    for col in range(1, len(HEADERS) + 1):
        letter = get_column_letter(col)
        width = max(len(HEADERS[col - 1]), 12)
        sheet.column_dimensions[letter].width = min(width + 6, 42)
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(".tmp.xlsx")
    book.save(tmp)
    tmp.replace(path)


async def scrape_page(page, page_num: int, stop) -> tuple[str, list[dict]]:
    url = BASE if page_num == 1 else f"{BASE}page/{page_num}/"
    log(f"Открываю страницу {page_num}")
    await page.goto(url, timeout=90000, wait_until="domcontentloaded")
    try:
        await page.bring_to_front()
    except Exception:
        pass
    await asyncio.sleep(random.uniform(0.8, 1.4))
    if await is_challenge(page):
        if not await wait_user(page, stop, f"Капча на странице {page_num}."):
            return "stopped", []
    try:
        await page.wait_for_selector("article.product-card", timeout=12000)
    except PlaywrightTimeout:
        log(f"Карточек нет на странице {page_num}")
        return "empty", []
    if stop.is_set():
        return "stopped", []
    raw = await page.evaluate(
        """() => Array.from(document.querySelectorAll("article.product-card")).map((card) => {
            const link = card.querySelector("a.product-card__media-link");
            const article = (card.querySelector(".product-card__article, [class*='article']") || {}).innerText || "";
            const name = (card.querySelector(".product-card__title, [class*='title']") || {}).innerText || "";
            const text = (card.innerText || "").replace(/\\s+/g, " ").trim();
            const href = link ? link.getAttribute("href") || "" : "";
            const code = href.split("/product/")[-1].replace(/\\/$/, "") || "";
            return {
                article: article.trim(),
                name: name.trim(),
                code,
                text
            };
        })"""
    )
    items = []
    for card in raw or []:
        article = re.sub(r"\s+", " ", card.get("article") or "").strip()
        name = re.sub(r"\s+", " ", card.get("name") or "").strip()
        code = card.get("code") or ""
        if not article or not code:
            continue
        text = card.get("text") or ""
        prices = re.findall(r"[\d\s\xa0]+[.,]\d{2}\s*₽", text)
        prices = [parse_price(p) for p in prices]
        prices = [p for p in prices if p is not None]
        discount = ""
        match = re.search(r"−\d+%|-\d+%", text)
        if match:
            discount = match.group(0)
        ship = ""
        if "Отгрузка" in text:
            after = text.split("Отгрузка", 1)[-1].strip()
            parts = after.split()[:3]
            ship = " ".join(parts)
        current = prices[-1] if prices else None
        old = prices[0] if len(prices) > 1 else None
        items.append(
            {
                "Артикул": article,
                "Название": name,
                "Бренд": "JOHN DEERE",
                "Цена": current,
                "Цена_старая": old,
                "Скидка": discount,
                "Количество": None,
                "Отгрузка": ship or "—",
                "_code": code,
            }
        )
    return "items", items


async def fetch_stock(code: str) -> int | None:
    url = API_PRODUCT.format(code)
    def _get():
        req = urllib.request.Request(url, headers={"User-Agent": "pult"})
        try:
            with urllib.request.urlopen(req, timeout=20) as resp:
                data = json.loads(resp.read().decode("utf-8"))
                return data.get("data", {}).get("stock")
        except Exception:
            return None
    return await asyncio.to_thread(_get)


async def enrich_stocks(rows: list[dict], stop) -> None:
    codes = [(i, row["_code"]) for i, row in enumerate(rows) if row.get("_code")]
    if not codes:
        return
    log(f"Забираю точные остатки через API ({len(codes)} позиций, {CONCURRENCY} параллельно)…")
    sem = asyncio.Semaphore(CONCURRENCY)

    async def one(index: int, code: str):
        if stop.is_set():
            return
        async with sem:
            stock = await fetch_stock(code)
            if stock is not None:
                rows[index]["Количество"] = stock
            await asyncio.sleep(random.uniform(0.04, 0.12))

    tasks = [one(i, code) for i, code in codes]
    done = 0
    for coro in asyncio.as_completed(tasks):
        await coro
        done += 1
        if done % 200 == 0:
            log(f"Остатки: {done}/{len(codes)}")
        if stop.is_set():
            break
    log("Остатки собраны")


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
    LOG_FILE = out.with_name("grenar.log")
    CONTINUE = out.with_name("continue.flag")
    PAUSE = out.with_name("pause.flag")
    for flag in (CONTINUE, PAUSE):
        if flag.exists():
            flag.unlink()
    log(f"Лог ошибок: {LOG_FILE}")

    stop = asyncio.Event()
    loop = asyncio.get_running_loop()

    def request_stop(*_args):
        loop.call_soon_threadsafe(stop.set)

    try:
        loop.add_signal_handler(signal.SIGINT, stop.set)
        loop.add_signal_handler(signal.SIGTERM, stop.set)
    except NotImplementedError:
        signal.signal(signal.SIGINT, request_stop)
        if hasattr(signal, "SIGTERM"):
            signal.signal(signal.SIGTERM, request_stop)

    rows: list[dict] = []
    browsers = Path("/opt/pw-browsers")
    if browsers.is_dir() and "PLAYWRIGHT_BROWSERS_PATH" not in os.environ:
        os.environ["PLAYWRIGHT_BROWSERS_PATH"] = str(browsers)
    chrome = os.environ.get("GNG_CHROME") or CHROME
    launch = {
        "headless": not args.show,
        "slow_mo": 150 if args.show else 0,
        "args": ["--disable-blink-features=AutomationControlled", "--no-sandbox", "--disable-dev-shm-usage"],
    }
    if not args.show and chrome and Path(chrome).is_file():
        launch["executable_path"] = chrome
    if args.show:
        log("Браузер откроется отдельным окном — там видно, что делает парсер")

    async with async_playwright() as pw:
        browser = await pw.chromium.launch(**launch)
        context = await browser.new_context(
            locale="ru-RU",
            viewport={"width": 1400, "height": 900},
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                "(KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36"
            ),
        )
        login = os.environ.get("PULT_LOGIN", os.environ.get("GNG_LOGIN", "")).strip()
        password = os.environ.get("PULT_PASSWORD", os.environ.get("GNG_PASSWORD", ""))
        page = await context.new_page()
        if login and password:
            await try_login(page, login, password, stop)
        if stop.is_set():
            await browser.close()
            return

        seen: set[str] = set()
        page_num = 1
        while page_num <= 2000:
            if stop.is_set():
                log("Остановка по запросу")
                break
            if not await honor_pause(page, stop):
                log("Остановка по запросу")
                break
            if limit is not None and page_num > limit:
                log(f"Достигнут лимит {limit}")
                break
            try:
                status, batch = await scrape_page(page, page_num, stop)
            except Exception as error:
                if stop.is_set() or "has been closed" in str(error):
                    log("Остановка по запросу")
                    break
                log_error(f"Страница {page_num}: {error}")
                page_num += 1
                continue
            if status == "stopped":
                log("Остановка по запросу")
                break
            if status == "empty":
                log("Дальше страниц нет")
                break
            fresh = [row for row in batch if row.get("Артикул") and row["Артикул"] not in seen]
            if batch and not fresh:
                log("Страница повторяет уже собранное — это конец")
                break
            for row in fresh:
                seen.add(row["Артикул"])
            rows.extend(fresh)
            log(f"Страница {page_num}: карточек {len(batch)}, новых {len(fresh)}. Всего {len(rows)}")
            page_num += 1
            await asyncio.sleep(random.uniform(0.6, 1.2))
        await browser.close()

    if rows and not stop.is_set():
        await enrich_stocks(rows, stop)

    # убираем служебное поле
    for row in rows:
        row.pop("_code", None)

    write_xlsx(out, rows)
    log(f"ROWS {len(rows)}")
    log("DONE")


if __name__ == "__main__":
    try:
        asyncio.run(main())
    except Exception as error:
        log_error(f"Парсер остановился: {error}")
        sys.exit(1)
