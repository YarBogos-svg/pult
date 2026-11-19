#!/usr/bin/env python3
"""Неинтерактивный парсер каталога John Deere на gngzap.com."""

import argparse
import asyncio
import os
import random
import re
import signal
import sys
import traceback
from datetime import datetime
from pathlib import Path

from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from playwright.async_api import TimeoutError as PlaywrightTimeout
from playwright.async_api import async_playwright

CHROME = "/opt/pw-browsers/chromium_headless_shell-1248/chrome-headless-shell-linux64/chrome-headless-shell"
BASE = "https://gngzap.com/mycatalog/zapchasti-dlia-selkhoztekhniki/john-deere?start={}"
STEP = 40
HEADERS = ["Название", "Артикул", "Бренд", "Тип", "№ предложения", "Цена", "Наличие", "Срок"]
LOG_FILE: Path | None = None
CONTINUE: Path | None = None
PAUSE: Path | None = None


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


def parse_stock(value: str):
    if not value:
        return None
    text = value.lower()
    if any(word in text for word in ("под заказ", "нет в наличии", "отсутствует")):
        return 0
    match = re.search(r"(\d+)", text)
    return int(match.group(1)) if match else None


def total_pages(html: str) -> int | None:
    soup = BeautifulSoup(html, "lxml")
    starts = []
    for link in soup.select("a[href*='start=']"):
        href = link.get("href") or ""
        match = re.search(r"[?&]start=(\d+)", href)
        if match:
            starts.append(int(match.group(1)))
    if starts:
        return max(starts) // STEP + 1
    numbered = []
    for link in soup.select(".pagination a, .pager a, .pages a, ul.pagination a"):
        text = link.get_text(strip=True)
        if text.isdigit():
            numbered.append(int(text))
    if numbered:
        return max(numbered)
    text = soup.get_text(" ", strip=True)
    match = re.search(r"из\s+(\d{1,5})\s*стр", text, re.I)
    if match:
        return int(match.group(1))
    return None


def sku_count(rows: list[dict]) -> int:
    return len({row["Артикул"] for row in rows if row.get("Артикул")})


def captcha_wall(html: str) -> bool:
    soup = BeautifulSoup(html, "lxml")
    if soup.select_one("iframe[src*='recaptcha'], iframe[src*='hcaptcha'], iframe[src*='captcha']"):
        return True
    text = soup.get_text(" ", strip=True).lower()
    return "я не робот" in text or "подтвердите, что вы не робот" in text


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
    stock_col = HEADERS.index("Наличие") + 1
    for index in range(2, sheet.max_row + 1):
        price = sheet.cell(index, price_col)
        stock = sheet.cell(index, stock_col)
        if isinstance(price.value, float):
            price.number_format = '#,##0.00" ₽"'
        if isinstance(stock.value, int):
            stock.number_format = "#,##0"
    for col in range(1, len(HEADERS) + 1):
        letter = get_column_letter(col)
        width = max(len(HEADERS[col - 1]), 12)
        sheet.column_dimensions[letter].width = min(width + 4, 42)
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(".tmp.xlsx")
    book.save(tmp)
    tmp.replace(path)


def offers_from_html(html: str) -> list[dict]:
    soup = BeautifulSoup(html, "lxml")
    found = []
    for item in soup.select("li.item"):
        sku_input = item.select_one('input[name="number"]')
        sku = sku_input.get("value", "").strip() if sku_input else ""
        if not sku:
            continue
        brand_input = item.select_one('input[name="brand"]')
        brand = brand_input.get("value", "").strip() if brand_input else ""
        name_el = item.select_one(".articleDesc h3")
        name = name_el.get_text(strip=True) if name_el else ""
        desc_el = item.select_one(".articleDesc")
        kind = ""
        if desc_el:
            text = desc_el.get_text(" ", strip=True).upper()
            if "ОРИГИНАЛ" in text:
                kind = "Оригинал"
            elif "АНАЛОГ" in text:
                kind = "Аналог"
        tables = item.select("div.price_block table") or [None]
        for index, table in enumerate(tables, start=1):
            price = stock = deadline = ""
            if table is not None:
                for tr in table.select("tr"):
                    th, td = tr.find("th"), tr.find("td")
                    if not th or not td:
                        continue
                    key = th.get_text(strip=True).lower()
                    value = td.get_text(strip=True)
                    if "срок" in key:
                        deadline = value
                    elif "наличие" in key:
                        stock = value
                    elif "цена" in key:
                        price = value
            if not price and table is not None:
                continue
            if table is None:
                continue
            found.append(
                {
                    "Название": name,
                    "Артикул": sku,
                    "Бренд": brand,
                    "Тип": kind,
                    "№ предложения": index,
                    "Цена": parse_price(price),
                    "Наличие": parse_stock(stock),
                    "Срок": deadline,
                }
            )
    return found


async def page_text(page) -> str:
    try:
        return (await page.content()).lower()
    except Exception:
        return ""


async def is_challenge(page) -> bool:
    html = await page_text(page)
    if not html:
        return False
    if await page.locator("li.item").count():
        return False
    if captcha_wall(html):
        return True
    markers = (
        "access restricted",
        "security check",
        "pass browser security",
        "подтвердите, что вы не робот",
        "проверка безопасности",
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
        log("Открываю сайт и ищу кнопку «Авторизация»")
        await page.goto("https://gngzap.com/", timeout=90000, wait_until="domcontentloaded")
        await page.bring_to_front()
        await asyncio.sleep(1.2)
        if await is_challenge(page):
            if not await wait_user(page, stop, "Сайт просит проверку перед входом."):
                return
        for selector in ("text=Авторизация", "text=Вход", "text=Войти", "a[href*='login']", "a[href*='auth']"):
            locator = page.locator(selector).first
            if await locator.count() and await locator.is_visible():
                await locator.click()
                await page.wait_for_timeout(800)
                break
        secret = page.locator("input[type='password']").first
        if not await secret.count():
            if not await wait_user(page, stop, "Форму входа не нашёл. Войди вручную."):
                return
        else:
            form = page.locator("form").filter(has=page.locator("input[type='password']")).first
            scope = form if await form.count() else page
            user = scope.locator(
                "input[type='email'], input[name*='login'], input[name*='mail'], input[name*='user'], input[type='text']"
            ).first
            if not await user.count():
                if not await wait_user(page, stop, "Поле логина не нашёл. Войди вручную."):
                    return
            else:
                await user.fill(login)
                await secret.fill(password)
                submit = scope.locator("button[type='submit'], input[type='submit'], button:has-text('Войти'), button:has-text('Авторизация')").first
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
        password_box = page.locator("input[type='password']").first
        still_login = await password_box.count() and await password_box.is_visible()
        if "выйти" in html or "выход" in html or "logout" in html:
            log("Вход выполнен")
        elif still_login or await is_challenge(page):
            if not await wait_user(page, stop, "Вход не подтвердился. Закончи его в браузере."):
                return
        else:
            log("Форма входа закрыта")
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


async def scrape_page(page, page_num: int, known: int | None, stop) -> tuple[str, list[dict], int | None]:
    url = BASE.format((page_num - 1) * STEP)
    log(f"Открываю страницу {page_num}" + (f" из {known}" if known else ""))
    await page.goto(url, timeout=90000, wait_until="domcontentloaded")
    try:
        await page.bring_to_front()
    except Exception:
        pass
    await asyncio.sleep(random.uniform(1.0, 2.0))
    html = await page.content()
    found = total_pages(html) or known
    while not stop.is_set():
        try:
            await page.wait_for_selector("li.item", timeout=8000)
            break
        except PlaywrightTimeout:
            html = await page.content()
            found = total_pages(html) or found
            if await is_challenge(page):
                if not await wait_user(page, stop, f"Капча на странице {page_num}."):
                    return "stopped", [], found
                continue
            log(f"Карточек нет на странице {page_num}")
            return "empty", [], found
    if stop.is_set():
        return "stopped", [], found
    await page.evaluate(
        """() => document.querySelectorAll('.dd-pointer').forEach((el) => { try { el.click() } catch (e) {} })"""
    )
    log(f"Страница {page_num}: раскрываю цены")
    await asyncio.sleep(random.uniform(0.6, 1.2))
    html = await page.content()
    found = total_pages(html) or found
    return "items", offers_from_html(html), found


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
    LOG_FILE = out.with_name("gng.log")
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
        "slow_mo": 250 if args.show else 0,
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
        login = os.environ.get("GNG_LOGIN", "").strip()
        password = os.environ.get("GNG_PASSWORD", "")
        page = await context.new_page()
        if login and password:
            await try_login(page, login, password, stop)
        if stop.is_set():
            await browser.close()
            return

        seen: set[str] = set()
        known: int | None = None
        page_num = 1
        while page_num <= 5000:
            if stop.is_set():
                log("Остановка по запросу")
                break
            if not await honor_pause(page, stop):
                log("Остановка по запросу")
                break
            if limit is not None and page_num > limit:
                log(f"Достигнут лимит {limit}")
                break
            if known is not None and page_num > known:
                log(f"Каталог закончился на {known}")
                break
            try:
                status, batch, found = await scrape_page(page, page_num, known, stop)
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
            if found and found != known:
                known = found
                log(f"PAGES {known}")
            if status == "empty":
                log("Дальше страниц нет")
                break
            fresh = {row["Артикул"] for row in batch if row.get("Артикул")} - seen
            if batch and not fresh:
                log("Страница повторяет уже собранное — это конец")
                break
            seen.update(fresh)
            rows.extend(batch)
            write_xlsx(out, rows)
            log(f"Страница {page_num}: карточек {sku_count(batch)}, предложений {len(batch)}. Всего строк {len(rows)}")
            log(f"ROWS {len(rows)}")
            cards = sku_count(batch)
            more_known = known is not None and page_num < known
            if cards < STEP and not more_known:
                log("Последняя неполная страница — каталог собран")
                break
            if cards < STEP and more_known:
                log(f"На странице {page_num} карточек меньше {STEP}, но каталог ещё не кончился — иду дальше")
            page_num += 1
            await asyncio.sleep(random.uniform(1.4, 3.2))
        await browser.close()

    write_xlsx(out, rows)
    log(f"ROWS {len(rows)}")
    log("DONE")


if __name__ == "__main__":
    try:
        asyncio.run(main())
    except Exception as error:
        log_error(f"Парсер остановился: {error}")
        sys.exit(1)
