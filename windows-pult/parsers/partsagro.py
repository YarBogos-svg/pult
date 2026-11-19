#!/usr/bin/env python3
"""Парсер каталога data.partsagro.ru: все бренды, по листу на бренд."""

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
BASE = "https://data.partsagro.ru"
COLUMNS = ["Название", "Артикул", "Бренд", "Цена", "Наличие", "OEM-номер", "Бренд OEM"]
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


def sheet_name(name: str, used: set[str]) -> str:
    clean = re.sub(r"[\[\]:*?/\\]", " ", name).strip() or "Бренд"
    clean = re.sub(r"\s+", " ", clean)[:31]
    base = clean
    number = 2
    while clean in used:
        suffix = f" {number}"
        clean = (base[: 31 - len(suffix)] + suffix).strip()
        number += 1
    used.add(clean)
    return clean


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
        log("Открываю PartsAgro")
        await page.goto(BASE, timeout=120000, wait_until="domcontentloaded")
        await page.bring_to_front()
        await asyncio.sleep(2)
        if await is_challenge(page):
            if not await wait_user(page, stop, "Сайт просит проверку перед входом."):
                return
        if login and password:
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
        if "выйти" in html or "выход" in html or "logout" in html:
            log("Вход выполнен")
        else:
            if not await wait_user(page, stop, "Войди в аккаунт PartsAgro, если ещё не вошёл."):
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


async def list_brands(page) -> list[dict]:
    raw = await page.evaluate(
        """() => {
            const select = document.querySelector("select[ng-model='searchData.brand_id']");
            if (!select) return [];
            return Array.from(select.options).map((option) => ({
                value: option.value,
                name: (option.textContent || "").replace(/\\s+/g, " ").trim(),
            }));
        }"""
    )
    brands = []
    seen = set()
    for item in raw or []:
        value = str(item.get("value") or "").strip()
        name = str(item.get("name") or "").strip()
        if not value or value in {"0", "null", "undefined"}:
            continue
        if not name or name.lower() in {"бренд", "все", "все бренды", "выберите", "выберите бренд", "-"}:
            continue
        key = name.lower()
        if key in seen:
            continue
        seen.add(key)
        brands.append({"value": value, "name": name})
    return brands


async def choose_brand(page, brand: dict) -> None:
    await page.evaluate(
        """({value}) => {
            const select = document.querySelector("select[ng-model='searchData.brand_id']");
            if (!select) return;
            select.value = value;
            select.dispatchEvent(new Event("change", {bubbles: true}));
            select.dispatchEvent(new Event("input", {bubbles: true}));
            if (window.angular) {
                try {
                    const scope = window.angular.element(select).scope();
                    if (scope) {
                        scope.searchData = scope.searchData || {};
                        scope.searchData.brand_id = value;
                        scope.$applyAsync();
                    }
                } catch (error) {}
            }
        }""",
        brand,
    )
    try:
        await page.select_option("select[ng-model='searchData.brand_id']", value=brand["value"])
    except Exception:
        pass
    await asyncio.sleep(0.6)
    button = page.locator("button[ng-click='search()']").first
    if await button.count():
        await button.click()
    await page.wait_for_timeout(1500)


async def scroll_all(page, stop) -> int:
    previous = 0
    idle = 0
    for step in range(1, 301):
        if stop.is_set():
            return previous
        if not await honor_pause(page, stop):
            return previous
        await page.evaluate(
            """() => {
                const box = document.querySelector(".dataTables_scrollBody");
                if (box) box.scrollTop = box.scrollHeight;
                window.scrollTo(0, document.body.scrollHeight);
            }"""
        )
        await asyncio.sleep(random.uniform(1.2, 2.0))
        count = await page.locator("table#table tbody tr").count()
        if step == 1 or step % 5 == 0 or count != previous:
            log(f"Прокрутка {step}: строк {count}")
        if count == previous:
            idle += 1
            if idle >= 3:
                return count
        else:
            idle = 0
        previous = count
    return previous


async def row_item(row, brand_name: str) -> dict | None:
    name_el = row.locator("b.catalog-name span").first
    name = ""
    if await name_el.count():
        name = (await name_el.get_attribute("title")) or (await name_el.inner_text())
    else:
        fallback = row.locator("b.catalog-name").first
        if await fallback.count():
            name = await fallback.inner_text()
    name = re.sub(r"\s+", " ", name or "").strip()
    sku_el = row.locator("span.selectable-art").first
    article = re.sub(r"\s+", " ", (await sku_el.inner_text()) if await sku_el.count() else "").strip()
    if brand_name and article.lower().endswith(" " + brand_name.lower()):
        article = article[: -(len(brand_name) + 1)].strip()
    price = ""
    for cell in await row.locator("td.text-right span.copy.selectable").all():
        classes = (await cell.get_attribute("class")) or ""
        if "total-sum" in classes:
            continue
        text = (await cell.inner_text()).strip()
        if "₽" in text or re.search(r"\d", text):
            price = text
            break
    availability = ""
    for cell in await row.locator("td.text-center").all():
        text = re.sub(r"\s+", " ", (await cell.inner_text())).strip()
        match = re.search(r"(\d+)\s*/\s*(\d+)", text)
        if match:
            availability = f"{match.group(1)} / {match.group(2)}"
            break
    if not name or not article:
        return None
    return {"Название": name, "Артикул": article, "Цена": price, "Наличие": availability}


async def oem_from_popup(page, row) -> list[dict]:
    found = []
    try:
        name = row.locator("b.catalog-name").first
        if not await name.count():
            return found
        await name.scroll_into_view_if_needed()
        await page.wait_for_timeout(200)
        await name.click()
        await page.wait_for_selector(".modal.show .sku-right, .modal.in .sku-right", timeout=8000)
        await page.wait_for_timeout(500)
        for item in await page.locator(".modal.show .sku-right ul li span.ng-binding, .modal.in .sku-right ul li span.ng-binding").all():
            raw = re.sub(r"\s+", " ", (await item.inner_text())).strip()
            if not raw:
                continue
            number, _, brand = raw.partition(" ")
            number = number.strip()
            brand = brand.strip()
            if number and (number, brand) not in {(row["oem"], row["brand"]) for row in found}:
                found.append({"oem": number, "brand": brand})
        close = page.locator(".modal.show .modal-footer button[data-dismiss='modal'], .modal.in .modal-footer button[data-dismiss='modal']").first
        if await close.count():
            await close.click()
        else:
            await page.keyboard.press("Escape")
        await page.wait_for_timeout(700)
    except Exception as error:
        log(f"OEM не открылся: {str(error)[:80]}")
        try:
            await page.keyboard.press("Escape")
            await page.wait_for_timeout(500)
        except Exception:
            pass
    return found


def write_xlsx(path: Path, by_brand: dict[str, list[dict]]) -> None:
    book = Workbook()
    summary = book.active
    summary.title = "Сводка"
    summary.append(["Бренд", "Товаров", "Строк с OEM"])
    header_fill = PatternFill("solid", fgColor="1E40AF")
    header_font = Font(bold=True, color="FFFFFF")
    main_fill = PatternFill("solid", fgColor="D9E1F2")
    main_font = Font(bold=True)
    line = Border(bottom=Side(style="medium", color="555555"))
    used = {"Сводка"}
    total_items = 0
    total_rows = 0
    for brand, items in by_brand.items():
        name = sheet_name(brand, used)
        sheet = book.create_sheet(name)
        sheet.append(COLUMNS)
        for cell in sheet[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center")
        expanded = 0
        for item in items:
            price = parse_price(item.get("Цена", ""))
            sheet.append([item.get("Название", ""), item.get("Артикул", ""), brand, price if price is not None else item.get("Цена", ""), item.get("Наличие", ""), "", ""])
            main_row = sheet.max_row
            for cell in sheet[main_row]:
                cell.font = main_font
                cell.fill = main_fill
            expanded += 1
            for oem in item.get("OEM", []):
                sheet.append(["", "", "", None, "", oem.get("oem", ""), oem.get("brand", "")])
                expanded += 1
            for cell in sheet[sheet.max_row]:
                cell.border = line
        for column in range(1, len(COLUMNS) + 1):
            letter = get_column_letter(column)
            width = max(len(COLUMNS[column - 1]), 12)
            sheet.column_dimensions[letter].width = min(width + 8, 48)
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = sheet.dimensions
        summary.append([brand, len(items), expanded])
        total_items += len(items)
        total_rows += expanded
    summary.append(["Итого", total_items, total_rows])
    for cell in summary[1]:
        cell.fill = header_fill
        cell.font = header_font
    for column in range(1, 4):
        summary.column_dimensions[get_column_letter(column)].width = 28
    summary.freeze_panes = "A2"
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(".tmp.xlsx")
    book.save(tmp)
    tmp.replace(path)


def load_progress(path: Path) -> dict[str, list[dict]]:
    if not path.is_file():
        return {}
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def save_progress(path: Path, data: dict[str, list[dict]]) -> None:
    path.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")


async def scrape_brand(page, brand: dict, limit: int | None, seen: set[str], stop) -> list[dict]:
    log(f"Бренд: {brand['name']}")
    await choose_brand(page, brand)
    if await is_challenge(page):
        if not await wait_user(page, stop, f"Капча на бренде {brand['name']}."):
            return []
    try:
        await page.wait_for_selector("table#table tbody tr", timeout=20000)
    except PlaywrightTimeout:
        log(f"У {brand['name']} таблица пустая")
        return []
    total = await scroll_all(page, stop)
    log(f"{brand['name']}: строк в таблице {total}")
    rows = page.locator("table#table tbody tr")
    count = await rows.count()
    items = []
    for index in range(count):
        if stop.is_set():
            break
        if not await honor_pause(page, stop):
            break
        if limit is not None and len(items) >= limit:
            log(f"Лимит {limit} на бренд")
            break
        row = rows.nth(index)
        item = await row_item(row, brand["name"])
        if not item or item["Артикул"] in seen:
            continue
        seen.add(item["Артикул"])
        item["OEM"] = await oem_from_popup(page, row)
        items.append(item)
        log(f"{brand['name']} [{len(items)}] {item['Артикул']} | {item['Цена']} | {item['Наличие']} | OEM {len(item['OEM'])}")
        await asyncio.sleep(random.uniform(0.6, 1.4))
    return items


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
    LOG_FILE = out.with_name("partsagro.log")
    CONTINUE = out.with_name("continue.flag")
    PAUSE = out.with_name("pause.flag")
    for flag in (CONTINUE, PAUSE):
        if flag.exists():
            flag.unlink()
    progress_path = out.with_name("progress.json")
    by_brand = load_progress(progress_path)
    if by_brand:
        log(f"Продолжаю с сохранённого: брендов {len(by_brand)}")
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

    browsers = Path("/opt/pw-browsers")
    if browsers.is_dir() and "PLAYWRIGHT_BROWSERS_PATH" not in os.environ:
        os.environ["PLAYWRIGHT_BROWSERS_PATH"] = str(browsers)
    chrome = os.environ.get("GNG_CHROME") or CHROME
    launch = {
        "headless": not args.show,
        "slow_mo": 200 if args.show else 0,
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
        page = await context.new_page()
        login = os.environ.get("PARTS_LOGIN", "").strip()
        password = os.environ.get("PARTS_PASSWORD", "")
        await try_login(page, login, password, stop)
        if stop.is_set():
            await browser.close()
            return
        brands = await list_brands(page)
        if not brands:
            log("Список брендов не нашёл. Открой каталог в браузере.")
            if not await wait_user(page, stop, "Не вижу список брендов. Открой поиск и нажми «Продолжить»."):
                await browser.close()
                return
            brands = await list_brands(page)
        log(f"Брендов: {len(brands)}")
        log("BRANDS " + ", ".join(item["name"] for item in brands))
        done = set(by_brand)
        for brand in brands:
            if stop.is_set():
                break
            if brand["name"] in done and by_brand.get(brand["name"]):
                log(f"{brand['name']} уже собран, пропускаю")
                continue
            seen = {item.get("Артикул") for item in by_brand.get(brand["name"], [])}
            items = await scrape_brand(page, brand, limit, seen, stop)
            if items:
                by_brand.setdefault(brand["name"], []).extend(items)
                save_progress(progress_path, by_brand)
                write_xlsx(out, by_brand)
                total = sum(len(rows) for rows in by_brand.values())
                log(f"ROWS {total}")
        await browser.close()

    write_xlsx(out, by_brand)
    total = sum(len(rows) for rows in by_brand.values())
    log(f"ROWS {total}")
    log("DONE")


if __name__ == "__main__":
    try:
        asyncio.run(main())
    except Exception as error:
        log_error(f"Парсер остановился: {error}")
        sys.exit(1)
