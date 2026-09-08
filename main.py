"""
Бот мониторинга тендеров на вентиляционное оборудование.

Источник: RSS-лента РосТендер (агрегатор 8000+ источников, включая 44-ФЗ,
223-ФЗ и коммерческие закупки), ветка «Вентиляционное оборудование и материалы».

Логика:
1. Скачиваем RSS.
2. Отсеиваем тендеры, которые уже видели раньше (по ID из ссылки).
3. Дополнительно фильтруем по ключевым словам — ветка РосТендера широкая
   и иногда цепляет тендеры на технику/запчасти, которые лишь косвенно
   попали в категорию.
4. Новые релевантные тендеры отправляем в Telegram.
5. Сохраняем обновлённый список увиденных ID.

Работает только на стандартной библиотеке Python — не нужно ничего
устанавливать через pip.
"""

import os
import re
import json
import html
import time
import urllib.request
import urllib.parse
from pathlib import Path
from xml.etree import ElementTree as ET

# ---------------------------------------------------------------------------
# Настройки
# ---------------------------------------------------------------------------

TELEGRAM_TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN", "")
CHAT_ID = os.environ.get("TELEGRAM_CHAT_ID", "")

# Если true — просто помечаем все текущие тендеры как увиденные,
# ничего не отправляя. Полезно для первого запуска, чтобы не получить
# разом сотни сообщений с историческими тендерами.
SEED_ONLY = os.environ.get("SEED_ONLY", "false").strip().lower() == "true"

# RSS-ленты для мониторинга.
# rss-branch-317.xml — вся ветка "Вентиляционное оборудование и материалы"
# (установки, вентиляторы, воздуховоды, решётки, калориферы и т.д.)
FEEDS = [
    "https://rostender.info/rss-branch-317.xml",
]
FEEDS = ["https://rostender.info/rss-branch-317.xml", "https://rostender.info/rss-category-313.xml",]

# Тендер считается релевантным, если в названии/описании встречается
# хотя бы одно слово из этого списка.
INCLUDE_KEYWORDS = [
    "вентиля",       # вентиляция, вентилятор, вентиляционный...
    "воздуховод",
    "приточ",        # приточная вентиляция
    "вытяжн",        # вытяжная вентиляция
    "аспираци",      # аспирационные установки
    "дымоудал",
    "калорифер",
    "диффузор",
    "решетк",        # вентиляционные решётки
    "клапан огнезадерж",
    "фильтровальн",
    "овк",           # отопление-вентиляция-кондиционирование
    "кондицион",
    "климат",
    "ПВУ",
    "рекуп",
    "ПУ",

]

# Явно исключаем — если встречается, тендер не релевантен, даже если
# случайно попал в ветку категории (частая ситуация для многолотовых закупок).
EXCLUDE_KEYWORDS = [
    "опрыскиватель",
    "caterpillar",
    "john deere",
    "бульдозер",
]

STATE_FILE = Path(__file__).parent / "seen_tenders.json"

# ---------------------------------------------------------------------------
# Работа с состоянием (какие тендеры уже видели)
# ---------------------------------------------------------------------------


def load_seen() -> set:
    if STATE_FILE.exists():
        try:
            return set(json.loads(STATE_FILE.read_text(encoding="utf-8")))
        except json.JSONDecodeError:
            return set()
    return set()


def save_seen(seen: set) -> None:
    STATE_FILE.write_text(
        json.dumps(sorted(seen), ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


# ---------------------------------------------------------------------------
# RSS
# ---------------------------------------------------------------------------


def fetch_feed(url: str) -> ET.Element:
    req = urllib.request.Request(
        url, headers={"User-Agent": "Mozilla/5.0 (TenderMonitorBot/1.0)"}
    )
    with urllib.request.urlopen(req, timeout=30) as resp:
        data = resp.read()
    return ET.fromstring(data)


def extract_tender_id(link: str) -> str:
    """Достаём числовой ID тендера из ссылки РосТендера."""
    m = re.search(r"/(\d{5,})-tender", link)
    if m:
        return m.group(1)
    m = re.search(r"/tender/(\d+)", link)
    if m:
        return m.group(1)
    return link  # запасной вариант — используем саму ссылку как ключ


def is_relevant(title: str, description: str) -> bool:
    text = f"{title} {description}".lower()
    if any(bad in text for bad in EXCLUDE_KEYWORDS):
        return False
    return any(good in text for good in INCLUDE_KEYWORDS)


# ---------------------------------------------------------------------------
# Telegram
# ---------------------------------------------------------------------------


def send_telegram_message(text: str) -> None:
    if not TELEGRAM_TOKEN or not CHAT_ID:
        raise RuntimeError(
            "Не заданы TELEGRAM_BOT_TOKEN / TELEGRAM_CHAT_ID (переменные окружения)."
        )
    url = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendMessage"
    data = urllib.parse.urlencode(
        {
            "chat_id": CHAT_ID,
            "text": text,
            "parse_mode": "HTML",
            "disable_web_page_preview": "false",
        }
    ).encode()
    req = urllib.request.Request(url, data=data)
    with urllib.request.urlopen(req, timeout=30) as resp:
        resp.read()


def format_message(title: str, description: str, link: str, pub_date: str) -> str:
    # description выглядит так:
    # "Предмет тендера: ...;<br />Место поставки: ...;<br />Цена: ... руб."
    desc = description.replace("<br />", "\n").replace("<br/>", "\n")
    desc = html.unescape(desc)
    title = html.unescape(title)
    return f"🆕 <b>{title}</b>\n{desc}\n📅 {pub_date}\n🔗 {link}"


# ---------------------------------------------------------------------------
# Основная логика
# ---------------------------------------------------------------------------


def main() -> None:
    seen = load_seen()
    updated_seen = set(seen)
    sent_count = 0
    skipped_irrelevant = 0

    for feed_url in FEEDS:
        try:
            root = fetch_feed(feed_url)
        except Exception as exc:  # noqa: BLE001
            print(f"[ошибка] Не удалось загрузить {feed_url}: {exc}")
            continue

        items = root.findall(".//item")
        print(f"[инфо] {feed_url}: получено {len(items)} тендеров в ленте")

        for item in items:
            title = (item.findtext("title") or "").strip()
            description = (item.findtext("description") or "").strip()
            link = (item.findtext("link") or "").strip()
            pub_date = (item.findtext("pubDate") or "").strip()

            tender_id = extract_tender_id(link)
            if tender_id in seen:
                continue

            if not is_relevant(title, description):
                skipped_irrelevant += 1
                updated_seen.add(tender_id)
                continue

            if SEED_ONLY:
                updated_seen.add(tender_id)
                continue

            message = format_message(title, description, link, pub_date)
            try:
                send_telegram_message(message)
                sent_count += 1
                time.sleep(1)  # не превышаем лимиты Telegram API
            except Exception as exc:  # noqa: BLE001
                print(f"[ошибка] Не отправлено ({tender_id}): {exc}")
                continue  # не помечаем как увиденный — попробуем в следующий раз

            updated_seen.add(tender_id)

    save_seen(updated_seen)

    if SEED_ONLY:
        print(f"[готово] Режим первого запуска: помечено как увиденные {len(updated_seen)} тендеров. Ничего не отправлено.")
    else:
        print(f"[готово] Отправлено новых тендеров: {sent_count}. Отфильтровано как нерелевантные: {skipped_irrelevant}.")


if __name__ == "__main__":
    main()
