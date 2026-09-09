"""
Бот мониторинга тендеров на вентиляционное оборудование.

Источник: RSS-ленты РосТендера (агрегатор 8000+ источников, включая 44-ФЗ,
223-ФЗ и коммерческие закупки):
  - вся ветка "Вентиляционное оборудование и материалы"
  - категория "Вентиляция и кондиционирование" (шире: + монтаж/обслуживание, кондиционеры)

Логика:
1. Скачиваем обе RSS-ленты.
2. Отсеиваем тендеры, которые уже видели раньше (по ID из ссылки).
3. Отбрасываем совсем нерелевантные тендеры (случайно попавшие в категорию —
   например техника, не имеющая отношения к вентиляции) и тендеры без
   ключевых слов по теме вообще.
4. Оставшиеся тендеры СОРТИРУЕМ на подходящие и неподходящие, но не отбрасываем:
     ✅ — подходящий (поставка оборудования)
     🛑 — неподходящий (обслуживание или ремонт)
   Сортировку делает ИИ (Claude Haiku 4.5, понимает смысл текста), а не просто
   поиск слов — но если ключ ANTHROPIC_API_KEY не задан или ИИ недоступен,
   бот автоматически переключается на резервную сортировку по ключевым словам.
   Отправляем в Telegram ВСЕ тендеры по теме, с пометкой и пояснением, почему.
5. Для подходящих (✅) тендеров ДОПОЛНИТЕЛЬНО пытаемся найти документацию:
   если тендер размещён на ЕИС (44-ФЗ/223-ФЗ — там документация публична и
   бесплатна по закону), бот находит номер извещения, скачивает вложенные
   PDF/Word-файлы, вытаскивает текст и просит ИИ выделить упомянутые модели
   оборудования и материалы. Если тендер не с ЕИС, файлы не нашлись или
   что-то пошло не так — этот шаг просто пропускается, само уведомление
   всё равно уходит.
6. Сохраняем обновлённый список увиденных ID.

Требует установки библиотек pypdf и python-docx (см. requirements.txt) —
без них чтение PDF/Word не сработает, но остальной бот продолжит работать.
"""

import os
import re
import io
import json
import html
import time
import zipfile
import datetime
import urllib.request
import urllib.parse
import urllib.error
from pathlib import Path
from xml.etree import ElementTree as ET

# ---------------------------------------------------------------------------
# Настройки
# ---------------------------------------------------------------------------

TELEGRAM_TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN", "")
CHAT_ID = os.environ.get("TELEGRAM_CHAT_ID", "")

# Ключ для ИИ-классификации тендеров (Claude API). Если не задан — бот
# автоматически использует старую сортировку по ключевым словам.
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")
AI_MODEL = "claude-haiku-4-5-20251001"  # самая быстрая и дешёвая модель — для этой задачи достаточно

# Если true — просто помечаем все текущие тендеры как увиденные,
# ничего не отправляя. Полезно для первого запуска / после смены источников,
# чтобы не получить разом пачку сообщений с историческими тендерами.
SEED_ONLY = os.environ.get("SEED_ONLY", "false").strip().lower() == "true"

# RSS-ленты для мониторинга.
FEEDS = [
    "https://rostender.info/rss-branch-317.xml",
    "https://rostender.info/rss-category-313.xml",
]

# Тендер в принципе считается по теме, если в названии/описании встречается
# хотя бы одно слово из этого списка. Не полный фильтр "подходит/не подходит" —
# только грубое отсечение совсем не по теме тендеров.
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
]

# Явно исключаем — если встречается, тендер отбрасывается совсем, даже без
# пометки 🛑. Это шум из смежных категорий, не имеющий отношения к вентиляции
# вообще (например, многолотовые закупки спецтехники).
EXCLUDE_KEYWORDS = [
    "опрыскиватель",
    "caterpillar",
    "john deere",
    "бульдозер",
]

# Слова, указывающие, что тендер — на ОБСЛУЖИВАНИЕ/РЕМОНТ (🛑 неподходящий),
# а не на поставку оборудования (✅ подходящий).
SERVICE_KEYWORDS = [
    "техническое обслуживание",
    "техобслуживание",
    "обслуживание",
    "ремонт",
    "чистка",
    "прочистка",
    "диагностика",
    "пусконаладоч",  # пусконаладочные работы
]

STATE_FILE = Path(__file__).parent / "seen_tenders.json"

# "Тихие часы" по московскому времени — в этот период бот не присылает
# уведомления. Тендеры, найденные в это время, никуда не пропадают: они
# просто не помечаются как отправленные и уйдут одним пакетом, как только
# тихие часы закончатся.
QUIET_HOURS_START_MSK = 22  # с 22:00 МСК
QUIET_HOURS_END_MSK = 8     # до 08:00 МСК

# ---------------------------------------------------------------------------
# Анализ документации тендера (ЕИС) — экспериментальная функция
# ---------------------------------------------------------------------------
# Работает только для тендеров с ЕИС (44-ФЗ/223-ФЗ) — там документация
# публична и бесплатна по закону. Для тендеров с коммерческих площадок
# (не ЕИС) документы обычно закрыты подпиской конкретной площадки —
# бот их пропускает и просто не добавляет раздел с моделями оборудования.
#
# Делается только для тендеров с меткой ✅ (подходящих), чтобы не тратить
# время и деньги на анализ тендеров по обслуживанию/ремонту.
ANALYZE_DOCUMENTS = os.environ.get("ANALYZE_DOCUMENTS", "true").strip().lower() == "true"
MAX_DOC_FILES = 5              # не более стольки файлов на тендер
MAX_DOC_BYTES_TOTAL = 15_000_000  # суммарно не больше ~15 МБ на тендер
MAX_DOC_TEXT_CHARS = 12_000     # сколько текста максимум отдаём ИИ (экономия токенов)

EIS_DOC_EXTENSIONS = (".pdf", ".doc", ".docx", ".xls", ".xlsx", ".zip", ".rtf")

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


def is_on_topic(title: str, description: str) -> bool:
    """Грубая проверка: тендер вообще имеет отношение к вентиляции/климату?"""
    text = f"{title} {description}".lower()
    if any(bad in text for bad in EXCLUDE_KEYWORDS):
        return False
    return any(good in text for good in INCLUDE_KEYWORDS)


def is_quiet_hours() -> bool:
    """Сейчас 'тихие часы' по московскому времени (UTC+3, без перехода на летнее)?"""
    utc_now = datetime.datetime.now(datetime.timezone.utc)
    msk_hour = (utc_now.hour + 3) % 24
    if QUIET_HOURS_START_MSK > QUIET_HOURS_END_MSK:
        # период переходит через полночь, например 22:00 — 08:00
        return msk_hour >= QUIET_HOURS_START_MSK or msk_hour < QUIET_HOURS_END_MSK
    return QUIET_HOURS_START_MSK <= msk_hour < QUIET_HOURS_END_MSK


def classify_by_keywords(title: str, description: str) -> tuple:
    """
    Резервная сортировка по ключевым словам (используется, если ИИ недоступен).
    Возвращает (маркер, короткое пояснение).
    """
    text = f"{title} {description}".lower()
    has_service = any(word in text for word in SERVICE_KEYWORDS)

    if has_service:
        return "🛑", "Тендер на обслуживание/ремонт — не подходит под поставку оборудования"
    return "✅", "Подходит: похоже на поставку оборудования, а не на обслуживание/ремонт"


def classify_with_ai(title: str, description: str) -> tuple:
    """
    Отправляет тендер на классификацию модели Claude (Haiku 4.5). Возвращает
    (маркер, короткое пояснение). Бросает исключение при любой проблеме —
    вызывающий код должен подстраховаться резервной сортировкой по словам.
    """
    prompt = (
        "Ты помогаешь отобрать тендеры на поставку вентиляционного оборудования "
        "для производителя (центробежные ЕС-вентиляторы, вентиляционные установки).\n\n"
        f"Название тендера: {title}\n"
        f"Описание: {description}\n\n"
        "Реши: тендер ПОДХОДИТ (это закупка/поставка оборудования, даже если "
        "попутно требуется монтаж) или НЕ ПОДХОДИТ (это по сути услуга — "
        "техническое обслуживание, ремонт, чистка, диагностика, без закупки "
        "нового оборудования).\n\n"
        "Ответь СТРОГО в формате JSON, без пояснений вне JSON и без markdown-разметки:\n"
        '{"ok": true или false, "reason": "одно короткое предложение на русском"}'
    )

    body = json.dumps(
        {
            "model": AI_MODEL,
            "max_tokens": 200,
            "messages": [{"role": "user", "content": prompt}],
        }
    ).encode("utf-8")

    req = urllib.request.Request(
        "https://api.anthropic.com/v1/messages",
        data=body,
        headers={
            "content-type": "application/json",
            "x-api-key": ANTHROPIC_API_KEY,
            "anthropic-version": "2023-06-01",
        },
    )
    with urllib.request.urlopen(req, timeout=30) as resp:
        data = json.loads(resp.read().decode("utf-8"))

    reply_text = "".join(
        block.get("text", "") for block in data.get("content", []) if block.get("type") == "text"
    ).strip()

    # На случай, если модель всё же обернёт ответ в ```json ... ```
    reply_text = re.sub(r"^```(json)?|```$", "", reply_text.strip(), flags=re.MULTILINE).strip()

    parsed = json.loads(reply_text)
    if parsed.get("ok"):
        return "✅", parsed.get("reason", "ИИ счёл тендер подходящим")
    return "🛑", parsed.get("reason", "ИИ счёл тендер неподходящим")


def classify(title: str, description: str) -> tuple:
    """
    Сортирует тендер на подходящий (поставка оборудования) и неподходящий
    (обслуживание/ремонт). Использует ИИ, если задан ANTHROPIC_API_KEY,
    иначе — резервную сортировку по ключевым словам.
    """
    if ANTHROPIC_API_KEY:
        try:
            return classify_with_ai(title, description)
        except Exception as exc:  # noqa: BLE001
            print(f"[предупреждение] ИИ-классификация не сработала ({exc}), использую резервную по словам")
    return classify_by_keywords(title, description)


# ---------------------------------------------------------------------------
# Анализ документации тендера (ЕИС)
# ---------------------------------------------------------------------------


def fetch_url_bytes(url: str, timeout: int = 20, max_bytes: int = 20_000_000) -> bytes:
    """Скачивает URL, ограничивая размер (защита от случайно огромных файлов)."""
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0 (TenderMonitorBot/1.0)"})
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        data = resp.read(max_bytes + 1)
    if len(data) > max_bytes:
        raise ValueError(f"Файл больше лимита {max_bytes} байт, пропускаем")
    return data


def find_eis_reg_number(tender_page_url: str) -> tuple:
    """
    Открывает страницу тендера на РосТендере и ищет ссылку/номер на ЕИС.
    Возвращает (reg_number, law_type) где law_type — "44" или "223",
    либо (None, None), если тендер не с ЕИС (например, коммерческая площадка).
    """
    try:
        page_bytes = fetch_url_bytes(tender_page_url, timeout=20, max_bytes=3_000_000)
        page_text = page_bytes.decode("utf-8", errors="ignore")
    except Exception:
        return None, None

    # Прямая ссылка на извещение 44-ФЗ на zakupki.gov.ru
    m = re.search(r"zakupki\.gov\.ru/epz/order/notice/[^\"'\s]*[?&]regNumber=(\d+)", page_text)
    if m:
        return m.group(1), "44"

    # Прямая ссылка на извещение 223-ФЗ
    m = re.search(r"zakupki\.gov\.ru/223/purchase/public/purchase/info[^\"'\s]*[?&]regNumber=(\d+)", page_text)
    if m:
        return m.group(1), "223"

    # Запасной вариант: 19-значный номер извещения 44-ФЗ где-то в тексте
    m = re.search(r"\b(0\d{18})\b", page_text)
    if m:
        return m.group(1), "44"

    return None, None


def get_eis_documents_page_url(reg_number: str, law_type: str) -> str:
    if law_type == "223":
        return f"https://zakupki.gov.ru/223/purchase/public/purchase/info/documents.html?regNumber={reg_number}"
    return f"https://zakupki.gov.ru/epz/order/notice/ea44/view/documents.html?regNumber={reg_number}"


def extract_document_links(page_html: str) -> list:
    """
    Ищет в HTML ссылки на файлы документации (PDF/DOC/XLS/ZIP и т.п.).
    Реальные ссылки на скачивание в ЕИС часто не содержат расширения в самом
    URL (там технический ID файла) — расширение видно только в тексте
    ссылки (например "Техническое задание.pdf"). Проверяем оба варианта.
    """
    pairs = re.findall(r'<a[^>]+href="([^"]+)"[^>]*>([^<]*)</a>', page_html, flags=re.IGNORECASE | re.DOTALL)

    doc_links = []
    for href, link_text in pairs:
        href_lower = href.lower()
        text_lower = link_text.lower().strip()
        looks_like_doc = href_lower.endswith(EIS_DOC_EXTENSIONS) or text_lower.endswith(EIS_DOC_EXTENSIONS)
        if not looks_like_doc:
            continue

        url = href
        if url.startswith("//"):
            url = "https:" + url
        elif url.startswith("/"):
            url = "https://zakupki.gov.ru" + url

        # если расширение видно только в тексте ссылки, а не в самом URL —
        # запоминаем его отдельно, чтобы потом правильно выбрать парсер
        filename_hint = link_text.strip() if text_lower.endswith(EIS_DOC_EXTENSIONS) else url.rsplit("/", 1)[-1]
        doc_links.append((url, filename_hint))

    # убираем дубликаты по URL, сохраняя порядок
    seen_urls = set()
    unique_links = []
    for url, filename_hint in doc_links:
        if url not in seen_urls:
            seen_urls.add(url)
            unique_links.append((url, filename_hint))
    return unique_links[:MAX_DOC_FILES]


def extract_text_from_file(filename: str, data: bytes) -> str:
    """Достаёт текст из файла по расширению. Возвращает '' при неудаче/неподдержке."""
    lower = filename.lower()
    try:
        if lower.endswith(".pdf"):
            from pypdf import PdfReader  # библиотека ставится в GitHub Actions через requirements.txt

            reader = PdfReader(io.BytesIO(data))
            return "\n".join((page.extract_text() or "") for page in reader.pages)

        if lower.endswith(".docx"):
            import docx  # python-docx

            doc = docx.Document(io.BytesIO(data))
            return "\n".join(p.text for p in doc.paragraphs)

        if lower.endswith(".zip"):
            texts = []
            with zipfile.ZipFile(io.BytesIO(data)) as zf:
                for name in zf.namelist()[:MAX_DOC_FILES]:
                    if name.lower().endswith((".pdf", ".docx")):
                        try:
                            inner_data = zf.read(name)
                            texts.append(extract_text_from_file(name, inner_data))
                        except Exception:
                            continue
            return "\n".join(texts)

        # .doc (старый бинарный формат) и .xls/.xlsx — пропускаем, слишком
        # ненадёжно доставать текст без тяжёлых зависимостей
        return ""
    except Exception:
        return ""


def gather_tender_documents_text(tender_page_url: str) -> str:
    """
    Полный цикл: найти номер ЕИС → открыть страницу документов → скачать
    файлы → извлечь текст. Возвращает '' при любой неудаче (тендер не с
    ЕИС, файлы не найдены, скачивание не удалось и т.п.) — это ожидаемо
    для значительной части тендеров и не считается ошибкой.
    """
    reg_number, law_type = find_eis_reg_number(tender_page_url)
    if not reg_number:
        return ""

    docs_page_url = get_eis_documents_page_url(reg_number, law_type)
    try:
        docs_page_bytes = fetch_url_bytes(docs_page_url, timeout=20, max_bytes=3_000_000)
        docs_page_html = docs_page_bytes.decode("utf-8", errors="ignore")
    except Exception:
        return ""

    doc_links = extract_document_links(docs_page_html)
    if not doc_links:
        return ""

    all_text = []
    total_bytes = 0
    for doc_url, filename_hint in doc_links:
        if total_bytes >= MAX_DOC_BYTES_TOTAL:
            break
        try:
            data = fetch_url_bytes(doc_url, timeout=25, max_bytes=MAX_DOC_BYTES_TOTAL - total_bytes)
            total_bytes += len(data)
            text = extract_text_from_file(filename_hint, data)
            if text.strip():
                all_text.append(text)
        except Exception:
            continue

    combined = "\n\n---\n\n".join(all_text)
    return combined[:MAX_DOC_TEXT_CHARS]


def summarize_equipment_with_ai(documents_text: str) -> str:
    """
    Просит ИИ выделить из текста документации конкретные модели оборудования
    и материалы. Возвращает готовый текст для вставки в сообщение, либо ''
    если ничего не нашлось или ИИ недоступен.
    """
    if not documents_text.strip() or not ANTHROPIC_API_KEY:
        return ""

    prompt = (
        "Ниже — текст, извлечённый из документации тендера на поставку "
        "вентиляционного оборудования (техническое задание, спецификация "
        "и т.п.). Найди в нём конкретные упомянутые модели оборудования, "
        "марки материалов, производителей и ключевые технические параметры "
        "(производительность, диаметр, мощность и т.п.), если они указаны.\n\n"
        f"Текст документации:\n{documents_text}\n\n"
        "Ответь СТРОГО в формате JSON, без пояснений вне JSON и без markdown:\n"
        '{"found": true или false, "summary": "краткий список моделей/материалов '
        'в 2-5 строк, каждая с новой строки через \\n, на русском"}\n'
        'Если ничего конкретного не нашлось, верни {"found": false, "summary": ""}'
    )

    body = json.dumps(
        {
            "model": AI_MODEL,
            "max_tokens": 500,
            "messages": [{"role": "user", "content": prompt}],
        }
    ).encode("utf-8")

    req = urllib.request.Request(
        "https://api.anthropic.com/v1/messages",
        data=body,
        headers={
            "content-type": "application/json",
            "x-api-key": ANTHROPIC_API_KEY,
            "anthropic-version": "2023-06-01",
        },
    )
    try:
        with urllib.request.urlopen(req, timeout=40) as resp:
            data = json.loads(resp.read().decode("utf-8"))
        reply_text = "".join(
            block.get("text", "") for block in data.get("content", []) if block.get("type") == "text"
        ).strip()
        reply_text = re.sub(r"^```(json)?|```$", "", reply_text.strip(), flags=re.MULTILINE).strip()
        parsed = json.loads(reply_text)
        if parsed.get("found") and parsed.get("summary"):
            return parsed["summary"]
    except Exception as exc:  # noqa: BLE001
        print(f"[предупреждение] Не удалось получить сводку по оборудованию: {exc}")
    return ""


def analyze_tender_documents(tender_page_url: str) -> str:
    """
    Оркестрирует весь цикл анализа документации. Никогда не бросает
    исключение наружу — при любой проблеме просто возвращает ''.
    """
    if not ANALYZE_DOCUMENTS:
        return ""
    try:
        documents_text = gather_tender_documents_text(tender_page_url)
        if not documents_text:
            return ""
        return summarize_equipment_with_ai(documents_text)
    except Exception as exc:  # noqa: BLE001
        print(f"[предупреждение] Анализ документации тендера не удался: {exc}")
        return ""


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


def format_message(
    title: str,
    description: str,
    link: str,
    pub_date: str,
    marker: str,
    reason: str,
    equipment_summary: str = "",
) -> str:
    desc = description.replace("<br />", "\n").replace("<br/>", "\n")
    desc = html.unescape(desc)
    title = html.unescape(title)
    message = f"{marker} <b>{title}</b>\n<i>{reason}</i>\n{desc}\n📅 {pub_date}\n🔗 {link}"
    if equipment_summary:
        message += f"\n\n📋 <b>Найдено в документации:</b>\n{html.escape(equipment_summary)}"
    return message


# ---------------------------------------------------------------------------
# Основная логика
# ---------------------------------------------------------------------------


def main() -> None:
    seen = load_seen()
    updated_seen = set(seen)
    sent_count = 0
    sent_ok = 0
    sent_stop = 0
    skipped_off_topic = 0
    postponed_quiet_hours = 0

    quiet_now = is_quiet_hours()
    if quiet_now and not SEED_ONLY:
        print(
            f"[инфо] Сейчас тихие часы ({QUIET_HOURS_START_MSK}:00–{QUIET_HOURS_END_MSK}:00 МСК) — "
            "новые тендеры не отправляются, будут накоплены и присланы утром."
        )

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

            if not is_on_topic(title, description):
                skipped_off_topic += 1
                updated_seen.add(tender_id)
                continue

            marker, reason = classify(title, description)

            if SEED_ONLY:
                updated_seen.add(tender_id)
                continue

            if quiet_now:
                # Не отправляем и не помечаем как увиденный — тендер будет
                # обработан заново на следующей проверке (в том числе тогда,
                # когда тихие часы уже закончатся).
                postponed_quiet_hours += 1
                continue

            message = format_message(title, description, link, pub_date, marker, reason)

            if marker == "✅":
                equipment_summary = analyze_tender_documents(link)
                if equipment_summary:
                    message = format_message(
                        title, description, link, pub_date, marker, reason, equipment_summary
                    )

            try:
                send_telegram_message(message)
                sent_count += 1
                if marker == "✅":
                    sent_ok += 1
                else:
                    sent_stop += 1
                time.sleep(1)  # не превышаем лимиты Telegram API
            except Exception as exc:  # noqa: BLE001
                print(f"[ошибка] Не отправлено ({tender_id}): {exc}")
                continue  # не помечаем как увиденный — попробуем в следующий раз

            updated_seen.add(tender_id)

    save_seen(updated_seen)

    if SEED_ONLY:
        print(
            f"[готово] Режим первого запуска: помечено как увиденные {len(updated_seen)} тендеров. Ничего не отправлено."
        )
    elif quiet_now:
        print(
            f"[готово] Тихие часы: отложено до утра {postponed_quiet_hours} тендеров. "
            f"Отфильтровано как не по теме: {skipped_off_topic}."
        )
    else:
        print(
            f"[готово] Отправлено тендеров: {sent_count} (✅ подходящие: {sent_ok}, 🛑 обслуживание/ремонт: {sent_stop}). "
            f"Отфильтровано как не по теме: {skipped_off_topic}."
        )


if __name__ == "__main__":
    main()
