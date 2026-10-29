"""Правила автоматической оценки («Корректно») для HDE AI Analysis.

Модуль не ходит в сеть и не знает про HDE: получает тексты и ячейки, возвращает решение.
Все фразы и стемы вынесены в константы вверху, чтобы их можно было дополнять.
"""
import re
from datetime import datetime, time as dtime

from openpyxl.utils import column_index_from_string

OK = "Корректно"

# ---------------------------------------------------------------- настройки
MAIN_F_COL = "F"        # колонка с решением/обработкой в основном листе
MAIN_RESULT_COL = "H"   # сюда пишем «Корректно» в основном листе

ESC_SHEET = "Эскалации"
ESC_TIME_COL = "B"      # время (UTC+3)
ESC_REGION_COL = "D"    # регион
ESC_EMAIL_COL = "E"     # email отправителя
ESC_ROLE_COL = "G"      # «Управляющий офисом»
ESC_EMAIL2_COL = "H"    # email получателя
ESC_STATUS_COL = "I"    # «Ждёт рабочего времени офиса»
ESC_RESULT_COL = "K"    # сюда пишем «Корректно» на листе эскалаций

WORK_START = dtime(9, 0)
WORK_END = dtime(18, 0)
SOURCE_UTC_OFFSET = 3   # время в колонке B указано в UTC+3

# ---------------------------------------------------------------- нормализация
_ARROWS = re.compile(r"\s*(?:→|->|=>|⇒|➔|➜|—>|–>)\s*")


def norm(s) -> str:
    """lower, ё→е, неразрывные пробелы, единые стрелки, схлопнутые пробелы."""
    if s is None:
        return ""
    s = str(s).replace("\xa0", " ").lower().replace("ё", "е")
    s = _ARROWS.sub(" → ", s)
    return re.sub(r"\s+", " ", s).strip()


# ---------------------------------------------------------------- основной лист
# Клиент просит отмену / возврат (Отмена, Отмените, Возврат, ...)
CANCEL_RE = re.compile(r"(?<!\w)(?:отмен|возврат)\w*")

# Ответ ИИ, подтверждающий оформление возврата
AI_RETURN_PHRASES = tuple(norm(p) for p in (
    "Информация о необходимости возврата внесена в накладную",
    "По данному отправлению возврат уже оформлен ранее, повторная заявка не требуется",
))
# Значение колонки F, подтверждающее отмену
F_CANCEL_PHRASES = tuple(norm(p) for p in (
    "Отменена часть списка → остальное оператору",
))

# Клиент просит изменить / перенести дату доставки
DATE_CHANGE_RES = tuple(re.compile(p) for p in (
    # «Просьба доставить 30.10», «Простьба перенести доставку на ...», «просьба изменить дату»
    r"(?<!\w)прост?ьб\w*[^.\n]{0,40}?(?:достав|перенес|перенос|измен)",
    # «Изменить дату доставки», «перенесите доставку», «поменять дату доставки»
    r"(?<!\w)(?:измен|перенес|перенос|поменя|сдвин)\w*[^.\n]{0,25}?достав",
    # «доставить 30.10», «доставка на 30.10.2026»
    r"(?<!\w)достав\w*\s+(?:на\s+|до\s+|к\s+)?\d{1,2}[./-]\d{1,2}",
))
# Значения колонки F, подтверждающие обработку даты
F_DATE_PHRASES = tuple(norm(p) for p in (
    "Дата доставки изменена + ответ клиенту",
    "Дата уже стоит в УПД → ответ клиенту по шаблону",
    "Дата недоступна → клиенту отправлен список вариантов",
))


def _has_any(text: str, phrases) -> bool:
    return any(p in text for p in phrases)


def is_correct_cancel(client_text, ai_text, f_text) -> bool:
    """Клиент просит отмену/возврат И (ответ ИИ — одна из двух фраз ИЛИ в колонке F «Отменена часть списка»)."""
    if not CANCEL_RE.search(norm(client_text)):
        return False
    return _has_any(norm(ai_text), AI_RETURN_PHRASES) or _has_any(norm(f_text), F_CANCEL_PHRASES)


def is_correct_date_change(client_text, f_text) -> bool:
    """Клиент просит изменить дату доставки И колонка F — одно из трёх значений."""
    c = norm(client_text)
    if not any(rx.search(c) for rx in DATE_CHANGE_RES):
        return False
    return _has_any(norm(f_text), F_DATE_PHRASES)


def is_correct_main(client_text, ai_text, f_text) -> bool:
    return is_correct_cancel(client_text, ai_text, f_text) or is_correct_date_change(client_text, f_text)


# ---------------------------------------------------------------- эскалации
_EMAIL_DOMAIN = re.compile(r"[\w.+-]+@([\w-]+(?:\.[\w-]+)+)")


def email_domain(value):
    m = _EMAIL_DOMAIN.search(str(value or ""))
    return m.group(1).lower() if m else None


# Смещение региона от UTC (часы). Ищется по началу слова; неизвестный регион → правило пропускается.
REGION_UTC = {
    # +2
    "калининград": 2,
    # +3
    "москв": 3, "московск": 3, "санкт-петербург": 3, "петербург": 3, "ленинградск": 3,
    "нижегородск": 3, "нижний новгород": 3, "новгород": 3, "казан": 3, "татарстан": 3, "воронеж": 3,
    "ростов": 3, "краснодар": 3, "сочи$": 3, "ставрополь": 3, "тула$": 3, "тульск": 3, "рязан": 3,
    "ярославл": 3, "владимир": 3, "иваново$": 3, "ивановск": 3, "калуг": 3, "смолен": 3, "брянск": 3,
    "курск": 3, "белгород": 3, "липецк": 3, "тамбов": 3, "волгоград": 3, "киров$": 3, "кировск": 3,
    "чебоксар": 3, "чуваш": 3, "марий": 3, "йошкар": 3, "саранск": 3, "мордов": 3, "псков": 3,
    "тверь$": 3, "тверск": 3, "вологд": 3, "архангельск": 3, "мурманск": 3, "карели": 3,
    "петрозаводск": 3, "сыктывкар": 3, "коми$": 3, "орел$": 3, "орлов": 3, "крым$": 3,
    "симферопол": 3, "севастопол": 3, "костром": 3, "пенз": 3, "адыгея$": 3, "майкоп": 3,
    # +4
    "самар": 4, "тольятти": 4, "ульяновск": 4, "саратов": 4, "астрахан": 4, "удмурт": 4, "ижевск": 4,
    # +5
    "екатеринбург": 5, "свердловск": 5, "челябинск": 5, "тюмен": 5, "перм": 5, "уфа$": 5, "уфим": 5,
    "башкир": 5, "башкортостан": 5, "оренбург": 5, "курган": 5, "ханты": 5, "югра$": 5,
    "сургут": 5, "нижневартовск": 5, "ямал": 5,
    # +6
    "омск": 6,
    # +7
    "новосибирск": 7, "красноярск": 7, "томск": 7, "кемеров": 7, "кузбасс": 7, "новокузнецк": 7,
    "алтайск": 7, "алтай$": 7, "барнаул": 7, "абакан": 7, "хакас": 7, "горно-алтайск": 7,
    "тыва$": 7, "кызыл": 7,
    # +8
    "иркутск": 8, "бурят": 8, "улан-удэ": 8,
    # +9
    "забайкал": 9, "чита$": 9, "якут": 9, "саха$": 9, "амурск": 9, "благовещенск": 9,
    # +10
    "приморск": 10, "владивосток": 10, "хабаровск": 10,
    # +11
    "сахалин": 11, "южно-сахалинск": 11, "магадан": 11,
    # +12
    "камчат": 12, "петропавловск": 12, "чукот": 12, "анадырь": 12,
}
_REGION_KEYS = sorted(REGION_UTC, key=len, reverse=True)


def region_offset(region):
    """UTC-смещение региона или None, если регион неизвестен."""
    r = norm(region)
    if not r:
        return None
    for key in _REGION_KEYS:
        whole = key.endswith("$")          # «$» в ключе = слово целиком, а не начало слова
        pat = r"(?<!\w)" + re.escape(key.rstrip("$")) + (r"(?!\w)" if whole else "")
        if re.search(pat, r):
            return REGION_UTC[key]
    return None


def parse_time(value):
    """Время суток из ячейки: datetime / time / доля суток Excel / строка «17:18» или «29.10.2026 17:18»."""
    if value is None or value == "":
        return None
    if isinstance(value, datetime):
        return value.time()
    if isinstance(value, dtime):
        return value
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        minutes = round((float(value) % 1) * 1440) % 1440
        return dtime(minutes // 60, minutes % 60)
    m = re.search(r"(?<!\d)(\d{1,2}):(\d{2})", str(value))
    if m:
        h, mi = int(m.group(1)), int(m.group(2))
        if 0 <= h < 24 and 0 <= mi < 60:
            return dtime(h, mi)
    return None


def local_time(msk_time: dtime, offset_utc: int) -> dtime:
    """Переводит время UTC+3 (колонка B) в местное время региона."""
    minutes = (msk_time.hour * 60 + msk_time.minute + (offset_utc - SOURCE_UTC_OFFSET) * 60) % 1440
    return dtime(minutes // 60, minutes % 60)


def is_outside_work_hours(local: dtime) -> bool:
    return local < WORK_START or local >= WORK_END


MANAGER = norm("Управляющий офисом")
WAITING = norm("Ждёт рабочего времени офиса")


def escalation_verdict(time_v, region_v, email_e, role_g, email_h, status_i):
    """Возвращает (is_correct: bool, reason: str). reason нужен для лога.

    Правило 1: G «Управляющий офисом», I НЕ «Ждёт рабочего времени офиса», домен E == домен H.
    Правило 2 (если возможно): то же, но I «Ждёт...», и время из B в регионе D вне 9:00–18:00.
    """
    if MANAGER not in norm(role_g):
        return False, "не управляющий офисом"
    d_e, d_h = email_domain(email_e), email_domain(email_h)
    if not d_e or not d_h:
        return False, "нет email в E или H"
    if d_e != d_h:
        return False, "домены различаются"
    if WAITING not in norm(status_i):
        return True, "правило 1"
    t = parse_time(time_v)
    if t is None:
        return False, "правило 2 невозможно: не удалось прочитать время в B"
    off = region_offset(region_v)
    if off is None:
        return False, f"правило 2 невозможно: неизвестный регион {region_v!r}"
    if is_outside_work_hours(local_time(t, off)):
        return True, "правило 2"
    return False, "правило 2: время в регионе рабочее"


def find_sheet(wb, name):
    want = norm(name)
    return next((ws for ws in wb.worksheets if norm(ws.title) == want), None)


def apply_escalations(wb, log=print):
    """Проходит лист «Эскалации» и ставит «Корректно» в колонку K. Существующие значения K не затирает."""
    ws = find_sheet(wb, ESC_SHEET)
    if ws is None:
        log(f"Лист «{ESC_SHEET}» не найден — правила эскалаций пропущены.")
        return {"marked": 0}
    ci = {k: column_index_from_string(v) for k, v in dict(
        t=ESC_TIME_COL, reg=ESC_REGION_COL, e=ESC_EMAIL_COL, g=ESC_ROLE_COL,
        h=ESC_EMAIL2_COL, i=ESC_STATUS_COL, res=ESC_RESULT_COL).items()}
    stats = {"marked": 0, "rule1": 0, "rule2": 0, "occupied": 0, "unknown_regions": {}, "bad_time": 0}
    for row in range(2, ws.max_row + 1):
        g = ws.cell(row, ci["g"]).value
        if MANAGER not in norm(g):
            continue
        ok, why = escalation_verdict(
            ws.cell(row, ci["t"]).value, ws.cell(row, ci["reg"]).value, ws.cell(row, ci["e"]).value,
            g, ws.cell(row, ci["h"]).value, ws.cell(row, ci["i"]).value)
        if why.startswith("правило 2 невозможно: неизвестный регион"):
            reg = str(ws.cell(row, ci["reg"]).value)
            stats["unknown_regions"][reg] = stats["unknown_regions"].get(reg, 0) + 1
        elif why.startswith("правило 2 невозможно: не удалось"):
            stats["bad_time"] += 1
        if not ok:
            continue
        cur = ws.cell(row, ci["res"]).value
        if cur not in (None, "") and norm(cur) != norm(OK):
            stats["occupied"] += 1      # там уже что-то написано — не трогаем
            continue
        ws.cell(row, ci["res"], OK)
        stats["marked"] += 1
        stats["rule1" if why == "правило 1" else "rule2"] += 1
    log(f"Эскалации: «{OK}» — {stats['marked']} (правило 1: {stats['rule1']}, правило 2: {stats['rule2']}); "
        f"K уже занята другим значением: {stats['occupied']}; не прочитано время: {stats['bad_time']}")
    if stats["unknown_regions"]:
        log("Регионы без известного часового пояса (правило 2 для них пропущено; добавьте в REGION_UTC): "
            + ", ".join(f"{k} ×{v}" for k, v in stats["unknown_regions"].items()))
    return stats
