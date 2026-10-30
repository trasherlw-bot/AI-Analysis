"""Правила автоматической оценки («Корректно») для готового отчёта HDE AI Analysis.

Модуль не ходит в сеть и не знает про HDE: получает ячейки, возвращает решение.
Все фразы и стемы вынесены в константы вверху, чтобы их можно было дополнять.
Колонки ищутся по ЗАГОЛОВКУ, а не по букве: если в отчёте поменяют порядок колонок, всё продолжит работать.
"""
import re
from datetime import datetime, time as dtime

from openpyxl.utils import get_column_letter

OK = "Корректно"

# ---------------------------------------------------------------- настройки
REPORT_SHEET = "Обращения"
# ключ -> заголовок колонки в отчёте (после ключа в скобках — буква в отчёте на 27.10.2026, только для справки)
REPORT_COLS = {
    "type": "Тип запроса",              # C
    "subject": "Тема",                  # D
    "result": "Итог",                   # F — что сделал ИИ
    "client": "Запрос клиента",         # G
    "note": "Причина / примечание",     # H
    "opinion": "Комментарий КСЭ",       # I — сюда пишем «Корректно»
}
REPORT_REQUIRED = ("result", "client")   # без этих колонок лист не считается готовым отчётом
SUMMARY_SHEET = "Сводка"

ESC_SHEET = "Эскалации"
ESC_COLS = {
    "time": "Время",                    # B — время UTC+3
    "region": "Офис",                   # D
    "to": "Изначальные адресаты",       # E
    "role": "Кого подключили",          # G
    "addr": "Адрес",                    # H
    "status": "Итог",                   # I
    "opinion": "Комментарий КСЭ",       # K — сюда пишем «Корректно»
}
ESC_REQUIRED = ("time", "region", "to", "role", "addr", "status")

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


# ---------------------------------------------------------------- готовый отчёт («Обращения»)
# Клиент просит отменить / вернуть заказ. Ищется в тексте запроса (G) И в теме (D).
CANCEL_RE = re.compile(
    r"(?<!\w)(?:отмен|отказ|возврат|возвращ|вернуть|верните|аннулир|не актуальн|передума)\w*")

# Значения «Итог» (F), при которых отмена считается обработанной верно
F_CANCEL_PHRASES = tuple(norm(p) for p in (
    "Возврат оформлен + ответ клиенту",
    "Отмена невозможна → клиенту дан ответ + оператору",
    "Сбор по заказу ещё не забран → оператору",
    "Отменена часть списка → остальное оператору",
))
# Если ИИ написал это в «Причина / примечание» (H) — тоже считаем отмену обработанной
NOTE_RETURN_PHRASES = tuple(norm(p) for p in (
    "Информация о необходимости возврата внесена в накладную",
    "По данному отправлению возврат уже оформлен ранее, повторная заявка не требуется",
))

# «Итог» (F) для подтверждения отмены: достаточно одного этого значения
F_PROOF_PHRASES = tuple(norm(p) for p in (
    "Запрос подтверждения отмены → ответ клиенту",
))

# Клиент просит изменить / перенести дату доставки (запрос G или тема D)
DATE_CHANGE_RES = tuple(re.compile(p) for p in (
    # «Просьба доставить 30.10», «Простьба перенести доставку на ...», «просьба изменить дату»
    r"(?<!\w)прост?ьб\w*[^.\n]{0,40}?(?:достав|перенес|перенос|измен)",
    # «Изменить дату доставки», «перенесите доставку», «поменять дату доставки»
    r"(?<!\w)(?:измен|перенес|перенос|поменя|сдвин)\w*[^.\n]{0,25}?достав",
    # «доставить 30.10», «доставка на 30.10.2026»
    r"(?<!\w)достав\w*\s+(?:на\s+|до\s+|к\s+)?\d{1,2}[./-]\d{1,2}",
    # «согласован перенос на 28.10», «перенести на 02.11»
    r"(?<!\w)перенос\w*\s+(?:достав\w*\s+)?(?:на|до|с)\s+\d",
    r"(?<!\w)перенес\w*\s+(?:достав\w*\s+)?(?:на|до|с)\s+\d",
))
# «Итог» (F) для даты доставки
F_DATE_PHRASES = tuple(norm(p) for p in (
    "Дата доставки изменена + ответ клиенту",
    "Дата уже стоит в УПД → ответ клиенту по шаблону",
    "Дата недоступна → клиенту отправлен список вариантов",
))


def _has_any(text: str, phrases) -> bool:
    return any(p in text for p in phrases)


def report_verdict(client, subject, note, result):
    """Возвращает (is_correct: bool, rule: str).

    1) Подтверждение отмены: «Итог» = «Запрос подтверждения отмены → ответ клиенту».
    2) Отмена: в запросе клиента ИЛИ в теме есть слова об отмене/отказе/возврате
       И («Итог» = «Возврат оформлен…» / «Отменена часть списка…» ИЛИ в примечании фраза про возврат).
    3) Перенос даты: в запросе ИЛИ теме просьба изменить дату доставки И «Итог» — одно из трёх значений.
    """
    res = norm(result)
    if _has_any(res, F_PROOF_PHRASES):
        return True, "подтверждение отмены"
    text = norm(client) + " \n " + norm(subject)
    if CANCEL_RE.search(text) and (_has_any(res, F_CANCEL_PHRASES) or _has_any(norm(note), NOTE_RETURN_PHRASES)):
        return True, "отмена"
    if any(rx.search(text) for rx in DATE_CHANGE_RES) and _has_any(res, F_DATE_PHRASES):
        return True, "перенос даты"
    return False, ""


def find_columns(ws, spec, required=()):
    """Ищет колонки по заголовку в первой строке. -> ({ключ: номер}, [ненайденные обязательные заголовки])."""
    heads = {}
    for c in ws[1]:
        if c.value not in (None, ""):
            heads.setdefault(norm(c.value), c.column)
    found = {k: heads[norm(h)] for k, h in spec.items() if norm(h) in heads}
    return found, [spec[k] for k in required if k not in found]


def ensure_column(ws, cols, key, header):
    """Если колонки нет — добавляет её в конец с этим заголовком."""
    if key not in cols:
        cols[key] = ws.max_column + 1
        ws.cell(1, cols[key], header)
    return cols[key]


def is_report(wb) -> bool:
    ws = find_sheet(wb, REPORT_SHEET)
    return bool(ws is not None and not find_columns(ws, REPORT_COLS, REPORT_REQUIRED)[1])


def report_date(wb):
    """Дата из листа «Сводка» (ячейка рядом с подписью «Дата») как строка, либо None."""
    ws = find_sheet(wb, SUMMARY_SHEET)
    if ws is None:
        return None
    for row in ws.iter_rows(min_row=1, max_row=5, max_col=2, values_only=True):
        if row[0] is not None and norm(row[0]) == "дата" and row[1] is not None:
            return str(row[1]).strip()
    return None


def _write_ok(ws, row, col, stats):
    """Пишет «Корректно» в пустую ячейку. Чужие значения не трогает. -> True, если записали."""
    cur = ws.cell(row, col).value
    if cur in (None, ""):
        ws.cell(row, col, OK)
        return True
    if norm(cur) != norm(OK):
        stats["occupied"] += 1
    return False


def apply_report(wb, log=print):
    """Лист «Обращения»: «Корректно» в «Комментарий КСЭ» по report_verdict."""
    ws = find_sheet(wb, REPORT_SHEET)
    if ws is None:
        log(f"Лист «{REPORT_SHEET}» не найден.")
        return None
    cols, missing = find_columns(ws, REPORT_COLS, REPORT_REQUIRED)
    if missing:
        log(f"Лист «{REPORT_SHEET}»: нет колонок {', '.join('«'+m+'»' for m in missing)} — пропущено.")
        return None
    ensure_column(ws, cols, "opinion", REPORT_COLS["opinion"])
    get = lambda r, k: ws.cell(r, cols[k]).value if k in cols else None
    stats = {"rows": 0, "marked": 0, "occupied": 0, "by_rule": {}, "left": {}}
    for row in range(2, ws.max_row + 1):
        if get(row, "result") in (None, "") and get(row, "client") in (None, ""):
            continue
        stats["rows"] += 1
        ok, rule = report_verdict(get(row, "client"), get(row, "subject"), get(row, "note"), get(row, "result"))
        if ok:
            if _write_ok(ws, row, cols["opinion"], stats):
                stats["marked"] += 1
                stats["by_rule"][rule] = stats["by_rule"].get(rule, 0) + 1
        elif get(row, "opinion") in (None, ""):
            key = str(get(row, "result") or "—").strip()
            stats["left"][key] = stats["left"].get(key, 0) + 1
    letter = get_column_letter(cols["opinion"])
    log(f"«{REPORT_SHEET}»: строк {stats['rows']}; «{OK}» проставлено {stats['marked']} в колонке {letter} "
        f"({', '.join(f'{k}: {v}' for k, v in stats['by_rule'].items()) or '—'}); "
        f"ячейка уже занята вашим текстом: {stats['occupied']}")
    left = sum(stats["left"].values())
    if left:
        log(f"Осталось для ручной проверки: {left}")
        for k, v in sorted(stats["left"].items(), key=lambda x: -x[1]):
            log(f"   {v} × {k}")
    return stats


# ---------------------------------------------------------------- эскалации
_EMAIL_DOMAIN = re.compile(r"[\w.+-]+@([\w-]+(?:\.[\w-]+)+)")


def email_domain(value):
    m = _EMAIL_DOMAIN.search(str(value or ""))
    return m.group(1).lower() if m else None


_TO_RE = re.compile(r"кому\s*:(.*?)(?=копия\s*:|скрыт\w*\s*:|$)", re.I | re.S)


def to_domains(value):
    """Домены адресов из поля «Кому:» (то, что стоит после «Копия:» не учитывается).
    Если слова «Кому:» нет — берётся вся строка."""
    text = str(value or "")
    m = _TO_RE.search(text)
    part = m.group(1) if m else text
    return {d.lower() for d in _EMAIL_DOMAIN.findall(part)}


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
    # добавлено по отчёту 27.10.2026
    "пятигорск": 3, "грозн": 3, "нальчик": 3, "магнитогорск": 5, "норильск": 7,
    "ангарск": 8, "братск": 8, "артем$": 10, "комсомольск": 10, "находк": 10,
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

    Правило 1: «Кого подключили» = «Управляющий офисом», «Итог» НЕ «Ждёт рабочего времени офиса»,
               домен адреса из «Изначальные адресаты» (Кому) == домен из «Адрес».
    Правило 2 (если возможно): то же, но «Итог» = «Ждёт…», и время из «Время» (UTC+3) в регионе «Офис»
               вне 9:00–18:00.
    """
    if MANAGER not in norm(role_g):
        return False, "не управляющий офисом"
    d_e, d_h = to_domains(email_e), email_domain(email_h)
    if not d_e or not d_h:
        return False, "нет email в «Изначальные адресаты» или «Адрес»"
    if d_h not in d_e:
        return False, "домены различаются"
    if WAITING not in norm(status_i):
        return True, "правило 1"
    t = parse_time(time_v)
    if t is None:
        return False, "правило 2 невозможно: не удалось прочитать время в «Время»"
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
    """Лист «Эскалации»: «Корректно» в «Комментарий КСЭ». Существующие значения не затирает."""
    ws = find_sheet(wb, ESC_SHEET)
    if ws is None:
        log(f"Лист «{ESC_SHEET}» не найден — правила эскалаций пропущены.")
        return {"marked": 0}
    cols, missing = find_columns(ws, ESC_COLS, ESC_REQUIRED)
    if missing:
        log(f"Лист «{ESC_SHEET}»: нет колонок {', '.join('«'+m+'»' for m in missing)} — пропущено.")
        return {"marked": 0}
    ensure_column(ws, cols, "opinion", ESC_COLS["opinion"])
    g = lambda r, k: ws.cell(r, cols[k]).value
    stats = {"marked": 0, "rule1": 0, "rule2": 0, "occupied": 0, "unknown_regions": {}, "bad_time": 0,
             "other_roles": {}}
    for row in range(2, ws.max_row + 1):
        role = g(row, "role")
        if MANAGER not in norm(role):
            if role not in (None, "") and g(row, "opinion") in (None, ""):
                stats["other_roles"][str(role).strip()] = stats["other_roles"].get(str(role).strip(), 0) + 1
            continue
        ok, why = escalation_verdict(g(row, "time"), g(row, "region"), g(row, "to"), role,
                                     g(row, "addr"), g(row, "status"))
        if why.startswith("правило 2 невозможно: неизвестный регион"):
            reg = str(g(row, "region"))
            stats["unknown_regions"][reg] = stats["unknown_regions"].get(reg, 0) + 1
        elif why.startswith("правило 2 невозможно: не удалось"):
            stats["bad_time"] += 1
        if not ok:
            continue
        if _write_ok(ws, row, cols["opinion"], stats):
            stats["marked"] += 1
            stats["rule1" if why == "правило 1" else "rule2"] += 1
    log(f"«{ESC_SHEET}»: «{OK}» — {stats['marked']} (правило 1: {stats['rule1']}, правило 2: {stats['rule2']}); "
        f"ячейка уже занята вашим текстом: {stats['occupied']}; не прочитано время: {stats['bad_time']}")
    if stats["other_roles"]:
        log("Не проверяются автоматически (правило только для «Управляющий офисом»): "
            + ", ".join(f"{k} ×{v}" for k, v in stats["other_roles"].items()))
    if stats["unknown_regions"]:
        log("Регионы без известного часового пояса (правило 2 для них пропущено; добавьте в REGION_UTC): "
            + ", ".join(f"{k} ×{v}" for k, v in stats["unknown_regions"].items()))
    return stats
