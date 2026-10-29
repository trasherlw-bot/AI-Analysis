import sys
import unittest
from datetime import datetime, time as dtime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "app"))
import hde_rules as r  # noqa: E402
from openpyxl import Workbook  # noqa: E402

F_CANCEL = "Отменена часть списка → остальное оператору"
RET1 = "Информация о необходимости возврата внесена в накладную."
RET2 = "По данному отправлению возврат уже оформлен ранее, повторная заявка не требуется"


class MainRules(unittest.TestCase):
    def test_cancel_ai_phrase(self):
        for c in ("Отмена заявки", "Отмените пожалуйста", "нужен ВОЗВРАТ груза", "Просим отменить"):
            self.assertTrue(r.is_correct_cancel(c, RET1, ""), c)
            self.assertTrue(r.is_correct_cancel(c, "Ответ. " + RET2, ""), c)

    def test_cancel_f_column(self):
        self.assertTrue(r.is_correct_cancel("Отмена", "что-то другое", F_CANCEL))
        # стрелка в другом написании и лишние пробелы
        self.assertTrue(r.is_correct_cancel("Отмена", "", "Отменена часть списка->остальное  оператору"))

    def test_cancel_negative(self):
        self.assertFalse(r.is_correct_cancel("Где моя посылка?", RET1, F_CANCEL))   # клиент не просит отмену
        self.assertFalse(r.is_correct_cancel("Отмена", "Принято", "Дата доставки изменена + ответ клиенту"))
        self.assertFalse(r.is_correct_cancel("", RET1, F_CANCEL))

    def test_date_change(self):
        f = "Дата доставки изменена + ответ клиенту"
        for c in ("Простьба доставить 30.10.2026", "Просьба перенести доставку на 02.11",
                  "Просьба доставить завтра", "Изменить дату доставки на пятницу",
                  "Прошу изменить дату доставки", "Перенесите доставку, пожалуйста",
                  "доставить 31.10"):
            self.assertTrue(r.is_correct_date_change(c, f), c)
        for f2 in ("Дата уже стоит в УПД → ответ клиенту по шаблону",
                   "Дата недоступна → клиенту отправлен список вариантов"):
            self.assertTrue(r.is_correct_date_change("Просьба доставить 30.10", f2), f2)

    def test_date_change_negative(self):
        f = "Дата доставки изменена + ответ клиенту"
        self.assertFalse(r.is_correct_date_change("Где мой груз?", f))
        self.assertFalse(r.is_correct_date_change("Просьба доставить 30.10", "Другое"))
        self.assertFalse(r.is_correct_date_change("Просьба доставить 30.10", ""))

    def test_is_correct_main_none_values(self):
        self.assertFalse(r.is_correct_main(None, None, None))


class Escalations(unittest.TestCase):
    def test_domain(self):
        self.assertEqual(r.email_domain("test@skx.cse.ru"), "skx.cse.ru")
        self.assertEqual(r.email_domain("Иван <Test2@SKX.cse.ru>"), "skx.cse.ru")
        self.assertIsNone(r.email_domain("нет почты"))
        self.assertIsNone(r.email_domain(None))

    def test_rule1_same_domain(self):
        ok, why = r.escalation_verdict("10:00", "Москва", "test@skx.cse.ru", "Управляющий офисом",
                                       "test2@skx.cse.ru", "В работе")
        self.assertTrue(ok)
        self.assertEqual(why, "правило 1")

    def test_rule1_different_domain(self):
        ok, _ = r.escalation_verdict("10:00", "Москва", "a@skx.cse.ru", "Управляющий офисом", "b@other.ru", "")
        self.assertFalse(ok)

    def test_rule1_not_manager(self):
        ok, _ = r.escalation_verdict("10:00", "Москва", "a@x.ru", "Менеджер", "b@x.ru", "")
        self.assertFalse(ok)

    def test_rule2_chelyabinsk_example(self):
        # из задания: B=17:18 (UTC+3), Челябинск (+5) → 19:18 → вне 9–18 → корректно
        ok, why = r.escalation_verdict("17:18", "Челябинск", "a@skx.cse.ru", "Управляющий офисом",
                                       "b@skx.cse.ru", "Ждёт рабочего времени офиса")
        self.assertTrue(ok)
        self.assertEqual(why, "правило 2")

    def test_rule2_inside_hours(self):
        ok, why = r.escalation_verdict("12:00", "Челябинск", "a@skx.cse.ru", "Управляющий офисом",
                                       "b@skx.cse.ru", "Ждёт рабочего времени офиса")
        self.assertFalse(ok)  # 14:00 местного
        self.assertIn("рабочее", why)

    def test_rule2_boundaries(self):
        f = lambda t, reg: r.escalation_verdict(t, reg, "a@x.ru", "Управляющий офисом", "b@x.ru",
                                                "ждет рабочего времени офиса")[0]
        self.assertFalse(f("09:00", "Москва"))     # ровно 9:00 — рабочее
        self.assertTrue(f("08:59", "Москва"))
        self.assertTrue(f("18:00", "Москва"))      # 18:00 уже нерабочее
        self.assertFalse(f("17:59", "Москва"))
        self.assertTrue(f("23:30", "Владивосток"))  # +7ч → 06:30 следующего дня
        self.assertFalse(f("03:00", "Владивосток"))  # → 10:00
        self.assertTrue(f("08:00", "Калининград"))   # 07:00 местного

    def test_rule2_unknown_region_and_bad_time(self):
        a = ("a@x.ru", "Управляющий офисом", "b@x.ru", "Ждёт рабочего времени офиса")
        ok, why = r.escalation_verdict("17:00", "Атлантида", *a)
        self.assertFalse(ok)
        self.assertIn("неизвестный регион", why)
        ok, why = r.escalation_verdict("не время", "Москва", *a)
        self.assertFalse(ok)

    def test_time_parsing(self):
        self.assertEqual(r.parse_time("29.10.2026 17:18"), dtime(17, 18))
        self.assertEqual(r.parse_time(datetime(2026, 10, 29, 9, 5)), dtime(9, 5))
        self.assertEqual(r.parse_time(dtime(7, 0)), dtime(7, 0))
        self.assertEqual(r.parse_time(0.75), dtime(18, 0))
        self.assertIsNone(r.parse_time(""))
        self.assertIsNone(r.parse_time("25:99"))

    def test_regions(self):
        cases = {"Челябинск": 5, "Челябинская область": 5, "Томск": 7, "Омск": 6, "Москва": 3,
                 "Нижневартовск": 5, "Нижний Новгород": 3, "Киров": 3, "Кировград": None,
                 "Уфа": 5, "Сахалинская область": 11, "Республика Саха (Якутия)": 9, "Чита": 9, "": None}
        for reg, off in cases.items():
            self.assertEqual(r.region_offset(reg), off, reg)

    def test_apply_sheet(self):
        wb = Workbook()
        ws = wb.active
        ws.title = "Эскалации"
        ws.append(["h"] * 11)
        # B  C  D  E  F  G  H  I  J  K
        rows = [
            (None, "10:00", None, "Москва", "a@skx.cse.ru", None, "Управляющий офисом", "b@skx.cse.ru", "", None, None),   # 1 → K
            (None, "10:00", None, "Москва", "a@skx.cse.ru", None, "Управляющий офисом", "b@zzz.ru", "", None, None),      # домен
            (None, "17:18", None, "Челябинск", "a@skx.cse.ru", None, "Управляющий офисом", "b@skx.cse.ru", "Ждёт рабочего времени офиса", None, None),  # 2 → K
            (None, "12:00", None, "Челябинск", "a@skx.cse.ru", None, "Управляющий офисом", "b@skx.cse.ru", "Ждёт рабочего времени офиса", None, None),  # рабочее
            (None, "10:00", None, "Москва", "a@skx.cse.ru", None, "Менеджер", "b@skx.cse.ru", "", None, None),            # не управляющий
            (None, "10:00", None, "Москва", "a@skx.cse.ru", None, "Управляющий офисом", "b@skx.cse.ru", "", None, "Ручная пометка"),  # K занята
        ]
        for i, row in enumerate(rows, start=2):
            for j, v in enumerate(row, start=1):
                if v is not None:
                    ws.cell(i, j, v)
        # сдвиг: в rows первая None = колонка A, дальше B..; проверяем соответствие букв
        self.assertEqual(ws["B2"].value, "10:00")
        self.assertEqual(ws["D2"].value, "Москва")
        self.assertEqual(ws["E2"].value, "a@skx.cse.ru")
        self.assertEqual(ws["G2"].value, "Управляющий офисом")
        self.assertEqual(ws["H2"].value, "b@skx.cse.ru")
        logs = []
        st = r.apply_escalations(wb, logs.append)
        self.assertEqual([ws.cell(i, 11).value for i in range(2, 8)],
                         [r.OK, None, r.OK, None, None, "Ручная пометка"])
        self.assertEqual((st["marked"], st["rule1"], st["rule2"], st["occupied"]), (2, 1, 1, 1))

    def test_missing_sheet(self):
        wb = Workbook()
        self.assertEqual(r.apply_escalations(wb, lambda m: None), {"marked": 0})


if __name__ == "__main__":
    unittest.main()
