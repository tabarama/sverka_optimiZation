# -*- coding: utf-8 -*-
"""
sverka.py — автоматическая «сверка» отчётов о продажах в Excel.

Скрипт кладётся в папку рядом с исходными .xlsx и при запуске:
  • обрабатывает каждый  X.xlsx  →  создаёт  сверка_X.xlsx  в той же папке;
  • перемещает нетронутый оригинал  X.xlsx  в подпапку  old/;
  • файлы с префиксом  сверка_  и временные файлы Excel (~$...) пропускает;
  • если  сверка_X.xlsx  уже существует — X пропускается целиком.

Для каждого файла воспроизводится ручной алгоритм сверки:
  1. Последний (пустой) столбец отчёта получает заголовок «Комментарий».
  2. Данные превращаются в умную таблицу «Таблица1».
  3. Справа добавляется вычисляемый столбец «ЛПР+ЮЛ» = «КВ ЛПР Итого» + «КВ ЮЛ Итого».
  4. Обновление данных (Обновить всё).
  5. Строятся три сводные таблицы:
       — Партнёры/ЛПР;
       — Агенты;
       — Операторы (классический макет, без промежуточных итогов, кроме «Агент ЮЛ»).
  6. Автоподбор ширины столбцов области сводных.
  7. Сохранение в  сверка_X.xlsx.

ТРЕБОВАНИЯ: Windows + установленный Microsoft Excel + пакет pywin32.
Установка зависимости:   pip install pywin32
(Желательно один раз выполнить, чтобы ускорить работу:
    python -m win32com.client.makepy "Microsoft Excel 16.0 Object Library")

Запуск:   python sverka.py
Лог работы пишется в файл  sverka.log  рядом со скриптом.
"""

import os
import sys
import gc
import shutil
import logging
import subprocess
from pathlib import Path

# win32com доступен только на Windows. Импортируем аккуратно, чтобы файл можно было
# хотя бы синтаксически проверить на другой ОС; реальная работа требует Windows+Excel.
try:
    import win32com.client as _wc
    import pythoncom as _pythoncom
    _WIN32_ERR = None
except Exception as _imp_err:            # pragma: no cover  (не-Windows окружение)
    _wc = None
    _pythoncom = None
    _WIN32_ERR = _imp_err


# ---------------------------------------------------------------------------
# Константы
# ---------------------------------------------------------------------------

SVERKA_PREFIX = "сверка_"                 # кириллица!
OLD_DIRNAME = "old"
LOG_NAME = "sverka.log"

TABLE_NAME = "Таблица1"
TABLE_STYLE = "TableStyleMedium9"
MONEY_FMT = "#,##0.00"

# Заголовки столбцов (точная кириллица). Используются для поиска и для записи.
H_ID_TO      = "ID ТО"
H_UL         = "Юридическое лицо"
H_AGENT_UL   = "Агент ЮЛ"
H_ID_OFORM   = "ID Оформителя"
H_OFORMIL    = "Оформил заявку"
H_KV_OPER    = "КВ Оператор Итого"
H_OPER_RANEE = "Оператор оплачено ранее"
H_KV_AGENT   = "КВ Агент Итого"
H_KV_UL      = "КВ ЮЛ Итого"
H_KV_LPR     = "КВ ЛПР Итого"
H_COMMENT    = "Комментарий"
H_LPR_UL     = "ЛПР+ЮЛ"

# Обязательные заголовки — если хоть одного нет, файл не обрабатываем.
REQUIRED_HEADERS = [
    H_ID_TO, H_UL, H_AGENT_UL, H_ID_OFORM, H_OFORMIL,
    H_KV_OPER, H_OPER_RANEE, H_KV_AGENT, H_KV_UL, H_KV_LPR,
]

# Значения перечислений Excel (заданы явно — не полагаемся на win32com.client.constants,
# который пуст без сгенерированной обёртки типов).
XL_SRC_RANGE        = 1        # xlSrcRange
XL_YES              = 1        # xlYes
XL_DATABASE         = 1        # xlDatabase (тип источника кэша сводной)
XL_ROW_FIELD        = 1        # xlRowField
XL_PAGE_FIELD       = 3        # xlPageField
XL_SUM              = -4157    # xlSum
XL_TABULAR_ROW      = 1        # xlTabularRow (аргумент RowAxisLayout)
XL_OPENXML_WORKBOOK = 51       # xlOpenXMLWorkbook (.xlsx)
XL_UP               = -4162    # xlUp

log = logging.getLogger("sverka")


# ---------------------------------------------------------------------------
# Вспомогательные функции
# ---------------------------------------------------------------------------

def norm(value):
    """Нормализация текста заголовка/значения для сравнения:
    None → '', неразрывные пробелы → обычные, схлопывание пробелов, trim, casefold."""
    if value is None:
        return ""
    s = str(value).replace("\xa0", " ").replace(" ", " ").replace("\t", " ")
    s = " ".join(s.split())
    return s.strip().casefold()


def _num(x):
    """Число из значения ячейки (None/текст → 0.0, где возможно распарсить)."""
    if x is None:
        return 0.0
    if isinstance(x, bool):
        return 0.0
    if isinstance(x, (int, float)):
        return float(x)
    s = str(x).strip().replace("\xa0", "").replace(" ", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0


def _is_zero(v):
    """True, если значение фильтра означает ноль (0 / 0,00 / 0.00 / «0»)."""
    if v is None or isinstance(v, bool):
        return False
    if isinstance(v, (int, float)):
        return abs(v) < 1e-9
    s = str(v).strip().replace("\xa0", "").replace(" ", "").replace(",", ".")
    try:
        return abs(float(s)) < 1e-9
    except ValueError:
        return s in ("0", "0.0", "0.00")


def setup_logging(folder):
    log.setLevel(logging.INFO)
    log.handlers.clear()
    fmt = logging.Formatter("%(asctime)s  %(levelname)-7s  %(message)s", "%Y-%m-%d %H:%M:%S")
    try:
        fh = logging.FileHandler(str(folder / LOG_NAME), encoding="utf-8")
        fh.setFormatter(fmt)
        log.addHandler(fh)
    except Exception:
        pass
    ch = logging.StreamHandler()
    ch.setFormatter(fmt)
    log.addHandler(ch)


# ---------------------------------------------------------------------------
# Оркестрация файловой системы
# ---------------------------------------------------------------------------

def should_skip(path, folder):
    """(skip: bool, причина: str) — стоит ли пропустить файл."""
    name = path.name
    if name.startswith("~$"):
        return True, "временный файл Excel"
    if name.startswith(SVERKA_PREFIX):
        return True, "уже файл сверки"
    if path.parent.name == OLD_DIRNAME:
        return True, "внутри old/"
    if (folder / (SVERKA_PREFIX + name)).exists():
        return True, "сверка_%s уже существует" % name
    return False, ""


def discover_files(folder):
    """Список .xlsx в папке, подлежащих обработке (нерекурсивно, с учётом правил пропуска)."""
    result = []
    for p in sorted(folder.glob("*.xlsx")):
        skip, why = should_skip(p, folder)
        if skip:
            log.info("Пропуск: %s (%s)", p.name, why)
        else:
            result.append(p)
    return result


def ensure_old_dir(folder):
    old = folder / OLD_DIRNAME
    old.mkdir(exist_ok=True)
    return old


def move_to_old(src, old_dir):
    """Переместить оригинал в old/. При совпадении имени не затираем прежний
    резерв, а добавляем суффикс ' (N)'."""
    dest = old_dir / src.name
    if dest.exists():
        stem, suf = src.stem, src.suffix
        i = 1
        while True:
            cand = old_dir / ("%s (%d)%s" % (stem, i, suf))
            if not cand.exists():
                dest = cand
                break
            i += 1
    shutil.move(str(src), str(dest))
    return dest


# ---------------------------------------------------------------------------
# Жизненный цикл Excel
# ---------------------------------------------------------------------------

def make_excel():
    """Запуск отдельного (не пользовательского) экземпляра Excel. Возвращает (app, pid)."""
    if _wc is None:
        raise RuntimeError(
            "Не удалось импортировать win32com (pywin32). Скрипт работает только на "
            "Windows с установленным Excel. Установите зависимость: pip install pywin32. "
            "Исходная ошибка импорта: %s" % _WIN32_ERR
        )
    _pythoncom.CoInitialize()
    app = _wc.DispatchEx("Excel.Application")   # отдельный процесс — не трогаем открытый Excel
    app.Visible = False
    app.DisplayAlerts = False
    app.ScreenUpdating = False
    app.EnableEvents = False
    try:
        app.AskToUpdateLinks = False
    except Exception:
        pass
    pid = None
    try:
        import win32process
        _, pid = win32process.GetWindowThreadProcessId(app.Hwnd)
    except Exception:
        pid = None
    return app, pid


def quit_excel(app):
    if app is None:
        return
    for setter in (
        lambda: setattr(app, "ScreenUpdating", True),
        lambda: setattr(app, "DisplayAlerts", True),
        lambda: setattr(app, "EnableEvents", True),
    ):
        try:
            setter()
        except Exception:
            pass
    try:
        app.Quit()
    except Exception:
        pass
    try:
        _pythoncom.CoUninitialize()
    except Exception:
        pass
    gc.collect()


def kill_pid(pid):
    """Последнее средство: снять зависший процесс Excel по PID (только если это EXCEL.EXE)."""
    if not pid:
        return
    try:
        subprocess.run(
            ["taskkill", "/F", "/PID", str(pid), "/FI", "IMAGENAME eq EXCEL.EXE"],
            capture_output=True, check=False,
        )
    except Exception:
        pass


# ---------------------------------------------------------------------------
# Чтение заголовков / геометрии листа
# ---------------------------------------------------------------------------

def read_headers(ws):
    """Возвращает (header_row, by_norm, by_col) для листа.
       header_row — абсолютный номер строки заголовков (ищется в первых 5 строках по «ID ТО»);
       by_norm    — {норм. заголовок → абсолютный номер столбца};
       by_col     — {абсолютный номер столбца → исходный текст заголовка}.
    Бросает исключение, если «ID ТО» не найден."""
    ur = ws.UsedRange
    base_row = ur.Row
    base_col = ur.Column
    vals = ur.Value
    if vals is None:
        raise ValueError("пустой лист")
    if not isinstance(vals, (tuple, list)):
        vals = ((vals,),)
    elif not isinstance(vals[0], (tuple, list)):
        vals = (vals,)

    target = norm(H_ID_TO)
    for i, row in enumerate(vals[:5]):
        cells = row if isinstance(row, (tuple, list)) else (row,)
        if any(norm(c) == target for c in cells):
            header_row = base_row + i
            by_norm, by_col = {}, {}
            for j, c in enumerate(cells):
                n = norm(c)
                if n:
                    col = base_col + j
                    by_norm.setdefault(n, col)
                    by_col[col] = str(c)
            return header_row, by_norm, by_col
    raise ValueError("не найден заголовок «%s» в первых 5 строках" % H_ID_TO)


def find_data_sheet(wb):
    """Первый лист с данными (в заголовке которого есть «ID ТО»). Обычно единственный."""
    last_err = None
    for i in range(1, wb.Worksheets.Count + 1):
        ws = wb.Worksheets(i)
        try:
            header_row, by_norm, by_col = read_headers(ws)
            return ws, header_row, by_norm, by_col
        except Exception as e:
            last_err = e
    raise ValueError("не найден лист с данными (заголовок «%s»): %s" % (H_ID_TO, last_err))


def require_headers(by_norm):
    missing = [h for h in REQUIRED_HEADERS if norm(h) not in by_norm]
    if missing:
        raise ValueError("отсутствуют обязательные столбцы: " + ", ".join(missing))


# ---------------------------------------------------------------------------
# Шаги трансформации
# ---------------------------------------------------------------------------

def ensure_comment_column(ws, header_row, by_norm, by_col):
    """Шаг 1. Определяет столбец комментариев (последний столбец отчёта) и, если нужно,
    пишет ему заголовок «Комментарий». Возвращает индекс столбца комментариев."""
    if norm(H_COMMENT) in by_norm:                      # уже есть «Комментарий»
        return by_norm[norm(H_COMMENT)]
    last_hdr_col = max(by_col.keys())
    if norm(by_col[last_hdr_col]).startswith("коммент"):  # назван похоже — переименуем строго
        ws.Cells(header_row, last_hdr_col).Value = H_COMMENT
        return last_hdr_col
    # последний заголовок — обычный столбец данных → комментарий идёт следующим (пустой столбец)
    comment_col = last_hdr_col + 1
    ws.Cells(header_row, comment_col).Value = H_COMMENT
    return comment_col


def create_table(ws, header_row, left_col, right_col, last_row):
    """Шаг 2. Создаёт (или переиспользует) умную таблицу «Таблица1»."""
    # Если таблица уже есть на листе — переиспользуем первую.
    try:
        if ws.ListObjects.Count >= 1:
            lo = ws.ListObjects(1)
            lo.Name = TABLE_NAME
            try:
                lo.TableStyle = TABLE_STYLE
            except Exception:
                pass
            return lo
    except Exception:
        pass
    rng = ws.Range(ws.Cells(header_row, left_col), ws.Cells(last_row, right_col))
    lo = ws.ListObjects.Add(XL_SRC_RANGE, rng, None, XL_YES)
    lo.Name = TABLE_NAME
    try:
        lo.TableStyle = TABLE_STYLE
    except Exception:
        pass
    return lo


def _column_values(rng):
    """Значения одностолбцового диапазона → плоский список."""
    v = rng.Value
    if v is None:
        return []
    if not isinstance(v, (tuple, list)):
        return [v]
    out = []
    for row in v:
        out.append(row[0] if isinstance(row, (tuple, list)) else row)
    return out


def add_lpr_ul_column(app, ws, lo, by_norm, by_col):
    """Шаг 3. Добавляет справа вычисляемый столбец «ЛПР+ЮЛ» = КВ ЛПР Итого + КВ ЮЛ Итого."""
    newcol = lo.ListColumns.Add()
    newcol.Name = H_LPR_UL
    body = newcol.DataBodyRange

    lpr_name = by_col.get(by_norm[norm(H_KV_LPR)], H_KV_LPR)
    ul_name = by_col.get(by_norm[norm(H_KV_UL)], H_KV_UL)
    formula = "=%s[[#This Row],[%s]]+%s[[#This Row],[%s]]" % (
        TABLE_NAME, lpr_name, TABLE_NAME, ul_name)
    try:
        body.Formula = formula
        app.Calculate()
    except Exception as e:
        log.warning("Формула «ЛПР+ЮЛ» не применилась (%s) — записываю значениями", e)
        first = lo.HeaderRowRange.Row + 1
        last = lo.Range.Row + lo.Range.Rows.Count - 1
        lpr_col = by_norm[norm(H_KV_LPR)]
        ul_col = by_norm[norm(H_KV_UL)]
        lpr_vals = _column_values(ws.Range(ws.Cells(first, lpr_col), ws.Cells(last, lpr_col)))
        ul_vals = _column_values(ws.Range(ws.Cells(first, ul_col), ws.Cells(last, ul_col)))
        out = tuple((_num(a) + _num(b),) for a, b in zip(lpr_vals, ul_vals))
        body.Value = out
    try:
        body.NumberFormat = MONEY_FMT
    except Exception:
        pass
    return newcol


def refresh_all(app, wb):
    """Шаг 4. «Обновить всё» + дождаться пересчёта."""
    try:
        wb.RefreshAll()
    except Exception:
        pass
    try:
        app.CalculateUntilAsyncQueriesDone()
    except Exception:
        pass
    try:
        app.Calculate()
    except Exception:
        pass


# --- построители сводных таблиц ---
#
# Настройка полей выполняется на «живой» сводной (pt.ManualUpdate = False).
# Порядок строгий и повторяет действия человека: строки → ЗНАЧЕНИЯ → фильтры
# (страничные поля). Это принципиально:
#   • страничное поле нельзя поставить на «скелет» сводной без значений —
#     Excel бросает «Нельзя установить свойство Orientation класса PivotField»
#     (так срывался фильтр «ЛПР+ЮЛ»);
#   • для поля, которое одновременно и значение, и фильтр («КВ Оператор Итого»),
#     сначала добавляем его как значение, ЗАТЕМ делаем страничным — так
#     сохраняются обе роли (как в ручном файле: axis="axisPage" + dataField="1").
#     При обратном порядке AddDataField снимал роль фильтра.

def _pivot_items(pf):
    """Список пунктов поля (через .Item(k) — работает и при раннем связывании COM)."""
    coll = pf.PivotItems()
    return [coll.Item(k) for k in range(1, coll.Count + 1)]


def _add_value(pt, source_name):
    df = pt.AddDataField(pt.PivotFields(source_name), "Сумма по полю " + source_name, XL_SUM)
    try:
        df.NumberFormat = MONEY_FMT
    except Exception:
        pass
    return df


def _is_page(pt, name):
    try:
        return pt.PivotFields(name).Orientation == XL_PAGE_FIELD
    except Exception:
        return False


def _set_page_orientation(pt, name):
    """Сделать поле страничным (область фильтров).

    Способ 1 — прямое присвоение Orientation: надёжно для ТЕКСТОВЫХ полей
    («Оператор оплачено ранее») и для поля, уже добавленного как значение
    («КВ Оператор Итого»).

    Способ 2 — метод AddFields(PageFields=...): это ровно то, что делает
    перетаскивание поля в область «Фильтры» в интерфейсе Excel. Другой путь в
    COM, чем присвоение Orientation, и он проходит там, где Orientation срывается
    («Нельзя установить свойство Orientation класса PivotField») — а именно на
    «свежем» числовом поле-фильтре («ЛПР+ЮЛ»).

    Возвращает True только если поле действительно стало страничным."""
    errs = []

    # Способ 1 — прямое присвоение Orientation.
    try:
        pt.PivotFields(name).Orientation = XL_PAGE_FIELD
        if _is_page(pt, name):
            return True
    except Exception as e:
        errs.append("orient:%r" % (e,))

    # Способ 2 — AddFields (UI-эквивалент перетаскивания в «Фильтры»).
    # AddToTable=True: добавляем страничное поле, НЕ трогая строки/значения.
    try:
        pt.AddFields(PageFields=name, AddToTable=True)          # именованные аргументы
        if _is_page(pt, name):
            return True
    except Exception as e:
        errs.append("addfields_kw:%r" % (e,))
        try:
            pt.AddFields(None, None, name, True)                # позиционно (запасной путь)
            if _is_page(pt, name):
                return True
        except Exception as e2:
            errs.append("addfields_pos:%r" % (e2,))

    log.error("ФИЛЬТР НЕ УСТАНОВЛЕН: поле «%s» не стало страничным (%s)",
              name, " | ".join(errs) or "неизвестно")
    return False


def _hide_zero_item(pf):
    """Скрыть в фильтре пункт «0» (показать всё, кроме нуля). Нельзя скрыть все пункты."""
    zero_item = None
    visible_non_zero = 0
    for it in _pivot_items(pf):
        try:
            if _is_zero(it.Value):
                zero_item = it
            else:
                it.Visible = True
                visible_non_zero += 1
        except Exception:
            pass
    if zero_item is not None and visible_non_zero >= 1:
        try:
            zero_item.Visible = False
        except Exception:
            pass


def _show_only(pf, wanted):
    """Оставить в СТРАНИЧНОМ поле единственный выбранный пункт == wanted.

    Одиночный выбор делаем через CurrentPage, а НЕ переключением .Visible у пунктов.
    Если выбирать пункт через .Visible, Excel включает режим множественного выбора
    (EnableMultiplePageItems) и в самом фильтре показывает «(несколько элементов)»,
    даже когда реально виден один пункт. CurrentPage даёт одиночный выбор — в фильтре
    отображается подпись самого пункта («Нет»).

    Запасной путь (если CurrentPage не сработал) — старый способ через .Visible:
    подпись может стать «(несколько элементов)», но данные отфильтруются правильно."""
    target = norm(wanted)
    exact = None
    for it in _pivot_items(pf):
        try:
            if norm(it.Value) == target:
                exact = it.Value                # точная подпись пункта (реальный регистр/пробелы)
                break
        except Exception:
            pass
    if exact is None:
        raise ValueError("пункт «%s» не найден в поле «%s»" % (wanted, pf.Name))

    # Основной путь: одиночный выбор через CurrentPage.
    try:
        try:
            pf.EnableMultiplePageItems = False   # иначе CurrentPage недоступен
        except Exception:
            pass
        pf.CurrentPage = str(exact)
        return
    except Exception as e:
        log.warning("CurrentPage для «%s» не сработал (%s) — переключаюсь на .Visible", pf.Name, e)

    # Запасной путь: показать нужный пункт, скрыть остальные.
    items = _pivot_items(pf)
    for it in items:
        try:
            if norm(it.Value) == target:
                it.Visible = True
        except Exception:
            pass
    for it in items:
        try:
            if norm(it.Value) != target:
                it.Visible = False
        except Exception:
            pass


def _make_pivot(cache, dest, name, rows, values, pages, classic=False, keep_subtotal=None):
    """Универсальный построитель сводной.
       rows     — список полей строк (по порядку);
       values   — список полей значений (сумма);
       pages    — список (поле, режим, аргумент), режим: 'exclude_zero' | 'only';
       classic  — классический (табличный) макет + перетаскивание в сетке;
       keep_subtotal — единственное поле строк, у которого оставить промежуточные итоги."""
    pt = cache.CreatePivotTable(dest, name)
    # Работаем на «живой» сводной — так установка страничных полей надёжна.
    try:
        pt.ManualUpdate = False
    except Exception:
        pass

    # 1) Поля строк (строго по возрастанию позиции)
    for i, nm in enumerate(rows, start=1):
        f = pt.PivotFields(nm)
        f.Orientation = XL_ROW_FIELD
        f.Position = i

    # 2) Значения (сумма) — ДО фильтров. Значения «материализуют» сводную,
    #    иначе установка страничного поля падает с ошибкой Orientation.
    for nm in values:
        _add_value(pt, nm)

    # 3) Поля фильтров — ПОСЛЕ значений. Для поля, которое одновременно значение
    #    и фильтр («КВ Оператор Итого»), это сохраняет обе роли.
    for nm, _mode, _arg in pages:
        _set_page_orientation(pt, nm)

    # 4) Классический (табличный) макет
    if classic:
        try:
            pt.RowAxisLayout(XL_TABULAR_ROW)
        except Exception as e:
            log.warning("Табличный макет «%s» не применился: %s", name, e)
        try:
            pt.InGridDropZones = True
        except Exception as e:
            log.warning("Классический макет «%s» не применился: %s", name, e)

    # 5) Промежуточные итоги: убрать у всех полей строк, кроме keep_subtotal
    if keep_subtotal is not None:
        no_sub = tuple([False] * 12)
        for nm in rows:
            if nm != keep_subtotal:
                try:
                    pt.PivotFields(nm).Subtotals = no_sub
                except Exception as e:
                    log.warning("Не удалось убрать итоги у «%s»: %s", nm, e)

    # 6) Применяем фильтры по пунктам (на «живой» сводной пункты доступны)
    for nm, mode, arg in pages:
        try:
            pf = pt.PivotFields(nm)
            if mode == "exclude_zero":
                _hide_zero_item(pf)
            elif mode == "only":
                _show_only(pf, arg)
        except Exception as e:
            log.warning("Фильтр «%s» не применился: %s", nm, e)

    # 7) Финальный пересчёт — фиксируем итоговую геометрию сводной
    try:
        pt.Update()
    except Exception:
        pass
    return pt


def _pivot_partners(cache, ws, dest):
    """Сводная 1 — Партнёры/ЛПР."""
    return _make_pivot(
        cache, dest, "свПартнеры",
        rows=[H_UL, H_ID_TO, H_COMMENT],
        values=[H_KV_UL, H_KV_LPR],
        pages=[(H_LPR_UL, "exclude_zero", None)],      # фильтр ЛПР+ЮЛ: всё, кроме 0,00
    )


def _pivot_agents(cache, ws, dest):
    """Сводная 2 — Агенты."""
    return _make_pivot(
        cache, dest, "свАгенты",
        rows=[H_AGENT_UL],
        values=[H_KV_AGENT],
        pages=[],
    )


def _pivot_operators(cache, ws, dest):
    """Сводная 3 — Операторы (классический макет, итоги только у «Агент ЮЛ»)."""
    return _make_pivot(
        cache, dest, "свОператоры",
        rows=[H_AGENT_UL, H_OFORMIL, H_ID_OFORM, H_UL, H_COMMENT],
        values=[H_KV_OPER],
        pages=[
            (H_KV_OPER, "exclude_zero", None),         # КВ Оператор Итого: всё, кроме 0,00
            (H_OPER_RANEE, "only", "Нет"),             # Оператор оплачено ранее = «Нет»
        ],
        classic=True,
        keep_subtotal=H_AGENT_UL,
    )


def _autofit_pivots(ws, start_row, pivots):
    """Шаг 6. Автоподбор ширины столбцов области сводных таблиц."""
    try:
        right = 1
        bottom = start_row
        for pt in pivots:
            r = pt.TableRange2
            right = max(right, r.Column + r.Columns.Count - 1)
            bottom = max(bottom, r.Row + r.Rows.Count - 1)
        ws.Range(ws.Cells(start_row, 1), ws.Cells(bottom, right)).EntireColumn.AutoFit()
    except Exception as e:
        log.warning("Автоподбор ширины не выполнен: %s", e)


def build_pivots(app, wb, ws, lo):
    """Шаг 5. Один кэш → три сводные, размещённые под таблицей.

    Раскладка сверху вниз: ВЕРХ — Партнёры/ЛПР (столбец B) и Агенты (справа от них);
    НИЖЕ — Операторы.

    Между умной таблицей и верхней сводной оставляем фиксированный зазор в
    4 пустые строки. Зазор обязателен из-за страничных полей (фильтров): их строки
    Excel вставляет НАД телом сводной (строки фильтров + пустая строка-разделитель).
    Если поставить сводную вплотную к «Таблица1», этим строкам некуда встать, и
    попытка сделать поле страничным срывается ошибкой «Не допускается перекрытие
    отчёта сводной таблицы и таблицы» — так у верхней сводной срывался фильтр (у
    Партнёров — «ЛПР+ЮЛ»; у Операторов при диагностическом обмене — «КВ Оператор
    Итого» и «Оператор оплачено ранее»). Диагностика (обмен верхней и нижней
    сводных местами) подтвердила: проблема чисто позиционная — фильтры срывались у
    той сводной, что стояла вплотную к таблице. Тот же зазор в 4 строки давно и
    надёжно работает у нижней сводной (bottom + 5).
    """
    cache = wb.PivotCaches().Create(XL_DATABASE, TABLE_NAME)
    tbl = lo.Range
    tbl_bottom = tbl.Row + tbl.Rows.Count - 1
    # Верхняя сводная на 5 строк ниже таблицы = 4 пустые строки-зазор над ней.
    # Было tbl_bottom + 2 (лишь 1 пустая строка) → строке фильтра некуда встать,
    # Excel бросал «перекрытие» и фильтр «ЛПР+ЮЛ» срывался. +3 строки к прежнему
    # отступу гарантированно освобождают место под страничное поле верхней сводной.
    start_row = tbl_bottom + 5
    col_b = 2

    # Сводная 1 (ВЕРХ) — Партнёры/ЛПР (столбец B). Строится и ФИЛЬТРУЕТСЯ полностью,
    # только потом читаем её реальную (уже сжатую фильтром) геометрию.
    p1 = _pivot_partners(cache, ws, ws.Cells(start_row, col_b))
    ext1 = p1.TableRange2                               # весь диапазон, включая фильтр
    body_top = p1.TableRange1.Row                       # верх тела (под строкой фильтра)
    p2_col = ext1.Column + ext1.Columns.Count + 1       # справа, через один пустой столбец

    # Сводная 2 — Агенты (справа, верх тела выровнен с телом верхней сводной)
    p2 = _pivot_agents(cache, ws, ws.Cells(body_top, p2_col))

    # Сводная 3 (НИЗ) — Операторы. Позицию считаем ПОСЛЕ того, как верхние
    # сводные полностью построены и отфильтрованы — перечитываем их низ заново.
    # Первая пустая строка под сводными и три следующие пропускаются → сводная
    # начинается на (низ + 5). Здесь над сводной те же 4 пустые строки, чтобы
    # страничные поля «КВ Оператор Итого» и «Оператор оплачено ранее» встали без
    # перекрытия.
    ext1 = p1.TableRange2
    ext2 = p2.TableRange2
    bottom = max(ext1.Row + ext1.Rows.Count - 1, ext2.Row + ext2.Rows.Count - 1)
    p3 = _pivot_operators(cache, ws, ws.Cells(bottom + 5, col_b))

    _autofit_pivots(ws, start_row, [p1, p2, p3])


# ---------------------------------------------------------------------------
# Обработка одного файла
# ---------------------------------------------------------------------------

def transform_file(app, src_path, out_path):
    """Открыть исходник, выполнить шаги 1–7, сохранить как out_path."""
    wb = app.Workbooks.Open(os.path.abspath(str(src_path)), 0)   # 0 = не обновлять связи
    try:
        ws, header_row, by_norm, by_col = find_data_sheet(wb)
        require_headers(by_norm)

        id_to_col = by_norm[norm(H_ID_TO)]
        last_data_row = ws.Cells(ws.Rows.Count, id_to_col).End(XL_UP).Row
        if last_data_row <= header_row:
            raise ValueError("в таблице нет строк данных")

        comment_col = ensure_comment_column(ws, header_row, by_norm, by_col)      # шаг 1
        lo = create_table(ws, header_row, id_to_col, comment_col, last_data_row)  # шаг 2
        add_lpr_ul_column(app, ws, lo, by_norm, by_col)                           # шаг 3
        refresh_all(app, wb)                                                      # шаг 4
        build_pivots(app, wb, ws, lo)                                             # шаги 5–6

        wb.SaveAs(os.path.abspath(str(out_path)), XL_OPENXML_WORKBOOK)            # шаг 7
    finally:
        try:
            wb.Close(False)
        except Exception:
            pass


def process_one(app, src_path, folder):
    out = folder / (SVERKA_PREFIX + src_path.name)
    transform_file(app, src_path, out)
    dest = move_to_old(src_path, ensure_old_dir(folder))
    log.info("Готово: %s → %s ; оригинал → %s",
             src_path.name, out.name, os.path.join(OLD_DIRNAME, dest.name))


def main():
    folder = Path(__file__).resolve().parent
    setup_logging(folder)
    log.info("=== Запуск сверки. Папка: %s ===", folder)

    files = discover_files(folder)
    if not files:
        log.info("Нет файлов для обработки.")
        return 0
    log.info("К обработке: %d файл(ов): %s", len(files), ", ".join(f.name for f in files))

    try:
        app, pid = make_excel()
    except Exception:
        log.exception("Не удалось запустить Excel (нужен Windows + Excel + pywin32).")
        return 2

    ok = failed = 0
    try:
        for src in files:
            try:
                process_one(app, src, folder)
                ok += 1
            except Exception:
                failed += 1
                log.exception("ОШИБКА при обработке %s — оставлен на месте, пропущен", src.name)
    finally:
        quit_excel(app)
        kill_pid(pid)

    log.info("=== Итог: успешно %d, с ошибкой %d ===", ok, failed)
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
