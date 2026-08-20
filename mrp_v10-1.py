# -*- coding: utf-8 -*-
# ============================================================================
# MRP System — расчётный движок.  Build SVK002.10.
# (c) S.V. Kamalov.  Developed by S.V. Kamalov.
# Проприетарный код. Копирование, распространение и повторное использование
# без письменного разрешения автора запрещено.
# ============================================================================
"""
MRP v9 — устранение замечаний
=============================
[1] остатки бамперов = улица + линия + буфер (а не один столбец)
[2] цветная декомпозиция плана для бамперов и красок через
    Order_calculation_statistics_UPDATED.xlsx (батч → цвет → шт)
[3] дорестайл-бамперы XST33/AST33 (передние+задние) — только B02 2WD premium
[4] B16 elite (подголовники) — ТОЛЬКО B16_4WD_elite
[5] потребность в краске = норма × кол-во кузовов нужных цветов
[6] BOM_Детальный — корректный offset столбцов (A01_2WD_comfort не пуст)
[7] case B04_4WD_Premium → premium (нормализация)
[8] Сводка_3мес: порядок колонок данных = порядок шапки
[9] Риск_Дефицита: «Ср/день» = потребность месяца / число дней месяца
[10] авто-сканирование всех xlsx/xls в папке (все вкладки; типовые группы + новые файлы)
"""
import pandas as pd
import re
import math, os, json, datetime, warnings, glob, sys
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.utils.datetime import from_excel
warnings.filterwarnings('ignore')

# ── Консоль Windows (cp1251) падает при печати китайских имён файлов и эмодзи
#    (UnicodeEncodeError). Переключаем stdout/stderr на UTF-8 с заменой
#    непечатаемых символов, чтобы скрипт не прерывался на print(). ──
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass

# ══════════════════════════════════════════════════════════════
# АВТО-ОПРЕДЕЛЕНИЕ ФАЙЛОВ
# Все входящие файлы ищутся по паттерну — просто положи новый файл
# в папку, старый можно оставить или удалить. Скрипт возьмёт
# самый свежий файл по каждому паттерну (по дате изменения).
# ══════════════════════════════════════════════════════════════

# ROOT = рабочая папка MRP (данные). Задаётся через mrp_root.txt, MRP_ROOT или рядом со скриптом.
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__)) if '__file__' in globals() else os.getcwd()

def _resolve_root():
    """Папка данных: env MRP_ROOT / MRP_DATA_DIR → mrp_root.txt → папка скрипта."""
    candidates = []
    for key in ('MRP_ROOT', 'MRP_DATA_DIR'):
        val = os.environ.get(key, '').strip()
        if val:
            candidates.append(val)
    cfg = os.path.join(SCRIPT_DIR, 'mrp_root.txt')
    if os.path.isfile(cfg):
        try:
            with open(cfg, 'r', encoding='utf-8') as f:
                for line in f:
                    line = line.strip()
                    if line and not line.startswith('#'):
                        candidates.append(line)
                        break
        except OSError:
            pass
    candidates.append(SCRIPT_DIR)
    for path in candidates:
        if path and os.path.isdir(path):
            return os.path.abspath(path)
    return os.path.abspath(SCRIPT_DIR)

ROOT = _resolve_root()

# ── Сетевой замок: один расчёт MRP на папку ───────────────────
if SCRIPT_DIR not in sys.path:
    sys.path.insert(0, SCRIPT_DIR)
import atexit as _atexit
import mrp_locks as _mrp_locks

_MRP_RUN_LOCK = _mrp_locks.run_lock_path(ROOT)
_MRP_LOCK_HELD = False
if os.environ.get("MRP_SKIP_RUN_LOCK", "").strip().lower() not in ("1", "true", "yes"):
    _out_xlsx = os.path.join(ROOT, "MRP_System_v9.xlsx")
    _block = _mrp_locks.active_lock_message(ROOT, _out_xlsx)
    if _block:
        print(f"[LOCK] {_block}")
        sys.exit(2)
    _lock_ok, _lock_info = _mrp_locks.try_acquire(_MRP_RUN_LOCK, "mrp_run")
    if not _lock_ok:
        print(f"[LOCK] Расчёт уже выполняется:\n{_mrp_locks.format_lock_holder(_lock_info)}")
        sys.exit(2)
    _MRP_LOCK_HELD = True

    def _release_mrp_run_lock():
        if _MRP_LOCK_HELD:
            _mrp_locks.release(_MRP_RUN_LOCK, os.getpid())

    _atexit.register(_release_mrp_run_lock)

def _is_excel_lock_file(path):
    """Excel lock/temp files (~$...) are not real workbooks."""
    return os.path.basename(path).startswith('~$')

def _is_valid_xlsx(path):
    """True если файл — целый xlsx (не OneDrive-placeholder / обрезанный csv)."""
    if not path or not os.path.isfile(path) or _is_excel_lock_file(path):
        return False
    try:
        if os.path.getsize(path) < 512:
            return False
        with open(path, 'rb') as _f:
            if _f.read(2) != b'PK':
                return False
        import zipfile
        with zipfile.ZipFile(path, 'r') as _z:
            _z.namelist()
        _wb = load_workbook(path, read_only=True, data_only=True)
        _wb.close()
        return True
    except Exception:
        return False

_CKD_PLAN_NAME_RE = re.compile(
    r'(?:'
    r'rus\s*\([^)]*chn\s*\)|'
    r"sen['\u2019]|сен['\u2019]|фев['\u2019]|"
    r'sgk[_\s-]*delivery|stamping\s*plan|production\s*request|'
    r'^\d+\s*(?:ckd|jolion|dargo|f\s*series|h3|h7|f7)'
    r')',
    re.I,
)
_MONTH_RANGE_RE = re.compile(
    r'(?:jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec|'
    r'янв|фев|мар|апр|май|июн|июл|авг|сен|окт|ноя|дек)'
    r'\s*[-–—]\s*'
    r'(?:jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec|'
    r'янв|фев|мар|апр|май|июн|июл|авг|сен|окт|ноя|дек)',
    re.I,
)

def _is_non_stock_excel(path):
    """Планы CKD/производства и поставки — не файлы остатков."""
    parts_lower = [p.lower() for p in os.path.normpath(path).split(os.sep)]
    for seg in parts_lower:
        if 'ckd' in seg:
            return True
    if any(seg in ('замечания', 'графики_поставщиков') for seg in parts_lower):
        return True
    base = os.path.basename(path)
    bl = base.lower()
    if _CKD_PLAN_NAME_RE.search(base):
        return True
    if _MONTH_RANGE_RE.search(bl):
        return True
    if any(k in base for k in ('计划', '产能满足', '计划执行')):
        return True
    return False

def _find_gwm_batch_file():
    """GWM 各车型成套批次统计表 — источник paint statistics."""
    _exclude = ('order_calculation', 'plan-fact', 'plan_fact')
    _patterns = (
        '*批次*統計*.xlsx', '*各车型*.xlsx', '*批次*统计*.xlsx',
        '*GWM*批次*.xlsx', '*成套*批次*.xlsx', '*车型*批次*.xlsx',
        '*统计表*.xlsx',
    )
    _folders = [
        ROOT, os.path.join(ROOT, 'Входные данные'),
        os.path.join(ROOT, 'Замечания'),
        os.path.join(ROOT, 'Замечания', 'Paiint calcualtion'),
        os.path.dirname(ROOT),
    ]
    _cands = []
    _seen = set()
    for folder in _folders:
        if not folder or not os.path.isdir(folder):
            continue
        for pat in _patterns:
            for p in glob.glob(os.path.join(folder, pat)):
                p = os.path.abspath(p)
                if p in _seen or _is_excel_lock_file(p):
                    continue
                _seen.add(p)
                bl = os.path.basename(p).lower()
                if any(x in bl for x in _exclude):
                    continue
                bn = os.path.basename(p)
                if not any(k in bn for k in ('批次', '各车型', '涂装')) and 'gwm' not in bl:
                    continue
                if _is_valid_xlsx(p):
                    _cands.append(p)
            for p in glob.glob(os.path.join(folder, '**', pat), recursive=True):
                p = os.path.abspath(p)
                if p in _seen or _is_excel_lock_file(p):
                    continue
                _seen.add(p)
                bl = os.path.basename(p).lower()
                if any(x in bl for x in _exclude):
                    continue
                bn = os.path.basename(p)
                if not any(k in bn for k in ('批次', '各车型', '涂装')) and 'gwm' not in bl:
                    continue
                if _is_valid_xlsx(p):
                    _cands.append(p)
    return max(_cands, key=os.path.getmtime) if _cands else None

def _find_build_order_calc_module():
    import sys as _sys
    _boc_candidates = [
        os.path.join(SCRIPT_DIR, 'build_order_calc_v2.py'),
        os.path.join(ROOT, 'build_order_calc_v2.py'),
        os.path.join(os.path.dirname(ROOT), 'build_order_calc_v2.py'),
        os.path.join(os.path.dirname(ROOT), 'files FINAL', 'build_order_calc_v2.py'),
        os.path.join(os.path.dirname(ROOT), 'files', 'build_order_calc_v2.py'),
        os.path.join(ROOT, 'Замечания', 'Paiint calcualtion', 'build_order_calc_v2.py'),
        os.path.join(os.path.expanduser('~'), 'Downloads', 'files FINAL', 'build_order_calc_v2.py'),
        os.path.join(os.path.expanduser('~'), 'Downloads', 'files', 'build_order_calc_v2.py'),
    ]
    return next((p for p in _boc_candidates if os.path.isfile(p)), None)

def _try_regenerate_paint_stats(gwm_path, out_path):
    """Сгенерировать Order_calculation_statistics из GWM."""
    _boc_path = _find_build_order_calc_module()
    if not _boc_path or not gwm_path or not _is_valid_xlsx(gwm_path):
        return False
    try:
        import importlib.util as _ilu
        import sys as _sys
        _boc_dir = os.path.dirname(_boc_path)
        if _boc_dir not in _sys.path:
            _sys.path.insert(0, _boc_dir)
        _spec = _ilu.spec_from_file_location('build_order_calc_v2', _boc_path)
        _boc_mod = _ilu.module_from_spec(_spec)
        _spec.loader.exec_module(_boc_mod)
        _boc_mod.build_output(gwm_path, out_path)
        return _is_valid_xlsx(out_path)
    except Exception as _e:
        print(f'[INIT] ⚠️  Ошибка генерации paint statistics: {_e}')
        return False

def _resolve_paint_stats_file():
    """Самый свежий валидный файл paint statistics."""
    _raw = [
        os.path.join(ROOT, 'Order_calculation_statistics_UPDATED.xlsx'),
        _latest(ROOT, '*Order*calculation*.xlsx', '*涂装*.xlsx'),
        os.path.join(ROOT, 'Замечания', 'Paiint calcualtion',
                     'Order_calculation_statistics_UPDATED.xlsx'),
        _latest(os.path.dirname(ROOT), '*Order*calculation*.xlsx', '*涂装*.xlsx'),
    ]
    _seen, _cands = set(), []
    for p in _raw:
        if not p or not os.path.isfile(p) or _is_excel_lock_file(p):
            continue
        p = os.path.abspath(p)
        if p in _seen:
            continue
        _seen.add(p)
        if _is_valid_xlsx(p):
            _cands.append(p)
    return max(_cands, key=os.path.getmtime) if _cands else None

def _resolve_ecoalliance_use_file():
    candidates = [
        os.path.join(os.path.expanduser('~'), 'Documents', 'EcoAlliance use.xlsx'),
        os.path.join(ROOT, 'EcoAlliance use.xlsx'),
        os.path.join(ROOT, 'Входные данные', 'EcoAlliance use.xlsx'),
    ]
    candidates.extend(glob.glob(os.path.join(ROOT, '*EcoAlliance*use*.xlsx')))
    candidates.extend(glob.glob(os.path.join(ROOT, '*EcoAlliance*use*.xls')))
    candidates.extend(glob.glob(os.path.join(ROOT, '**', '*EcoAlliance*use*.xlsx'), recursive=True))
    candidates.extend(glob.glob(os.path.join(ROOT, '**', '*EcoAlliance*use*.xls'), recursive=True))
    seen = set()
    for p in candidates:
        if not p or p in seen:
            continue
        seen.add(p)
        if os.path.isfile(p) and not _is_excel_lock_file(p):
            return p
    return ''

ECOALLIANCE_USE_FILE = _resolve_ecoalliance_use_file()

def _resolve_excel_path(path):
    """If a lock file was picked, resolve to the real workbook next to it."""
    if not path:
        return path
    if not _is_excel_lock_file(path):
        return path
    real_name = os.path.basename(path)[2:]
    real_path = os.path.join(os.path.dirname(path), real_name)
    return real_path if os.path.isfile(real_path) else path

def _save_workbook_safe(wb, target_path, label='результат MRP'):
    """Сохраняет workbook; если файл открыт в Excel — пробует альтернативное имя."""
    base, ext = os.path.splitext(target_path)
    folder = os.path.dirname(target_path) or ROOT
    ts = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
    candidates = [
        target_path,
        os.path.join(folder, f"MRP_System_v9_new{ext}"),
        os.path.join(folder, f"MRP_System_v9_{ts}{ext}"),
    ]
    last_err = None
    for path in candidates:
        try:
            wb.save(path)
            if os.path.normcase(path) != os.path.normcase(target_path):
                print(f"\n  ⚠️  {os.path.basename(target_path)} открыт в Excel — сохранено как:")
                print(f"       {path}")
                print(f"       Закройте Excel и откройте этот файл, либо скопируйте лист BOM_Детальный в {os.path.basename(target_path)}")
            return path
        except PermissionError as e:
            last_err = e
            print(f"  ⚠️  Нет доступа к «{os.path.basename(path)}» — закройте файл в Excel")
        except OSError as e:
            if getattr(e, 'errno', None) == 13:
                last_err = e
                print(f"  ⚠️  Нет доступа к «{os.path.basename(path)}» — закройте файл в Excel")
            else:
                raise
    print(f"\n  ❌ Не удалось сохранить {label}.")
    print("     Закройте MRP_System_v9.xlsx (и Excel полностью, если нужно) и запустите скрипт снова.")
    raise last_err

def _latest(folder, *patterns):
    """Возвращает самый свежий файл среди всех паттернов, или None."""
    candidates = []
    for pat in patterns:
        candidates.extend(glob.glob(os.path.join(folder, pat)))
        candidates.extend(glob.glob(os.path.join(folder, '**', pat), recursive=True))
    candidates = [
        p for p in candidates
        if os.path.isfile(p) and not _is_excel_lock_file(p)
    ]
    return max(candidates, key=os.path.getmtime) if candidates else None

def _require(path, label):
    if path and os.path.exists(path): return path
    print(f"  ⚠️  Файл не найден: {label}")
    return path or ''

# ── Основные файлы ────────────────────────────────────────────
PF_FILE  = _resolve_excel_path(_require(_latest(ROOT, "*Plan-Fact*.xlsx", "*Plan_Fact*.xlsx",
                                                  "*план-факт*.xlsx", "*планфакт*.xlsx"),
                                       "Plan-Fact (план производства)"))

BOM_FILE = _require(_latest(ROOT, "Master_BOM_актуальный.xlsx",
                             "Master_BOM*.xlsx"),
                    "Master_BOM")
BOM_NEW  = _latest(ROOT, "Master_BOM_актуальный.xlsx") or BOM_FILE

PKG_FILE = _require(_latest(ROOT, "Упаковка локала*.xlsx", "Упаковка_локала*.xlsx",
                             "*Упаковка*локала*.xlsx"),
                    "Упаковка локала")

ADD_FILE = _require(_latest(ROOT, "Additional.xlsx", "Additional*.xlsx"),
                    "Additional")

# ── Теоретические остатки (высший приоритет, перекрывает все другие) ──
THEOR_STOCK_FILE = _latest(ROOT,
                            "Теоретические остатки*.xlsx",
                            "теоретические*остатки*.xlsx",
                            "Теор*остатки*.xlsx")

OUT      = os.path.join(ROOT, "MRP_System_v9.xlsx")

def _resolve_mrp_live_input_path():
    """Файл с LIVE BOM_Детальный: сначала основной MRP_System_v9.xlsx, иначе запасные копии."""
    primary = _resolve_excel_path(OUT)
    if primary and os.path.isfile(primary) and not _is_excel_lock_file(primary):
        try:
            wb = load_workbook(primary, read_only=True, data_only=True)
            ok = 'BOM_Детальный' in wb.sheetnames
            wb.close()
            if ok:
                return primary
        except Exception:
            pass
    cands = []
    for p in glob.glob(os.path.join(ROOT, 'MRP_System_v9*.xlsx')):
        if _is_excel_lock_file(p):
            continue
        try:
            wb = load_workbook(p, read_only=True, data_only=True)
            if 'BOM_Детальный' not in wb.sheetnames:
                wb.close()
                continue
            wb.close()
            cands.append(p)
        except Exception:
            continue
    if not cands:
        return OUT
    best = max(cands, key=os.path.getmtime)
    if os.path.normcase(os.path.abspath(best)) != os.path.normcase(os.path.abspath(OUT)):
        print(f"[INIT] LIVE BOM: {os.path.basename(best)} (основной файл недоступен)")
    return best

OUT_LIVE_INPUT = _resolve_mrp_live_input_path()

# Снимок ✓ из LIVE BOM_Детальный — восстанавливается после авто-правок скрипта
live_bom_configs_snapshot = {}

# ── Paint statistics: авто-генерация из GWM-файла ─────────────
# 1. Ищем входной GWM-файл (各车型成套批次统计表) в папке MRP
# 2. Если найден — запускаем build_order_calc_v2 для генерации статистики
# 3. Используем свежий Order_calculation_statistics_UPDATED.xlsx
_PAINT_OUT = os.path.join(ROOT, "Order_calculation_statistics_UPDATED.xlsx")
_GWM_INPUT = _find_gwm_batch_file()

if _GWM_INPUT:
    print(f"[INIT] GWM-файл найден: {os.path.basename(_GWM_INPUT)}")
    _need_paint_regen = (
        not _is_valid_xlsx(_PAINT_OUT)
        or os.path.getmtime(_GWM_INPUT) > os.path.getmtime(_PAINT_OUT)
    )
    if _need_paint_regen:
        print(f"[INIT] Генерация paint statistics...")
        if _try_regenerate_paint_stats(_GWM_INPUT, _PAINT_OUT):
            print(f"[INIT] ✅ Paint statistics обновлена → {os.path.basename(_PAINT_OUT)}")
        elif not _find_build_order_calc_module():
            print(f"[INIT] ⚠️  build_order_calc_v2.py не найден — используем готовый файл статистики")
        else:
            print(f"[INIT] ⚠️  Не удалось сгенерировать paint statistics — ищем готовый файл")
else:
    print(f"[INIT] ℹ️  GWM-файл не найден — используем существующий файл статистики")

PAINT_STATS = _resolve_paint_stats_file()
if not PAINT_STATS and _GWM_INPUT:
    print(f"[INIT] Paint stats повреждён/отсутствует — повторная генерация из GWM...")
    if _try_regenerate_paint_stats(_GWM_INPUT, _PAINT_OUT):
        PAINT_STATS = _PAINT_OUT
if PAINT_STATS and not _is_valid_xlsx(_PAINT_OUT):
    try:
        import shutil as _shutil
        _shutil.copy2(PAINT_STATS, _PAINT_OUT)
        print(f"[INIT] ♻️  Восстановлен {os.path.basename(_PAINT_OUT)} "
              f"← {os.path.basename(PAINT_STATS)}")
    except OSError as _cp_e:
        print(f"[INIT] ⚠️  Не удалось скопировать paint stats локально: {_cp_e}")

# ── Бамперы ───────────────────────────────────────────────────
BUMPER_FILE = _latest(ROOT, "*бампер*.xlsx", "*Bumper*.xlsx",
                      "*Пересчет бамперов*.xlsx") or ''

# ── Файлы остатков — авто-сканирование ───────────────────────
# Сканируем все xlsx в папке: типовые группы (AAT, GSK…) + любые новые файлы.
# Из каждой группы / серии имён берётся только самый свежий файл.
# После загрузки BOM остаются только остатки деталей, входящих в расчёт MRP.
_SYSTEM_FILES = {
    os.path.normcase(p) for p in [
        OUT, PF_FILE, BOM_FILE, BOM_NEW, PKG_FILE, ADD_FILE,
        THEOR_STOCK_FILE,                          # теоретические остатки — отдельная загрузка
        _PAINT_OUT, PAINT_STATS, _GWM_INPUT,
        os.path.join(ROOT, "MRP_System_v8.xlsx"),
        os.path.join(ROOT, "MRP_System_v9_new.xlsx"),
        os.path.join(ROOT, "Stock in days.xlsx"),  # справочник норм — не файл остатков
    ] if p
}
# Группы типов файлов остатков: (название_группы, [ключевые_слова])
_STOCK_GROUPS = [
    ('AAT',         ['aat', 'аат']),
    ('Warehouse',   ['warehouse']),
    ('GSK',         ['gsk', 'гск']),
    ('Purem',       ['пюрэм', 'пюрем', 'purem']),
    ('Bumpers',     ['бампер', 'bumper', 'пересчет бампер', 'пересчёт бампер']),
    ('ComboRecalc', ['жуйчуан', 'zhuichuan', 'jui chuan', 'msa перес', 'мса перес']),
    ('MSA',         ['msa', 'мса']),
    ('Accounting',  ['accounting']),
    ('Остатки',     ['остатк', 'остаток']),
    ('Lear',        ['lear']),
    ('Fuyao',       ['fuyao']),
    ('Yapp',        ['yapp']),
    ('STP',         ['stp']),
    ('GSK2',        ['gsk']),
]
_SKIP_IN_STOCK = ['mrp_system','plan-fact','plan_fact','master_bom',
                   'упаковка','additional','order_calc','order calculation','涂装',
                   'statistics','stock in days','теорет','批次','gwm','замечан',
                   'страх', 'ecoalliance', 'eco alliance',
                   'stamping',   # Stamping plan — план SGK, не остатки
                   'сводная',    # сводные/аналитические файлы
                   'fuel',       # «Fuel warehouse checklist» — чек-лист топлива, не остатки
                   'чек-лист', 'checklist',
                   'клея', 'клей',  # «Данные по остаткам клея» — журнал норм расхода (клей берётся из warehouse 16)
                   'sgk_delivery', 'sgk delivery', 'delivery 11', 'jolion', 'dargo']

_STOCK_SCAN_SKIP_DIRS = frozenset({
    'графики_поставщиков', 'логи', 'logs', 'assets', '__pycache__', 'замечания',
})

def _under_skipped_stock_dir(path):
    parts = {p.lower() for p in os.path.normpath(path).split(os.sep)}
    return bool(parts & _STOCK_SCAN_SKIP_DIRS)

def _iter_excel_under(root):
    """Все xlsx/xls в root и подпапках (кроме служебных каталогов)."""
    seen = set()
    for pat in ('*.xlsx', '*.xls'):
        for p in glob.glob(os.path.join(root, '**', pat), recursive=True):
            if _is_excel_lock_file(p) or _under_skipped_stock_dir(p):
                continue
            key = os.path.normcase(os.path.abspath(p))
            if key not in seen:
                seen.add(key)
                yield p

# Сортируем xlsx/xls по дате изменения (сначала новые) — все вкладки каждого файла
_all_xlsx = sorted(_iter_excel_under(ROOT), key=os.path.getmtime, reverse=True)

def _stock_file_stem(filename):
    """Имя файла без даты (.ДД.ММ.ГГ) — для группировки версий одного источника."""
    stem = os.path.splitext(filename)[0].lower()
    stem = re.sub(r'[._]\d{2}[._]\d{2}[._]\d{2}(?:[._]\d+)?', '', stem, flags=re.I)
    return stem.strip('._- ')

def _match_stock_group(base_lower):
    for grp_name, kws in _STOCK_GROUPS:
        if any(kw in base_lower for kw in kws):
            return grp_name
    return None

STOCK_FILES = []
_seen_groups = set()   # типовые группы, для которых файл уже найден
_seen_auto_stems = set()  # новые файлы без ключевых слов — по stem имени
_skipped_old = []      # старые файлы — пропускаем
_skipped_non_stock = []  # планы CKD / производства — не остатки
_auto_stock_files = []  # новые файлы вне типовых групп

for p in _all_xlsx:
    if os.path.normcase(p) in _SYSTEM_FILES:
        continue
    if ECOALLIANCE_USE_FILE and os.path.normcase(p) == os.path.normcase(ECOALLIANCE_USE_FILE):
        continue
    if _is_non_stock_excel(p):
        _skipped_non_stock.append(os.path.basename(p))
        continue
    base = os.path.basename(p).lower()
    if any(x in base for x in _SKIP_IN_STOCK): continue
    matched_group = _match_stock_group(base)
    if matched_group:
        if matched_group in _seen_groups:
            _skipped_old.append(os.path.basename(p))  # уже есть более свежий
        else:
            _seen_groups.add(matched_group)
            STOCK_FILES.append(p)
        continue
    # Любой новый xlsx/xls без ключевых слов — отдельная группа по stem (без даты в имени)
    stem = _stock_file_stem(os.path.basename(p))
    if not stem:
        continue
    if stem in _seen_auto_stems:
        _skipped_old.append(os.path.basename(p))
    else:
        _seen_auto_stems.add(stem)
        _auto_stock_files.append(p)
        STOCK_FILES.append(p)

print(f"[INIT] ROOT = {ROOT}")
print(f"[INIT] Plan-Fact: {os.path.basename(PF_FILE) if PF_FILE else '❌ НЕ НАЙДЕН'}")
print(f"[INIT] BOM:       {os.path.basename(BOM_FILE) if BOM_FILE else '❌ НЕ НАЙДЕН'}")
print(f"[INIT] Теор.ост.: {os.path.basename(THEOR_STOCK_FILE) if THEOR_STOCK_FILE else 'не найден (необязательно)'}")
print(f"[INIT] Бамперы:   {os.path.basename(BUMPER_FILE) if BUMPER_FILE else '❌ НЕ НАЙДЕН'}")
print(f"[INIT] PaintStats:{os.path.basename(PAINT_STATS) if PAINT_STATS and os.path.exists(PAINT_STATS) else '❌ НЕ НАЙДЕН'}")
if ECOALLIANCE_USE_FILE:
    print(f"[INIT] EcoAlliance use: {os.path.basename(ECOALLIANCE_USE_FILE)}")
print(f"[INIT] Найдено {len(STOCK_FILES)} файлов остатков "
      f"(типовые: {len(STOCK_FILES) - len(_auto_stock_files)}, новые: {len(_auto_stock_files)}):")
for p in STOCK_FILES:
    _tag = ' 🆕' if p in _auto_stock_files else ''
    print(f"       ✅ {os.path.basename(p)}{_tag}")
if _skipped_non_stock:
    print(f"[INIT] Исключено планов/CKD (не остатки): {len(_skipped_non_stock)}")
    for nm in sorted(set(_skipped_non_stock))[:12]:
        print(f"       ⏭ {nm}")
if _skipped_old:
    print(f"[INIT] Пропущено устаревших:")
    for nm in _skipped_old:
        print(f"       ⏩ {nm}")

# ── Устаревшие коды (не поставляются, исключаем из расчёта) ──
OBSOLETE_CODES = {
    '7005110XKJ22A',     # больше не поставляется (заменён на другой)
    '3101100xst33A',     # дубль 3101100XST33A (lowercase) — удалить
    # ── 2804111/2804112 XKN61 → перенесены в REFERENCE_ONLY_CODES (справочные,
    #    видны в BOM, но потребность/заказы/график для них не считаются) ──
    # ── Больше не заказываем (замечания Корнеевой, Замечания-3) ──
    '5301120XST11A',     # A01 comfort — больше не применяем
    '5301501XST11A',     # A01 — больше не применяем
    '5401571XST33A',     # A01 — больше не применяем
    '5401572XST33A',     # A01 — больше не применяем
    '5401871XST33A',     # A01 — больше не применяем
    '5401872XST33A',     # A01 — больше не применяем
    '8400451XST11A',     # A01 — больше не применяем
    '8400452XST11A',     # A01 — больше не применяем
    '8400751XST11A',     # A01 — больше не применяем
    '8400752XST11A',     # A01 — больше не применяем
    # Purem — больше не используем
    '1201UQ8XGW01A', '1201UQ9XGW01A', '1201UR1XGW01A', '1201UR2XGW01A',
    '1205656XKJ22A', '1205657XKJ22A', '1205658XKJ22A', '1205659XKJ22A', '1205660XKJ22A',
}

# ── Справочные коды (REFERENCE-ONLY): присутствуют в BOM для справки, но
#    потребность, заказы и график поставок для них НЕ рассчитываются.
#    2804111/2804112 XKN61 — дубли задних бамперов 2804KN260004/2804KN260005
#    (та же применяемость), заказываются по KN-кодам, эти — только справочно.
REFERENCE_ONLY_CODES = {
    '2804111XKN61A8T', '2804111XKN61A9C', '2804111XKN61AC3', '2804111XKN61AH4',
    '2804112XKN61A8T', '2804112XKN61A9C', '2804112XKN61AC3', '2804112XKN61AH4',
}

# ── Переопределение размеров упаковки (замечания Корнеевой, Замечания-3) ──
# Значения взяты из листа «Упаковка» файла «Замечания Корнеевой.xlsx» (колонка Примечание)
PACKAGE_OVERRIDES = {
    # Поставщик Пюрэм — упаковки не были проставлены
    '1201UR5XGW01A': 600,   # AIR INLET PIPE, 4B15 2WD
    '1205100XKN63A': 16,    # B02/B04/B06/B16 4*4
    '1205102XST1AA': 384,   # A01/A08 (Ecoal'yance)
    '1205112XST11A': 432,   # Ecoal'yance, A01 4*2 / A08 4*4 old / B02 4*2 old
    '1205113XST11A': 432,   # Экоальянс, A01 4*2 / A08 4*4 old / B02 4*2 old
    '1205526XGW01B': 432,   # Экоальянс, B02/B04 4*4 old
    '1205527XGW01A': 432,   # Экоальянс, B02/B04 4*4 old
    '5120101XGW01A': 200,   # Пюрэм
    # GSK / SGK — упаковки не были проставлены
    '5401188XKN01B': 120,   # B02, B04
    'LK011475':      240,   # A01/A08, SGK
    'LK012148':      360,   # A01/A08, SGK
    'LK012154':      360,   # A01, SGK
    'LK014401':      450,   # A01, SGK
    'LK014403':      240,   # A01, SGK
}

DEFAULT_BUMPER_PKG = 8   # если в «Упаковка локала» для бампера не задано иначе

# ── Нормализация имён поставщиков (Ecokhim = Ecohim и т.п.) ──
SUPPLIER_ALIASES = {
    'ECOHIM':    'Ecokhim',
    'ECOKHIM':   'Ecokhim',
    'ЭКОХИМ':    'Ecokhim',
    'ЭКОКХИМ':   'Ecokhim',
    'ECHOKHIM':  'Ecokhim',
    'ECHOHIM':   'Ecokhim',
    'ECOALLIANCE': "Ecoal'yance",
    'ECO ALLIANCE': "Ecoal'yance",
    'ECOALYANCE': "Ecoal'yance",
    "ECOAL'YANCE": "Ecoal'yance",
    'ЭКОАЛЬЯНС': "Ecoal'yance",
    'ECOTECHSYS': 'Ecotexis',
    'ECOTEXIS':   'Ecotexis',
    'LEAR':       'Lear',
}

# ── Правила графиков поставок по поставщикам (канонические имена из BOM) ──
ECOALYANCE_SUPPLIER = "Ecoal'yance"
ECOTEXIS_SUPPLIER   = 'Ecotexis'
SMC_SUPPLIER        = 'SMC'
PUREM_SUPPLIER      = 'Purem'
# ── Эфтек: загрузка ТС не более 18 т (согласовано). Коды Эфтек ведутся
#    в граммах (вес единицы = 1 г; подтверждено Таблицей warehouse 16) ──
EFTEC_SUPPLIER = 'Eftec'
EFTEC_MAX_TRUCK_WEIGHT_G = 18_000_000   # 18 тонн
# ── Эфтек, замечания планировщика (согласовано) ──
#    ALAC011940 / ALAC011941 отгружаются кратно 4 бочкам; размер бочки —
#    это упаковка кода из листа «Упаковка» (218 и 246 кг соответственно).
EFTEC_DRUM_MULTIPLE_CODES = {'ALAC011940', 'ALAC011941'}
EFTEC_DRUMS_PER_SHIPMENT = 4
#    Поставки Эфтек не ставятся на понедельник (и на воскресенье — оно и так
#    не рейсовый день, но до правки прорывалось после консолидации).
EFTEC_FORBIDDEN_WEEKDAYS = {0, 6}   # 0 = понедельник, 6 = воскресенье

def _eftec_unit_weight_g(code):
    """Вес единицы кода Эфтек в граммах (ед.изм. из BOM: g/кг/мл)."""
    u = str(bom.get(code, {}).get('unit', '') or '').strip().lower()
    if u in ('kg', 'кг', 'kg.'):
        return 1000.0
    return 1.0   # g / мл / прочее — 1 г за единицу
SMC_MAX_PALLETS_PER_TRUCK = 7
# ── MSA: фура = 21 контейнер (3 яруса × 7 рядов). 1 контейнер бампера = 1 упаковка.
#    Накладки: мин. отгрузка передних 480 шт; задние кратно 180 шт.
#    Бамперы XST33/AST33: суммарно ≤ 480 шт/день в группе (передние / задние). ──
MSA_SUPPLIER = 'MSA'
MSA_MAX_CONTAINERS_PER_TRUCK = 21
# ── NDKT (B06 колёса): на начало дня — 100% потребности дня ПЕРВЫХ
#    производственных партий плана (по одной на вкладку AS_in_F_A / AS_in_H_B),
#    остальные колёса — 80 шт (+20% допуск). Поставка ставится в день
#    потребности: Del[d] = potr[d] + цель утра d+1 − Ss[d] (≈ потребность
#    ТЕКУЩЕГО дня, не следующего); стандартный приход Del[d] → Ss[d+1] ──
NDKT_SUPPLIER = 'NDKT'
NDKT_AS_TABS = ('AS_in_F_A', 'AS_in_H_B')
NDKT_MIN_TRANSITION_QTY = 80       # переходящий запас каждого активного колеса, допуск +20%
NDKT_SS_TOLERANCE_UP = 0.20        # допустимое отклонение вверх (+20%)
# ── NDKT: ежедневный страховой запас по кодам. Перепланирован РАЗОВО по
#    скриншоту «Safety stock required» (09.06.2026); значения зафиксированы
#    здесь, скрипт больше не ссылается на файл-источник. ──
NDKT_SAFETY_BY_CODE = {
    '3101100XKN4AA': 240,   # Колесо в сборе B06 NEW, elite/premium
    '3101101XKN4AA': 240,   # Колесо в сборе B06 NEW, Tech Plus
    '3101101XST33A': 480,   # R-18, A01 techplus/premium
    '3101100XST33A': 480,   # R-17, A01 elite/comfort
    '3101101XKN61A': 480,   # R-19, B02+B04(old), Techplus/Premium
    '3101100XKN63A': 480,   # B04 NEW, premium/Tech Plus
    '3101100XKN61A': 480,   # R-18, B02, Elite
    '3101100XKJ23A': 240,   # 18", A08
}

def _ndkt_safety_base(code):
    """Ежедневный страховой запас колеса NDKT (по коду; default 80 шт)."""
    return float(NDKT_SAFETY_BY_CODE.get(code, NDKT_MIN_TRANSITION_QTY))

# ── NDKT: упаковка и фуры (упрощённая логика поставок).
#    Отдельный вид упаковки только у 3101100XST33A — бокс 30 шт.
#    Вместимость фуры: 3101100XST33A — 540 шт (18 боксов), остальные колёса — 384 шт.
#    Колёса МИКСУЮТСЯ в одной машине: загрузка = x/540 + y/384 ≤ 1,
#    где x — шт 3101100XST33A, y — шт прочих колёс (в т.ч. 3101101XST33A).
#    Микс 3101100XST33A ↔ 3101101XST33A: 1 бокс (30 шт) занимает 1/18 фуры
#    = место 21,3 шт прочих колёс; целые опорные комбинации — каждые 3 бокса:
#      боксы(шт) 3101100XST33A → шт 3101101XST33A (или других колёс):
#      0(0)→384 | 3(90)→320 | 6(180)→256 | 9(270)→192 | 12(360)→128 |
#      15(450)→64 | 18(540)→0   (т.е. 3 бокса = 90 шт ↔ 64 шт прочих).
#    Переходящий остаток приводится к цели = страх.запас по коду (240/480). ──
NDKT_PACKAGE_FIX = {
    '3101100XST33A': 30,   # отдельный вид упаковки — бокс 30 шт (всегда 30)
}
NDKT_WHEEL_PKG_DEFAULT = 24   # упаковка (стопка) всех остальных колёс — 24 шт
NDKT_TRUCK_CAP_DEFAULT = 384  # = 16 стопок по 24
NDKT_TRUCK_CAP = {'3101100XST33A': 540}   # = 18 боксов по 30
# Согласовано по эталону Delivery NDKT (23CW): максимум машин в день
# («вчера было 7, план растёт — даже 8»); пики предвозятся на ранние дни.
NDKT_MAX_TRUCKS_PER_DAY = 8

# Производственный календарь рейсов: воскресенья и праздники РФ — рейсов нет
# (по эталону НДКТ: нет 07.06, 12–14.06; суббота — рабочая).
_RU_HOLIDAYS_MD = {
    (1, 1), (1, 2), (1, 3), (1, 4), (1, 5), (1, 6), (1, 7), (1, 8),
    (2, 23), (3, 8), (5, 1), (5, 9), (6, 12), (11, 4),
}

def _is_ru_ship_day(dt):
    """Рейсовый день: не воскресенье и не гос. праздник РФ."""
    return dt.weekday() != 6 and (dt.month, dt.day) not in _RU_HOLIDAYS_MD

def _ndkt_truck_cap(code):
    return float(NDKT_TRUCK_CAP.get(code, NDKT_TRUCK_CAP_DEFAULT))
# ── Накладки MSA — пластиковые боксы, поставляются ПАРНО (комплектами):
#    5402411XKN02B + 5402422XKN02B — по 4 шт в боксе, партия 480+480;
#    5402433XKN02A + 5402444XKN02A — по 3 шт в боксе, партия 180+180.
#    Комплект пары (480+480 или 180+180) занимает в фуре место 3 СТАЛЬНЫХ тар.
MSA_FRONT_TRIM_CODES = {'5402411XKN02B', '5402422XKN02B'}
MSA_REAR_TRIM_CODES  = {'5402433XKN02A', '5402444XKN02A'}
MSA_FRONT_TRIM_MIN_QTY = 480
MSA_FRONT_TRIM_MULTIPLE = 480       # партия передних накладок (тара 120, возят по 480)
# Согласовано по эталону MSA: задние накладки кратны ТАРЕ 60 (в графиках 180/240/360),
# минимальный комплект 180; парность LH=RH сохраняется.
MSA_REAR_TRIM_MULTIPLE = 60
MSA_REAR_TRIM_MIN_QTY = 180
MSA_TRIM_SET_STEEL_SLOTS = 3        # комплект пары (480+480 или 180+180) = 3 стальных места
MSA_TRIM_SAFETY_QTY = 0             # запас накладок не раздуваем: партия — сама буфер
# ── Страховые (переходящие) запасы бамперов MSA — согласовано:
#    база «полпартии 100%» (партия 120 → 60 шт);
#    A01 (XST33/AST33): ходовые цвета чёрный/белый/серый (8T/9C/C3) — 90 шт,
#    неходовые голубой/красный (5B/GN) — в 2 раза меньше (45 шт);
#    B02/B04 (XKN61/KN26000x): все цвета равномерно,
#    Premium+TechPlus (…105…/…0005…) — 90 шт, Elite (…104…/…0004…) — 60 шт. ──
def _msa_bumper_safety_qty(code):
    if 'XST33' in code or 'AST33' in code:      # A01
        if code.endswith(('A5B', 'AGN')):
            return 45.0
        return 90.0
    if '2803105' in code or 'KN260005' in code:  # B02/B04 Premium+TechPlus
        return 90.0
    if '2803104' in code or 'KN260004' in code:  # B02 Elite
        return 60.0
    return 60.0
_MSA_BUMPER_COLOR_SUFFIXES = ('A5B', 'A8T', 'A9C', 'AC3', 'AGN')
MSA_XST33_FRONT_BUMPER_CODES = frozenset(
    f'{p}{s}' for p in ('2803120XST33', '2803130XST33') for s in _MSA_BUMPER_COLOR_SUFFIXES)
MSA_AST33_REAR_BUMPER_CODES = frozenset(
    f'{p}{s}' for p in ('2804104AST33', '2804105AST33') for s in _MSA_BUMPER_COLOR_SUFFIXES)
MSA_BUMPER_GROUP_MAX_PCS_PER_DAY = 480
# ── Цикл поставки MSA = 2 дня: страховой запас стальной тары (бамперы)
#    держим на уровне 2 средних дневных потребностей (по дням производства). ──
MSA_DELIVERY_CYCLE_DAYS = 2
# ── 5304100XKN02A: отгрузка кратно ФУРАМ (не полибоксам) — 528 или 576 шт,
#    1–2 раза в неделю, отдельной машиной (НЕ входит в фуру 21 стальной тары). ──
MSA_FULL_TRUCK_CODE = '5304100XKN02A'
MSA_FULL_TRUCK_SIZES = (528, 576)
# Калуга: одна поставка фиксированного объёма на период.
# 5304100XKN02A убран — теперь возится кратно фурам 528/576 (см. MSA_FULL_TRUCK_CODE).
KALUGA_SINGLE_DELIVERY_QTY = {}
# ── Lear: фура 18 паллетомест; на место — 2 уп. пены ИЛИ 3 уп. подголовника (36 пен / машину) ──
LEAR_SUPPLIER = 'Lear'
LEAR_MAX_PALLET_SLOTS = 18
LEAR_FOAM_PKGS_PER_SLOT = 2
LEAR_HEADREST_PKGS_PER_SLOT = 3
LEAR_FOAM_PKGS_PER_TRUCK = LEAR_MAX_PALLET_SLOTS * LEAR_FOAM_PKGS_PER_SLOT  # 36
LEAR_HEADREST_PKG_MULTIPLE = 3      # упаковки подголовников отгружаются кратно 3
# ── Lear: страховые запасы и вместимость тары — из файла «страх.запасы ЛИР.xlsx»
#    (детали в этой таре не миксуются: 1 упаковка = 1 код — уже общее правило). ──
LEAR_SAFETY_FILE_PATTERNS = ("*страх*запасы*ЛИР*.xlsx", "*страх*ЛИР*.xlsx", "*ЛИР*страх*.xlsx")
LEAR_SAFETY_BY_CODE = {}   # code -> страх.запас (шт) из файла
LEAR_PKG_BY_CODE = {}      # code -> вместимость тары (шт) из файла

def _load_lear_safety_file():
    """Читает «страх.запасы ЛИР.xlsx»: колонки партномер / УПАКОВКА / страх.запас."""
    path = _latest(ROOT, *LEAR_SAFETY_FILE_PATTERNS)
    if not path or not os.path.isfile(path):
        print("  ⚠️  Файл «страх.запасы ЛИР» не найден — страх.запасы Lear по умолчанию")
        return 0
    try:
        wb_l = load_workbook(path, read_only=True, data_only=True)
    except Exception as _e:
        print(f"  ⚠️  Не удалось прочитать «{os.path.basename(path)}»: {_e}")
        return 0
    ws_l = wb_l.worksheets[0]
    col_code = col_pkg = col_saf = None
    n_rows = 0
    for row in ws_l.iter_rows(values_only=True):
        if col_code is None:
            joined = [str(v or '').strip().lower() for v in row]
            for ci, v in enumerate(joined):
                if 'партномер' in v or 'код' == v:
                    col_code = ci
                elif 'упаковка' in v:
                    col_pkg = ci
                elif 'страх' in v:
                    col_saf = ci
            continue
        code = str(row[col_code]).strip() if col_code < len(row) and row[col_code] else ''
        if not code or not re.match(r'^[0-9A-ZА-Я]', code, re.I):
            continue
        try:
            pkg_v = float(row[col_pkg]) if col_pkg is not None and col_pkg < len(row) and row[col_pkg] not in (None, '') else None
        except (TypeError, ValueError):
            pkg_v = None
        try:
            saf_v = float(row[col_saf]) if col_saf is not None and col_saf < len(row) and row[col_saf] not in (None, '') else None
        except (TypeError, ValueError):
            saf_v = None
        if pkg_v and pkg_v >= 1:
            LEAR_PKG_BY_CODE[code] = int(pkg_v)
        if saf_v is not None and saf_v >= 0:
            LEAR_SAFETY_BY_CODE[code] = float(saf_v)
        n_rows += 1
    wb_l.close()
    print(f"  Lear из «{os.path.basename(path)}»: {n_rows} строк "
          f"(тара: {len(LEAR_PKG_BY_CODE)}, страх.запас: {len(LEAR_SAFETY_BY_CODE)})")
    return n_rows
SMC_WEEKDAY         = 0   # понедельник
PUREM_WEEKDAY       = 0
ECOTEXIS_ORDER_LAG_DAYS = 35
# шт (или кг) на 1 паллету; если нет — 1 паллета = 1 упаковка (pkg)
SMC_UNITS_PER_PALLET = {}

def normalize_supplier(s):
    if not s: return ''
    s = str(s).strip()
    key = s.upper()
    return SUPPLIER_ALIASES.get(key, s)

# ── Карта поставщик → вкладки плана. Перекрывает part_tab_map. ──
# По замечаниям: некоторые поставщики неправильно маппятся на сварку (WS) вместо сборки (AS).
SUPPLIER_TAB_MAP = {
    # Сборочные F_A — для деталей, идущих на B02/B04 (или универсальных)
    'MSA':         'AS_in_F_A',
    # Сборочные обе ветки (F_A + H_B)
    'SHUOWEIJIA':  'AS_in_F_A AS_in_H_B',
    'VIGOR':       'AS_in_F_A AS_in_H_B',
    'TEEC':        'AS_in_F_A AS_in_H_B',
    'TECHNOFORM':  'AS_in_F_A AS_in_H_B',
    'ITELMA':      'AS_in_F_A AS_in_H_B',
    'GANTO':       'AS_in_F_A AS_in_H_B',
    # NDKT: вкладки задаются TAB_OVERRIDES_PREFIX (B06 → AS_in_H_B), не обе AS сразу
    # Сварочные (Welding Shop) — обе ветки
    'SGK':         'WS_in WS_in_H',
    # Покрасочные (краски/химия) — обе ветки
    'LAB INDUSTRIES': 'PS_in PS_in_H_B',
    'LAB':            'PS_in PS_in_H_B',
    'LITUM':       'PS_in PS_in_H_B',
    'RESMETAL':    'PS_in PS_in_H_B',
    'ECOKHIM':     'PS_in PS_in_H_B',
    'YIERTE':      'PS_in PS_in_H_B',
    'EFTEC':       'PS_in PS_in_H_B',
}

# ── Переопределение применяемости BOM (замечания Корнеевой, Замечания-3) ──
# Формат: код → множество разрешённых конфигурационных ключей.
# Применяется ПОСЛЕ загрузки BOM — перезаписывает данные из xlsx-файла.
BOM_APPLICABILITY_OVERRIDES = {
    # Дорестайл-бамперы XST33/AST33 — overrides дополняются автоматически из bumper_clr_map
    # LK015530 (SGK): в BOM ошибочно помечен на B02+B04; должен только B04
    'LK015530': {'B04_4WD_TechPlus', 'B04_4WD_premium'},
    # 1101100XKJ23A (Yapp): в BOM нет отметок о применяемости → без override
    # код использует все конфиги (некорректно). XKJ23A = A08, применяется на всех A08.
    '1101100XKJ23A': {'A08_2WD_elite', 'A08_2WD_premium',
                      'A08_4WD_elite', 'A08_4WD_premium', 'A08_4WD_TechPlus'},
    # TODO (Корнеева): 1205526XGW01B / 1205527XGW01A — неверные остатки в файлах.
    #   Требует корректировки данных в источниках остатков.
    # TODO (Ильина): ALAB001234 — расчёт не соответствует действительности.
    #   Требует уточнения норм расхода (chem_norms) у ответственного.
}

# ── Ecoal'yance: матрица применяемости из «EcoAlliance use.xlsx» ──
# ECOALLIANCE_USE_FILE — см. авто-поиск после ROOT

# Колонки E–R (5–18): модель + привод из шапки «EcoAlliance use.xlsx»
ECO_COLUMN_SPECS = {
    5:  {'models': ['A01'], 'drive': '4WD'},
    6:  {'models': ['A01'], 'drive': '2WD'},
    7:  {'models': ['A01'], 'drive': '4WD'},
    8:  {'models': ['A01'], 'drive': '2WD'},
    9:  {'models': ['A08'], 'drive': '2WD'},
    10: {'models': ['A08'], 'drive': '4WD'},
    11: {'models': ['A08'], 'drive': '2WD'},
    12: {'models': ['A08'], 'drive': '4WD'},
    13: {'models': ['B02'], 'drive': '2WD'},
    14: {'models': ['B02'], 'drive': '4WD'},
    15: {'models': ['B02'], 'drive': '2WD'},
    16: {'models': ['B02'], 'drive': '4WD'},
    17: {'models': ['B04', 'B06', 'B16'], 'drive': '4WD'},
    18: {'models': ['B04', 'B06', 'B16'], 'drive': '4WD'},
}

ECOALLIANCE_PAIR_CODES = (
    frozenset({'1205112XST11A', '1205113XST11A'}),
)

eco_demand_rules = {}       # code → [{col, switch_batch, china_only}, ...]
eco_applicability = {}      # code → set(CFG_KEYS)
eco_superseded_codes = set()
_eco_batch_order_cache = {}  # (tab, month_num) → [(first_day, bv), ...]

_ECO_PART_CODE_RE = re.compile(r'\b(\d{7}[A-Z0-9]{4,10})\b')
_ECO_SWITCH_BATCH_RE = re.compile(r'\b(R[A-Z]{2}\d{4})\b')


def _eco_col_to_cfg_keys(col_num):
    spec = ECO_COLUMN_SPECS.get(col_num)
    if not spec:
        return set()
    out = set()
    for ck in CFG_KEYS:
        parts = ck.split('_')
        if len(parts) < 3:
            continue
        if parts[0] in spec['models'] and parts[1] == spec['drive']:
            out.add(ck)
    return out


def _parse_eco_cell_part_codes(text):
    return list(dict.fromkeys(_ECO_PART_CODE_RE.findall(str(text or '').upper())))


def _parse_eco_switch_batch(text):
    m = _ECO_SWITCH_BATCH_RE.search(str(text or '').upper())
    return _normalize_batch_id(m.group(1)) if m else None


def _eco_is_china_shipment(text):
    s = str(text or '')
    return '中国发货' in s or 'china' in s.lower() and 'ship' in s.lower()


def _eco_add_rule(code, col, switch_batch=None, china_only=False):
    code = normalize_code(str(code).strip())
    if not code:
        return
    eco_demand_rules.setdefault(code, [])
    rule = {'col': col, 'switch_batch': switch_batch, 'china_only': china_only}
    if rule not in eco_demand_rules[code]:
        eco_demand_rules[code].append(rule)
    eco_applicability.setdefault(code, set()).update(_eco_col_to_cfg_keys(col))


def load_ecoalliance_use(path):
    """Читает «EcoAlliance use.xlsx» → eco_demand_rules, eco_applicability."""
    global eco_demand_rules, eco_applicability, eco_superseded_codes
    eco_demand_rules = {}
    eco_applicability = {}
    eco_superseded_codes = set()
    if not path or not os.path.isfile(path):
        return 0, 0
    try:
        wb = load_workbook(path, data_only=True)
        ws = wb.active
        rows = list(ws.iter_rows(values_only=True))
        wb.close()
    except Exception as e:
        print(f"  ⚠️  EcoAlliance use: ошибка чтения {os.path.basename(path)}: {e}")
        return 0, 0

    supersession = {}
    data_rows = []
    for row in rows:
        row = list(row) if row else []
        while len(row) < 19:
            row.append(None)
        supp = str(row[0] or '').strip()
        code = normalize_code(str(row[1]).strip()) if row[1] else ''
        if not code or len(code) < 8:
            continue
        supp_l = supp.lower().replace("'", "'")
        if 'ecoalliance' not in supp_l and 'ecol' not in supp_l:
            continue
        note = str(row[3] or '')
        if '1205279' in note.upper() and code in ('1205526XGW01B', '1205527XGW01A'):
            supersession[code] = '1205279XGW01A'
        data_rows.append(row)

    for row in data_rows:
        code = normalize_code(str(row[1]).strip())
        for col in range(5, 19):
            cell = row[col - 1]
            if cell is None or str(cell).strip() in ('', 'nan', 'None', '——', '—'):
                continue
            text = str(cell)
            switch = _parse_eco_switch_batch(text)
            china = _eco_is_china_shipment(text)
            parts = _parse_eco_cell_part_codes(text)
            if not parts and china:
                _eco_add_rule(code, col, switch_batch=switch, china_only=True)
                continue
            for pc in parts:
                _eco_add_rule(pc, col, switch_batch=switch, china_only=china and len(parts) == 0)

    # 1205113 — зеркало 1205112, если в строке матрицы пусто
    if '1205112XST11A' in eco_demand_rules and '1205113XST11A' not in eco_demand_rules:
        eco_demand_rules['1205113XST11A'] = [dict(r) for r in eco_demand_rules['1205112XST11A']]
        eco_applicability['1205113XST11A'] = set(eco_applicability.get('1205112XST11A', set()))

    for old, new in supersession.items():
        eco_superseded_codes.add(old)
        if old in eco_demand_rules:
            eco_demand_rules.setdefault(new, [])
            for rule in eco_demand_rules.pop(old):
                if rule not in eco_demand_rules[new]:
                    eco_demand_rules[new].append(rule)
        if old in eco_applicability:
            eco_applicability.setdefault(new, set()).update(eco_applicability.pop(old))

    n_rules = sum(len(v) for v in eco_demand_rules.values())
    n_codes = len(eco_demand_rules)
    return n_codes, n_rules


def _eco_batch_order(tab, month_num):
    key = (tab, month_num)
    if key in _eco_batch_order_cache:
        return _eco_batch_order_cache[key]
    order = []
    if tab in plan_batches and month_num in plan_batches[tab]:
        for bv, dq in plan_batches[tab][month_num].items():
            first_day = min(dq.keys()) if dq else 999
            order.append((first_day, _normalize_batch_id(bv)))
    order.sort(key=lambda x: (x[0], x[1]))
    _eco_batch_order_cache[key] = order
    return order


def _eco_batch_at_or_after_switch(tab, month_num, bv, switch_batch):
    if not switch_batch:
        return True
    switch_batch = _normalize_batch_id(switch_batch)
    order = _eco_batch_order(tab, month_num)
    switch_idx = None
    for i, (_, b) in enumerate(order):
        if b == switch_batch or b.endswith(switch_batch[3:]):
            switch_idx = i
            break
    if switch_idx is None:
        return True
    bv_n = _normalize_batch_id(bv)
    for i, (_, b) in enumerate(order):
        if b == bv_n:
            return i >= switch_idx
    return False


def _batch_matches_eco_col(bi, col_num):
    spec = ECO_COLUMN_SPECS.get(col_num)
    if not spec or not bi:
        return False
    model = bi.get('model', '') or ''
    drive = bi.get('drive', '') or ''
    if model not in spec['models']:
        return False
    if spec['drive'] and drive and drive != spec['drive']:
        return False
    return True


def apply_ecoalliance_applicability(bom):
    """Переносит применяемость из EcoAlliance use → bom[code]['configs']."""
    updated = 0
    cleared = 0
    for code, cfgs in eco_applicability.items():
        if code in eco_superseded_codes:
            continue
        if code not in bom:
            bom[code] = {
                'name': code, 'unit': 'шт', 'supplier': ECOALYANCE_SUPPLIER,
                'configs': {}, 'model_qty': {}, 'section': 'EcoAlliance',
                'package': PACKAGE_OVERRIDES.get(code, 1), 'notes': 'EcoAlliance use',
            }
        bom[code]['configs'] = {k: 1 for k in sorted(cfgs)}
        bom[code]['supplier'] = normalize_supplier(bom[code].get('supplier') or ECOALYANCE_SUPPLIER)
        updated += 1
    for code in eco_superseded_codes:
        if code in bom:
            bom[code]['configs'] = {}
            cleared += 1
    return updated, cleared


def get_ecoalliance_daily_demand(code, month_num):
    """Потребность по матрице EcoAlliance (с партиями переключения RAS/RBU/RBA)."""
    rules = eco_demand_rules.get(code)
    if not rules:
        return None
    tabs = resolve_plan_tabs(part_tab_map.get(code, ''))
    if not tabs:
        return None
    daily = {}
    for rule in rules:
        if rule.get('china_only'):
            continue
        col = rule['col']
        col_cfgs = _eco_col_to_cfg_keys(col)
        if not col_cfgs:
            continue
        applicable = {k: 1 for k in col_cfgs}
        for tab in tabs:
            if tab not in plan_batches or month_num not in plan_batches[tab]:
                continue
            for bv, day_qty in plan_batches[tab][month_num].items():
                bi = _get_batch_info(tab, bv)
                if not _batch_matches_eco_col(bi, col):
                    continue
                switch = rule.get('switch_batch')
                if switch and not _eco_batch_at_or_after_switch(tab, month_num, bv, switch):
                    continue
                b_model = bi.get('model', '')
                b_drive = bi.get('drive', '')
                b_config = bi.get('config', '')
                cfg_key = f"{b_model}_{b_drive}_{b_config}"
                matched, q = cfg_match(cfg_key, applicable)
                if not matched:
                    continue
                for day, cars in day_qty.items():
                    daily[day] = daily.get(day, 0) + cars * q
    return daily if daily else {}

# ── Исключение из расчёта потребности (применяемость принудительно пуста) ──
# Имеет наивысший приоритет: перекрывает B02_2WD_elite, merge с прошлым выводом и overrides.
BOM_NO_DEMAND_CODES_DEFAULT = {
    '3101100XKN46A', '3101102XKN46A', '3101105XKN46A',
}
BOM_NO_DEMAND_EXCLUDE_SHEET = 'Исключить_применяемость'
BOM_NO_DEMAND_CODES = set(BOM_NO_DEMAND_CODES_DEFAULT)

def _is_no_demand_code(code):
    return code in BOM_NO_DEMAND_CODES

def _load_no_demand_codes_from_excel():
    """Читает лист Исключить_применяемость из MRP_System_v9.xlsx (ручные исключения)."""
    codes = set(BOM_NO_DEMAND_CODES_DEFAULT)
    paths = []
    if OUT_LIVE_INPUT and os.path.exists(OUT_LIVE_INPUT):
        paths.append(OUT_LIVE_INPUT)
    if os.path.exists(OUT) and OUT not in paths:
        paths.append(OUT)
    _mb = BOM_NEW if (BOM_NEW and os.path.exists(BOM_NEW)) else BOM_FILE
    if _mb and os.path.exists(_mb):
        paths.append(_mb)
    sheet_found = False
    for path in paths:
        try:
            wb = load_workbook(path, data_only=True)
        except Exception:
            continue
        if BOM_NO_DEMAND_EXCLUDE_SHEET not in wb.sheetnames:
            wb.close()
            continue
        sheet_found = True
        for row in wb[BOM_NO_DEMAND_EXCLUDE_SHEET].iter_rows(min_row=2, values_only=True):
            if not row:
                continue
            code = str(row[0]).strip() if row[0] else ''
            if not code or code.lower() in ('код', 'code', 'nan', 'none'):
                continue
            flag = row[1] if len(row) > 1 else None
            if flag is not None and str(flag).strip().lower() in ('да', 'yes', '1', '✓', 'v', 'y'):
                codes.discard(code)
            else:
                codes.add(code)
        wb.close()
        break
    return codes, sheet_found

def _apply_no_demand_exclusions(bom_dict):
    """Обнуляет применяемость для кодов без расчёта потребности."""
    cleared = 0
    for code in BOM_NO_DEMAND_CODES:
        if code in bom_dict:
            bom_dict[code]['configs'] = {}
            cleared += 1
    return cleared

# ── Точечные переопределения для конкретных кодов (выше supplier-карты) ──
# Точечные правила имеют наивысший приоритет.
CODE_TAB_OVERRIDES = {
    '1205100XKN63A': 'AS_in_F_A AS_in_H_B',   # Purem, по замечанию
    'ALAC012327':    'AS_in_F_A AS_in_H_B',   # по замечанию
    '5120107XST11A': 'WS_in (5)',             # по замечанию
}
def supplier_tab(supp):
    """Возвращает строку вкладок плана для данного поставщика, либо None."""
    if not supp: return None
    return SUPPLIER_TAB_MAP.get(str(supp).strip().upper())

# ── Применяемость x2 для задних подголовников (по 2 шт на машину) ──
HEADREST_X2_CODES = {
    '7008101XKN61A8P', '7008102XKN61A8P',
    '7008110XST11A8P', '7008112XKN61A8P',
    '7008120XST11A8P', '7008120XST11AZ6', '7008120XST11AQT',
}

# ── Применяемость (шт на машину) для кодов, где Master_BOM хранит только
#    галочку ✓ и количество выразить нечем. Замечание планировщика:
#    Уралэластотехника, 8402106AKJ20A — применяемость 2 ──
QTY_PER_CAR_OVERRIDE = {
    '8402106AKJ20A': 2,   # Уралэластотехника, 2 шт на машину (согласовано)
}

# ── Уралэластотехника: отгрузка целыми фурами, все коды группы едут вместе
#    (таблица «комплекты», согласовано). Упаковка кодов A01 = 360, но фура
#    только A01 везёт 480 — 480 НЕ кратно 360, поэтому коды этого поставщика
#    выведены из общей проверки кратности упаковки: здесь правило фуры главнее.
URAL_SUPPLIER = 'Uralelastotechnika'
URAL_A01_CODES = [
    '6107107AST01A', '6107300AST01A', '6207107AST01A', '6207100AST01A',
    '6107108AST01A', '6107400AST01A', '6207108AST01A', '6207200AST01A',
    '6307100AST01A', '8402110AST01A',
]
URAL_A08_CODES = [
    '6107107AKJ20A', '6107100AKJ20A', '6207107AKJ20A', '6207100AKJ20A',
    '6107108AKJ20A', '6107200AKJ20A', '6207108AKJ20A', '6207200AKJ20A',
    '6307101AKJ20A', '8402105AKJ20A', '8402106AKJ20A',
]
# шт на код в СМЕШАННОЙ фуре (A01 + A08)
URAL_MIXED_QTY = {c: 360 for c in URAL_A01_CODES}
URAL_MIXED_QTY.update({c: 120 for c in URAL_A08_CODES})
URAL_MIXED_QTY['8402106AKJ20A'] = 240      # применяемость 2 → двойная норма
# шт на код в фуре ТОЛЬКО A01
URAL_A01_ONLY_QTY = 480
# только A08 — не менее 4 комплектов (комплект = норма смешанной фуры)
URAL_A08_MIN_KITS = 4
URAL_ALL_CODES = set(URAL_A01_CODES) | set(URAL_A08_CODES)

# ── ВМ Авто: отгрузка кратна РЯДУ в кузове; сумма длин рядов ≤ длины кузова.
#    Источник: «Расчет упаковки в Фуру.xlsx» (иерархия короб → поддон → пачка → ряд).
#    шт в ряду = шт в коробе × коробов на поддоне × поддонов в пачке × пачек в ряду
VM_SUPPLIER = 'VM Auto'
VM_ROW_QTY = {           # штук в ряду
    '5006700AST01B': 640,    # порог левый:  20×8×2×2
    '5006800AST01B': 640,    # порог правый: 20×8×2×2
    '5173102XGW01B': 180,    # 102: 20×1×3×3
    '5174104XGW01E': 135,    # 104: 15×1×3×3
    '5304101AST01A': 216,    # А01 жабо: 6×6×3×2 (в калькуляторе код 5304100AST01A)
    '5304100XKJ20A': 144,    # А08 жабо: 12×2×3×2
}
VM_ROW_LEN_M = {         # длина ряда по полу кузова, м
    '5006700AST01B': 1.92, '5006800AST01B': 1.92,
    '5173102XGW01B': 1.30, '5174104XGW01E': 1.30,
    '5304101AST01A': 1.45, '5304100XKJ20A': 1.68,
}
VM_TRUCK_LEN_M = 13.4    # длина кузова

# ── МТС-авто: отгрузка целыми паллетами (16 коробов), поставки 1–2 раза в неделю.
#    Источник: «ГАБАРИТЫ ЗАКАЗОВ_HMMR_калькулятор.xlsx» (кол-во в 1 коробе по модели);
#    код → модель взяты из применяемости Master_BOM.
MTS_SUPPLIER = 'MTS Auto'
MTS_BOX_QTY = {          # штук в коробе
    '7925100XST33A': 14,   # A01
    '7925101XKJ23A': 10,   # A08
    '7925102XKN83A': 40,   # B16 / AVTOTOR
    '7925109XKN61A': 14,   # B02 / B04
    '7925103XKN46A': 14,   # B06
    '7925103XKN4AA': 14,   # B06
}
MTS_BOXES_PER_PALLET = 16
MTS_WEEKDAYS = (1, 4)    # вторник и пятница — 2 отгрузки в неделю

# ── 22 конфигурации BOM (B02_2WD_elite добавлена с v10) ──
CFG_KEYS = [
    'A01_2WD_comfort', 'A01_2WD_elite', 'A01_2WD_premium',
    'A01_4WD_elite', 'A01_4WD_premium', 'A01_4WD_TechPlus',
    'A08_2WD_elite', 'A08_2WD_premium', 'A08_4WD_elite',
    'A08_4WD_premium', 'A08_4WD_TechPlus',
    'B02_2WD_premium', 'B02_2WD_elite', 'B02_4WD_elite', 'B02_4WD_TechPlus',
    'B04_4WD_TechPlus', 'B04_4WD_premium',
    'B06_4WD_elite', 'B06_4WD_premium', 'B06_4WD_TechPlus',
    'B16_4WD_premium', 'B16_4WD_TechPlus',
]
CFG_KEYS_LEGACY_21 = [k for k in CFG_KEYS if k != 'B02_2WD_elite']
BOM_CONFIGS = {5 + i: k for i, k in enumerate(CFG_KEYS)}
BOM_LAST_CFG_COL = 5 + len(CFG_KEYS) - 1

B02_ALL_THREE = frozenset({'B02_2WD_premium', 'B02_4WD_elite', 'B02_4WD_TechPlus'})
# Дорестайл-бампера B02: только premium 2WD — не переносить на B02_2WD_elite
B02_PREM_2WD_ONLY_FRAGMENTS = ('XST33', 'AST33', 'AKN02')
# Рестайл B02 4WD: бамперы/кузовные XKN61 — только 4WD, не elite 2WD
B02_4WD_RESTYLE_FRAGMENTS = ('XKN61', 'KN260004', 'KN260005', 'AKN61')
# Дорестайл-бамперы XST33/AST33 (передние 2803120/2804104 и задние 2803130/2804105)
# — только B02 2WD premium из AS_in_F_A; рестайл XKN61 не входит
PRERESTYLE_B02_CONFIGS = {'premium'}
PRERESTYLE_B02_DRIVES = {'2WD'}
# Дорестайл-бамперы XST33/AST33 (2803120/2804104) ставятся на A01 2WD premium
# (подтверждено по плану: на 06.06 идут A01 4x2 premium, цвет C3 → бампер AC3).
PRERESTYLE_PREMIUM_MODEL = 'A01'
PRERESTYLE_PREMIUM_CFG = 'A01_2WD_premium'
PRERESTYLE_BUMPER_MARKERS = ('XST33', 'AST33')
PRERESTYLE_BUMPER_EXCLUDE = ('XKN61', 'KN260004', 'KN260005', 'AKN61')
# Дорестайл-бамперы B02 2WD premium, которые РЕАЛЬНО заказываются:
#   передний 2803120XST33*  и  задний 2804104AST33*
# Варианты 2803130XST33* / 2804105AST33* — это другая (не premium) модификация,
# на B02 2WD premium не ставится → потребность 0 (замечания пользователя 06.06).
PRERESTYLE_PREMIUM_BUMPER_PREFIXES = ('2803120XST33', '2804104AST33')


def _is_prerestyle_bumper_code(code):
    c = str(code)
    if any(x in c for x in PRERESTYLE_BUMPER_EXCLUDE):
        return False
    return any(m in c for m in PRERESTYLE_BUMPER_MARKERS)


def _is_prerestyle_premium_bumper_code(code):
    """Дорестайл-бампер, который ставится на B02 2WD premium (2803120/2804104)."""
    if not _is_prerestyle_bumper_code(code):
        return False
    c = str(code)
    return any(c.startswith(p) for p in PRERESTYLE_PREMIUM_BUMPER_PREFIXES)


def _normalize_batch_id(bv):
    return re.sub(r'\s+', '', str(bv or '').strip().upper())


def _map_plan_config(value):
    s = str(value or '').strip()
    if not s:
        return ''
    if s in CONFIG_MAP:
        return CONFIG_MAP[s]
    sl = s.lower()
    for k, v in CONFIG_MAP.items():
        if k.lower() == sl:
            return v
    if 'tech' in sl and 'plus' in sl:
        return 'TechPlus'
    if 'premium' in sl or '高配' in s or '中高配' in s:
        return 'premium'
    if 'elite' in sl or '中配' in s:
        return 'elite'
    if 'comfort' in sl or '低配' in s:
        return 'comfort'
    return ''


def _bom_header_has_b02_2wd_elite(hdr_row):
    for v in (hdr_row or []):
        vs = str(v or '').lower().replace('\n', ' ')
        if 'b02' in vs and '2wd' in vs and 'elite' in vs and 'prem' not in vs:
            return True
    return False


def detect_cfg_keys_from_header(hdr_row):
    """22 колонки в новом BOM; 21 — без B02_2WD_elite (сдвиг не нужен при чтении)."""
    if _bom_header_has_b02_2wd_elite(hdr_row):
        return list(CFG_KEYS)
    return list(CFG_KEYS_LEGACY_21)


def should_mark_b02_2wd_elite(code, cfgs):
    """Деталь для всех B02 / всех B02 2WD / всех B02 Elite → B02_2WD_elite."""
    if _is_prerestyle_bumper_code(code):
        return False
    if not cfgs or 'B02_2WD_elite' in cfgs:
        return False
    if not any(k.startswith('B02_') for k in cfgs):
        return False
    has_prem = 'B02_2WD_premium' in cfgs
    has_elite = 'B02_4WD_elite' in cfgs
    has_tech = 'B02_4WD_TechPlus' in cfgs
    # все B02 (3 существующие конфигурации до elite 2WD)
    if B02_ALL_THREE <= set(cfgs):
        if any(p in code for p in B02_PREM_2WD_ONLY_FRAGMENTS) and 'XKN61' not in code:
            return False
        return True
    # все B02 2WD (premium; не дорестайл-only)
    if has_prem and not has_elite and not has_tech:
        if any(p in code for p in B02_PREM_2WD_ONLY_FRAGMENTS) and 'XKN61' not in code:
            return False
        return True
    # все B02 Elite (4WD elite; не рестайл 4WD-only)
    if has_elite and not has_tech and not has_prem:
        if any(p in code for p in B02_4WD_RESTYLE_FRAGMENTS):
            return False
        return True
    # premium 2WD + elite 4WD (без Tech+) — elite-линейка + все 2WD
    if has_prem and has_elite and not has_tech:
        return True
    return False


def apply_b02_2wd_elite_marks(cfgs, code):
    if not isinstance(cfgs, dict):
        cfgs = {k: 1 for k in (cfgs or [])}
    if should_mark_b02_2wd_elite(code, cfgs):
        qty = 1
        for src in ('B02_2WD_premium', 'B02_4WD_elite'):
            if src in cfgs and isinstance(cfgs[src], (int, float)) and cfgs[src] != 1:
                qty = cfgs[src]
                break
        cfgs['B02_2WD_elite'] = qty
    return cfgs


# ── Карта fuzzy-match для нестандартных конфигов ──
# При несовпадении прямого ключа пробуем альтернативный (та же модель+комплектация, другой привод).
CFG_FUZZY_ALT = {
    'B02_2WD_TechPlus': ['B02_4WD_TechPlus'],
    'B04_2WD_elite':    ['B04_4WD_premium','B04_4WD_TechPlus'],
    'B04_2WD_premium':  ['B04_4WD_premium'],
    'B06_2WD_elite':    ['B06_4WD_elite'],
    'B16_2WD_elite':    ['B16_4WD_premium','B16_4WD_TechPlus'],
}

def cfg_match(cfg_key, applicable):
    """Возвращает (matched_key, qty) если cfg_key подходит под BOM-applicability,
    учитывая fuzzy-match для нестандартных конфигов. Иначе (None, 0)."""
    if not applicable:
        return None, 0
    if cfg_key in applicable:
        qty = applicable[cfg_key] if isinstance(applicable, dict) else 1
        return cfg_key, qty
    for alt in CFG_FUZZY_ALT.get(cfg_key, []):
        if alt in applicable:
            qty = applicable[alt] if isinstance(applicable, dict) else 1
            return alt, qty
    return None, 0


def _cfg_model_config(key):
    """('A01_2WD_premium') → ('A01','premium'); привод (2WD/4WD) игнорируется."""
    parts = str(key).split('_')
    if len(parts) >= 3:
        return (parts[0], parts[2])
    if len(parts) == 2:
        return (parts[0], parts[1])
    return (key, '')


def _bumper_cfg_match(code, cfg_key, applicable_b):
    """Бамперы сопоставляются по (модель, комплектация) без учёта привода.

    Применяемость берётся из BOM (лист BOM_Детальный). 21-конфигурационный BOM
    не содержит колонок B02_2WD_elite и B02_4WD_premium, но эти кузова есть в плане
    и используют тот же бампер, что и B02 4WD elite / B02 2WD premium соответственно
    — поэтому сравниваем по (модель+комплектация), привод не важен для бампера.
    """
    if not applicable_b:
        return None, 0
    mc = _cfg_model_config(cfg_key)
    for ak in applicable_b:
        if _cfg_model_config(ak) == mc:
            qty = applicable_b[ak] if isinstance(applicable_b, dict) else 1
            if not (isinstance(qty, (int, float)) and qty > 0):
                qty = 1
            return ak, qty
    return None, 0





def fill(h): return PatternFill("solid", fgColor=h)
H_FILL=fill("1F3864"); H2=fill("2E75B6"); H3=fill("375623")
MAY_F=fill("2E75B6"); JUN_F=fill("1F7A4C"); JUL_F=fill("7B3F91")
YEL_F=fill("FFFF00"); RED_F=fill("FFD7D7"); GRN_F=fill("E2EFDA")
BLU_F=fill("D9E1F2"); GRY_F=fill("F2F2F2"); ORG_F=fill("FFF0CC")
CYN_F=fill("E0F7FF"); MAN_F=fill("FEFFD0"); NRM_F=fill("EDF7E0")
NO_F=PatternFill(fill_type=None)
# M_FILL строится динамически после определения MONTHS (см. ниже)
_M_COLORS = [MAY_F, JUN_F, JUL_F]  # цвета для 3 месяцев
# M_FILL заполняется после MONTHS (строка ниже — placeholder, перезаписывается)
M_FILL = {5:MAY_F,6:JUN_F,7:JUL_F}  # будет перезаписан

def hcell(ws,r,c,txt,bg=None,fc="FFFFFF",bold=True,sz=9,wrap=True,align='center'):
    cl=ws.cell(r,c); cl.value=txt
    cl.font=Font(bold=bold,color=fc,size=sz,name="Arial")
    if bg: cl.fill=bg
    cl.alignment=Alignment(horizontal=align,vertical='center',wrap_text=wrap)
    return cl

MODEL_MAP={'New A01':'A01','New A08':'A08','New B02':'B02','New B04':'B04',
           'New B06':'B06','New B16':'B16','A01':'A01','A08':'A08',
           'B02':'B02','B04':'B04','B06':'B06','B06X':'B06','B16':'B16'}
DRIVE_MAP={'4x2':'2WD','4x4':'4WD','4X2':'2WD','4X4':'4WD',
           '6MT两驱':'2WD','两驱':'2WD','四驱':'4WD','6MT 4x2':'2WD',
           '4x2':'2WD','4x4':'4WD',
           '2WD':'2WD','4WD':'4WD','FWD':'2WD','AWD':'4WD',
           '前驱':'2WD','后驱':'4WD','前轮驱动':'2WD','四轮驱动':'4WD'}
CONFIG_MAP={'comfort低配':'comfort','elite中配':'elite','Elite低配':'elite',
            'elite超低配':'elite','premium高配':'premium','Premium中配':'premium',
            'Premium中高配':'premium','Premium':'premium','Premium低配':'premium',
            'Premium高配':'premium',
            'Tech Plus顶配':'TechPlus','Tech+高配':'TechPlus',
            '顶配Tech Plus':'TechPlus','Tech plus顶配':'TechPlus'}
# BOM_CONFIGS — см. CFG_KEYS выше (22 конфигурации)
B16_ELITE={'6803112XKN08A','6903112XKN08A'}
# ══════════════════════════════════════════════════════════════
# АВТО-ОПРЕДЕЛЕНИЕ РАСЧЁТНОГО ПЕРИОДА (3 месяца, скользящий)
# Стартовый месяц = месяц актуальности остатков (из имени файла)
# или текущий месяц если имя не распознано.
# ══════════════════════════════════════════════════════════════
import calendar as _cal, re as _re

_MONTH_EN_FULL = {1:'January',2:'February',3:'March',4:'April',5:'May',6:'June',
                  7:'July',8:'August',9:'September',10:'October',11:'November',12:'December'}
_MONTH_EN_SHORT = {m: _MONTH_EN_FULL[m][:3] for m in range(1,13)}

# Авто-определяем дату актуальности остатков из имени файла (шаблон ХХ.ДД.ММ.ГГ.xlsx)
_today = datetime.date.today()
_stock_date = None
for _sf in sorted(STOCK_FILES, key=os.path.getmtime, reverse=True):
    _bn = os.path.basename(_sf)
    _mm = _re.search(r'[._](\d{2})[._](\d{2})[._](\d{2})\.(xlsx|xls)$', _bn, _re.IGNORECASE)
    if _mm:
        try:
            _d, _m, _y = int(_mm.group(1)), int(_mm.group(2)), 2000+int(_mm.group(3))
            if 1<=_d<=31 and 1<=_m<=12:
                _stock_date = datetime.date(_y, _m, _d)
                break
        except: pass
if _stock_date is None:
    _stock_date = _today  # fallback — текущая дата

# Дата запуска расчёта — от неё считается оставшаяся потребность в 1-м расчётном месяце
CALC_RUN_DATE   = _today
CALC_RUN_DAY    = CALC_RUN_DATE.day
CALC_RUN_MONTH  = CALC_RUN_DATE.month
CALC_RUN_YEAR   = CALC_RUN_DATE.year

# ── Режимы запуска (галочки в run_mrp_gui / переменные окружения) ──
# Всегда: ежедневная потребность + графики поставок.
# Опционально: месячные заказы (листы Заказы_*, Сводка_3мес, Прогноз_CKD)
# и выгрузка графиков по поставщикам.
def _env_flag(name, default='0'):
    return os.environ.get(name, default).strip().lower() in ('1', 'true', 'да', 'yes', 'y', 'вкл')

CALC_ORDERS = _env_flag('MRP_CALC_ORDERS', '0')
EXPORT_SUPPLIERS = _env_flag('MRP_EXPORT_SUPPLIERS', '0')
print(f"[INIT] Режим запуска: потребность + графики поставок (всегда) | "
      f"месячные заказы: {'ВКЛ' if CALC_ORDERS else 'выкл'} | "
      f"выгрузка поставщикам: {'ВКЛ' if EXPORT_SUPPLIERS else 'выкл'}")

STOCK_AS_OF_MONTH = _stock_date.month
STOCK_AS_OF_DAY   = _stock_date.day

# ── Скользящее окно: если текущий месяц опередил дату остатков — сдвигаем ──
# Например: остатки от 30.05, запускаем 01.06 → начинаем с июня (→ июн/июл/авг).
_start_year = _stock_date.year
if (_today.year > _stock_date.year or
        (_today.year == _stock_date.year and _today.month > _stock_date.month)):
    STOCK_AS_OF_MONTH = _today.month
    STOCK_AS_OF_DAY   = _today.day
    _start_year       = _today.year
    print(f"[INIT] Скользящий сдвиг: остатки от {_stock_date}, сегодня {_today} "
          f"→ расчётный период начинается с {STOCK_AS_OF_MONTH:02d}.{_start_year}")

# 3 расчётных месяца начиная с STOCK_AS_OF_MONTH
def _make_months(start_month, start_year, n=3):
    result = []
    for i in range(n):
        _m = (start_month - 1 + i) % 12 + 1
        _y = start_year + (start_month - 1 + i) // 12
        _nd = _cal.monthrange(_y, _m)[1]
        _label = f"{_MONTH_EN_SHORT[_m]} {_y}"   # e.g. "May 2026", "Aug 2026", "Jan 2027"
        result.append((_m, _label, _nd))
    return result

MONTHS = _make_months(STOCK_AS_OF_MONTH, _start_year)

# MONTH_KW — ключевые слова для поиска месяца в Plan-Fact файле
MONTH_KW = {mnum: f"{mnum:02d} {_MONTH_EN_FULL[mnum]} {int(mlabel.split()[1])}"
            for mnum, mlabel, _ in MONTHS}

# Год каждого расчётного месяца (для datetime.date)
MONTH_YEAR = {mnum: int(mlabel.split()[1]) for mnum, mlabel, _ in MONTHS}

# Короткие имена (3 буквы) для имён листов: Заказы_May, Потребность_Jun и т.д.
MONTH_SHORT = {mnum: mlabel[:3] for mnum, mlabel, _ in MONTHS}

# Диапазон для заголовков (напр. "Май–Июль 2026" или "Авг 2026–Окт 2026")
_M_RU = {1:'Янв',2:'Фев',3:'Мар',4:'Апр',5:'Май',6:'Июн',7:'Июл',8:'Авг',9:'Сен',10:'Окт',11:'Ноя',12:'Дек'}
PERIOD_LABEL = (f"{_M_RU[MONTHS[0][0]]}–{_M_RU[MONTHS[2][0]]} {int(MONTHS[0][1].split()[1])}"
                if int(MONTHS[0][1].split()[1]) == int(MONTHS[2][1].split()[1])
                else f"{_M_RU[MONTHS[0][0]]} {int(MONTHS[0][1].split()[1])}–{_M_RU[MONTHS[2][0]]} {int(MONTHS[2][1].split()[1])}")

print(f"[INIT] Расчётный период: {PERIOD_LABEL}  (остатки: {STOCK_AS_OF_DAY:02d}.{STOCK_AS_OF_MONTH:02d}.{_stock_date.year}, "
      f"расчёт с {CALC_RUN_DAY:02d}.{CALC_RUN_MONTH:02d}.{CALC_RUN_YEAR})")

# Теперь обновляем M_FILL и WEEK_CLR под реальные номера месяцев
M_FILL = {mnum: _M_COLORS[i] for i,(mnum,_,_) in enumerate(MONTHS)}
# WEEK_CLR обновится после build_weeks()

# ── TAB_COLS verified against actual xlsx ─────────────────────
TAB_COLS={
    'WS_in (5)': dict(batch=3,model=5,drive=8,config=9,day1=11),  # FIXED was wrong
    'WS_in_H':   dict(batch=1,model=3,drive=6,config=7,day1=9),
    'PS_in':     dict(batch=2,model=3,drive=6,config=7,day1=9),
    'PS_in_H_B': dict(batch=2,model=3,drive=6,config=7,day1=9),
    'AS_in_F_A': dict(batch=2,model=3,drive=6,config=7,day1=9),
    'AS_in_H_B': dict(batch=2,model=3,drive=6,config=7,day1=9),
}
# Имена листов Plan-Fact (точные)
_PLAN_TAB_SHEETS = frozenset({
    'WS_in (5)', 'WS_in_H', 'PS_in', 'PS_in_H_B', 'AS_in_F_A', 'AS_in_H_B',
})

def resolve_plan_tabs(plan_tab_str):
    """
    Вкладки Plan-Fact для расчёта потребности.
    Одно имя (напр. 'AS_in_F_A', 'WS_in (5)') = только этот лист.
    Явные пары в part_tab_map ('WS_in WS_in_H', 'AS_in_F_A AS_in_H_B') = оба листа.
    'WS_in' без уточнения = обе сварочные (WS_in (5) + WS_in_H).
    """
    if not plan_tab_str:
        return []
    s = str(plan_tab_str).strip()
    _MULTI = {
        'SW_in_F_A SW_IN_H_B': ['AS_in_F_A', 'AS_in_H_B'],
        'WS_in WS_in_H': ['WS_in (5)', 'WS_in_H'],
        'PS_in PS_in_H_B': ['PS_in', 'PS_in_H_B'],
        'AS_in_F_A AS_in_H_B': ['AS_in_F_A', 'AS_in_H_B'],
        'AS_in_H_B AS_in_F_A': ['AS_in_F_A', 'AS_in_H_B'],
        'AS_in_H_B  AS_in_F_A': ['AS_in_F_A', 'AS_in_H_B'],
    }
    if s in _MULTI:
        return _MULTI[s]
    if s in _PLAN_TAB_SHEETS:
        return [s]
    if s == 'WS_in':
        return ['WS_in (5)', 'WS_in_H']
    if s == 'PS_in':
        return ['PS_in']
    return [s] if s else []
SAFETY_DAYS=7
SAFETY_DAYS_FILE = (_latest(ROOT, "Stock in days*.xlsx")
                    or os.path.join(ROOT, "Stock in days.xlsx"))  # ищется и в подпапках
# Минимальная доля дневной потребности, которая должна быть перекрыта остатком
# на НАЧАЛО дня. Поставка ставится в ПРЕДЫДУЩИЙ день (рассчитывается под спрос
# следующего дня), чтобы к началу дня спроса остаток уже покрывал ≥ этой доли —
# без отрицательных остатков на начало дня.
DELIVERY_START_COVER = 0.5
# Допуск превышения номинала фуры (0 = строго по лимиту, без перегруза).
TRUCK_CAPACITY_TOLERANCE = 0.0
# Допуск ниже целевого Ss (safety + cover×next_dem). Жёсткий пол: Ss ≥ potr без допуска.
DELIVERY_SS_TARGET_TOLERANCE = 0.20

def _soft_truck_cap(nominal):
    return float(nominal) * (1.0 + TRUCK_CAPACITY_TOLERANCE)

def _soft_ss_target(full_target, code=None, supplier=None):
    """Мягкая цель остатка: допускаем отклонение вниз от полной цели."""
    return float(full_target) * (1.0 - DELIVERY_SS_TARGET_TOLERANCE)

def build_weeks():
    weeks=[]; wnum=1
    _yr0 = MONTH_YEAR[MONTHS[0][0]]; _yr2 = MONTH_YEAR[MONTHS[2][0]]
    d=datetime.date(_yr0, MONTHS[0][0], 1)
    end=datetime.date(_yr2, MONTHS[2][0], MONTHS[2][2])
    while d<=end:
        wd=d.weekday()
        wend=min(d+datetime.timedelta(days=6-wd),end)
        weeks.append({'label':f'W{wnum}','start':d,'end':wend})
        wnum+=1; d=wend+datetime.timedelta(days=1)
    return weeks
WEEKS=build_weeks()
_MONTH_COLOR_HEX = {mnum: ['2E75B6','1F7A4C','7B3F91'][i] for i,(mnum,_,_) in enumerate(MONTHS)}
WEEK_CLR=[_MONTH_COLOR_HEX.get(w['start'].month,'1F3864') for w in WEEKS]

def date_to_week(dt):
    for i,w in enumerate(WEEKS):
        if w['start']<=dt<=w['end']: return i
    return -1

# ═══ STEP 0: Load V7 reference data ══════════════════════════
print("Loading V7 reference data (embedded)...")
part_tab_map   = {"1101100AGW01A": "AS_in_F_A", "1101100BGW02A": "AS_in_F_A", "1101100XKJ23A": "AS_in_F_A", "1101101AGW01A": "AS_in_F_A", "1101116AGW02A": "AS_in_F_A", "1201593XGW02A": "WS_in", "1201594XGW02A": "WS_in", "1201729XGW02A": "WS_in", "1201730XGW02A": "WS_in", "1201731XGW02A": "WS_in", "1201732XGW02A": "WS_in", "1201733XGW02A": "WS_in", "1201734XGW02A": "WS_in", "1201735XGW02A": "WS_in", "1201UQ5XGW01A": "WS_in", "1201UQ6XGW01A": "WS_in", "1201UQ7XGW01A": "WS_in", "1201UQ8XGW01A": "WS_in", "1201UQ9XGW01A": "WS_in", "1201UR1XGW01A": "WS_in", "1201UR2XGW01A": "WS_in", "1201UR3XGW01A": "WS_in", "1201UR4XGW01A": "WS_in", "1201UR5XGW01A": "WS_in", "1201UR6XGW01A": "WS_in", "1201UR7XGW01A": "WS_in", "1205102XST1AA": "WS_in WS_in_H", "1205105XST1AA": "WS_in (5)", "1205112XST11A": "WS_in WS_in_H", "1205113XST11A": "WS_in WS_in_H", "1205526XGW01B": "WS_in WS_in_H", "1205527XGW01A": "WS_in WS_in_H", "1205653XKJ22A": "WS_in", "1205654XKJ22A": "WS_in", "1205655XKJ22A": "WS_in", "1205656XKJ22A": "WS_in", "1205657XKJ22A": "WS_in", "1205658XKJ22A": "WS_in", "1205659XKJ22A": "WS_in", "1205660XKJ22A": "WS_in", "2803104XKN61A8T": "AS_in_F_A", "2803104XKN61A9C": "AS_in_F_A", "2803104XKN61AC3": "AS_in_F_A", "2803104XKN61AH4": "AS_in_F_A", "2803105XKN61A8T": "AS_in_F_A", "2803105XKN61A9C": "AS_in_F_A", "2803105XKN61AC3": "AS_in_F_A", "2803105XKN61AH4": "AS_in_F_A", "2803120XST33A5B": "AS_in_F_A", "2803120XST33A8T": "AS_in_F_A", "2803120XST33A9C": "AS_in_F_A", "2803120XST33AC3": "AS_in_F_A", "2803120XST33AGN": "AS_in_F_A", "2803130XST33A5B": "AS_in_F_A", "2803130XST33A8T": "AS_in_F_A", "2803130XST33A9C": "AS_in_F_A", "2803130XST33AC3": "AS_in_F_A", "2803130XST33AGN": "AS_in_F_A", "2804104AST33A5B": "AS_in_F_A", "2804104AST33A8T": "AS_in_F_A", "2804104AST33A9C": "AS_in_F_A", "2804104AST33AC3": "AS_in_F_A", "2804104AST33AGN": "AS_in_F_A", "2804105AST33A5B": "AS_in_F_A", "2804105AST33A8T": "AS_in_F_A", "2804105AST33A9C": "AS_in_F_A", "2804105AST33AC3": "AS_in_F_A", "2804105AST33AGN": "AS_in_F_A", "2804111XKN61A8T": "AS_in_F_A", "2804111XKN61A9C": "AS_in_F_A", "2804111XKN61AC3": "AS_in_F_A", "2804111XKN61AH4": "AS_in_F_A", "2804112XKN61A8T": "AS_in_F_A", "2804112XKN61A9C": "AS_in_F_A", "2804112XKN61AC3": "AS_in_F_A", "2804112XKN61AH4": "AS_in_F_A", "2804KN260004A8T": "AS_in_F_A", "2804KN260004A9C": "AS_in_F_A", "2804KN260004AC3": "AS_in_F_A", "2804KN260004AH4": "AS_in_F_A", "2804KN260005A8T": "AS_in_F_A", "2804KN260005A9C": "AS_in_F_A", "2804KN260005AC3": "AS_in_F_A", "2804KN260005AH4": "AS_in_F_A", "2906101BGW02A": "AS_in_F_A", "2906113BGW02A": "AS_in_F_A", "2916107BGW02A": "AS_in_F_A", "3101100XKJ23A": "AS_in_F_A", "3101100XKN46A": "AS_in_H_B", "3101100xst33A": "AS_in_F_A", "3101101XKN61A": "AS_in_F_A", "3101101XST33A": "AS_in_F_A", "3101102XKN46A": "AS_in_H_B", "3101105XKN46A": "AS_in_H_B", "3703100XKQ04A": "AS_in_F_A", "3904101XK82XA": "SW_in_F_A SW_IN_H_B", "5010211XST11A": "WS_in (5)", "5010212XST11A": "WS_in (5)", "5010213XST11A": "WS_in (5)", "5010214XST11A": "WS_in (5)", "5010217XST11A": "WS_in (5)", "5010218XST11A": "WS_in (5)", "5010223XST11A": "WS_in (5)", "5010224XST11A": "WS_in (5)", "5010226XST11A": "WS_in (5)", "5010228XST11A": "WS_in (5)", "5010232XST11A": "WS_in (5)", "5010241XST11A": "WS_in (5)", "5010242XST11A": "WS_in (5)", "5010800XST11A": "WS_in (5)", "5010806XST11A": "WS_in (5)", "5010831XST11A": "WS_in (5)", "5109101AST11A86": "AS_in_F_A", "5109125XKN61A8P": "AS_in_F_A", "5109201AST11A86": "AS_in_F_A", "5109201XKN61A8P": "AS_in_F_A", "5109501AST11A86": "AS_in_F_A", "5109800XKN46A": "AS_in_H_B", "5109800XKN83A": "AS_in_H_B", "5120101XGW01A": "WS_in", "5120102XGW01A": "WS_in", "5120107XST11А": "WS_in (5)", "5120108XST11B": "WS_in (5)", "5120364XGW01A": "WS_in (5)", "5120365XGW01A": "WS_in (5)", "5120383XGW01A": "WS_in", "5120384XGW01A": "WS_in", "5120807XST11A": "WS_in (5)", "5120808XST11A": "WS_in (5)", "5122105XST11A": "WS_in (5)", "5122106XST11A": "WS_in (5)", "5122113XST11A": "WS_in (5)", "5122114XST11A": "WS_in (5)", "5122121XST11A": "WS_in (5)", "5122122XST11A": "WS_in (5)", "5122211XST11A": "WS_in (5)", "5122212XST11A": "WS_in (5)", "5122213XST11A": "WS_in (5)", "5122214XST11A": "WS_in (5)", "5122218XGW01A": "WS_in", "5122221XGW01A": "WS_in", "5130103XGW02A": "WS_in (5)", "5130104XGW02A": "WS_in (5)", "5130121XGW01A": "WS_in", "5130127XGW01A": "WS_in", "5130128XGW01A": "WS_in", "5130137XST11A": "WS_in (5)", "5130138XST11A": "WS_in (5)", "5130167XGW01A": "WS_in", "5130199XGW01B": "WS_in", "5130200XGW01B": "WS_in", "5130203XGW01A": "WS_in", "5130219XST11A": "WS_in (5)", "5130248XST11A": "WS_in (5)", "5130253XST11A": "WS_in (5)", "5130381XST11A": "WS_in (5)", "5130397XST11A": "WS_in (5)", "5130448XST11A": "WS_in (5)", "5130449XST11A": "WS_in (5)", "5130461XST11A": "WS_in (5)", "5130477XST11A": "WS_in (5)", "5130478XST11A": "WS_in (5)", "5130555XGW01B": "WS_in", "5130621XGW02A": "WS_in (5)", "5130807XST11A": "WS_in (5)", "5130851XST11A": "WS_in (5)", "5130852XST11A": "WS_in (5)", "5130883XST11A": "WS_in (5)", "5130884XST11A": "WS_in (5)", "5130891XST11A": "WS_in (5)", "5173102XGW01B": "AS_in_F_A", "5174104XGW01E": "AS_in_F_A", "5206100XKN83A": "AS_in_H_B", "5206101XKN46A": "AS_in_H_B", "5206102XKN46A": "AS_in_H_B", "5206103XKN46A": "AS_in_H_B", "5206104XKN61A": "AS_in_F_A", "5206105XKN61A": "AS_in_F_A", "5206200XKN83A": "AS_in_H_B", "5206300XKN83A": "AS_in_H_B", "5206300XST10A": "AS_in_F_A", "5206400XKJ23A": "AS_in_F_A", "5206500XKJ23A": "AS_in_F_A", "5206603XST10A": "AS_in_F_A", "5206605XST10A": "AS_in_F_A", "5300103XGW01A": "WS_in", "5300120XGW01B": "WS_in", "5300123XGW01B": "WS_in", "5300164XGW01B": "WS_in", "5301102XST33A": "WS_in (5)", "5301105XST33A": "WS_in (5)", "5301111XST11A": "WS_in (5)", "5301113XST11A": "WS_in (5)", "5301114XST11A": "WS_in (5)", "5301120XST11A": "WS_in", "5301122XST11A": "WS_in (5)", "5301126XKN02A": "WS_in", "5301131XST11A": "WS_in (5)", "5301133XST11A": "WS_in (5)", "5301134XKN02A": "WS_in", "5301151XST11A": "WS_in (5)", "5301152XST11A": "WS_in (5)", "5301153XST11A": "WS_in (5)", "5301163XST11A": "WS_in (5)", "5301302XST11A": "WS_in (5)", "5301501XST11A": "WS_in", "5301507XST11A": "WS_in (5)", "5301551XST11A": "WS_in (5)", "5301552XST11A": "WS_in (5)", "5304100XKN02A": "AS_in_F_A", "5401105XKN06A": "WS_in (5)", "5401106XKN06A": "WS_in (5)", "5401107XKN06A": "WS_in (5)", "5401108XKN06A": "WS_in (5)", "5401109XKN06A": "WS_in (5)", "5401110XKN06A": "WS_in (5)", "5401119XKN06A": "WS_in", "5401120XKN06A": "WS_in", "5401121XKN06A": "WS_in", "5401122XKN06A": "WS_in", "5401187XKN01B": "WS_in", "5401261XKN02A": "WS_in", "5401262XKN02A": "WS_in", "5401289XKN61A": "WS_in (5)", "5401291XKN61A": "WS_in (5)", "5401301XST11A": "WS_in (5)", "5401302XST11A": "WS_in (5)", "5401311XST11A": "WS_in (5)", "5401312XST11A": "WS_in (5)", "5401413XST11A": "WS_in (5)", "5401414XST11A": "WS_in (5)", "5401417XKN02A": "WS_in", "5401418XKN02A": "WS_in", "5401421XKN02A": "WS_in", "5401422XKN02A": "WS_in", "5401517XST11A": "WS_in (5)", "5401518XST11A": "WS_in (5)", "5401523XST11A": "WS_in (5)", "5401524XST11A": "WS_in (5)", "5401571XST33A": "WS_in", "5401572XST33A": "WS_in", "5401601XKN02A": "WS_in", "5401602XKN02A": "WS_in", "5401701XKN02A": "WS_in", "5401702XKN02A": "WS_in", "5401717XST01A": "WS_in (5)", "5401718XST01A": "WS_in (5)", "5401733XST11A": "WS_in (5)", "5401734XST11A": "WS_in (5)", "5401831XST11A": "WS_in (5)", "5401832XST11A": "WS_in (5)", "5401871XST33A": "WS_in", "5401872XST33A": "WS_in", "5401901XKN02A": "WS_in (5)", "5401902XKN02A": "WS_in (5)", "5401907XKN02A": "WS_in", "5401908XKN02A": "WS_in", "5402411XKN02B": "AS_in_F_A", "5402422XKN02B": "AS_in_F_A", "5402433XKN02A": "AS_in_F_A", "5402444XKN02A": "AS_in_F_A", "5531015AKN02A": "AS_in_F_A", "5531106XGW01B": "PS_in", "5531107XGW01B": "PS_in", "5531109XGW01B": "PS_in", "5531114XGW01B": "PS_in", "5531115XGW01B": "PS_in", "5531117XGW01B": "PS_in", "5531118XGW01B": "PS_in", "5531121XST11A": "PS_in", "5531122XST11A": "PS_in", "5531123XST11A": "PS_in", "5531131XST11A": "PS_in", "5531132XST11A": "PS_in", "5531133XST11A": "PS_in", "5531160XKN01B": "PS_in", "5531161XKN01B": "PS_in", "5531311XST11A": "PS_in", "5531313XST11A": "PS_in", "5601115xkn02a": "WS_in (5)", "5601117xkn02a": "WS_in (5)", "5601120XKN02A": "WS_in", "5601123XKN02A": "WS_in", "5604100XKN46A": "AS_in_F_A", "5604100XKN83A": "AS_in_H_B", "5604101AKN02A": "PS_in", "5604101XKN46A": "AS_in_F_A", "5604101XST11A": "AS_in_F_A", "5604102AKN02A": "PS_in", "5604102XKN46A": "AS_in_F_A", "5604105XST11A": "AS_in_F_A", "5604106XST11A": "AS_in_F_A", "5604107AKN02A": "PS_in", "5604108AKN02A": "PS_in", "5604300XKJ23A": "AS_in_F_A", "5604400XKJ23A": "AS_in_F_A", "5701102xkn02a": "WS_in (5)", "5701103xkn02a": "WS_in (5)", "5701109XKN06A": "WS_in", "5701151XST11A": "WS_in (5)", "6101135XKN02A": "WS_in", "6101136XKN02A": "WS_in", "6103100AKN02B": "AS_in_F_A", "6103100XKN46A": "AS_in_F_A", "6103101XKJ23A": "AS_in_F_A", "6103102XKJ23A": "AS_in_F_A", "6103200AKN02B": "AS_in_F_A", "6103200XKN46A": "AS_in_F_A", "6103300XST11A": "AS_in_F_A", "6103400XST11A": "AS_in_F_A", "6107100AKJ20A": "AS_in_F_A", "6107107AKJ20A": "AS_in_F_A", "6107107AST01A": "AS_in_F_A", "6107108AKJ20A": "AS_in_F_A", "6107108AST01A": "AS_in_F_A", "6107200AKJ20A": "AS_in_F_A", "6107300AST01A": "AS_in_F_A", "6107400AST01A": "AS_in_F_A", "6201103XST11A": "WS_in (5)", "6201104XST11A": "WS_in (5)", "6201121XKN02A": "WS_in", "6201122XKN02A": "WS_in", "6203100AKN02B": "AS_in_F_A", "6203100AKN61A": "AS_in_F_A", "6203100AST11A": "AS_in_F_A", "6203100XKJ23A": "AS_in_F_A", "6203100XKN46A": "AS_in_F_A", "6203100XKN83A": "AS_in_F_A", "6203100XST11A": "AS_in_F_A", "6203101XKJ23A": "AS_in_F_A", "6203102XKJ23A": "AS_in_F_A", "6203200AKN02B": "AS_in_F_A", "6203200AKN61A": "AS_in_F_A", "6203200AST11A": "AS_in_F_A", "6203200XKJ23A": "AS_in_F_A", "6203200XKN46A": "AS_in_F_A", "6203200XKN83A": "AS_in_F_A", "6203200XST11A": "AS_in_F_A", "6203300XKN46A": "AS_in_F_A", "6203300XKN83A": "AS_in_F_A", "6203400XKN46A": "AS_in_F_A", "6203400XKN83A": "AS_in_F_A", "6207100AKJ20A": "AS_in_F_A", "6207100AST01A": "AS_in_F_A", "6207107AKJ20A": "AS_in_F_A", "6207107AST01A": "AS_in_F_A", "6207108AKJ20A": "AS_in_F_A", "6207108AST01A": "AS_in_F_A", "6207200AKJ20A": "AS_in_F_A", "6207200AST01A": "AS_in_F_A", "6303100AKN04B": "AS_in_H_B", "6303100XKN02B": "AS_in_F_A", "6303100XKN83A": "AS_in_H_B", "6303100XST11A": "AS_in_F_A", "6303101XST11A": "AS_in_F_A", "6303200AKN04B": "AS_in_H_B", "6303200XKN61A": "AS_in_F_A", "6303200XKN62A": "AS_in_F_A", "6303200XKN83A": "AS_in_H_B", "6307100AST01A": "AS_in_F_A", "6307101AKJ20A": "AS_in_F_A", "6802101XKN61A8P": "AS_in_F_A", "6802102XKN61A8P": "AS_in_F_A", "6802105XKN61A8P": "AS_in_F_A", "6802107XKN61A8P": "AS_in_F_A", "6802110XST33A8P": "AS_in_F_A", "6802111XKN61A8P": "AS_in_F_A", "6802300XST33A8P": "AS_in_F_A", "6802300XST33AQT": "AS_in_F_A", "6802300XST33AZ6": "AS_in_F_A", "6802410XST33A8P": "AS_in_F_A", "6802500XST33A8P": "AS_in_F_A", "6802700XST33A8P": "AS_in_F_A", "6802700XST33AQT": "AS_in_F_A", "6802700XST33AZ6": "AS_in_F_A", "6803110XKJ20A": "AS_in_F_A", "6803111XKN02A": "AS_in_F_A", "6803112XKN02A": "AS_in_F_A", "6803112XKN08A": "AS_in_H_B", "6803113XKN01C": "AS_in_F_A", "6803113XST01A": "AS_in_F_A", "6803114XKN08A": "AS_in_H_B", "6805108XST33A": "AS_in_F_A", "6805117XKN61A": "AS_in_F_A", "6805210XKJ22A": "AS_in_F_A", "6805210XST01A": "AS_in_F_A", "6805268XKN81A": "AS_in_H_B", "6808101XKN61A8P": "AS_in_F_A", "6808201XKN61A8P": "AS_in_F_A", "6902101XKN61A8P": "AS_in_F_A", "6902102XKN61A8P": "AS_in_F_A", "6902110XST33A8P": "AS_in_F_A", "6902300XST33A8P": "AS_in_F_A", "6902300XST33AQT": "AS_in_F_A", "6902300XST33AZ6": "AS_in_F_A", "6902410XST33A8P": "AS_in_F_A", "6902500XST33A8P": "AS_in_F_A", "6902700XST33A8P": "AS_in_F_A", "6902700XST33AQT": "AS_in_F_A", "6902700XST33AZ6": "AS_in_F_A", "6903101XKN61A8P": "AS_in_F_A", "6903102XKN61A8P": "AS_in_F_A", "6903103XKN61A8P": "AS_in_F_A", "6903110XKJ20A": "AS_in_F_A", "6903112XKN02A": "AS_in_F_A", "6903112XKN08A": "AS_in_H_B", "6903113XKN02A": "AS_in_F_A", "6903113XKN08A": "AS_in_H_B", "6903116XKN01C": "AS_in_F_A", "6905107XST33A": "AS_in_F_A", "6905111XKN61A": "AS_in_F_A", "6905210XKJ22A": "AS_in_F_A", "6905210XST01A": "AS_in_F_A", "6905273XKN81A": "AS_in_H_B", "6908108XKN61A8P": "AS_in_F_A", "6908208XKN61A8P": "AS_in_F_A", "7002101XKN61A8P": "AS_in_F_A", "7002102XKN61A8P": "AS_in_F_A", "7002105XKN61A8P": "AS_in_F_A", "7002106XKN61A8P": "AS_in_F_A", "7002130XST33A8P": "AS_in_F_A", "7002130XST33AQT": "AS_in_F_A", "7002130XST33AZ6": "AS_in_F_A", "7002141XST33A8P": "AS_in_F_A", "7002211XST33A8P": "AS_in_F_A", "7002230XST33A8P": "AS_in_F_A", "7002230XST33AQT": "AS_in_F_A", "7002230XST33AZ6": "AS_in_F_A", "7002310XST33A8P": "AS_in_F_A", "7002440XST33A8P": "AS_in_F_A", "7002440XST33AQT": "AS_in_F_A", "7002440XST33AZ6": "AS_in_F_A", "7003100XKJ22A": "AS_in_F_A", "7003101XKN61A": "AS_in_F_A", "7003110XST01A": "AS_in_F_A", "7003110XST11A": "AS_in_F_A", "7003112XKN81A": "AS_in_H_B", "7005104XKN07A": "AS_in_F_A", "7005107XKN81A": "AS_in_H_B", "7005110XKJ22A": "AS_in_F_A", "7005111XKN07A": "AS_in_F_A", "7005113XKN81B": "AS_in_H_B", "7005127XKN61A8P": "AS_in_F_A", "7005128XKN61A8P": "AS_in_F_A", "7005310XKJ22A": "AS_in_F_A", "7005320XST01A": "AS_in_F_A", "7005410XST01A": "AS_in_F_A", "7008100XST11A8P": "AS_in_F_A", "7008101XKN61A8P": "AS_in_F_A", "7008102XKN61A8P": "AS_in_F_A", "7008110XST11A8P": "AS_in_F_A", "7008112XKN61A8P": "AS_in_F_A", "7008113XKN61A8P": "AS_in_F_A", "7008120XST11A8P": "AS_in_F_A", "7008120XST11AQT": "AS_in_F_A", "7008120XST11AZ6": "AS_in_F_A", "7008200XST11A8P": "AS_in_F_A", "7008200XST11AZ6": "AS_in_F_A", "7925100XKJ23A": "AS_in_F_A", "7925100XKN83A": "AS_in_H_B", "7925100XST33A": "AS_in_F_A", "7925101XKJ23A": "AS_in_F_A", "7925101XKN83A": "AS_in_H_B AS_in_F_A", "7925102XKN46A": "AS_in_H_B", "7925102XKN61A": "AS_in_F_A", "7925102XKN83A": "AS_in_H_B", "7925103XKN46A": "AS_in_H_B", "7925104XKQ41B": "AS_in_F_A", "7925105XKQ41A": "AS_in_F_A", "7925108XST11A": "AS_in_F_A", "7925109XKN61A": "AS_in_F_A", "7925109XST11A": "AS_in_F_A", "7925113XST11A": "AS_in_F_A", "8400131XGW01A": "WS_in", "8400138XGW01B": "WS_in", "8400141XST11A": "WS_in (5)", "8400142XST11A": "WS_in (5)", "8400156XGW01B": "WS_in", "8400157XST11A": "WS_in (5)", "8400158XST11A": "WS_in (5)", "8400171XST11A": "WS_in (5)", "8400172XST11A": "WS_in (5)", "8400173XST11A": "WS_in (5)", "8400174XST11A": "WS_in (5)", "8400181XST11A": "WS_in (5)", "8400195XGW01B": "WS_in", "8400202XGW01B": "WS_in", "8400220XGW01B": "WS_in", "8400323XST11A": "WS_in (5)", "8400324XST11A": "WS_in (5)", "8400331XST11A": "WS_in (5)", "8400332XST11A": "WS_in (5)", "8400343XST11A": "WS_in (5)", "8400344XST11A": "WS_in (5)", "8400371XST11A": "WS_in (5)", "8400372XST11A": "WS_in (5)", "8400411XST11A": "WS_in (5)", "8400412XST11A": "WS_in (5)", "8400413XST11A": "WS_in (5)", "8400414XST11A": "WS_in (5)", "8400421XST11A": "WS_in (5)", "8400422XST11A": "WS_in (5)", "8400451XST11A": "WS_in", "8400452XST11A": "WS_in", "8400473XST11A": "WS_in (5)", "8400474XST11A": "WS_in (5)", "8400751XST11A": "WS_in", "8400752XST11A": "WS_in", "8402105AKJ20A": "AS_in_F_A", "8402106AKJ20A": "AS_in_F_A", "8402110AST01A": "AS_in_F_A", "8402115XST11A": "WS_in (5)", "8402116XST11A": "WS_in (5)", "ALAA001234": "AS_in_F_A AS_in_H_B", "ALAA003255": "PS_in PS_in_H_B", "ALAA003256": "PS_in PS_in_H_B", "ALAA003257": "PS_in PS_in_H_B", "ALAA003286": "PS_in PS_in_H_B", "ALAA003287": "PS_in PS_in_H_B", "ALAA003288": "PS_in PS_in_H_B", "ALAA003289": "PS_in PS_in_H_B", "ALAA003338": "PS_in PS_in_H_B", "ALAA003339": "PS_in PS_in_H_B", "ALAA003526": "PS_in PS_in_H_B", "ALAA004586": "PS_in PS_in_H_B", "ALAA004871": "PS_in PS_in_H_B", "ALAA004873": "PS_in PS_in_H_B", "ALAA005669": "PS_in PS_in_H_B", "ALAA005672": "PS_in PS_in_H_B", "ALAA006324": "PS_in PS_in_H_B", "ALAA006522": "PS_in PS_in_H_B", "ALAA006523": "PS_in PS_in_H_B", "ALAA006524": "PS_in PS_in_H_B", "ALAA006526": "PS_in PS_in_H_B", "ALAA007195": "PS_in PS_in_H_B", "ALAA007196": "PS_in PS_in_H_B", "ALAA007197": "PS_in PS_in_H_B", "ALAA007198": "PS_in PS_in_H_B", "ALAA007199": "PS_in PS_in_H_B", "ALAA007200": "PS_in PS_in_H_B", "ALAA007201": "PS_in PS_in_H_B", "ALAA007202": "PS_in PS_in_H_B", "ALAA007267": "PS_in PS_in_H_B", "ALAA007434": "PS_in PS_in_H_B", "ALAA007435": "PS_in PS_in_H_B", "ALAA007440": "PS_in PS_in_H_B", "ALAA007466": "PS_in PS_in_H_B", "ALAA007467": "PS_in PS_in_H_B", "ALAA007732": "PS_in PS_in_H_B", "ALAA007765": "PS_in PS_in_H_B", "ALAA007787": "PS_in PS_in_H_B", "ALAA008937": "PS_in PS_in_H_B", "ALAB000025": "SW_in_F_A SW_IN_H_B", "ALAB000028": "SW_in_F_A SW_IN_H_B", "ALAB000046": "SW_in_F_A SW_IN_H_B", "ALAB000049": "SW_in_F_A SW_IN_H_B", "ALAB000052": "SW_in_F_A SW_IN_H_B", "ALAB000055": "SW_in_F_A SW_IN_H_B", "ALAB000091": "SW_in_F_A SW_IN_H_B", "ALAB000208": "SW_in_F_A SW_IN_H_B", "ALAB000321": "SW_in_F_A SW_IN_H_B", "ALAB000663": "SW_in_F_A SW_IN_H_B", "ALAB000664": "SW_in_F_A SW_IN_H_B", "ALAB000709": "SW_in_F_A SW_IN_H_B", "ALAB000986": "SW_in_F_A SW_IN_H_B", "ALAB001037": "SW_in_F_A SW_IN_H_B", "ALAB001069": "SW_in_F_A SW_IN_H_B", "ALAB001078": "AS_in_F_A AS_in_H_B", "ALAB001079": "AS_in_F_A AS_in_H_B", "ALAB001080": "AS_in_F_A AS_in_H_B", "ALAB001507": "AS_in_H_B AS_in_F_A", "ALAB001571": "PS_in PS_in_H_B", "ALAB001572": "PS_in PS_in_H_B", "ALAB001573": "PS_in PS_in_H_B", "ALAC000077": "AS_in_F_A", "ALAC000386": "AS_in_H_B  AS_in_F_A", "ALAC011940": "PS_in PS_in_H_B", "ALAC011941": "PS_in PS_in_H_B", "ALAC011971": "AS_in_H_B  AS_in_F_A", "ALAC011977": "AS_in_H_B  AS_in_F_A", "ALAC012158": "SW_in_F_A SW_IN_H_B", "ALAC012327": "SW_in_F_A SW_IN_H_B", "ALAC012335": "AS_in_F_A AS_in_H_B", "ALAC012350": "AS_in_H_B  AS_in_F_A", "ALAC013461": "AS_in_H_B  AS_in_F_A", "ALAC013503": "AS_in_F_A", "ALAC013836": "AS_in_F_A", "LK012145": "WS_in", "LK012146": "WS_in", "LK012147": "WS_in", "LK012155": "WS_in", "LK012156": "WS_in", "LK012714": "WS_in", "LK013036": "WS_in", "LK014388": "WS_in", "LK014389": "WS_in", "LK014390": "WS_in", "LK014391": "WS_in", "LK014393": "WS_in", "LK014396": "WS_in", "LK014399": "WS_in", "LK014400": "WS_in", "LK015519": "WS_in", "LK015520": "WS_in", "LK015521": "WS_in", "LK015522": "WS_in", "LK015524": "WS_in", "LK015525": "WS_in", "LK015526": "WS_in", "LK015528": "WS_in", "LK015529": "WS_in", "LK015530": "WS_in", "LK015560": "WS_in", "LK015561": "WS_in", "LK015567": "WS_in", "LK015569": "WS_in", "LK015570": "WS_in", "LK015571": "WS_in", "LK015572": "WS_in", "LP000293": "WS_in", "LP000294": "WS_in", "LP000295": "WS_in", "LP000339": "WS_in", "LP000340": "WS_in"}
chem_norms     = {"ALAA003255": {"A01": 1.14, "A08": 1.14, "B16": 1.14, "B06": 1.14, "B04": 1.14, "B02": 1.14}, "ALAA003256": {"A01": 1.18, "A08": 1.18, "B16": 1.18, "B06": 1.18, "B04": 1.18, "B02": 1.18}, "ALAA003257": {"A01": 3.15, "A08": 3.15, "B16": 3.15, "B06": 3.15, "B04": 3.15, "B02": 3.15}, "ALAA003286": {"A01": 0.79, "A08": 0.79, "B16": 0.79, "B06": 0.79, "B04": 0.79, "B02": 0.79}, "ALAA003287": {"A01": 1.56, "A08": 1.56, "B16": 1.56, "B06": 1.56, "B04": 1.56, "B02": 1.56}, "ALAA003288": {"A01": 0.613, "A08": 0.613, "B16": 0.613, "B06": 0.613, "B04": 0.613, "B02": 0.613}, "ALAA003289": {"A01": 0.08, "A08": 0.08, "B16": 0.08, "B06": 0.08, "B04": 0.08, "B02": 0.08}, "ALAA003338": {"A01": 9.1, "A08": 9.1, "B16": 9.1, "B06": 9.1, "B04": 9.1, "B02": 9.1}, "ALAA003339": {"A01": 1.5, "A08": 1.5, "B16": 1.5, "B06": 1.5, "B04": 1.5, "B02": 1.5}, "ALAA004586": {"A01": 2.75, "A08": 2.75, "B16": 2.75, "B06": 2.75, "B04": 2.75, "B02": 2.75}, "ALAA005672": {"A08": 4.14, "B16": 4.14, "B06": 4.14}, "ALAA006324": {"A01": 0.12, "A08": 0.12, "B16": 0.12, "B06": 0.12, "B04": 0.12, "B02": 0.12}, "ALAA006522": {"A01": 0.26, "A08": 0.26, "B16": 0.26, "B06": 0.26, "B04": 0.26, "B02": 0.26}, "ALAA006524": {"A01": 0.007, "A08": 0.007, "B16": 0.007, "B06": 0.007, "B04": 0.007, "B02": 0.007}, "ALAA006526": {"A01": 0.28, "A08": 0.28, "B16": 0.28, "B06": 0.28, "B04": 0.28, "B02": 0.28}, "ALAA007195": {"A01": 1.13, "A08": 1.13, "B16": 1.13, "B06": 1.13, "B04": 1.13, "B02": 1.13}, "ALAA007196": {"A01": 0.11, "A08": 0.11, "B16": 0.11, "B06": 0.11, "B04": 0.11, "B02": 0.11}, "ALAA007197": {"A01": 0.03, "A08": 0.03, "B16": 0.03, "B06": 0.03, "B04": 0.03, "B02": 0.03}, "ALAA007200": {"A01": 0.05, "A08": 0.05, "B16": 0.05, "B06": 0.05, "B04": 0.05, "B02": 0.05}, "ALAA007201": {"A01": 0.2, "A08": 0.2, "B16": 0.2, "B06": 0.2, "B04": 0.2, "B02": 0.2}, "ALAA007202": {"A01": 1.06, "A08": 1.06, "B16": 1.06, "B06": 1.06, "B04": 1.06, "B02": 1.06}, "ALAA007267": {"A01": 1.8, "A08": 1.8, "B16": 1.8, "B06": 1.8, "B04": 1.8, "B02": 1.8}, "ALAA007466": {"A01": 2.82, "B16": 2.82, "B06": 2.82, "B04": 2.82, "B02": 2.82}, "ALAA007467": {"B04": 2.68, "B02": 2.68}, "ALAA007765": {"A01": 0.3, "A08": 0.3, "B16": 0.3, "B06": 0.3, "B04": 0.3, "B02": 0.3}, "ALAA007787": {"A01": 0.049, "A08": 0.049, "B16": 0.049, "B06": 0.049, "B04": 0.049, "B02": 0.049}, "ALAA008937": {"A01": 1.56, "A08": 1.56, "B16": 1.56, "B06": 1.56, "B04": 1.56, "B02": 1.56}, "ALAB000025": {"A01": 0.059, "B04": 0.05, "B02": 0.05}, "ALAB000028": {"A01": 0.19, "B04": 0.2, "B02": 0.2}, "ALAB000046": {"A01": 126.0, "A08": 0.0718, "B04": 65.6, "B02": 65.6}, "ALAB000049": {"A01": 1.4, "A08": 1.02, "B04": 1.56, "B02": 1.49}, "ALAB000052": {"A01": 0.11, "A08": 0.11, "B04": 0.11, "B02": 0.11}, "ALAB000208": {"A01": 0.566, "B04": 0.596, "B02": 0.596}, "ALAB000321": {"A01": 0.047, "B04": 0.03, "B02": 0.05}, "ALAB000663": {"A01": 0.456, "A08": 0.593, "B04": 0.599, "B02": 0.593}, "ALAB000664": {"B16": 0.347, "B06": 0.277}, "ALAB000986": {"A01": 0.0143, "A08": 0.089, "B04": 0.0303, "B02": 0.0292}, "ALAB001037": {"A01": 38.4, "A08": 67.7, "B16": 24.0, "B06": 24.0, "B04": 35.8, "B02": 36.5}, "ALAB001571": {"A01": 8400.0, "A08": 8400.0, "B16": 8400.0, "B06": 8400.0, "B04": 8400.0, "B02": 8400.0}, "ALAB001572": {"A01": 6000.0, "A08": 6000.0, "B16": 6000.0, "B06": 6000.0, "B04": 6000.0, "B02": 6000.0}, "ALAB001573": {"A01": 700.0, "A08": 700.0, "B16": 700.0, "B06": 700.0, "B04": 700.0, "B02": 700.0}, "ALAC000386": {"A01": 8.5, "A08": 8.5, "B16": 8.5, "B06": 8.5, "B04": 8.5, "B02": 8.5}, "ALAC011940": {"A01": 300.0, "A08": 300.0, "B16": 300.0, "B06": 300.0, "B04": 300.0, "B02": 300.0}, "ALAC011941": {"A01": 450.0, "A08": 450.0, "B16": 450.0, "B06": 450.0, "B04": 450.0, "B02": 450.0}, "ALAC011977": {"A01": 0.65, "A08": 0.65, "B16": 0.65, "B06": 0.65, "B04": 0.65, "B02": 0.65}, "ALAC012335": {"A01": 0.8, "A08": 0.8, "B16": 0.8, "B06": 0.8, "B04": 0.8, "B02": 0.8}, "ALAC013461": {"A01": 8.0, "A08": 8.0, "B16": 8.0, "B06": 8.0, "B04": 8.0, "B02": 8.0}}
paint_colors   = {"ALAA003255": ["GN_RED", "BLUE_5B"], "ALAA003256": ["ORANGE", "FU_GREY"], "ALAA003257": ["WHITE_C1"], "ALAA003286": ["WHITE_C1"], "ALAA003287": False, "ALAA003288": False, "ALAA003289": False, "ALAA003338": False, "ALAA003339": False, "ALAA004586": ["GOLDEN_BLACK", "CRYSTAL_BLACK"], "ALAA005672": ["ORANGE"], "ALAA006324": False, "ALAA006522": False, "ALAA006524": False, "ALAA006526": False, "ALAA007195": False, "ALAA007196": False, "ALAA007197": False, "ALAA007200": False, "ALAA007201": False, "ALAA007202": False, "ALAA007267": False, "ALAA007466": ["C3_GREY"], "ALAA007467": ["ATLANTIS"], "ALAA007765": False, "ALAA007787": False, "ALAA008937": False, "ALAB000025": False, "ALAB000028": False, "ALAB000046": False, "ALAB000049": False, "ALAB000052": False, "ALAB000208": False, "ALAB000321": False, "ALAB000663": False, "ALAB000664": False, "ALAB000986": False, "ALAB001037": False, "ALAB001571": False, "ALAB001572": False, "ALAB001573": False, "ALAC000386": False, "ALAC011940": False, "ALAC011941": False, "ALAC011977": False, "ALAC012335": False, "ALAC013461": False, "ALAA004871": ["GN_RED"], "ALAA004873": ["BLUE_5B"], "ALAA007732": ["KU_GREY"]}
bumper_clr_map = {"2803104XKN61A8T": "GOLDEN_BLACK", "2803104XKN61A9C": "WHITE_C1", "2803104XKN61AC3": "C3_GREY", "2803104XKN61AH4": "ATLANTIS", "2803105XKN61A8T": "GOLDEN_BLACK", "2803105XKN61A9C": "WHITE_C1", "2803105XKN61AC3": "C3_GREY", "2803105XKN61AH4": "ATLANTIS", "2803120XST33A5B": "BLUE_5B", "2803120XST33A8T": "GOLDEN_BLACK", "2803120XST33A9C": "WHITE_C1", "2803120XST33AC3": "C3_GREY", "2803120XST33AGN": "GN_RED", "2803130XST33A5B": "BLUE_5B", "2803130XST33A8T": "GOLDEN_BLACK", "2803130XST33A9C": "WHITE_C1", "2803130XST33AC3": "C3_GREY", "2803130XST33AGN": "GN_RED", "2804104AST33A5B": "BLUE_5B", "2804104AST33A8T": "GOLDEN_BLACK", "2804104AST33A9C": "WHITE_C1", "2804104AST33AC3": "C3_GREY", "2804104AST33AGN": "GN_RED", "2804105AST33A5B": "BLUE_5B", "2804105AST33A8T": "GOLDEN_BLACK", "2804105AST33A9C": "WHITE_C1", "2804105AST33AC3": "C3_GREY", "2804105AST33AGN": "GN_RED", "2804111XKN61A8T": "GOLDEN_BLACK", "2804111XKN61A9C": "WHITE_C1", "2804111XKN61AC3": "C3_GREY", "2804111XKN61AH4": "ATLANTIS", "2804112XKN61A8T": "GOLDEN_BLACK", "2804112XKN61A9C": "WHITE_C1", "2804112XKN61AC3": "C3_GREY", "2804112XKN61AH4": "ATLANTIS", "2804KN260004A8T": "GOLDEN_BLACK", "2804KN260004A9C": "WHITE_C1", "2804KN260004AC3": "C3_GREY", "2804KN260004AH4": "ATLANTIS", "2804KN260005A8T": "GOLDEN_BLACK", "2804KN260005A9C": "WHITE_C1", "2804KN260005AC3": "C3_GREY", "2804KN260005AH4": "ATLANTIS"}
v7_daily_raw   = {"1101100AGW01A": {"5": {"4": 430.0, "5": 310.0, "6": 430.0, "7": 310.0, "8": 360.0, "11": 245.0, "12": 355.0, "13": 360.0, "14": 455.0, "15": 325.0, "16": 180.0, "18": 480.0, "19": 340.0, "20": 310.0, "21": 430.0, "22": 430.0, "23": 275.0, "25": 430.0, "26": 430.0, "27": 550.0, "28": 405.0, "29": 335.0, "30": 265.0}, "6": {"1": 770.0, "2": 630.0, "3": 715.0, "4": 450.0, "5": 450.0, "6": 435.0, "7": 310.0, "8": 700.0, "9": 745.0, "10": 755.0, "11": 490.0, "13": 125.0, "15": 405.0, "16": 330.0, "17": 450.0, "18": 390.0, "19": 465.0, "20": 240.0, "22": 350.0, "23": 390.0, "24": 460.0, "25": 380.0, "26": 340.0, "27": 205.0, "29": 275.0, "30": 325.0}}, "1101100BGW02A": {"5": {"4": 520.0, "5": 640.0, "6": 640.0, "7": 640.0, "8": 560.0, "11": 440.0, "12": 640.0, "13": 640.0, "14": 640.0, "15": 640.0, "16": 320.0, "18": 640.0, "19": 640.0, "20": 640.0, "21": 640.0, "22": 640.0, "23": 320.0, "25": 640.0, "26": 640.0, "27": 640.0, "28": 640.0, "29": 640.0, "30": 320.0}, "6": {"1": 1225.0, "2": 1185.0, "3": 925.0, "4": 540.0, "5": 530.0, "6": 705.0, "7": 520.0, "8": 940.0, "9": 1060.0, "10": 1170.0, "11": 735.0, "13": 270.0, "14": 95.0, "15": 570.0, "16": 560.0, "17": 815.0, "18": 390.0, "19": 550.0, "20": 285.0, "22": 560.0, "23": 530.0, "24": 550.0, "25": 470.0, "26": 340.0, "27": 205.0, "29": 330.0, "30": 450.0}}, "1101100XKJ23A": {"5": {"4": 190.0, "5": 90.0, "6": 170.0, "7": 70.0, "8": 60.0, "11": 65.0, "12": 115.0, "13": 120.0, "14": 95.0, "15": 85.0, "16": 60.0, "18": 120.0, "19": 100.0, "20": 80.0, "21": 180.0, "22": 70.0, "23": 50.0, "25": 60.0, "26": 65.0, "27": 115.0, "28": 120.0, "29": 60.0, "30": 60.0}, "6": {"1": 60.0, "2": 120.0, "3": 60.0, "4": 60.0, "5": 60.0, "6": 60.0, "8": 60.0, "9": 120.0, "10": 60.0}}, "1101101AGW01A": {"5": {"4": 430.0, "5": 310.0, "6": 430.0, "7": 310.0, "8": 360.0, "11": 245.0, "12": 355.0, "13": 360.0, "14": 455.0, "15": 325.0, "16": 180.0, "18": 480.0, "19": 340.0, "20": 310.0, "21": 430.0, "22": 430.0, "23": 275.0, "25": 430.0, "26": 430.0, "27": 550.0, "28": 405.0, "29": 335.0, "30": 265.0}, "6": {"1": 770.0, "2": 630.0, "3": 715.0, "4": 450.0, "5": 450.0, "6": 435.0, "7": 310.0, "8": 700.0, "9": 745.0, "10": 755.0, "11": 490.0, "13": 125.0, "15": 405.0, "16": 330.0, "17": 450.0, "18": 390.0, "19": 465.0, "20": 240.0, "22": 350.0, "23": 390.0, "24": 460.0, "25": 380.0, "26": 340.0, "27": 205.0, "29": 275.0, "30": 325.0}}, "1101116AGW02A": {"5": {"4": 430.0, "5": 310.0, "6": 430.0, "7": 310.0, "8": 360.0, "11": 245.0, "12": 355.0, "13": 360.0, "14": 455.0, "15": 325.0, "16": 180.0, "18": 480.0, "19": 340.0, "20": 310.0, "21": 430.0, "22": 430.0, "23": 275.0, "25": 430.0, "26": 430.0, "27": 550.0, "28": 405.0, "29": 335.0, "30": 265.0}, "6": {"1": 770.0, "2": 630.0, "3": 715.0, "4": 450.0, "5": 450.0, "6": 435.0, "7": 310.0, "8": 700.0, "9": 745.0, "10": 755.0, "11": 490.0, "13": 125.0, "15": 405.0, "16": 330.0, "17": 450.0, "18": 390.0, "19": 465.0, "20": 240.0, "22": 350.0, "23": 390.0, "24": 460.0, "25": 380.0, "26": 340.0, "27": 205.0, "29": 275.0, "30": 325.0}}, "1201593XGW02A": {"5": {"4": 100.0, "5": 160.0, "6": 80.0, "7": 60.0, "8": 120.0, "11": 180.0, "12": 120.0, "13": 205.0, "14": 95.0, "15": 405.0, "16": 255.0, "18": 330.0, "19": 310.0, "20": 310.0, "21": 430.0, "22": 190.0, "23": 170.0, "25": 295.0, "26": 485.0, "27": 360.0, "28": 85.0, "29": 190.0, "30": 35.0}, "6": {"1": 425.0, "2": 475.0, "3": 330.0, "4": 260.0, "5": 240.0, "6": 415.0, "7": 370.0, "8": 295.0, "9": 495.0, "10": 365.0, "11": 280.0, "16": 240.0, "17": 120.0, "26": 120.0, "29": 295.0, "30": 65.0}}, "1201594XGW02A": {"5": {"4": 100.0, "5": 160.0, "6": 80.0, "7": 60.0, "8": 120.0, "11": 180.0, "12": 120.0, "13": 205.0, "14": 95.0, "15": 405.0, "16": 255.0, "18": 330.0, "19": 310.0, "20": 310.0, "21": 430.0, "22": 190.0, "23": 170.0, "25": 295.0, "26": 485.0, "27": 360.0, "28": 85.0, "29": 190.0, "30": 35.0}, "6": {"1": 425.0, "2": 475.0, "3": 330.0, "4": 260.0, "5": 240.0, "6": 415.0, "7": 370.0, "8": 295.0, "9": 495.0, "10": 365.0, "11": 280.0, "16": 240.0, "17": 120.0, "26": 120.0, "29": 295.0, "30": 65.0}}, "1201729XGW02A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 120.0, "12": 240.0, "13": 240.0, "14": 240.0, "20": 120.0, "22": 240.0, "23": 105.0, "25": 135.0, "26": 65.0, "27": 55.0, "28": 240.0, "29": 240.0, "30": 120.0}, "6": {"1": 215.0, "2": 240.0, "3": 355.0, "4": 225.0, "5": 50.0, "6": 90.0, "8": 480.0, "9": 240.0, "10": 240.0, "11": 135.0, "15": 345.0, "16": 195.0, "17": 270.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 240.0, "27": 120.0, "30": 385.0}}, "1201730XGW02A": {"5": {"4": 310.0, "5": 430.0, "6": 310.0, "7": 430.0, "8": 240.0, "11": 300.0, "12": 360.0, "13": 445.0, "14": 335.0, "15": 405.0, "16": 255.0, "18": 330.0, "19": 310.0, "20": 430.0, "21": 430.0, "22": 430.0, "23": 275.0, "25": 430.0, "26": 550.0, "27": 415.0, "28": 325.0, "29": 430.0, "30": 155.0}, "6": {"1": 640.0, "2": 715.0, "3": 685.0, "4": 485.0, "5": 290.0, "6": 505.0, "7": 370.0, "8": 775.0, "9": 735.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "1201731XGW02A": {"5": {"4": 310.0, "5": 430.0, "6": 310.0, "7": 430.0, "8": 240.0, "11": 300.0, "12": 360.0, "13": 445.0, "14": 335.0, "15": 405.0, "16": 255.0, "18": 330.0, "19": 310.0, "20": 430.0, "21": 430.0, "22": 430.0, "23": 275.0, "25": 430.0, "26": 550.0, "27": 415.0, "28": 325.0, "29": 430.0, "30": 155.0}, "6": {"1": 640.0, "2": 715.0, "3": 685.0, "4": 485.0, "5": 290.0, "6": 505.0, "7": 370.0, "8": 775.0, "9": 735.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "1201732XGW02A": {"5": {"4": 310.0, "5": 430.0, "6": 310.0, "7": 430.0, "8": 240.0, "11": 300.0, "12": 360.0, "13": 445.0, "14": 335.0, "15": 405.0, "16": 255.0, "18": 330.0, "19": 310.0, "20": 430.0, "21": 430.0, "22": 430.0, "23": 275.0, "25": 430.0, "26": 550.0, "27": 415.0, "28": 325.0, "29": 430.0, "30": 155.0}, "6": {"1": 640.0, "2": 715.0, "3": 685.0, "4": 485.0, "5": 290.0, "6": 505.0, "7": 370.0, "8": 775.0, "9": 735.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "1201733XGW02A": {"5": {"4": 100.0, "5": 160.0, "6": 80.0, "7": 60.0, "8": 120.0, "11": 180.0, "12": 120.0, "13": 205.0, "14": 95.0, "15": 405.0, "16": 255.0, "18": 330.0, "19": 310.0, "20": 310.0, "21": 430.0, "22": 190.0, "23": 170.0, "25": 295.0, "26": 485.0, "27": 360.0, "28": 85.0, "29": 190.0, "30": 35.0}, "6": {"1": 425.0, "2": 475.0, "3": 330.0, "4": 260.0, "5": 240.0, "6": 415.0, "7": 370.0, "8": 295.0, "9": 495.0, "10": 365.0, "11": 280.0, "16": 240.0, "17": 120.0, "26": 120.0, "29": 295.0, "30": 65.0}}, "1201734XGW02A": {"5": {"4": 100.0, "5": 160.0, "6": 80.0, "7": 60.0, "8": 120.0, "11": 180.0, "12": 120.0, "13": 205.0, "14": 95.0, "15": 405.0, "16": 255.0, "18": 330.0, "19": 310.0, "20": 310.0, "21": 430.0, "22": 190.0, "23": 170.0, "25": 295.0, "26": 485.0, "27": 360.0, "28": 85.0, "29": 190.0, "30": 35.0}, "6": {"1": 425.0, "2": 475.0, "3": 330.0, "4": 260.0, "5": 240.0, "6": 415.0, "7": 370.0, "8": 295.0, "9": 495.0, "10": 365.0, "11": 280.0, "16": 240.0, "17": 120.0, "26": 120.0, "29": 295.0, "30": 65.0}}, "1201735XGW02A": {"5": {"4": 100.0, "5": 160.0, "6": 80.0, "7": 60.0, "8": 120.0, "11": 180.0, "12": 120.0, "13": 205.0, "14": 95.0, "15": 405.0, "16": 255.0, "18": 330.0, "19": 310.0, "20": 310.0, "21": 430.0, "22": 190.0, "23": 170.0, "25": 295.0, "26": 485.0, "27": 360.0, "28": 85.0, "29": 190.0, "30": 35.0}, "6": {"1": 425.0, "2": 475.0, "3": 330.0, "4": 260.0, "5": 240.0, "6": 415.0, "7": 370.0, "8": 295.0, "9": 495.0, "10": 365.0, "11": 280.0, "16": 240.0, "17": 120.0, "26": 120.0, "29": 295.0, "30": 65.0}}, "1201UQ5XGW01A": {"5": {"4": 330.0, "5": 330.0, "6": 450.0, "7": 210.0, "8": 440.0, "11": 140.0, "12": 280.0, "13": 315.0, "14": 425.0, "15": 235.0, "16": 65.0, "18": 310.0, "19": 330.0, "20": 330.0, "21": 210.0, "22": 450.0, "23": 150.0, "25": 345.0, "26": 155.0, "27": 280.0, "28": 555.0, "29": 450.0, "30": 285.0}, "6": {"1": 765.0, "2": 470.0, "3": 325.0, "4": 365.0, "5": 280.0, "6": 400.0, "7": 5.0, "8": 765.0, "9": 650.0, "10": 610.0, "11": 330.0, "13": 110.0, "14": 10.0, "15": 545.0, "16": 735.0, "17": 150.0, "18": 575.0, "19": 555.0, "20": 140.0, "22": 535.0, "23": 65.0, "24": 245.0, "26": 125.0, "27": 165.0, "29": 130.0, "30": 505.0}}, "1201UQ6XGW01A": {"5": {"4": 330.0, "5": 330.0, "6": 450.0, "7": 210.0, "8": 440.0, "11": 140.0, "12": 280.0, "13": 315.0, "14": 425.0, "15": 235.0, "16": 65.0, "18": 310.0, "19": 330.0, "20": 330.0, "21": 210.0, "22": 450.0, "23": 150.0, "25": 345.0, "26": 155.0, "27": 280.0, "28": 555.0, "29": 450.0, "30": 285.0}, "6": {"1": 765.0, "2": 470.0, "3": 325.0, "4": 365.0, "5": 280.0, "6": 400.0, "7": 5.0, "8": 765.0, "9": 650.0, "10": 610.0, "11": 330.0, "13": 110.0, "14": 10.0, "15": 545.0, "16": 735.0, "17": 150.0, "18": 575.0, "19": 555.0, "20": 140.0, "22": 535.0, "23": 65.0, "24": 245.0, "26": 125.0, "27": 165.0, "29": 130.0, "30": 505.0}}, "1201UQ7XGW01A": {"5": {"4": 330.0, "5": 330.0, "6": 450.0, "7": 210.0, "8": 440.0, "11": 140.0, "12": 280.0, "13": 315.0, "14": 425.0, "15": 235.0, "16": 65.0, "18": 310.0, "19": 330.0, "20": 330.0, "21": 210.0, "22": 450.0, "23": 150.0, "25": 345.0, "26": 155.0, "27": 280.0, "28": 555.0, "29": 450.0, "30": 285.0}, "6": {"1": 765.0, "2": 470.0, "3": 325.0, "4": 365.0, "5": 280.0, "6": 400.0, "7": 5.0, "8": 765.0, "9": 650.0, "10": 610.0, "11": 330.0, "13": 110.0, "14": 10.0, "15": 545.0, "16": 735.0, "17": 150.0, "18": 575.0, "19": 555.0, "20": 140.0, "22": 535.0, "23": 65.0, "24": 245.0, "26": 125.0, "27": 165.0, "29": 130.0, "30": 505.0}}, "1201UQ8XGW01A": {"5": {"4": 330.0, "5": 330.0, "6": 450.0, "7": 210.0, "8": 440.0, "11": 140.0, "12": 280.0, "13": 315.0, "14": 425.0, "15": 235.0, "16": 65.0, "18": 310.0, "19": 330.0, "20": 330.0, "21": 210.0, "22": 450.0, "23": 150.0, "25": 345.0, "26": 155.0, "27": 280.0, "28": 555.0, "29": 450.0, "30": 285.0}, "6": {"1": 765.0, "2": 470.0, "3": 325.0, "4": 365.0, "5": 280.0, "6": 400.0, "7": 5.0, "8": 765.0, "9": 650.0, "10": 610.0, "11": 330.0, "13": 110.0, "14": 10.0, "15": 545.0, "16": 735.0, "17": 150.0, "18": 575.0, "19": 555.0, "20": 140.0, "22": 535.0, "23": 65.0, "24": 245.0, "26": 125.0, "27": 165.0, "29": 130.0, "30": 505.0}}, "1201UQ9XGW01A": {"5": {"4": 330.0, "5": 330.0, "6": 450.0, "7": 210.0, "8": 440.0, "11": 140.0, "12": 280.0, "13": 315.0, "14": 425.0, "15": 235.0, "16": 65.0, "18": 310.0, "19": 330.0, "20": 330.0, "21": 210.0, "22": 450.0, "23": 150.0, "25": 345.0, "26": 155.0, "27": 280.0, "28": 555.0, "29": 450.0, "30": 285.0}, "6": {"1": 765.0, "2": 470.0, "3": 325.0, "4": 365.0, "5": 280.0, "6": 400.0, "7": 5.0, "8": 765.0, "9": 650.0, "10": 610.0, "11": 330.0, "13": 110.0, "14": 10.0, "15": 545.0, "16": 735.0, "17": 150.0, "18": 575.0, "19": 555.0, "20": 140.0, "22": 535.0, "23": 65.0, "24": 245.0, "26": 125.0, "27": 165.0, "29": 130.0, "30": 505.0}}, "1201UR1XGW01A": {"5": {"4": 330.0, "5": 330.0, "6": 450.0, "7": 210.0, "8": 440.0, "11": 140.0, "12": 280.0, "13": 315.0, "14": 425.0, "15": 235.0, "16": 65.0, "18": 310.0, "19": 330.0, "20": 330.0, "21": 210.0, "22": 450.0, "23": 150.0, "25": 345.0, "26": 155.0, "27": 280.0, "28": 555.0, "29": 450.0, "30": 285.0}, "6": {"1": 765.0, "2": 470.0, "3": 325.0, "4": 365.0, "5": 280.0, "6": 400.0, "7": 5.0, "8": 765.0, "9": 650.0, "10": 610.0, "11": 330.0, "13": 110.0, "14": 10.0, "15": 545.0, "16": 735.0, "17": 150.0, "18": 575.0, "19": 555.0, "20": 140.0, "22": 535.0, "23": 65.0, "24": 245.0, "26": 125.0, "27": 165.0, "29": 130.0, "30": 505.0}}, "1201UR2XGW01A": {"5": {"4": 330.0, "5": 330.0, "6": 450.0, "7": 210.0, "8": 440.0, "11": 140.0, "12": 280.0, "13": 315.0, "14": 425.0, "15": 235.0, "16": 65.0, "18": 310.0, "19": 330.0, "20": 330.0, "21": 210.0, "22": 450.0, "23": 150.0, "25": 345.0, "26": 155.0, "27": 280.0, "28": 555.0, "29": 450.0, "30": 285.0}, "6": {"1": 765.0, "2": 470.0, "3": 325.0, "4": 365.0, "5": 280.0, "6": 400.0, "7": 5.0, "8": 765.0, "9": 650.0, "10": 610.0, "11": 330.0, "13": 110.0, "14": 10.0, "15": 545.0, "16": 735.0, "17": 150.0, "18": 575.0, "19": 555.0, "20": 140.0, "22": 535.0, "23": 65.0, "24": 245.0, "26": 125.0, "27": 165.0, "29": 130.0, "30": 505.0}}, "1201UR3XGW01A": {"5": {"4": 330.0, "5": 330.0, "6": 450.0, "7": 210.0, "8": 440.0, "11": 140.0, "12": 280.0, "13": 315.0, "14": 425.0, "15": 235.0, "16": 65.0, "18": 310.0, "19": 330.0, "20": 330.0, "21": 210.0, "22": 450.0, "23": 150.0, "25": 345.0, "26": 155.0, "27": 280.0, "28": 555.0, "29": 450.0, "30": 285.0}, "6": {"1": 765.0, "2": 470.0, "3": 325.0, "4": 365.0, "5": 280.0, "6": 400.0, "7": 5.0, "8": 765.0, "9": 650.0, "10": 610.0, "11": 330.0, "13": 110.0, "14": 10.0, "15": 545.0, "16": 735.0, "17": 150.0, "18": 575.0, "19": 555.0, "20": 140.0, "22": 535.0, "23": 65.0, "24": 245.0, "26": 125.0, "27": 165.0, "29": 130.0, "30": 505.0}}, "1201UR4XGW01A": {"5": {"4": 330.0, "5": 330.0, "6": 450.0, "7": 210.0, "8": 440.0, "11": 140.0, "12": 280.0, "13": 315.0, "14": 425.0, "15": 235.0, "16": 65.0, "18": 310.0, "19": 330.0, "20": 330.0, "21": 210.0, "22": 450.0, "23": 150.0, "25": 345.0, "26": 155.0, "27": 280.0, "28": 555.0, "29": 450.0, "30": 285.0}, "6": {"1": 765.0, "2": 470.0, "3": 325.0, "4": 365.0, "5": 280.0, "6": 400.0, "7": 5.0, "8": 765.0, "9": 650.0, "10": 610.0, "11": 330.0, "13": 110.0, "14": 10.0, "15": 545.0, "16": 735.0, "17": 150.0, "18": 575.0, "19": 555.0, "20": 140.0, "22": 535.0, "23": 65.0, "24": 245.0, "26": 125.0, "27": 165.0, "29": 130.0, "30": 505.0}}, "1201UR5XGW01A": {"5": {"4": 330.0, "5": 330.0, "6": 450.0, "7": 210.0, "8": 440.0, "11": 140.0, "12": 280.0, "13": 315.0, "14": 425.0, "15": 235.0, "16": 65.0, "18": 310.0, "19": 330.0, "20": 330.0, "21": 210.0, "22": 450.0, "23": 150.0, "25": 345.0, "26": 155.0, "27": 280.0, "28": 555.0, "29": 450.0, "30": 285.0}, "6": {"1": 765.0, "2": 470.0, "3": 325.0, "4": 365.0, "5": 280.0, "6": 400.0, "7": 5.0, "8": 765.0, "9": 650.0, "10": 610.0, "11": 330.0, "13": 110.0, "14": 10.0, "15": 545.0, "16": 735.0, "17": 150.0, "18": 575.0, "19": 555.0, "20": 140.0, "22": 535.0, "23": 65.0, "24": 245.0, "26": 125.0, "27": 165.0, "29": 130.0, "30": 505.0}}, "1201UR6XGW01A": {"5": {"4": 330.0, "5": 330.0, "6": 450.0, "7": 210.0, "8": 440.0, "11": 140.0, "12": 280.0, "13": 315.0, "14": 425.0, "15": 235.0, "16": 65.0, "18": 310.0, "19": 330.0, "20": 330.0, "21": 210.0, "22": 450.0, "23": 150.0, "25": 345.0, "26": 155.0, "27": 280.0, "28": 555.0, "29": 450.0, "30": 285.0}, "6": {"1": 765.0, "2": 470.0, "3": 325.0, "4": 365.0, "5": 280.0, "6": 400.0, "7": 5.0, "8": 765.0, "9": 650.0, "10": 610.0, "11": 330.0, "13": 110.0, "14": 10.0, "15": 545.0, "16": 735.0, "17": 150.0, "18": 575.0, "19": 555.0, "20": 140.0, "22": 535.0, "23": 65.0, "24": 245.0, "26": 125.0, "27": 165.0, "29": 130.0, "30": 505.0}}, "1201UR7XGW01A": {"5": {"4": 330.0, "5": 330.0, "6": 450.0, "7": 210.0, "8": 440.0, "11": 140.0, "12": 280.0, "13": 315.0, "14": 425.0, "15": 235.0, "16": 65.0, "18": 310.0, "19": 330.0, "20": 330.0, "21": 210.0, "22": 450.0, "23": 150.0, "25": 345.0, "26": 155.0, "27": 280.0, "28": 555.0, "29": 450.0, "30": 285.0}, "6": {"1": 765.0, "2": 470.0, "3": 325.0, "4": 365.0, "5": 280.0, "6": 400.0, "7": 5.0, "8": 765.0, "9": 650.0, "10": 610.0, "11": 330.0, "13": 110.0, "14": 10.0, "15": 545.0, "16": 735.0, "17": 150.0, "18": 575.0, "19": 555.0, "20": 140.0, "22": 535.0, "23": 65.0, "24": 245.0, "26": 125.0, "27": 165.0, "29": 130.0, "30": 505.0}}, "1205102XST1AA": {"5": {"4": 330.0, "5": 210.0, "6": 330.0, "7": 210.0, "8": 320.0, "11": 140.0, "12": 280.0, "13": 195.0, "14": 305.0, "15": 235.0, "16": 65.0, "18": 310.0, "19": 330.0, "20": 210.0, "21": 210.0, "22": 210.0, "23": 45.0, "25": 210.0, "26": 90.0, "27": 225.0, "28": 315.0, "29": 210.0, "30": 165.0}, "6": {"1": 550.0, "2": 230.0, "3": 90.0, "4": 210.0, "5": 280.0, "6": 310.0, "7": 5.0, "8": 285.0, "9": 410.0, "10": 370.0, "11": 195.0, "13": 110.0, "14": 10.0, "15": 200.0, "16": 540.0, "18": 80.0, "19": 210.0, "20": 45.0, "22": 135.0, "23": 65.0, "24": 125.0, "26": 5.0, "27": 45.0, "29": 130.0, "30": 120.0}}, "1205105XST1AA": {"5": {"4": 310.0, "5": 430.0, "6": 310.0, "7": 430.0, "8": 240.0, "11": 300.0, "12": 360.0, "13": 445.0, "14": 335.0, "15": 405.0, "16": 255.0, "18": 330.0, "19": 310.0, "20": 430.0, "21": 430.0, "22": 430.0, "23": 275.0, "25": 430.0, "26": 550.0, "27": 415.0, "28": 325.0, "29": 430.0, "30": 155.0}, "6": {"1": 640.0, "2": 715.0, "3": 685.0, "4": 485.0, "5": 290.0, "6": 505.0, "7": 370.0, "8": 775.0, "9": 735.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "1205112XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "1205113XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "1205526XGW01B": {"5": {"4": 310.0, "5": 430.0, "6": 310.0, "7": 430.0, "8": 240.0, "11": 300.0, "12": 360.0, "13": 445.0, "14": 335.0, "15": 405.0, "16": 255.0, "18": 330.0, "19": 310.0, "20": 430.0, "21": 430.0, "22": 430.0, "23": 275.0, "25": 430.0, "26": 550.0, "27": 415.0, "28": 325.0, "29": 430.0, "30": 155.0}, "6": {"1": 640.0, "2": 715.0, "3": 685.0, "4": 485.0, "5": 290.0, "6": 505.0, "7": 370.0, "8": 775.0, "9": 735.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "1205527XGW01A": {"5": {"4": 310.0, "5": 430.0, "6": 310.0, "7": 430.0, "8": 240.0, "11": 300.0, "12": 360.0, "13": 445.0, "14": 335.0, "15": 405.0, "16": 255.0, "18": 330.0, "19": 310.0, "20": 430.0, "21": 430.0, "22": 430.0, "23": 275.0, "25": 430.0, "26": 550.0, "27": 415.0, "28": 325.0, "29": 430.0, "30": 155.0}, "6": {"1": 640.0, "2": 715.0, "3": 685.0, "4": 485.0, "5": 290.0, "6": 505.0, "7": 370.0, "8": 775.0, "9": 735.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "1205653XKJ22A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 120.0, "12": 240.0, "13": 240.0, "14": 240.0, "20": 120.0, "22": 240.0, "23": 105.0, "25": 135.0, "26": 65.0, "27": 55.0, "28": 240.0, "29": 240.0, "30": 120.0}, "6": {"1": 215.0, "2": 240.0, "3": 355.0, "4": 225.0, "5": 50.0, "6": 90.0, "8": 480.0, "9": 240.0, "10": 240.0, "11": 135.0, "15": 345.0, "16": 195.0, "17": 270.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 240.0, "27": 120.0, "30": 385.0}}, "1205654XKJ22A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 120.0, "12": 240.0, "13": 240.0, "14": 240.0, "20": 120.0, "22": 240.0, "23": 105.0, "25": 135.0, "26": 65.0, "27": 55.0, "28": 240.0, "29": 240.0, "30": 120.0}, "6": {"1": 215.0, "2": 240.0, "3": 355.0, "4": 225.0, "5": 50.0, "6": 90.0, "8": 480.0, "9": 240.0, "10": 240.0, "11": 135.0, "15": 345.0, "16": 195.0, "17": 270.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 240.0, "27": 120.0, "30": 385.0}}, "1205655XKJ22A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 120.0, "12": 240.0, "13": 240.0, "14": 240.0, "20": 120.0, "22": 240.0, "23": 105.0, "25": 135.0, "26": 65.0, "27": 55.0, "28": 240.0, "29": 240.0, "30": 120.0}, "6": {"1": 215.0, "2": 240.0, "3": 355.0, "4": 225.0, "5": 50.0, "6": 90.0, "8": 480.0, "9": 240.0, "10": 240.0, "11": 135.0, "15": 345.0, "16": 195.0, "17": 270.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 240.0, "27": 120.0, "30": 385.0}}, "1205656XKJ22A": {"5": {"11": 120.0, "13": 120.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 140.0, "21": 360.0, "22": 120.0, "23": 120.0, "25": 235.0, "26": 365.0, "27": 240.0, "28": 25.0, "29": 95.0, "30": 10.0}, "6": {"1": 325.0, "2": 395.0, "3": 270.0, "4": 200.0, "5": 120.0, "6": 415.0, "7": 370.0, "8": 175.0, "9": 435.0, "10": 365.0, "11": 280.0, "16": 240.0, "17": 120.0, "26": 120.0, "29": 295.0, "30": 65.0}}, "1205657XKJ22A": {"5": {"11": 120.0, "13": 120.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 140.0, "21": 360.0, "22": 120.0, "23": 120.0, "25": 235.0, "26": 365.0, "27": 240.0, "28": 25.0, "29": 95.0, "30": 10.0}, "6": {"1": 325.0, "2": 395.0, "3": 270.0, "4": 200.0, "5": 120.0, "6": 415.0, "7": 370.0, "8": 175.0, "9": 435.0, "10": 365.0, "11": 280.0, "16": 240.0, "17": 120.0, "26": 120.0, "29": 295.0, "30": 65.0}}, "1205658XKJ22A": {"5": {"5": 120.0, "6": 60.0, "7": 60.0, "11": 120.0, "12": 60.0, "13": 205.0, "14": 95.0, "15": 345.0, "16": 195.0, "18": 300.0, "19": 280.0, "20": 310.0, "21": 370.0, "22": 190.0, "23": 170.0, "25": 295.0, "26": 425.0, "27": 360.0, "28": 85.0, "29": 190.0, "30": 35.0}, "6": {"1": 425.0, "2": 475.0, "3": 330.0, "4": 260.0, "5": 180.0, "6": 415.0, "7": 370.0, "8": 295.0, "9": 435.0, "10": 365.0, "11": 280.0, "16": 240.0, "17": 120.0, "26": 120.0, "29": 295.0, "30": 65.0}}, "1205659XKJ22A": {"5": {"5": 120.0, "6": 60.0, "7": 60.0, "11": 120.0, "12": 60.0, "13": 205.0, "14": 95.0, "15": 345.0, "16": 195.0, "18": 300.0, "19": 280.0, "20": 310.0, "21": 370.0, "22": 190.0, "23": 170.0, "25": 295.0, "26": 425.0, "27": 360.0, "28": 85.0, "29": 190.0, "30": 35.0}, "6": {"1": 425.0, "2": 475.0, "3": 330.0, "4": 260.0, "5": 180.0, "6": 415.0, "7": 370.0, "8": 295.0, "9": 435.0, "10": 365.0, "11": 280.0, "16": 240.0, "17": 120.0, "26": 120.0, "29": 295.0, "30": 65.0}}, "1205660XKJ22A": {"5": {"11": 120.0, "13": 120.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 140.0, "21": 360.0, "22": 120.0, "23": 120.0, "25": 235.0, "26": 365.0, "27": 240.0, "28": 25.0, "29": 95.0, "30": 10.0}, "6": {"1": 325.0, "2": 395.0, "3": 270.0, "4": 200.0, "5": 120.0, "6": 415.0, "7": 370.0, "8": 175.0, "9": 435.0, "10": 365.0, "11": 280.0, "16": 240.0, "17": 120.0, "26": 120.0, "29": 295.0, "30": 65.0}}, "2803104XKN61A8T": {"5": {"4": 30.0, "5": 60.0, "6": 30.0, "8": 30.0, "11": 30.0, "12": 49.0, "13": 49.0, "14": 24.0, "15": 57.0, "16": 24.0, "18": 10.0, "19": 30.0, "20": 60.0, "22": 30.0, "25": 30.0, "26": 30.0, "28": 30.0, "30": 3.0}, "6": {"1": 88.0, "2": 63.0, "3": 58.0, "4": 83.0, "5": 30.0, "6": 38.0, "8": 90.0, "9": 49.0, "10": 12.0, "11": 49.0, "13": 17.0, "14": 14.0, "15": 12.0, "16": 60.0, "17": 30.0, "18": 30.0, "23": 30.0, "24": 13.0, "25": 48.0, "26": 28.0, "27": 3.0, "29": 30.0, "30": 60.0}}, "2803104XKN61A9C": {"5": {"4": 30.0, "5": 60.0, "6": 30.0, "8": 30.0, "11": 30.0, "12": 49.0, "13": 49.0, "14": 24.0, "15": 57.0, "16": 24.0, "18": 10.0, "19": 30.0, "20": 60.0, "22": 30.0, "25": 30.0, "26": 30.0, "28": 30.0, "30": 3.0}, "6": {"1": 88.0, "2": 63.0, "3": 58.0, "4": 83.0, "5": 30.0, "6": 38.0, "8": 90.0, "9": 49.0, "10": 12.0, "11": 49.0, "13": 17.0, "14": 14.0, "15": 12.0, "16": 60.0, "17": 30.0, "18": 30.0, "23": 30.0, "24": 13.0, "25": 48.0, "26": 28.0, "27": 3.0, "29": 30.0, "30": 60.0}}, "2803104XKN61AC3": {"5": {"4": 36.0, "5": 72.0, "6": 36.0, "8": 36.0, "11": 36.0, "12": 59.0, "13": 58.0, "14": 29.0, "15": 68.0, "16": 29.0, "18": 12.0, "19": 36.0, "20": 72.0, "22": 36.0, "25": 36.0, "26": 36.0, "28": 36.0, "30": 3.0}, "6": {"1": 105.0, "2": 75.0, "3": 69.0, "4": 99.0, "5": 36.0, "6": 45.0, "8": 108.0, "9": 59.0, "10": 14.0, "11": 59.0, "13": 20.0, "14": 17.0, "15": 14.0, "16": 72.0, "17": 36.0, "18": 36.0, "23": 36.0, "24": 15.0, "25": 57.0, "26": 33.0, "27": 3.0, "29": 36.0, "30": 72.0}}, "2803104XKN61AH4": {"5": {"4": 24.0, "5": 48.0, "6": 24.0, "8": 24.0, "11": 24.0, "12": 39.0, "13": 38.0, "14": 19.0, "15": 45.0, "16": 19.0, "18": 8.0, "19": 24.0, "20": 48.0, "22": 24.0, "25": 24.0, "26": 24.0, "28": 24.0, "30": 2.0}, "6": {"1": 70.0, "2": 50.0, "3": 46.0, "4": 66.0, "5": 24.0, "6": 30.0, "8": 72.0, "9": 39.0, "10": 9.0, "11": 39.0, "13": 13.0, "14": 11.0, "15": 9.0, "16": 48.0, "17": 24.0, "18": 24.0, "23": 24.0, "24": 10.0, "25": 38.0, "26": 22.0, "27": 2.0, "29": 24.0, "30": 48.0}}, "2803105XKN61A8T": {"5": {"4": 30.0, "5": 60.0, "6": 30.0, "7": 84.0, "8": 30.0, "11": 30.0, "12": 49.0, "13": 49.0, "14": 24.0, "15": 57.0, "16": 24.0, "18": 21.0, "19": 62.0, "20": 60.0, "21": 42.0, "22": 30.0, "25": 30.0, "26": 30.0, "28": 39.0, "29": 76.0, "30": 3.0}, "6": {"1": 88.0, "2": 105.0, "3": 100.0, "4": 83.0, "5": 30.0, "6": 80.0, "7": 42.0, "8": 90.0, "9": 79.0, "10": 85.0, "11": 65.0, "13": 47.0, "14": 14.0, "15": 54.0, "16": 60.0, "17": 72.0, "18": 30.0, "22": 42.0, "23": 30.0, "24": 13.0, "25": 48.0, "26": 28.0, "27": 3.0, "29": 30.0, "30": 60.0}}, "2803105XKN61A9C": {"5": {"4": 30.0, "5": 60.0, "6": 30.0, "7": 60.0, "8": 30.0, "11": 30.0, "12": 49.0, "13": 49.0, "14": 24.0, "15": 57.0, "16": 24.0, "18": 18.0, "19": 53.0, "20": 60.0, "21": 30.0, "22": 30.0, "25": 30.0, "26": 30.0, "28": 37.0, "29": 54.0, "30": 3.0}, "6": {"1": 88.0, "2": 93.0, "3": 88.0, "4": 83.0, "5": 30.0, "6": 68.0, "7": 30.0, "8": 90.0, "9": 79.0, "10": 62.0, "11": 59.0, "13": 47.0, "14": 14.0, "15": 42.0, "16": 60.0, "17": 60.0, "18": 30.0, "22": 30.0, "23": 30.0, "24": 13.0, "25": 48.0, "26": 28.0, "27": 3.0, "29": 30.0, "30": 60.0}}, "2803105XKN61AC3": {"5": {"4": 36.0, "5": 72.0, "6": 36.0, "7": 60.0, "8": 36.0, "11": 36.0, "12": 59.0, "13": 58.0, "14": 29.0, "15": 68.0, "16": 29.0, "18": 20.0, "19": 59.0, "20": 72.0, "21": 30.0, "22": 36.0, "25": 36.0, "26": 36.0, "28": 43.0, "29": 54.0, "30": 3.0}, "6": {"1": 105.0, "2": 105.0, "3": 99.0, "4": 99.0, "5": 36.0, "6": 75.0, "7": 30.0, "8": 108.0, "9": 89.0, "10": 64.0, "11": 69.0, "13": 50.0, "14": 17.0, "15": 44.0, "16": 72.0, "17": 66.0, "18": 36.0, "22": 30.0, "23": 36.0, "24": 15.0, "25": 57.0, "26": 33.0, "27": 3.0, "29": 36.0, "30": 72.0}}, "2803105XKN61AH4": {"5": {"4": 24.0, "5": 48.0, "6": 24.0, "7": 36.0, "8": 24.0, "11": 24.0, "12": 39.0, "13": 38.0, "14": 19.0, "15": 45.0, "16": 19.0, "18": 13.0, "19": 38.0, "20": 48.0, "21": 18.0, "22": 24.0, "25": 24.0, "26": 24.0, "28": 28.0, "29": 33.0, "30": 2.0}, "6": {"1": 70.0, "2": 68.0, "3": 64.0, "4": 66.0, "5": 24.0, "6": 48.0, "7": 18.0, "8": 72.0, "9": 69.0, "10": 39.0, "11": 45.0, "13": 43.0, "14": 11.0, "15": 27.0, "16": 48.0, "17": 42.0, "18": 24.0, "22": 18.0, "23": 24.0, "24": 10.0, "25": 38.0, "26": 22.0, "27": 2.0, "29": 24.0, "30": 48.0}}, "2803120XST33A5B": {"5": {"13": 120.0, "14": 60.0, "18": 60.0, "29": 35.0, "30": 85.0}, "6": {"3": 20.0, "4": 40.0, "7": 55.0, "8": 5.0, "9": 60.0, "17": 240.0, "27": 60.0}}, "2803120XST33A8T": {"5": {"4": 60.0, "5": 120.0, "6": 240.0, "7": 120.0, "8": 180.0, "20": 110.0, "21": 10.0, "22": 240.0, "23": 105.0, "25": 15.0, "27": 75.0, "28": 45.0, "29": 60.0, "30": 120.0}, "6": {"1": 190.0, "2": 190.0, "3": 15.0, "4": 190.0, "5": 230.0, "6": 75.0, "8": 105.0, "9": 480.0, "10": 173.0, "11": 186.0, "13": 123.0, "14": 360.0, "15": 325.0, "16": 95.0, "18": 135.0, "19": 345.0, "23": 120.0, "24": 240.0, "25": 300.0, "30": 145.0}}, "2803120XST33A9C": {"5": {"7": 120.0, "8": 120.0, "11": 180.0, "12": 120.0, "16": 120.0, "19": 120.0, "20": 120.0, "21": 120.0, "26": 240.0, "29": 120.0}, "6": {"1": 240.0, "3": 600.0, "4": 120.0, "6": 160.0, "7": 200.0, "10": 240.0, "11": 298.0, "13": 123.0, "15": 180.0, "16": 290.0, "17": 170.0, "19": 120.0, "20": 240.0, "22": 120.0, "25": 80.0, "26": 340.0, "29": 120.0}}, "2803120XST33AC3": {"5": {"4": 180.0, "5": 100.0, "6": 20.0, "12": 120.0, "14": 240.0, "15": 240.0, "18": 240.0, "19": 120.0, "21": 120.0, "22": 120.0, "23": 120.0, "25": 235.0, "26": 125.0, "27": 360.0, "28": 240.0, "29": 60.0}, "6": {"1": 280.0, "2": 320.0, "5": 160.0, "6": 140.0, "8": 360.0, "9": 10.0, "10": 283.0, "11": 8.0, "15": 225.0, "16": 195.0, "18": 240.0, "22": 230.0, "23": 270.0, "24": 220.0, "27": 85.0, "29": 155.0, "30": 180.0}}, "2803120XST33AGN": {"5": {"13": 120.0, "14": 60.0, "18": 60.0, "25": 120.0}, "6": {"3": 20.0, "4": 40.0, "7": 55.0, "8": 170.0, "9": 75.0, "17": 105.0, "18": 15.0, "27": 60.0}}, "2803130XST33A5B": {"5": {"13": 120.0, "14": 60.0, "18": 60.0, "29": 35.0, "30": 85.0}, "6": {"3": 20.0, "4": 40.0, "7": 55.0, "8": 5.0, "9": 60.0, "17": 240.0, "27": 60.0}}, "2803130XST33A8T": {"5": {"4": 60.0, "5": 120.0, "6": 240.0, "7": 120.0, "8": 180.0, "20": 110.0, "21": 10.0, "22": 240.0, "23": 105.0, "25": 15.0, "27": 75.0, "28": 45.0, "29": 60.0, "30": 120.0}, "6": {"1": 190.0, "2": 190.0, "3": 15.0, "4": 190.0, "5": 230.0, "6": 75.0, "8": 105.0, "9": 480.0, "10": 173.0, "11": 186.0, "13": 123.0, "14": 360.0, "15": 325.0, "16": 95.0, "18": 135.0, "19": 345.0, "23": 120.0, "24": 240.0, "25": 300.0, "30": 145.0}}, "2803130XST33A9C": {"5": {"7": 120.0, "8": 120.0, "11": 180.0, "12": 120.0, "16": 120.0, "19": 120.0, "20": 120.0, "21": 120.0, "26": 240.0, "29": 120.0}, "6": {"1": 240.0, "3": 600.0, "4": 120.0, "6": 160.0, "7": 200.0, "10": 240.0, "11": 298.0, "13": 123.0, "15": 180.0, "16": 290.0, "17": 170.0, "19": 120.0, "20": 240.0, "22": 120.0, "25": 80.0, "26": 340.0, "29": 120.0}}, "2803130XST33AC3": {"5": {"4": 180.0, "5": 100.0, "6": 20.0, "12": 120.0, "14": 240.0, "15": 240.0, "18": 240.0, "19": 120.0, "21": 120.0, "22": 120.0, "23": 120.0, "25": 235.0, "26": 125.0, "27": 360.0, "28": 240.0, "29": 60.0}, "6": {"1": 280.0, "2": 320.0, "5": 160.0, "6": 140.0, "8": 360.0, "9": 10.0, "10": 283.0, "11": 8.0, "15": 225.0, "16": 195.0, "18": 240.0, "22": 230.0, "23": 270.0, "24": 220.0, "27": 85.0, "29": 155.0, "30": 180.0}}, "2803130XST33AGN": {"5": {"13": 120.0, "14": 60.0, "18": 60.0, "25": 120.0}, "6": {"3": 20.0, "4": 40.0, "7": 55.0, "8": 170.0, "9": 75.0, "17": 105.0, "18": 15.0, "27": 60.0}}, "2804104AST33A5B": {"5": {"13": 120.0, "14": 60.0, "18": 60.0, "29": 35.0, "30": 85.0}, "6": {"3": 20.0, "4": 40.0, "7": 55.0, "8": 5.0, "9": 60.0, "17": 240.0, "27": 60.0}}, "2804104AST33A8T": {"5": {"4": 60.0, "5": 120.0, "6": 240.0, "7": 120.0, "8": 180.0, "20": 110.0, "21": 10.0, "22": 240.0, "23": 105.0, "25": 15.0, "27": 75.0, "28": 45.0, "29": 60.0, "30": 120.0}, "6": {"1": 190.0, "2": 190.0, "3": 15.0, "4": 190.0, "5": 230.0, "6": 75.0, "8": 105.0, "9": 480.0, "10": 173.0, "11": 186.0, "13": 123.0, "14": 360.0, "15": 325.0, "16": 95.0, "18": 135.0, "19": 345.0, "23": 120.0, "24": 240.0, "25": 300.0, "30": 145.0}}, "2804104AST33A9C": {"5": {"7": 120.0, "8": 120.0, "11": 180.0, "12": 120.0, "16": 120.0, "19": 120.0, "20": 120.0, "21": 120.0, "26": 240.0, "29": 120.0}, "6": {"1": 240.0, "3": 600.0, "4": 120.0, "6": 160.0, "7": 200.0, "10": 240.0, "11": 298.0, "13": 123.0, "15": 180.0, "16": 290.0, "17": 170.0, "19": 120.0, "20": 240.0, "22": 120.0, "25": 80.0, "26": 340.0, "29": 120.0}}, "2804104AST33AC3": {"5": {"4": 180.0, "5": 100.0, "6": 20.0, "12": 120.0, "14": 240.0, "15": 240.0, "18": 240.0, "19": 120.0, "21": 120.0, "22": 120.0, "23": 120.0, "25": 235.0, "26": 125.0, "27": 360.0, "28": 240.0, "29": 60.0}, "6": {"1": 280.0, "2": 320.0, "5": 160.0, "6": 140.0, "8": 360.0, "9": 10.0, "10": 283.0, "11": 8.0, "15": 225.0, "16": 195.0, "18": 240.0, "22": 230.0, "23": 270.0, "24": 220.0, "27": 85.0, "29": 155.0, "30": 180.0}}, "2804104AST33AGN": {"5": {"13": 120.0, "14": 60.0, "18": 60.0, "25": 120.0}, "6": {"3": 20.0, "4": 40.0, "7": 55.0, "8": 170.0, "9": 75.0, "17": 105.0, "18": 15.0, "27": 60.0}}, "2804105AST33A5B": {"5": {"13": 120.0, "14": 60.0, "18": 60.0, "29": 35.0, "30": 85.0}, "6": {"3": 20.0, "4": 40.0, "7": 55.0, "8": 5.0, "9": 60.0, "17": 240.0, "27": 60.0}}, "2804105AST33A8T": {"5": {"4": 60.0, "5": 120.0, "6": 240.0, "7": 120.0, "8": 180.0, "20": 110.0, "21": 10.0, "22": 240.0, "23": 105.0, "25": 15.0, "27": 75.0, "28": 45.0, "29": 60.0, "30": 120.0}, "6": {"1": 190.0, "2": 190.0, "3": 15.0, "4": 190.0, "5": 230.0, "6": 75.0, "8": 105.0, "9": 480.0, "10": 173.0, "11": 186.0, "13": 123.0, "14": 360.0, "15": 325.0, "16": 95.0, "18": 135.0, "19": 345.0, "23": 120.0, "24": 240.0, "25": 300.0, "30": 145.0}}, "2804105AST33A9C": {"5": {"7": 120.0, "8": 120.0, "11": 180.0, "12": 120.0, "16": 120.0, "19": 120.0, "20": 120.0, "21": 120.0, "26": 240.0, "29": 120.0}, "6": {"1": 240.0, "3": 600.0, "4": 120.0, "6": 160.0, "7": 200.0, "10": 240.0, "11": 298.0, "13": 123.0, "15": 180.0, "16": 290.0, "17": 170.0, "19": 120.0, "20": 240.0, "22": 120.0, "25": 80.0, "26": 340.0, "29": 120.0}}, "2804105AST33AC3": {"5": {"4": 180.0, "5": 100.0, "6": 20.0, "12": 120.0, "14": 240.0, "15": 240.0, "18": 240.0, "19": 120.0, "21": 120.0, "22": 120.0, "23": 120.0, "25": 235.0, "26": 125.0, "27": 360.0, "28": 240.0, "29": 60.0}, "6": {"1": 280.0, "2": 320.0, "5": 160.0, "6": 140.0, "8": 360.0, "9": 10.0, "10": 283.0, "11": 8.0, "15": 225.0, "16": 195.0, "18": 240.0, "22": 230.0, "23": 270.0, "24": 220.0, "27": 85.0, "29": 155.0, "30": 180.0}}, "2804105AST33AGN": {"5": {"13": 120.0, "14": 60.0, "18": 60.0, "25": 120.0}, "6": {"3": 20.0, "4": 40.0, "7": 55.0, "8": 170.0, "9": 75.0, "17": 105.0, "18": 15.0, "27": 60.0}}, "2804111XKN61A8T": {"5": {"4": 30.0, "5": 60.0, "6": 30.0, "8": 30.0, "11": 30.0, "12": 49.0, "13": 49.0, "14": 24.0, "15": 57.0, "16": 24.0, "18": 10.0, "19": 30.0, "20": 60.0, "22": 30.0, "25": 30.0, "26": 30.0, "28": 30.0, "30": 3.0}, "6": {"1": 88.0, "2": 63.0, "3": 58.0, "4": 83.0, "5": 30.0, "6": 38.0, "8": 90.0, "9": 49.0, "10": 12.0, "11": 49.0, "13": 17.0, "14": 14.0, "15": 12.0, "16": 60.0, "17": 30.0, "18": 30.0, "23": 30.0, "24": 13.0, "25": 48.0, "26": 28.0, "27": 3.0, "29": 30.0, "30": 60.0}}, "2804111XKN61A9C": {"5": {"4": 30.0, "5": 60.0, "6": 30.0, "8": 30.0, "11": 30.0, "12": 49.0, "13": 49.0, "14": 24.0, "15": 57.0, "16": 24.0, "18": 10.0, "19": 30.0, "20": 60.0, "22": 30.0, "25": 30.0, "26": 30.0, "28": 30.0, "30": 3.0}, "6": {"1": 88.0, "2": 63.0, "3": 58.0, "4": 83.0, "5": 30.0, "6": 38.0, "8": 90.0, "9": 49.0, "10": 12.0, "11": 49.0, "13": 17.0, "14": 14.0, "15": 12.0, "16": 60.0, "17": 30.0, "18": 30.0, "23": 30.0, "24": 13.0, "25": 48.0, "26": 28.0, "27": 3.0, "29": 30.0, "30": 60.0}}, "2804111XKN61AC3": {"5": {"4": 36.0, "5": 72.0, "6": 36.0, "8": 36.0, "11": 36.0, "12": 59.0, "13": 58.0, "14": 29.0, "15": 68.0, "16": 29.0, "18": 12.0, "19": 36.0, "20": 72.0, "22": 36.0, "25": 36.0, "26": 36.0, "28": 36.0, "30": 3.0}, "6": {"1": 105.0, "2": 75.0, "3": 69.0, "4": 99.0, "5": 36.0, "6": 45.0, "8": 108.0, "9": 59.0, "10": 14.0, "11": 59.0, "13": 20.0, "14": 17.0, "15": 14.0, "16": 72.0, "17": 36.0, "18": 36.0, "23": 36.0, "24": 15.0, "25": 57.0, "26": 33.0, "27": 3.0, "29": 36.0, "30": 72.0}}, "2804111XKN61AH4": {"5": {"4": 24.0, "5": 48.0, "6": 24.0, "8": 24.0, "11": 24.0, "12": 39.0, "13": 38.0, "14": 19.0, "15": 45.0, "16": 19.0, "18": 8.0, "19": 24.0, "20": 48.0, "22": 24.0, "25": 24.0, "26": 24.0, "28": 24.0, "30": 2.0}, "6": {"1": 70.0, "2": 50.0, "3": 46.0, "4": 66.0, "5": 24.0, "6": 30.0, "8": 72.0, "9": 39.0, "10": 9.0, "11": 39.0, "13": 13.0, "14": 11.0, "15": 9.0, "16": 48.0, "17": 24.0, "18": 24.0, "23": 24.0, "24": 10.0, "25": 38.0, "26": 22.0, "27": 2.0, "29": 24.0, "30": 48.0}}, "2804112XKN61A8T": {"5": {"4": 30.0, "5": 60.0, "6": 30.0, "7": 84.0, "8": 30.0, "11": 30.0, "12": 49.0, "13": 49.0, "14": 24.0, "15": 57.0, "16": 24.0, "18": 21.0, "19": 62.0, "20": 60.0, "21": 42.0, "22": 30.0, "25": 30.0, "26": 30.0, "28": 39.0, "29": 76.0, "30": 3.0}, "6": {"1": 88.0, "2": 105.0, "3": 100.0, "4": 83.0, "5": 30.0, "6": 80.0, "7": 42.0, "8": 90.0, "9": 79.0, "10": 85.0, "11": 65.0, "13": 47.0, "14": 14.0, "15": 54.0, "16": 60.0, "17": 72.0, "18": 30.0, "22": 42.0, "23": 30.0, "24": 13.0, "25": 48.0, "26": 28.0, "27": 3.0, "29": 30.0, "30": 60.0}}, "2804112XKN61A9C": {"5": {"4": 30.0, "5": 60.0, "6": 30.0, "7": 60.0, "8": 30.0, "11": 30.0, "12": 49.0, "13": 49.0, "14": 24.0, "15": 57.0, "16": 24.0, "18": 18.0, "19": 53.0, "20": 60.0, "21": 30.0, "22": 30.0, "25": 30.0, "26": 30.0, "28": 37.0, "29": 54.0, "30": 3.0}, "6": {"1": 88.0, "2": 93.0, "3": 88.0, "4": 83.0, "5": 30.0, "6": 68.0, "7": 30.0, "8": 90.0, "9": 79.0, "10": 62.0, "11": 59.0, "13": 47.0, "14": 14.0, "15": 42.0, "16": 60.0, "17": 60.0, "18": 30.0, "22": 30.0, "23": 30.0, "24": 13.0, "25": 48.0, "26": 28.0, "27": 3.0, "29": 30.0, "30": 60.0}}, "2804112XKN61AC3": {"5": {"4": 36.0, "5": 72.0, "6": 36.0, "7": 60.0, "8": 36.0, "11": 36.0, "12": 59.0, "13": 58.0, "14": 29.0, "15": 68.0, "16": 29.0, "18": 20.0, "19": 59.0, "20": 72.0, "21": 30.0, "22": 36.0, "25": 36.0, "26": 36.0, "28": 43.0, "29": 54.0, "30": 3.0}, "6": {"1": 105.0, "2": 105.0, "3": 99.0, "4": 99.0, "5": 36.0, "6": 75.0, "7": 30.0, "8": 108.0, "9": 89.0, "10": 64.0, "11": 69.0, "13": 50.0, "14": 17.0, "15": 44.0, "16": 72.0, "17": 66.0, "18": 36.0, "22": 30.0, "23": 36.0, "24": 15.0, "25": 57.0, "26": 33.0, "27": 3.0, "29": 36.0, "30": 72.0}}, "2804112XKN61AH4": {"5": {"4": 24.0, "5": 48.0, "6": 24.0, "7": 36.0, "8": 24.0, "11": 24.0, "12": 39.0, "13": 38.0, "14": 19.0, "15": 45.0, "16": 19.0, "18": 13.0, "19": 38.0, "20": 48.0, "21": 18.0, "22": 24.0, "25": 24.0, "26": 24.0, "28": 28.0, "29": 33.0, "30": 2.0}, "6": {"1": 70.0, "2": 68.0, "3": 64.0, "4": 66.0, "5": 24.0, "6": 48.0, "7": 18.0, "8": 72.0, "9": 69.0, "10": 39.0, "11": 45.0, "13": 43.0, "14": 11.0, "15": 27.0, "16": 48.0, "17": 42.0, "18": 24.0, "22": 18.0, "23": 24.0, "24": 10.0, "25": 38.0, "26": 22.0, "27": 2.0, "29": 24.0, "30": 48.0}}, "2804KN260004A8T": {"5": {"4": 30.0, "5": 60.0, "6": 30.0, "8": 30.0, "11": 30.0, "12": 49.0, "13": 49.0, "14": 24.0, "15": 57.0, "16": 24.0, "18": 10.0, "19": 30.0, "20": 60.0, "22": 30.0, "25": 30.0, "26": 30.0, "28": 30.0, "30": 3.0}, "6": {"1": 88.0, "2": 63.0, "3": 58.0, "4": 83.0, "5": 30.0, "6": 38.0, "8": 90.0, "9": 49.0, "10": 12.0, "11": 49.0, "13": 17.0, "14": 14.0, "15": 12.0, "16": 60.0, "17": 30.0, "18": 30.0, "23": 30.0, "24": 13.0, "25": 48.0, "26": 28.0, "27": 3.0, "29": 30.0, "30": 60.0}}, "2804KN260004A9C": {"5": {"4": 30.0, "5": 60.0, "6": 30.0, "8": 30.0, "11": 30.0, "12": 49.0, "13": 49.0, "14": 24.0, "15": 57.0, "16": 24.0, "18": 10.0, "19": 30.0, "20": 60.0, "22": 30.0, "25": 30.0, "26": 30.0, "28": 30.0, "30": 3.0}, "6": {"1": 88.0, "2": 63.0, "3": 58.0, "4": 83.0, "5": 30.0, "6": 38.0, "8": 90.0, "9": 49.0, "10": 12.0, "11": 49.0, "13": 17.0, "14": 14.0, "15": 12.0, "16": 60.0, "17": 30.0, "18": 30.0, "23": 30.0, "24": 13.0, "25": 48.0, "26": 28.0, "27": 3.0, "29": 30.0, "30": 60.0}}, "2804KN260004AC3": {"5": {"4": 36.0, "5": 72.0, "6": 36.0, "8": 36.0, "11": 36.0, "12": 59.0, "13": 58.0, "14": 29.0, "15": 68.0, "16": 29.0, "18": 12.0, "19": 36.0, "20": 72.0, "22": 36.0, "25": 36.0, "26": 36.0, "28": 36.0, "30": 3.0}, "6": {"1": 105.0, "2": 75.0, "3": 69.0, "4": 99.0, "5": 36.0, "6": 45.0, "8": 108.0, "9": 59.0, "10": 14.0, "11": 59.0, "13": 20.0, "14": 17.0, "15": 14.0, "16": 72.0, "17": 36.0, "18": 36.0, "23": 36.0, "24": 15.0, "25": 57.0, "26": 33.0, "27": 3.0, "29": 36.0, "30": 72.0}}, "2804KN260004AH4": {"5": {"4": 24.0, "5": 48.0, "6": 24.0, "8": 24.0, "11": 24.0, "12": 39.0, "13": 38.0, "14": 19.0, "15": 45.0, "16": 19.0, "18": 8.0, "19": 24.0, "20": 48.0, "22": 24.0, "25": 24.0, "26": 24.0, "28": 24.0, "30": 2.0}, "6": {"1": 70.0, "2": 50.0, "3": 46.0, "4": 66.0, "5": 24.0, "6": 30.0, "8": 72.0, "9": 39.0, "10": 9.0, "11": 39.0, "13": 13.0, "14": 11.0, "15": 9.0, "16": 48.0, "17": 24.0, "18": 24.0, "23": 24.0, "24": 10.0, "25": 38.0, "26": 22.0, "27": 2.0, "29": 24.0, "30": 48.0}}, "2804KN260005A8T": {"5": {"4": 30.0, "5": 60.0, "6": 30.0, "7": 84.0, "8": 30.0, "11": 30.0, "12": 49.0, "13": 49.0, "14": 24.0, "15": 57.0, "16": 24.0, "18": 21.0, "19": 62.0, "20": 60.0, "21": 42.0, "22": 30.0, "25": 30.0, "26": 30.0, "28": 39.0, "29": 76.0, "30": 3.0}, "6": {"1": 88.0, "2": 105.0, "3": 100.0, "4": 83.0, "5": 30.0, "6": 80.0, "7": 42.0, "8": 90.0, "9": 79.0, "10": 85.0, "11": 65.0, "13": 47.0, "14": 14.0, "15": 54.0, "16": 60.0, "17": 72.0, "18": 30.0, "22": 42.0, "23": 30.0, "24": 13.0, "25": 48.0, "26": 28.0, "27": 3.0, "29": 30.0, "30": 60.0}}, "2804KN260005A9C": {"5": {"4": 30.0, "5": 60.0, "6": 30.0, "7": 60.0, "8": 30.0, "11": 30.0, "12": 49.0, "13": 49.0, "14": 24.0, "15": 57.0, "16": 24.0, "18": 18.0, "19": 53.0, "20": 60.0, "21": 30.0, "22": 30.0, "25": 30.0, "26": 30.0, "28": 37.0, "29": 54.0, "30": 3.0}, "6": {"1": 88.0, "2": 93.0, "3": 88.0, "4": 83.0, "5": 30.0, "6": 68.0, "7": 30.0, "8": 90.0, "9": 79.0, "10": 62.0, "11": 59.0, "13": 47.0, "14": 14.0, "15": 42.0, "16": 60.0, "17": 60.0, "18": 30.0, "22": 30.0, "23": 30.0, "24": 13.0, "25": 48.0, "26": 28.0, "27": 3.0, "29": 30.0, "30": 60.0}}, "2804KN260005AC3": {"5": {"4": 36.0, "5": 72.0, "6": 36.0, "7": 60.0, "8": 36.0, "11": 36.0, "12": 59.0, "13": 58.0, "14": 29.0, "15": 68.0, "16": 29.0, "18": 20.0, "19": 59.0, "20": 72.0, "21": 30.0, "22": 36.0, "25": 36.0, "26": 36.0, "28": 43.0, "29": 54.0, "30": 3.0}, "6": {"1": 105.0, "2": 105.0, "3": 99.0, "4": 99.0, "5": 36.0, "6": 75.0, "7": 30.0, "8": 108.0, "9": 89.0, "10": 64.0, "11": 69.0, "13": 50.0, "14": 17.0, "15": 44.0, "16": 72.0, "17": 66.0, "18": 36.0, "22": 30.0, "23": 36.0, "24": 15.0, "25": 57.0, "26": 33.0, "27": 3.0, "29": 36.0, "30": 72.0}}, "2804KN260005AH4": {"5": {"4": 24.0, "5": 48.0, "6": 24.0, "7": 36.0, "8": 24.0, "11": 24.0, "12": 39.0, "13": 38.0, "14": 19.0, "15": 45.0, "16": 19.0, "18": 13.0, "19": 38.0, "20": 48.0, "21": 18.0, "22": 24.0, "25": 24.0, "26": 24.0, "28": 28.0, "29": 33.0, "30": 2.0}, "6": {"1": 70.0, "2": 68.0, "3": 64.0, "4": 66.0, "5": 24.0, "6": 48.0, "7": 18.0, "8": 72.0, "9": 69.0, "10": 39.0, "11": 45.0, "13": 43.0, "14": 11.0, "15": 27.0, "16": 48.0, "17": 42.0, "18": 24.0, "22": 18.0, "23": 24.0, "24": 10.0, "25": 38.0, "26": 22.0, "27": 2.0, "29": 24.0, "30": 48.0}}, "2906101BGW02A": {"5": {"5": 220.0, "6": 260.0, "7": 240.0, "8": 300.0, "11": 180.0, "12": 120.0, "13": 240.0, "14": 240.0, "15": 240.0, "21": 120.0, "23": 120.0, "25": 235.0, "26": 125.0, "27": 75.0, "28": 45.0, "29": 240.0, "30": 120.0}, "6": {"1": 635.0, "2": 205.0, "3": 240.0, "4": 240.0, "5": 220.0, "6": 180.0, "7": 80.0, "9": 490.0, "10": 230.0, "11": 240.0, "15": 165.0, "16": 315.0, "17": 225.0, "18": 270.0, "19": 465.0, "20": 240.0, "22": 350.0, "23": 390.0, "24": 460.0, "25": 380.0, "26": 340.0, "27": 120.0, "29": 240.0}}, "2906113BGW02A": {"5": {"4": 430.0, "5": 310.0, "6": 430.0, "7": 310.0, "8": 360.0, "11": 245.0, "12": 355.0, "13": 360.0, "14": 455.0, "15": 325.0, "16": 180.0, "18": 480.0, "19": 340.0, "20": 310.0, "21": 430.0, "22": 430.0, "23": 275.0, "25": 430.0, "26": 430.0, "27": 550.0, "28": 405.0, "29": 335.0, "30": 265.0}, "6": {"1": 770.0, "2": 630.0, "3": 715.0, "4": 450.0, "5": 450.0, "6": 435.0, "7": 310.0, "8": 700.0, "9": 745.0, "10": 755.0, "11": 490.0, "13": 125.0, "15": 405.0, "16": 330.0, "17": 450.0, "18": 390.0, "19": 465.0, "20": 240.0, "22": 350.0, "23": 390.0, "24": 460.0, "25": 380.0, "26": 340.0, "27": 205.0, "29": 275.0, "30": 325.0}}, "2916107BGW02A": {"5": {"4": 430.0, "5": 310.0, "6": 430.0, "7": 310.0, "8": 360.0, "11": 245.0, "12": 355.0, "13": 360.0, "14": 455.0, "15": 325.0, "16": 180.0, "18": 480.0, "19": 340.0, "20": 310.0, "21": 430.0, "22": 430.0, "23": 275.0, "25": 430.0, "26": 430.0, "27": 550.0, "28": 405.0, "29": 335.0, "30": 265.0}, "6": {"1": 770.0, "2": 630.0, "3": 715.0, "4": 450.0, "5": 450.0, "6": 435.0, "7": 310.0, "8": 700.0, "9": 745.0, "10": 755.0, "11": 490.0, "13": 125.0, "15": 405.0, "16": 330.0, "17": 450.0, "18": 390.0, "19": 465.0, "20": 240.0, "22": 350.0, "23": 390.0, "24": 460.0, "25": 380.0, "26": 340.0, "27": 205.0, "29": 275.0, "30": 325.0}}, "3101100XKJ23A": {"5": {"4": 190.0, "5": 90.0, "6": 170.0, "7": 70.0, "8": 60.0, "11": 65.0, "12": 115.0, "13": 120.0, "14": 95.0, "15": 85.0, "16": 60.0, "18": 120.0, "19": 100.0, "20": 80.0, "21": 180.0, "22": 70.0, "23": 50.0, "25": 60.0, "26": 65.0, "27": 115.0, "28": 120.0, "29": 60.0, "30": 60.0}, "6": {"1": 60.0, "2": 120.0, "3": 60.0, "4": 60.0, "5": 60.0, "6": 60.0, "8": 60.0, "9": 120.0, "10": 60.0}}, "3101100XKN46A": {"5": {"15": 25.0, "16": 35.0, "18": 40.0, "19": 20.0, "22": 40.0, "23": 20.0, "30": 40.0}, "6": {"1": 20.0}}, "3101100xst33A": {"5": {"4": 240.0, "6": 120.0, "7": 120.0, "11": 120.0, "12": 120.0, "14": 240.0, "15": 120.0, "16": 120.0, "18": 120.0, "19": 240.0, "20": 230.0, "21": 250.0, "22": 120.0, "23": 225.0, "25": 370.0, "26": 125.0, "27": 75.0, "28": 45.0, "29": 275.0, "30": 205.0}, "6": {"1": 635.0, "2": 205.0, "3": 240.0, "4": 120.0, "5": 120.0, "6": 160.0, "7": 80.0, "9": 490.0, "10": 230.0, "11": 240.0, "15": 285.0, "16": 315.0, "17": 225.0, "18": 150.0, "19": 465.0, "20": 240.0, "22": 230.0, "23": 370.0, "25": 120.0, "29": 240.0}}, "3101101XKN61A": {"5": {"5": 240.0, "6": 120.0, "8": 120.0, "11": 120.0, "12": 195.0, "13": 190.0, "14": 95.0, "15": 225.0, "16": 95.0, "18": 40.0, "19": 120.0, "20": 240.0, "22": 120.0, "25": 120.0, "26": 120.0, "28": 120.0, "30": 10.0}, "6": {"1": 350.0, "2": 240.0, "6": 120.0, "8": 240.0, "9": 195.0, "10": 45.0, "11": 195.0, "15": 45.0, "16": 120.0, "17": 120.0, "23": 120.0, "24": 50.0, "25": 70.0, "30": 120.0}}, "3101101XST33A": {"5": {"5": 220.0, "6": 140.0, "7": 120.0, "8": 300.0, "11": 60.0, "12": 120.0, "13": 240.0, "14": 120.0, "15": 120.0, "18": 240.0, "22": 240.0, "26": 240.0, "27": 360.0, "28": 240.0}, "6": {"1": 75.0, "2": 305.0, "3": 415.0, "4": 270.0, "5": 270.0, "6": 215.0, "7": 230.0, "8": 640.0, "9": 135.0, "10": 465.0, "11": 250.0, "13": 125.0, "15": 120.0, "16": 15.0, "17": 225.0, "18": 240.0, "22": 120.0, "23": 20.0, "24": 460.0, "25": 260.0, "26": 340.0, "27": 205.0, "29": 35.0, "30": 325.0}}, "3101102XKN46A": {"5": {"4": 30.0, "5": 90.0, "6": 30.0, "7": 90.0, "8": 80.0, "11": 75.0, "12": 90.0, "13": 5.0, "14": 55.0, "15": 30.0, "16": 35.0, "18": 40.0, "19": 80.0, "21": 60.0, "22": 40.0, "23": 20.0, "26": 60.0, "27": 55.0, "28": 65.0, "29": 55.0, "30": 45.0}, "6": {"1": 20.0, "17": 60.0, "19": 60.0, "20": 10.0, "22": 90.0, "23": 20.0}}, "3101105XKN46A": {"5": {"4": 30.0, "5": 90.0, "6": 30.0, "7": 90.0, "8": 80.0, "11": 75.0, "12": 90.0, "13": 5.0, "14": 55.0, "15": 30.0, "16": 35.0, "18": 40.0, "19": 80.0, "21": 60.0, "22": 40.0, "23": 20.0, "26": 60.0, "27": 55.0, "28": 65.0, "29": 55.0, "30": 45.0}, "6": {"1": 20.0, "17": 60.0, "19": 60.0, "20": 10.0, "22": 90.0, "23": 20.0}}, "3703100XKQ04A": {"5": {"4": 240.0, "5": 220.0, "6": 260.0, "7": 240.0, "8": 300.0, "11": 180.0, "12": 240.0, "13": 240.0, "14": 360.0, "15": 240.0, "16": 120.0, "18": 360.0, "19": 240.0, "20": 230.0, "21": 250.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 365.0, "27": 435.0, "28": 285.0, "29": 275.0, "30": 205.0}, "6": {"1": 710.0, "2": 510.0, "3": 655.0, "4": 390.0, "5": 390.0, "6": 375.0, "7": 310.0, "8": 640.0, "9": 625.0, "10": 695.0, "11": 490.0, "13": 125.0, "15": 405.0, "16": 330.0, "17": 450.0, "18": 390.0, "19": 465.0, "20": 240.0, "22": 350.0, "23": 390.0, "24": 460.0, "25": 380.0, "26": 340.0, "27": 205.0, "29": 275.0, "30": 325.0}}, "3904101XK82XA": {"5": {"4": 520.0, "5": 640.0, "6": 640.0, "7": 640.0, "8": 560.0, "11": 440.0, "12": 640.0, "13": 640.0, "14": 640.0, "15": 640.0, "16": 320.0, "18": 640.0, "19": 640.0, "20": 640.0, "21": 640.0, "22": 640.0, "23": 320.0, "25": 640.0, "26": 640.0, "27": 640.0, "28": 640.0, "29": 640.0, "30": 320.0}, "6": {"1": 1225.0, "2": 1185.0, "3": 925.0, "4": 540.0, "5": 530.0, "6": 705.0, "7": 520.0, "8": 940.0, "9": 1060.0, "10": 1170.0, "11": 735.0, "13": 270.0, "14": 95.0, "15": 570.0, "16": 560.0, "17": 815.0, "18": 390.0, "19": 550.0, "20": 285.0, "22": 560.0, "23": 530.0, "24": 550.0, "25": 470.0, "26": 340.0, "27": 205.0, "29": 330.0, "30": 450.0}}, "5010211XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "5010212XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "5010213XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "5010214XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "5010217XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "5010218XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "5010223XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "5010224XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "5010226XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "5010228XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "5010232XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "5010241XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "5010242XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "5010800XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "5010806XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "5010831XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "5109101AST11A86": {"5": {"5": 220.0, "6": 260.0, "7": 240.0, "8": 300.0, "11": 180.0, "12": 120.0, "13": 240.0, "14": 240.0, "15": 240.0, "21": 120.0, "23": 120.0, "25": 235.0, "26": 125.0, "27": 75.0, "28": 45.0, "29": 240.0, "30": 120.0}, "6": {"1": 635.0, "2": 205.0, "3": 240.0, "4": 240.0, "5": 220.0, "6": 180.0, "7": 80.0, "9": 490.0, "10": 230.0, "11": 240.0, "15": 165.0, "16": 315.0, "17": 225.0, "18": 270.0, "19": 465.0, "20": 240.0, "22": 350.0, "23": 390.0, "24": 460.0, "25": 380.0, "26": 340.0, "27": 120.0, "29": 240.0}}, "5109125XKN61A8P": {"5": {"5": 240.0, "6": 120.0, "7": 240.0, "8": 120.0, "11": 120.0, "12": 195.0, "13": 190.0, "14": 95.0, "15": 225.0, "16": 95.0, "18": 70.0, "19": 210.0, "20": 240.0, "21": 120.0, "22": 120.0, "25": 120.0, "26": 120.0, "28": 145.0, "29": 215.0, "30": 10.0}, "6": {"1": 350.0, "2": 360.0, "3": 120.0, "6": 240.0, "7": 120.0, "8": 240.0, "9": 315.0, "10": 245.0, "11": 235.0, "13": 120.0, "15": 165.0, "16": 120.0, "17": 240.0, "22": 120.0, "23": 120.0, "24": 50.0, "25": 70.0, "30": 120.0}}, "5109201AST11A86": {"5": {"5": 220.0, "6": 260.0, "7": 240.0, "8": 300.0, "11": 180.0, "12": 120.0, "13": 240.0, "14": 240.0, "15": 240.0, "21": 120.0, "23": 120.0, "25": 235.0, "26": 125.0, "27": 75.0, "28": 45.0, "29": 240.0, "30": 120.0}, "6": {"1": 635.0, "2": 205.0, "3": 240.0, "4": 240.0, "5": 220.0, "6": 180.0, "7": 80.0, "9": 490.0, "10": 230.0, "11": 240.0, "15": 165.0, "16": 315.0, "17": 225.0, "18": 270.0, "19": 465.0, "20": 240.0, "22": 350.0, "23": 390.0, "24": 460.0, "25": 380.0, "26": 340.0, "27": 120.0, "29": 240.0}}, "5109201XKN61A8P": {"5": {"5": 240.0, "6": 120.0, "7": 240.0, "8": 120.0, "11": 120.0, "12": 195.0, "13": 190.0, "14": 95.0, "15": 225.0, "16": 95.0, "18": 70.0, "19": 210.0, "20": 240.0, "21": 120.0, "22": 120.0, "25": 120.0, "26": 120.0, "28": 145.0, "29": 215.0, "30": 10.0}, "6": {"1": 350.0, "2": 360.0, "3": 120.0, "6": 240.0, "7": 120.0, "8": 240.0, "9": 315.0, "10": 245.0, "11": 235.0, "13": 120.0, "15": 165.0, "16": 120.0, "17": 240.0, "22": 120.0, "23": 120.0, "24": 50.0, "25": 70.0, "30": 120.0}}, "5109501AST11A86": {"5": {"4": 240.0, "12": 120.0, "14": 120.0, "16": 120.0, "18": 360.0, "19": 240.0, "20": 230.0, "21": 130.0, "22": 360.0, "23": 105.0, "25": 135.0, "26": 240.0, "27": 360.0, "28": 240.0, "29": 35.0, "30": 85.0}, "6": {"1": 75.0, "2": 305.0, "3": 415.0, "4": 150.0, "5": 170.0, "6": 195.0, "7": 230.0, "8": 640.0, "9": 135.0, "10": 465.0, "11": 250.0, "13": 125.0, "15": 240.0, "16": 15.0, "17": 225.0, "18": 120.0, "27": 85.0, "29": 35.0, "30": 325.0}}, "5109800XKN46A": {"5": {"4": 30.0, "5": 90.0, "6": 30.0, "7": 90.0, "8": 80.0, "11": 75.0, "12": 90.0, "13": 5.0, "14": 55.0, "15": 30.0, "16": 35.0, "18": 40.0, "19": 80.0, "21": 60.0, "22": 40.0, "23": 20.0, "26": 60.0, "27": 55.0, "28": 65.0, "29": 55.0, "30": 45.0}, "6": {"1": 20.0, "17": 60.0, "19": 60.0, "20": 10.0, "22": 90.0, "23": 20.0}}, "5109800XKN83A": {"5": {"4": 60.0, "6": 60.0, "13": 85.0, "14": 35.0, "15": 60.0, "16": 10.0, "18": 50.0, "19": 10.0, "20": 90.0, "21": 30.0, "22": 50.0, "23": 25.0, "25": 90.0, "26": 30.0, "27": 35.0, "28": 25.0, "29": 35.0}, "6": {"1": 85.0, "2": 195.0, "3": 90.0, "4": 90.0, "5": 80.0, "6": 30.0, "7": 90.0, "10": 170.0, "11": 10.0, "13": 25.0, "14": 95.0, "16": 110.0, "17": 65.0, "19": 25.0, "20": 35.0, "24": 40.0, "25": 20.0, "29": 55.0, "30": 5.0}}, "5120101XGW01A": {"5": {"4": 240.0, "5": 120.0, "7": 120.0, "8": 240.0, "11": 65.0, "12": 190.0, "13": 105.0, "14": 215.0, "15": 145.0, "18": 120.0, "19": 240.0, "21": 120.0, "22": 120.0, "25": 120.0, "27": 120.0, "29": 120.0, "30": 120.0}, "6": {"1": 240.0, "4": 120.0, "5": 160.0, "6": 80.0, "8": 165.0, "9": 75.0, "10": 165.0, "11": 75.0, "15": 120.0, "16": 240.0, "22": 110.0, "23": 30.0, "24": 100.0, "29": 120.0, "30": 120.0}}, "5120102XGW01A": {"5": {"4": 240.0, "5": 120.0, "7": 120.0, "8": 240.0, "11": 65.0, "12": 190.0, "13": 105.0, "14": 215.0, "15": 145.0, "18": 120.0, "19": 240.0, "21": 120.0, "22": 120.0, "25": 120.0, "27": 120.0, "29": 120.0, "30": 120.0}, "6": {"1": 240.0, "4": 120.0, "5": 160.0, "6": 80.0, "8": 165.0, "9": 75.0, "10": 165.0, "11": 75.0, "15": 120.0, "16": 240.0, "22": 110.0, "23": 30.0, "24": 100.0, "29": 120.0, "30": 120.0}}, "5120107XST11А": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "5120108XST11B": {"5": {"4": 310.0, "5": 430.0, "6": 310.0, "7": 430.0, "8": 240.0, "11": 300.0, "12": 360.0, "13": 445.0, "14": 335.0, "15": 405.0, "16": 255.0, "18": 330.0, "19": 310.0, "20": 430.0, "21": 430.0, "22": 430.0, "23": 275.0, "25": 430.0, "26": 550.0, "27": 415.0, "28": 325.0, "29": 430.0, "30": 155.0}, "6": {"1": 640.0, "2": 715.0, "3": 685.0, "4": 485.0, "5": 290.0, "6": 505.0, "7": 370.0, "8": 775.0, "9": 735.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "5120364XGW01A": {"5": {"4": 240.0, "5": 120.0, "6": 240.0, "7": 120.0, "8": 240.0, "11": 65.0, "12": 190.0, "13": 105.0, "14": 215.0, "15": 145.0, "16": 20.0, "18": 220.0, "19": 240.0, "20": 120.0, "21": 120.0, "22": 120.0, "25": 120.0, "27": 135.0, "28": 225.0, "29": 120.0, "30": 120.0}, "6": {"1": 360.0, "2": 120.0, "4": 120.0, "5": 280.0, "6": 200.0, "8": 285.0, "9": 265.0, "10": 335.0, "11": 195.0, "15": 120.0, "16": 360.0, "19": 120.0, "22": 110.0, "23": 30.0, "24": 100.0, "29": 120.0, "30": 120.0}}, "5120365XGW01A": {"5": {"4": 240.0, "5": 120.0, "6": 240.0, "7": 120.0, "8": 240.0, "11": 65.0, "12": 190.0, "13": 105.0, "14": 215.0, "15": 145.0, "16": 20.0, "18": 220.0, "19": 240.0, "20": 120.0, "21": 120.0, "22": 120.0, "25": 120.0, "27": 135.0, "28": 225.0, "29": 120.0, "30": 120.0}, "6": {"1": 360.0, "2": 120.0, "4": 120.0, "5": 280.0, "6": 200.0, "8": 285.0, "9": 265.0, "10": 335.0, "11": 195.0, "15": 120.0, "16": 360.0, "19": 120.0, "22": 110.0, "23": 30.0, "24": 100.0, "29": 120.0, "30": 120.0}}, "5120383XGW01A": {"5": {"4": 240.0, "5": 120.0, "7": 120.0, "8": 240.0, "11": 65.0, "12": 190.0, "13": 105.0, "14": 215.0, "15": 145.0, "18": 120.0, "19": 240.0, "21": 120.0, "22": 120.0, "25": 120.0, "27": 120.0, "29": 120.0, "30": 120.0}, "6": {"1": 240.0, "4": 120.0, "5": 160.0, "6": 80.0, "8": 165.0, "9": 75.0, "10": 165.0, "11": 75.0, "15": 120.0, "16": 240.0, "22": 110.0, "23": 30.0, "24": 100.0, "29": 120.0, "30": 120.0}}, "5120384XGW01A": {"5": {"4": 240.0, "5": 120.0, "7": 120.0, "8": 240.0, "11": 65.0, "12": 190.0, "13": 105.0, "14": 215.0, "15": 145.0, "18": 120.0, "19": 240.0, "21": 120.0, "22": 120.0, "25": 120.0, "27": 120.0, "29": 120.0, "30": 120.0}, "6": {"1": 240.0, "4": 120.0, "5": 160.0, "6": 80.0, "8": 165.0, "9": 75.0, "10": 165.0, "11": 75.0, "15": 120.0, "16": 240.0, "22": 110.0, "23": 30.0, "24": 100.0, "29": 120.0, "30": 120.0}}, "5120807XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "5120808XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "5122105XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "5122106XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "5122113XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "5122114XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "5122121XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "5122122XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "5122211XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "5122212XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "5122213XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "5122214XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "5122218XGW01A": {"5": {"4": 240.0, "5": 120.0, "7": 120.0, "8": 240.0, "11": 65.0, "12": 190.0, "13": 105.0, "14": 215.0, "15": 145.0, "18": 120.0, "19": 240.0, "21": 120.0, "22": 120.0, "25": 120.0, "27": 120.0, "29": 120.0, "30": 120.0}, "6": {"1": 240.0, "4": 120.0, "5": 160.0, "6": 80.0, "8": 165.0, "9": 75.0, "10": 165.0, "11": 75.0, "15": 120.0, "16": 240.0, "22": 110.0, "23": 30.0, "24": 100.0, "29": 120.0, "30": 120.0}}, "5122221XGW01A": {"5": {"4": 240.0, "5": 120.0, "7": 120.0, "8": 240.0, "11": 65.0, "12": 190.0, "13": 105.0, "14": 215.0, "15": 145.0, "18": 120.0, "19": 240.0, "21": 120.0, "22": 120.0, "25": 120.0, "27": 120.0, "29": 120.0, "30": 120.0}, "6": {"1": 240.0, "4": 120.0, "5": 160.0, "6": 80.0, "8": 165.0, "9": 75.0, "10": 165.0, "11": 75.0, "15": 120.0, "16": 240.0, "22": 110.0, "23": 30.0, "24": 100.0, "29": 120.0, "30": 120.0}}, "5130103XGW02A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "5130104XGW02A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "5130121XGW01A": {"5": {"4": 240.0, "5": 120.0, "7": 120.0, "8": 240.0, "11": 65.0, "12": 190.0, "13": 105.0, "14": 215.0, "15": 145.0, "18": 120.0, "19": 240.0, "21": 120.0, "22": 120.0, "25": 120.0, "27": 120.0, "29": 120.0, "30": 120.0}, "6": {"1": 240.0, "4": 120.0, "5": 160.0, "6": 80.0, "8": 165.0, "9": 75.0, "10": 165.0, "11": 75.0, "15": 120.0, "16": 240.0, "22": 110.0, "23": 30.0, "24": 100.0, "29": 120.0, "30": 120.0}}, "5130127XGW01A": {"5": {"4": 240.0, "5": 120.0, "7": 120.0, "8": 240.0, "11": 65.0, "12": 190.0, "13": 105.0, "14": 215.0, "15": 145.0, "18": 120.0, "19": 240.0, "21": 120.0, "22": 120.0, "25": 120.0, "27": 120.0, "29": 120.0, "30": 120.0}, "6": {"1": 240.0, "4": 120.0, "5": 160.0, "6": 80.0, "8": 165.0, "9": 75.0, "10": 165.0, "11": 75.0, "15": 120.0, "16": 240.0, "22": 110.0, "23": 30.0, "24": 100.0, "29": 120.0, "30": 120.0}}, "5130128XGW01A": {"5": {"4": 240.0, "5": 120.0, "7": 120.0, "8": 240.0, "11": 65.0, "12": 190.0, "13": 105.0, "14": 215.0, "15": 145.0, "18": 120.0, "19": 240.0, "21": 120.0, "22": 120.0, "25": 120.0, "27": 120.0, "29": 120.0, "30": 120.0}, "6": {"1": 240.0, "4": 120.0, "5": 160.0, "6": 80.0, "8": 165.0, "9": 75.0, "10": 165.0, "11": 75.0, "15": 120.0, "16": 240.0, "22": 110.0, "23": 30.0, "24": 100.0, "29": 120.0, "30": 120.0}}, "5130137XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "5130138XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "5130167XGW01A": {"5": {"4": 240.0, "5": 120.0, "7": 120.0, "8": 240.0, "11": 65.0, "12": 190.0, "13": 105.0, "14": 215.0, "15": 145.0, "18": 120.0, "19": 240.0, "21": 120.0, "22": 120.0, "25": 120.0, "27": 120.0, "29": 120.0, "30": 120.0}, "6": {"1": 240.0, "4": 120.0, "5": 160.0, "6": 80.0, "8": 165.0, "9": 75.0, "10": 165.0, "11": 75.0, "15": 120.0, "16": 240.0, "22": 110.0, "23": 30.0, "24": 100.0, "29": 120.0, "30": 120.0}}, "5130199XGW01B": {"5": {"4": 240.0, "5": 120.0, "7": 120.0, "8": 240.0, "11": 65.0, "12": 190.0, "13": 105.0, "14": 215.0, "15": 145.0, "18": 120.0, "19": 240.0, "21": 120.0, "22": 120.0, "25": 120.0, "27": 120.0, "29": 120.0, "30": 120.0}, "6": {"1": 240.0, "4": 120.0, "5": 160.0, "6": 80.0, "8": 165.0, "9": 75.0, "10": 165.0, "11": 75.0, "15": 120.0, "16": 240.0, "22": 110.0, "23": 30.0, "24": 100.0, "29": 120.0, "30": 120.0}}, "5130200XGW01B": {"5": {"4": 240.0, "5": 120.0, "7": 120.0, "8": 240.0, "11": 65.0, "12": 190.0, "13": 105.0, "14": 215.0, "15": 145.0, "18": 120.0, "19": 240.0, "21": 120.0, "22": 120.0, "25": 120.0, "27": 120.0, "29": 120.0, "30": 120.0}, "6": {"1": 240.0, "4": 120.0, "5": 160.0, "6": 80.0, "8": 165.0, "9": 75.0, "10": 165.0, "11": 75.0, "15": 120.0, "16": 240.0, "22": 110.0, "23": 30.0, "24": 100.0, "29": 120.0, "30": 120.0}}, "5130203XGW01A": {"5": {"4": 240.0, "5": 120.0, "7": 120.0, "8": 240.0, "11": 65.0, "12": 190.0, "13": 105.0, "14": 215.0, "15": 145.0, "18": 120.0, "19": 240.0, "21": 120.0, "22": 120.0, "25": 120.0, "27": 120.0, "29": 120.0, "30": 120.0}, "6": {"1": 240.0, "4": 120.0, "5": 160.0, "6": 80.0, "8": 165.0, "9": 75.0, "10": 165.0, "11": 75.0, "15": 120.0, "16": 240.0, "22": 110.0, "23": 30.0, "24": 100.0, "29": 120.0, "30": 120.0}}, "5130219XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "5130248XST11A": {"5": {"11": 120.0, "13": 120.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 140.0, "21": 360.0, "22": 120.0, "23": 120.0, "25": 235.0, "26": 365.0, "27": 240.0, "28": 25.0, "29": 95.0, "30": 10.0}, "6": {"1": 325.0, "2": 395.0, "3": 270.0, "4": 200.0, "5": 120.0, "6": 415.0, "7": 370.0, "8": 175.0, "9": 435.0, "10": 365.0, "11": 280.0, "16": 240.0, "17": 120.0, "26": 120.0, "29": 295.0, "30": 65.0}}, "5130253XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "5130381XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "5130397XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "5130448XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "5130449XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "5130461XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "5130477XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "5130478XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "5130555XGW01B": {"5": {"4": 240.0, "5": 120.0, "7": 120.0, "8": 240.0, "11": 65.0, "12": 190.0, "13": 105.0, "14": 215.0, "15": 145.0, "18": 120.0, "19": 240.0, "21": 120.0, "22": 120.0, "25": 120.0, "27": 120.0, "29": 120.0, "30": 120.0}, "6": {"1": 240.0, "4": 120.0, "5": 160.0, "6": 80.0, "8": 165.0, "9": 75.0, "10": 165.0, "11": 75.0, "15": 120.0, "16": 240.0, "22": 110.0, "23": 30.0, "24": 100.0, "29": 120.0, "30": 120.0}}, "5130621XGW02A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "5130807XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "5130851XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "5130852XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "5130883XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "5130884XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "5130891XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 120.0, "12": 240.0, "13": 240.0, "14": 240.0, "20": 120.0, "22": 240.0, "23": 105.0, "25": 135.0, "26": 65.0, "27": 55.0, "28": 240.0, "29": 240.0, "30": 120.0}, "6": {"1": 215.0, "2": 240.0, "3": 355.0, "4": 225.0, "5": 50.0, "6": 90.0, "8": 480.0, "9": 240.0, "10": 240.0, "11": 135.0, "15": 345.0, "16": 195.0, "17": 270.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 240.0, "27": 120.0, "30": 385.0}}, "5173102XGW01B": {"5": {"4": 90.0, "5": 330.0, "6": 210.0, "7": 330.0, "8": 80.0, "11": 195.0, "12": 90.0, "13": 90.0, "14": 90.0, "15": 315.0, "16": 140.0, "18": 160.0, "19": 300.0, "20": 330.0, "21": 210.0, "22": 210.0, "23": 45.0, "25": 210.0, "26": 90.0, "27": 90.0, "28": 115.0, "29": 305.0, "30": 55.0}, "6": {"1": 455.0, "2": 555.0, "3": 210.0, "4": 90.0, "5": 80.0, "6": 270.0, "7": 210.0, "8": 120.0, "9": 315.0, "10": 415.0, "11": 245.0, "13": 145.0, "14": 95.0, "15": 165.0, "16": 110.0, "17": 365.0, "19": 85.0, "20": 45.0, "22": 210.0, "23": 140.0, "24": 90.0, "25": 90.0, "29": 55.0, "30": 5.0}}, "5174104XGW01E": {"5": {"4": 90.0, "5": 330.0, "6": 210.0, "7": 330.0, "8": 80.0, "11": 195.0, "12": 90.0, "13": 90.0, "14": 90.0, "15": 315.0, "16": 140.0, "18": 160.0, "19": 300.0, "20": 330.0, "21": 210.0, "22": 210.0, "23": 45.0, "25": 210.0, "26": 90.0, "27": 90.0, "28": 115.0, "29": 305.0, "30": 55.0}, "6": {"1": 455.0, "2": 555.0, "3": 210.0, "4": 90.0, "5": 80.0, "6": 270.0, "7": 210.0, "8": 120.0, "9": 315.0, "10": 415.0, "11": 245.0, "13": 145.0, "14": 95.0, "15": 165.0, "16": 110.0, "17": 365.0, "19": 85.0, "20": 45.0, "22": 210.0, "23": 140.0, "24": 90.0, "25": 90.0, "29": 55.0, "30": 5.0}}, "5206100XKN83A": {"5": {"4": 60.0, "6": 60.0, "13": 85.0, "14": 35.0, "15": 60.0, "16": 10.0, "18": 50.0, "19": 10.0, "20": 90.0, "21": 30.0, "22": 50.0, "23": 25.0, "25": 90.0, "26": 30.0, "27": 35.0, "28": 25.0, "29": 35.0}, "6": {"1": 85.0, "2": 195.0, "3": 90.0, "4": 90.0, "5": 80.0, "6": 30.0, "7": 90.0, "10": 170.0, "11": 10.0, "13": 25.0, "14": 95.0, "16": 110.0, "17": 65.0, "19": 25.0, "20": 35.0, "24": 40.0, "25": 20.0, "29": 55.0, "30": 5.0}}, "5206101XKN46A": {"5": {"4": 30.0, "5": 90.0, "6": 30.0, "7": 90.0, "8": 80.0, "11": 75.0, "12": 90.0, "13": 5.0, "14": 55.0, "15": 30.0, "16": 35.0, "18": 40.0, "19": 80.0, "21": 60.0, "22": 40.0, "23": 20.0, "26": 60.0, "27": 55.0, "28": 65.0, "29": 55.0, "30": 45.0}, "6": {"1": 20.0, "17": 60.0, "19": 60.0, "20": 10.0, "22": 90.0, "23": 20.0}}, "5206102XKN46A": {"5": {"4": 30.0, "5": 90.0, "6": 30.0, "7": 90.0, "8": 80.0, "11": 75.0, "12": 90.0, "13": 5.0, "14": 55.0, "15": 30.0, "16": 35.0, "18": 40.0, "19": 80.0, "21": 60.0, "22": 40.0, "23": 20.0, "26": 60.0, "27": 55.0, "28": 65.0, "29": 55.0, "30": 45.0}, "6": {"1": 20.0, "17": 60.0, "19": 60.0, "20": 10.0, "22": 90.0, "23": 20.0}}, "5206103XKN46A": {"5": {"4": 30.0, "5": 90.0, "6": 30.0, "7": 90.0, "8": 80.0, "11": 75.0, "12": 90.0, "13": 5.0, "14": 55.0, "15": 30.0, "16": 35.0, "18": 40.0, "19": 80.0, "21": 60.0, "22": 40.0, "23": 20.0, "26": 60.0, "27": 55.0, "28": 65.0, "29": 55.0, "30": 45.0}, "6": {"1": 20.0, "17": 60.0, "19": 60.0, "20": 10.0, "22": 90.0, "23": 20.0}}, "5206104XKN61A": {"5": {"5": 240.0, "6": 120.0, "7": 240.0, "8": 120.0, "11": 120.0, "12": 195.0, "13": 190.0, "14": 95.0, "15": 225.0, "16": 95.0, "18": 70.0, "19": 210.0, "20": 240.0, "21": 120.0, "22": 120.0, "25": 120.0, "26": 120.0, "28": 145.0, "29": 215.0, "30": 10.0}, "6": {"1": 350.0, "2": 360.0, "3": 120.0, "6": 240.0, "7": 120.0, "8": 240.0, "9": 315.0, "10": 245.0, "11": 235.0, "13": 120.0, "15": 165.0, "16": 120.0, "17": 240.0, "22": 120.0, "23": 120.0, "24": 50.0, "25": 70.0, "30": 120.0}}, "5206105XKN61A": {"5": {"5": 240.0, "6": 120.0, "7": 240.0, "8": 120.0, "11": 120.0, "12": 195.0, "13": 190.0, "14": 95.0, "15": 225.0, "16": 95.0, "18": 70.0, "19": 210.0, "20": 240.0, "21": 120.0, "22": 120.0, "25": 120.0, "26": 120.0, "28": 145.0, "29": 215.0, "30": 10.0}, "6": {"1": 350.0, "2": 360.0, "3": 120.0, "6": 240.0, "7": 120.0, "8": 240.0, "9": 315.0, "10": 245.0, "11": 235.0, "13": 120.0, "15": 165.0, "16": 120.0, "17": 240.0, "22": 120.0, "23": 120.0, "24": 50.0, "25": 70.0, "30": 120.0}}, "5206200XKN83A": {"5": {"4": 60.0, "6": 60.0, "13": 85.0, "14": 35.0, "15": 60.0, "16": 10.0, "18": 50.0, "19": 10.0, "20": 90.0, "21": 30.0, "22": 50.0, "23": 25.0, "25": 90.0, "26": 30.0, "27": 35.0, "28": 25.0, "29": 35.0}, "6": {"1": 85.0, "2": 195.0, "3": 90.0, "4": 90.0, "5": 80.0, "6": 30.0, "7": 90.0, "10": 170.0, "11": 10.0, "13": 25.0, "14": 95.0, "16": 110.0, "17": 65.0, "19": 25.0, "20": 35.0, "24": 40.0, "25": 20.0, "29": 55.0, "30": 5.0}}, "5206300XKN83A": {"5": {"4": 60.0, "6": 60.0, "13": 85.0, "14": 35.0, "15": 60.0, "16": 10.0, "18": 50.0, "19": 10.0, "20": 90.0, "21": 30.0, "22": 50.0, "23": 25.0, "25": 90.0, "26": 30.0, "27": 35.0, "28": 25.0, "29": 35.0}, "6": {"1": 85.0, "2": 195.0, "3": 90.0, "4": 90.0, "5": 80.0, "6": 30.0, "7": 90.0, "10": 170.0, "11": 10.0, "13": 25.0, "14": 95.0, "16": 110.0, "17": 65.0, "19": 25.0, "20": 35.0, "24": 40.0, "25": 20.0, "29": 55.0, "30": 5.0}}, "5206300XST10A": {"5": {"6": 120.0, "7": 120.0, "11": 120.0, "14": 120.0, "15": 120.0, "25": 115.0, "26": 5.0, "29": 240.0, "30": 120.0}, "6": {"1": 515.0, "2": 85.0, "3": 120.0, "5": 120.0, "6": 160.0, "7": 80.0, "9": 250.0, "10": 230.0, "29": 240.0}}, "5206400XKJ23A": {"5": {"4": 190.0, "5": 90.0, "6": 170.0, "7": 70.0, "8": 60.0, "11": 65.0, "12": 115.0, "13": 120.0, "14": 95.0, "15": 85.0, "16": 60.0, "18": 120.0, "19": 100.0, "20": 80.0, "21": 180.0, "22": 70.0, "23": 50.0, "25": 60.0, "26": 65.0, "27": 115.0, "28": 120.0, "29": 60.0, "30": 60.0}, "6": {"1": 60.0, "2": 120.0, "3": 60.0, "4": 60.0, "5": 60.0, "6": 60.0, "8": 60.0, "9": 120.0, "10": 60.0}}, "5206500XKJ23A": {"5": {"4": 190.0, "5": 90.0, "6": 170.0, "7": 70.0, "8": 60.0, "11": 65.0, "12": 115.0, "13": 120.0, "14": 95.0, "15": 85.0, "16": 60.0, "18": 120.0, "19": 100.0, "20": 80.0, "21": 180.0, "22": 70.0, "23": 50.0, "25": 60.0, "26": 65.0, "27": 115.0, "28": 120.0, "29": 60.0, "30": 60.0}, "6": {"1": 60.0, "2": 120.0, "3": 60.0, "4": 60.0, "5": 60.0, "6": 60.0, "8": 60.0, "9": 120.0, "10": 60.0}}, "5206603XST10A": {"5": {"4": 240.0, "12": 120.0, "14": 120.0, "16": 120.0, "18": 120.0, "19": 240.0, "20": 230.0, "21": 250.0, "22": 120.0, "23": 225.0, "25": 255.0, "26": 120.0, "27": 75.0, "28": 45.0, "29": 35.0, "30": 85.0}, "6": {"1": 120.0, "2": 120.0, "3": 120.0, "4": 120.0, "9": 240.0, "11": 240.0, "15": 285.0, "16": 315.0, "17": 225.0, "18": 150.0, "19": 465.0, "20": 240.0, "22": 230.0, "23": 370.0, "25": 120.0}}, "5206605XST10A": {"5": {"5": 220.0, "6": 140.0, "7": 120.0, "8": 300.0, "11": 60.0, "12": 120.0, "13": 240.0, "14": 120.0, "15": 120.0, "27": 120.0}, "6": {"1": 40.0, "2": 200.0, "3": 160.0, "4": 200.0, "5": 100.0, "6": 95.0, "7": 230.0, "8": 415.0, "10": 225.0, "11": 250.0, "13": 125.0, "18": 120.0, "22": 120.0, "23": 20.0, "24": 460.0, "25": 260.0, "26": 340.0, "27": 120.0}}, "5300103XGW01A": {"5": {"4": 240.0, "5": 120.0, "7": 120.0, "8": 240.0, "11": 65.0, "12": 190.0, "13": 105.0, "14": 215.0, "15": 145.0, "18": 120.0, "19": 240.0, "21": 120.0, "22": 120.0, "25": 120.0, "27": 120.0, "29": 120.0, "30": 120.0}, "6": {"1": 240.0, "4": 120.0, "5": 160.0, "6": 80.0, "8": 165.0, "9": 75.0, "10": 165.0, "11": 75.0, "15": 120.0, "16": 240.0, "22": 110.0, "23": 30.0, "24": 100.0, "29": 120.0, "30": 120.0}}, "5300120XGW01B": {"5": {"4": 240.0, "5": 120.0, "7": 120.0, "8": 240.0, "11": 65.0, "12": 190.0, "13": 105.0, "14": 215.0, "15": 145.0, "18": 120.0, "19": 240.0, "21": 120.0, "22": 120.0, "25": 120.0, "27": 120.0, "29": 120.0, "30": 120.0}, "6": {"1": 240.0, "4": 120.0, "5": 160.0, "6": 80.0, "8": 165.0, "9": 75.0, "10": 165.0, "11": 75.0, "15": 120.0, "16": 240.0, "22": 110.0, "23": 30.0, "24": 100.0, "29": 120.0, "30": 120.0}}, "5300123XGW01B": {"5": {"4": 240.0, "5": 120.0, "7": 120.0, "8": 240.0, "11": 65.0, "12": 190.0, "13": 105.0, "14": 215.0, "15": 145.0, "18": 120.0, "19": 240.0, "21": 120.0, "22": 120.0, "25": 120.0, "27": 120.0, "29": 120.0, "30": 120.0}, "6": {"1": 240.0, "4": 120.0, "5": 160.0, "6": 80.0, "8": 165.0, "9": 75.0, "10": 165.0, "11": 75.0, "15": 120.0, "16": 240.0, "22": 110.0, "23": 30.0, "24": 100.0, "29": 120.0, "30": 120.0}}, "5300164XGW01B": {"5": {"4": 240.0, "5": 120.0, "7": 120.0, "8": 240.0, "11": 65.0, "12": 190.0, "13": 105.0, "14": 215.0, "15": 145.0, "18": 120.0, "19": 240.0, "21": 120.0, "22": 120.0, "25": 120.0, "27": 120.0, "29": 120.0, "30": 120.0}, "6": {"1": 240.0, "4": 120.0, "5": 160.0, "6": 80.0, "8": 165.0, "9": 75.0, "10": 165.0, "11": 75.0, "15": 120.0, "16": 240.0, "22": 110.0, "23": 30.0, "24": 100.0, "29": 120.0, "30": 120.0}}, "5301102XST33A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "5301105XST33A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "5301111XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "5301113XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "5301114XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "5301120XST11A": {"5": {"5": 120.0, "6": 120.0, "8": 120.0, "13": 120.0, "14": 120.0, "23": 105.0, "25": 15.0, "28": 240.0, "29": 120.0, "30": 120.0}, "6": {"1": 95.0, "2": 120.0, "3": 115.0, "4": 155.0, "6": 90.0, "8": 240.0, "9": 240.0, "26": 120.0, "27": 120.0, "30": 385.0}}, "5301122XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "5301126XKN02A": {"5": {"4": 240.0, "5": 120.0, "7": 120.0, "8": 240.0, "11": 65.0, "12": 190.0, "13": 105.0, "14": 215.0, "15": 145.0, "18": 120.0, "19": 240.0, "21": 120.0, "22": 120.0, "25": 120.0, "27": 120.0, "29": 120.0, "30": 120.0}, "6": {"1": 240.0, "4": 120.0, "5": 160.0, "6": 80.0, "8": 165.0, "9": 75.0, "10": 165.0, "11": 75.0, "15": 120.0, "16": 240.0, "22": 110.0, "23": 30.0, "24": 100.0, "29": 120.0, "30": 120.0}}, "5301131XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "5301133XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "5301134XKN02A": {"5": {"4": 240.0, "5": 120.0, "7": 120.0, "8": 240.0, "11": 65.0, "12": 190.0, "13": 105.0, "14": 215.0, "15": 145.0, "18": 120.0, "19": 240.0, "21": 120.0, "22": 120.0, "25": 120.0, "27": 120.0, "29": 120.0, "30": 120.0}, "6": {"1": 240.0, "4": 120.0, "5": 160.0, "6": 80.0, "8": 165.0, "9": 75.0, "10": 165.0, "11": 75.0, "15": 120.0, "16": 240.0, "22": 110.0, "23": 30.0, "24": 100.0, "29": 120.0, "30": 120.0}}, "5301151XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "5301152XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "5301153XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "5301163XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "5301302XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "5301501XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "5301507XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "5301551XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "5301552XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "5304100XKN02A": {"5": {"5": 240.0, "6": 120.0, "8": 120.0, "11": 120.0, "12": 195.0, "13": 190.0, "14": 95.0, "15": 225.0, "16": 95.0, "18": 40.0, "19": 120.0, "20": 240.0, "22": 120.0, "25": 120.0, "26": 120.0, "28": 120.0, "30": 10.0}, "6": {"1": 350.0, "2": 240.0, "6": 120.0, "8": 240.0, "9": 195.0, "10": 45.0, "11": 195.0, "15": 45.0, "16": 120.0, "17": 120.0, "23": 120.0, "24": 50.0, "25": 70.0, "30": 120.0}}, "5401105XKN06A": {"5": {"6": 240.0, "16": 20.0, "18": 100.0, "20": 120.0, "27": 15.0, "28": 225.0}, "6": {"1": 120.0, "2": 120.0, "5": 120.0, "6": 120.0, "8": 120.0, "9": 190.0, "10": 170.0, "11": 120.0, "16": 120.0, "19": 120.0}}, "5401106XKN06A": {"5": {"6": 240.0, "16": 20.0, "18": 100.0, "20": 120.0, "27": 15.0, "28": 225.0}, "6": {"1": 120.0, "2": 120.0, "5": 120.0, "6": 120.0, "8": 120.0, "9": 190.0, "10": 170.0, "11": 120.0, "16": 120.0, "19": 120.0}}, "5401107XKN06A": {"5": {"6": 240.0, "16": 20.0, "18": 100.0, "20": 120.0, "27": 15.0, "28": 225.0}, "6": {"1": 120.0, "2": 120.0, "5": 120.0, "6": 120.0, "8": 120.0, "9": 190.0, "10": 170.0, "11": 120.0, "16": 120.0, "19": 120.0}}, "5401108XKN06A": {"5": {"6": 240.0, "16": 20.0, "18": 100.0, "20": 120.0, "27": 15.0, "28": 225.0}, "6": {"1": 120.0, "2": 120.0, "5": 120.0, "6": 120.0, "8": 120.0, "9": 190.0, "10": 170.0, "11": 120.0, "16": 120.0, "19": 120.0}}, "5401109XKN06A": {"5": {"6": 240.0, "16": 20.0, "18": 100.0, "20": 120.0, "27": 15.0, "28": 225.0}, "6": {"1": 120.0, "2": 120.0, "5": 120.0, "6": 120.0, "8": 120.0, "9": 190.0, "10": 170.0, "11": 120.0, "16": 120.0, "19": 120.0}}, "5401110XKN06A": {"5": {"6": 240.0, "16": 20.0, "18": 100.0, "20": 120.0, "27": 15.0, "28": 225.0}, "6": {"1": 120.0, "2": 120.0, "5": 120.0, "6": 120.0, "8": 120.0, "9": 190.0, "10": 170.0, "11": 120.0, "16": 120.0, "19": 120.0}}, "5401119XKN06A": {"5": {"6": 240.0, "16": 20.0, "18": 100.0, "20": 120.0, "27": 15.0, "28": 225.0}, "6": {"1": 120.0, "2": 120.0, "5": 120.0, "6": 120.0, "8": 120.0, "9": 190.0, "10": 170.0, "11": 120.0, "16": 120.0, "19": 120.0}}, "5401120XKN06A": {"5": {"6": 240.0, "16": 20.0, "18": 100.0, "20": 120.0, "27": 15.0, "28": 225.0}, "6": {"1": 120.0, "2": 120.0, "5": 120.0, "6": 120.0, "8": 120.0, "9": 190.0, "10": 170.0, "11": 120.0, "16": 120.0, "19": 120.0}}, "5401121XKN06A": {"5": {"6": 240.0, "16": 20.0, "18": 100.0, "20": 120.0, "27": 15.0, "28": 225.0}, "6": {"1": 120.0, "2": 120.0, "5": 120.0, "6": 120.0, "8": 120.0, "9": 190.0, "10": 170.0, "11": 120.0, "16": 120.0, "19": 120.0}}, "5401122XKN06A": {"5": {"6": 240.0, "16": 20.0, "18": 100.0, "20": 120.0, "27": 15.0, "28": 225.0}, "6": {"1": 120.0, "2": 120.0, "5": 120.0, "6": 120.0, "8": 120.0, "9": 190.0, "10": 170.0, "11": 120.0, "16": 120.0, "19": 120.0}}, "5401187XKN01B": {"5": {"4": 240.0, "5": 120.0, "7": 120.0, "8": 240.0, "11": 65.0, "12": 190.0, "13": 105.0, "14": 215.0, "15": 145.0, "18": 120.0, "19": 240.0, "21": 120.0, "22": 120.0, "25": 120.0, "27": 120.0, "29": 120.0, "30": 120.0}, "6": {"1": 240.0, "4": 120.0, "5": 160.0, "6": 80.0, "8": 165.0, "9": 75.0, "10": 165.0, "11": 75.0, "15": 120.0, "16": 240.0, "22": 110.0, "23": 30.0, "24": 100.0, "29": 120.0, "30": 120.0}}, "5401261XKN02A": {"5": {"4": 240.0, "5": 120.0, "7": 120.0, "8": 240.0, "11": 65.0, "12": 190.0, "13": 105.0, "14": 215.0, "15": 145.0, "18": 120.0, "19": 240.0, "21": 120.0, "22": 120.0, "25": 120.0, "27": 120.0, "29": 120.0, "30": 120.0}, "6": {"1": 240.0, "4": 120.0, "5": 160.0, "6": 80.0, "8": 165.0, "9": 75.0, "10": 165.0, "11": 75.0, "15": 120.0, "16": 240.0, "22": 110.0, "23": 30.0, "24": 100.0, "29": 120.0, "30": 120.0}}, "5401262XKN02A": {"5": {"4": 240.0, "5": 120.0, "7": 120.0, "8": 240.0, "11": 65.0, "12": 190.0, "13": 105.0, "14": 215.0, "15": 145.0, "18": 120.0, "19": 240.0, "21": 120.0, "22": 120.0, "25": 120.0, "27": 120.0, "29": 120.0, "30": 120.0}, "6": {"1": 240.0, "4": 120.0, "5": 160.0, "6": 80.0, "8": 165.0, "9": 75.0, "10": 165.0, "11": 75.0, "15": 120.0, "16": 240.0, "22": 110.0, "23": 30.0, "24": 100.0, "29": 120.0, "30": 120.0}}, "5401289XKN61A": {"5": {"4": 240.0, "5": 120.0, "6": 240.0, "7": 120.0, "8": 240.0, "11": 65.0, "12": 190.0, "13": 105.0, "14": 215.0, "15": 145.0, "16": 20.0, "18": 220.0, "19": 240.0, "20": 120.0, "21": 120.0, "22": 120.0, "25": 120.0, "27": 135.0, "28": 225.0, "29": 120.0, "30": 120.0}, "6": {"1": 360.0, "2": 120.0, "4": 120.0, "5": 280.0, "6": 200.0, "8": 285.0, "9": 265.0, "10": 335.0, "11": 195.0, "15": 120.0, "16": 360.0, "19": 120.0, "22": 110.0, "23": 30.0, "24": 100.0, "29": 120.0, "30": 120.0}}, "5401291XKN61A": {"5": {"4": 240.0, "5": 120.0, "6": 240.0, "7": 120.0, "8": 240.0, "11": 65.0, "12": 190.0, "13": 105.0, "14": 215.0, "15": 145.0, "16": 20.0, "18": 220.0, "19": 240.0, "20": 120.0, "21": 120.0, "22": 120.0, "25": 120.0, "27": 135.0, "28": 225.0, "29": 120.0, "30": 120.0}, "6": {"1": 360.0, "2": 120.0, "4": 120.0, "5": 280.0, "6": 200.0, "8": 285.0, "9": 265.0, "10": 335.0, "11": 195.0, "15": 120.0, "16": 360.0, "19": 120.0, "22": 110.0, "23": 30.0, "24": 100.0, "29": 120.0, "30": 120.0}}, "5401301XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "5401302XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "5401311XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "5401312XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "5401413XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "5401414XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "5401417XKN02A": {"5": {"4": 240.0, "5": 120.0, "7": 120.0, "8": 240.0, "11": 65.0, "12": 190.0, "13": 105.0, "14": 215.0, "15": 145.0, "18": 120.0, "19": 240.0, "21": 120.0, "22": 120.0, "25": 120.0, "27": 120.0, "29": 120.0, "30": 120.0}, "6": {"1": 240.0, "4": 120.0, "5": 160.0, "6": 80.0, "8": 165.0, "9": 75.0, "10": 165.0, "11": 75.0, "15": 120.0, "16": 240.0, "22": 110.0, "23": 30.0, "24": 100.0, "29": 120.0, "30": 120.0}}, "5401418XKN02A": {"5": {"4": 240.0, "5": 120.0, "7": 120.0, "8": 240.0, "11": 65.0, "12": 190.0, "13": 105.0, "14": 215.0, "15": 145.0, "18": 120.0, "19": 240.0, "21": 120.0, "22": 120.0, "25": 120.0, "27": 120.0, "29": 120.0, "30": 120.0}, "6": {"1": 240.0, "4": 120.0, "5": 160.0, "6": 80.0, "8": 165.0, "9": 75.0, "10": 165.0, "11": 75.0, "15": 120.0, "16": 240.0, "22": 110.0, "23": 30.0, "24": 100.0, "29": 120.0, "30": 120.0}}, "5401421XKN02A": {"5": {"4": 240.0, "5": 120.0, "7": 120.0, "8": 240.0, "11": 65.0, "12": 190.0, "13": 105.0, "14": 215.0, "15": 145.0, "18": 120.0, "19": 240.0, "21": 120.0, "22": 120.0, "25": 120.0, "27": 120.0, "29": 120.0, "30": 120.0}, "6": {"1": 240.0, "4": 120.0, "5": 160.0, "6": 80.0, "8": 165.0, "9": 75.0, "10": 165.0, "11": 75.0, "15": 120.0, "16": 240.0, "22": 110.0, "23": 30.0, "24": 100.0, "29": 120.0, "30": 120.0}}, "5401422XKN02A": {"5": {"4": 240.0, "5": 120.0, "7": 120.0, "8": 240.0, "11": 65.0, "12": 190.0, "13": 105.0, "14": 215.0, "15": 145.0, "18": 120.0, "19": 240.0, "21": 120.0, "22": 120.0, "25": 120.0, "27": 120.0, "29": 120.0, "30": 120.0}, "6": {"1": 240.0, "4": 120.0, "5": 160.0, "6": 80.0, "8": 165.0, "9": 75.0, "10": 165.0, "11": 75.0, "15": 120.0, "16": 240.0, "22": 110.0, "23": 30.0, "24": 100.0, "29": 120.0, "30": 120.0}}, "5401517XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "5401518XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "5401523XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "5401524XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "5401571XST33A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "5401572XST33A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "5401601XKN02A": {"5": {"4": 240.0, "5": 120.0, "7": 120.0, "8": 240.0, "11": 65.0, "12": 190.0, "13": 105.0, "14": 215.0, "15": 145.0, "18": 120.0, "19": 240.0, "21": 120.0, "22": 120.0, "25": 120.0, "27": 120.0, "29": 120.0, "30": 120.0}, "6": {"1": 240.0, "4": 120.0, "5": 160.0, "6": 80.0, "8": 165.0, "9": 75.0, "10": 165.0, "11": 75.0, "15": 120.0, "16": 240.0, "22": 110.0, "23": 30.0, "24": 100.0, "29": 120.0, "30": 120.0}}, "5401602XKN02A": {"5": {"4": 240.0, "5": 120.0, "7": 120.0, "8": 240.0, "11": 65.0, "12": 190.0, "13": 105.0, "14": 215.0, "15": 145.0, "18": 120.0, "19": 240.0, "21": 120.0, "22": 120.0, "25": 120.0, "27": 120.0, "29": 120.0, "30": 120.0}, "6": {"1": 240.0, "4": 120.0, "5": 160.0, "6": 80.0, "8": 165.0, "9": 75.0, "10": 165.0, "11": 75.0, "15": 120.0, "16": 240.0, "22": 110.0, "23": 30.0, "24": 100.0, "29": 120.0, "30": 120.0}}, "5401701XKN02A": {"5": {"4": 240.0, "5": 120.0, "7": 120.0, "8": 240.0, "11": 65.0, "12": 190.0, "13": 105.0, "14": 215.0, "15": 145.0, "18": 120.0, "19": 240.0, "21": 120.0, "22": 120.0, "25": 120.0, "27": 120.0, "29": 120.0, "30": 120.0}, "6": {"1": 240.0, "4": 120.0, "5": 160.0, "6": 80.0, "8": 165.0, "9": 75.0, "10": 165.0, "11": 75.0, "15": 120.0, "16": 240.0, "22": 110.0, "23": 30.0, "24": 100.0, "29": 120.0, "30": 120.0}}, "5401702XKN02A": {"5": {"4": 240.0, "5": 120.0, "7": 120.0, "8": 240.0, "11": 65.0, "12": 190.0, "13": 105.0, "14": 215.0, "15": 145.0, "18": 120.0, "19": 240.0, "21": 120.0, "22": 120.0, "25": 120.0, "27": 120.0, "29": 120.0, "30": 120.0}, "6": {"1": 240.0, "4": 120.0, "5": 160.0, "6": 80.0, "8": 165.0, "9": 75.0, "10": 165.0, "11": 75.0, "15": 120.0, "16": 240.0, "22": 110.0, "23": 30.0, "24": 100.0, "29": 120.0, "30": 120.0}}, "5401717XST01A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "5401718XST01A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "5401733XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "5401734XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "5401831XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "5401832XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "5401871XST33A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "5401872XST33A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "5401901XKN02A": {"5": {"4": 240.0, "5": 120.0, "7": 120.0, "8": 240.0, "11": 65.0, "12": 190.0, "13": 105.0, "14": 215.0, "15": 145.0, "18": 120.0, "19": 240.0, "21": 120.0, "22": 120.0, "25": 120.0, "27": 120.0, "29": 120.0, "30": 120.0}, "6": {"1": 240.0, "4": 120.0, "5": 160.0, "6": 80.0, "8": 165.0, "9": 75.0, "10": 165.0, "11": 75.0, "15": 120.0, "16": 240.0, "22": 110.0, "23": 30.0, "24": 100.0, "29": 120.0, "30": 120.0}}, "5401902XKN02A": {"5": {"4": 240.0, "5": 120.0, "7": 120.0, "8": 240.0, "11": 65.0, "12": 190.0, "13": 105.0, "14": 215.0, "15": 145.0, "18": 120.0, "19": 240.0, "21": 120.0, "22": 120.0, "25": 120.0, "27": 120.0, "29": 120.0, "30": 120.0}, "6": {"1": 240.0, "4": 120.0, "5": 160.0, "6": 80.0, "8": 165.0, "9": 75.0, "10": 165.0, "11": 75.0, "15": 120.0, "16": 240.0, "22": 110.0, "23": 30.0, "24": 100.0, "29": 120.0, "30": 120.0}}, "5401907XKN02A": {"5": {"4": 240.0, "5": 120.0, "7": 120.0, "8": 240.0, "11": 65.0, "12": 190.0, "13": 105.0, "14": 215.0, "15": 145.0, "18": 120.0, "19": 240.0, "21": 120.0, "22": 120.0, "25": 120.0, "27": 120.0, "29": 120.0, "30": 120.0}, "6": {"1": 240.0, "4": 120.0, "5": 160.0, "6": 80.0, "8": 165.0, "9": 75.0, "10": 165.0, "11": 75.0, "15": 120.0, "16": 240.0, "22": 110.0, "23": 30.0, "24": 100.0, "29": 120.0, "30": 120.0}}, "5401908XKN02A": {"5": {"4": 240.0, "5": 120.0, "7": 120.0, "8": 240.0, "11": 65.0, "12": 190.0, "13": 105.0, "14": 215.0, "15": 145.0, "18": 120.0, "19": 240.0, "21": 120.0, "22": 120.0, "25": 120.0, "27": 120.0, "29": 120.0, "30": 120.0}, "6": {"1": 240.0, "4": 120.0, "5": 160.0, "6": 80.0, "8": 165.0, "9": 75.0, "10": 165.0, "11": 75.0, "15": 120.0, "16": 240.0, "22": 110.0, "23": 30.0, "24": 100.0, "29": 120.0, "30": 120.0}}, "5402411XKN02B": {"5": {"5": 240.0, "6": 120.0, "7": 240.0, "8": 120.0, "11": 120.0, "12": 195.0, "13": 190.0, "14": 95.0, "15": 225.0, "16": 95.0, "18": 70.0, "19": 210.0, "20": 240.0, "21": 120.0, "22": 120.0, "25": 120.0, "26": 120.0, "28": 145.0, "29": 215.0, "30": 10.0}, "6": {"1": 350.0, "2": 360.0, "3": 120.0, "6": 240.0, "7": 120.0, "8": 240.0, "9": 315.0, "10": 245.0, "11": 235.0, "13": 120.0, "15": 165.0, "16": 120.0, "17": 240.0, "22": 120.0, "23": 120.0, "24": 50.0, "25": 70.0, "30": 120.0}}, "5402422XKN02B": {"5": {"5": 240.0, "6": 120.0, "7": 240.0, "8": 120.0, "11": 120.0, "12": 195.0, "13": 190.0, "14": 95.0, "15": 225.0, "16": 95.0, "18": 70.0, "19": 210.0, "20": 240.0, "21": 120.0, "22": 120.0, "25": 120.0, "26": 120.0, "28": 145.0, "29": 215.0, "30": 10.0}, "6": {"1": 350.0, "2": 360.0, "3": 120.0, "6": 240.0, "7": 120.0, "8": 240.0, "9": 315.0, "10": 245.0, "11": 235.0, "13": 120.0, "15": 165.0, "16": 120.0, "17": 240.0, "22": 120.0, "23": 120.0, "24": 50.0, "25": 70.0, "30": 120.0}}, "5402433XKN02A": {"5": {"5": 240.0, "6": 120.0, "7": 240.0, "8": 120.0, "11": 120.0, "12": 195.0, "13": 190.0, "14": 95.0, "15": 225.0, "16": 95.0, "18": 70.0, "19": 210.0, "20": 240.0, "21": 120.0, "22": 120.0, "25": 120.0, "26": 120.0, "28": 145.0, "29": 215.0, "30": 10.0}, "6": {"1": 350.0, "2": 360.0, "3": 120.0, "6": 240.0, "7": 120.0, "8": 240.0, "9": 315.0, "10": 245.0, "11": 235.0, "13": 120.0, "15": 165.0, "16": 120.0, "17": 240.0, "22": 120.0, "23": 120.0, "24": 50.0, "25": 70.0, "30": 120.0}}, "5402444XKN02A": {"5": {"5": 240.0, "6": 120.0, "7": 240.0, "8": 120.0, "11": 120.0, "12": 195.0, "13": 190.0, "14": 95.0, "15": 225.0, "16": 95.0, "18": 70.0, "19": 210.0, "20": 240.0, "21": 120.0, "22": 120.0, "25": 120.0, "26": 120.0, "28": 145.0, "29": 215.0, "30": 10.0}, "6": {"1": 350.0, "2": 360.0, "3": 120.0, "6": 240.0, "7": 120.0, "8": 240.0, "9": 315.0, "10": 245.0, "11": 235.0, "13": 120.0, "15": 165.0, "16": 120.0, "17": 240.0, "22": 120.0, "23": 120.0, "24": 50.0, "25": 70.0, "30": 120.0}}, "5531015AKN02A": {"5": {"5": 240.0, "6": 120.0, "7": 240.0, "8": 120.0, "11": 120.0, "12": 195.0, "13": 190.0, "14": 95.0, "15": 225.0, "16": 95.0, "18": 70.0, "19": 210.0, "20": 240.0, "21": 120.0, "22": 120.0, "25": 120.0, "26": 120.0, "28": 145.0, "29": 215.0, "30": 10.0}, "6": {"1": 350.0, "2": 360.0, "3": 120.0, "6": 240.0, "7": 120.0, "8": 240.0, "9": 315.0, "10": 245.0, "11": 235.0, "13": 120.0, "15": 165.0, "16": 120.0, "17": 240.0, "22": 120.0, "23": 120.0, "24": 50.0, "25": 70.0, "30": 120.0}}, "5531106XGW01B": {"5": {"4": 120.0, "5": 240.0, "7": 120.0, "8": 120.0, "11": 120.0, "12": 205.0, "13": 155.0, "14": 120.0, "15": 235.0, "16": 5.0, "18": 40.0, "19": 200.0, "20": 120.0, "21": 120.0, "22": 20.0, "23": 100.0, "25": 120.0, "27": 120.0, "29": 45.0, "30": 95.0}, "6": {"1": 325.0, "2": 135.0, "6": 325.0, "8": 155.0, "9": 120.0, "10": 120.0, "11": 120.0, "15": 120.0, "16": 120.0, "23": 120.0, "24": 120.0, "30": 120.0}}, "5531107XGW01B": {"5": {"4": 120.0, "5": 240.0, "7": 120.0, "8": 120.0, "11": 120.0, "12": 205.0, "13": 155.0, "14": 120.0, "15": 235.0, "16": 5.0, "18": 40.0, "19": 200.0, "20": 120.0, "21": 120.0, "22": 20.0, "23": 100.0, "25": 120.0, "27": 120.0, "29": 45.0, "30": 95.0}, "6": {"1": 325.0, "2": 135.0, "6": 325.0, "8": 155.0, "9": 120.0, "10": 120.0, "11": 120.0, "15": 120.0, "16": 120.0, "23": 120.0, "24": 120.0, "30": 120.0}}, "5531109XGW01B": {"5": {"4": 120.0, "5": 240.0, "7": 120.0, "8": 120.0, "11": 120.0, "12": 205.0, "13": 155.0, "14": 120.0, "15": 235.0, "16": 5.0, "18": 40.0, "19": 200.0, "20": 120.0, "21": 120.0, "22": 20.0, "23": 100.0, "25": 120.0, "27": 120.0, "29": 45.0, "30": 95.0}, "6": {"1": 325.0, "2": 135.0, "6": 325.0, "8": 155.0, "9": 120.0, "10": 120.0, "11": 120.0, "15": 120.0, "16": 120.0, "23": 120.0, "24": 120.0, "30": 120.0}}, "5531114XGW01B": {"5": {"4": 120.0, "5": 240.0, "7": 120.0, "8": 120.0, "11": 120.0, "12": 205.0, "13": 155.0, "14": 120.0, "15": 235.0, "16": 5.0, "18": 40.0, "19": 200.0, "20": 120.0, "21": 120.0, "22": 20.0, "23": 100.0, "25": 120.0, "27": 120.0, "29": 45.0, "30": 95.0}, "6": {"1": 325.0, "2": 135.0, "6": 325.0, "8": 155.0, "9": 120.0, "10": 120.0, "11": 120.0, "15": 120.0, "16": 120.0, "23": 120.0, "24": 120.0, "30": 120.0}}, "5531115XGW01B": {"5": {"4": 120.0, "5": 240.0, "7": 120.0, "8": 120.0, "11": 120.0, "12": 205.0, "13": 155.0, "14": 120.0, "15": 235.0, "16": 5.0, "18": 40.0, "19": 200.0, "20": 120.0, "21": 120.0, "22": 20.0, "23": 100.0, "25": 120.0, "27": 120.0, "29": 45.0, "30": 95.0}, "6": {"1": 325.0, "2": 135.0, "6": 325.0, "8": 155.0, "9": 120.0, "10": 120.0, "11": 120.0, "15": 120.0, "16": 120.0, "23": 120.0, "24": 120.0, "30": 120.0}}, "5531117XGW01B": {"5": {"4": 120.0, "5": 240.0, "7": 120.0, "8": 120.0, "11": 120.0, "12": 205.0, "13": 155.0, "14": 120.0, "15": 235.0, "16": 5.0, "18": 40.0, "19": 200.0, "20": 120.0, "21": 120.0, "22": 20.0, "23": 100.0, "25": 120.0, "27": 120.0, "29": 45.0, "30": 95.0}, "6": {"1": 325.0, "2": 135.0, "6": 325.0, "8": 155.0, "9": 120.0, "10": 120.0, "11": 120.0, "15": 120.0, "16": 120.0, "23": 120.0, "24": 120.0, "30": 120.0}}, "5531118XGW01B": {"5": {"4": 120.0, "5": 240.0, "7": 120.0, "8": 120.0, "11": 120.0, "12": 205.0, "13": 155.0, "14": 120.0, "15": 235.0, "16": 5.0, "18": 40.0, "19": 200.0, "20": 120.0, "21": 120.0, "22": 20.0, "23": 100.0, "25": 120.0, "27": 120.0, "29": 45.0, "30": 95.0}, "6": {"1": 325.0, "2": 135.0, "6": 325.0, "8": 155.0, "9": 120.0, "10": 120.0, "11": 120.0, "15": 120.0, "16": 120.0, "23": 120.0, "24": 120.0, "30": 120.0}}, "5531121XST11A": {"5": {"4": 240.0, "5": 230.0, "6": 250.0, "7": 250.0, "8": 300.0, "11": 185.0, "12": 225.0, "13": 275.0, "14": 325.0, "15": 240.0, "16": 240.0, "18": 240.0, "19": 240.0, "20": 240.0, "21": 250.0, "22": 470.0, "23": 120.0, "25": 365.0, "26": 475.0, "27": 325.0, "28": 370.0, "29": 265.0, "30": 120.0}, "6": {"1": 495.0, "2": 685.0, "3": 470.0, "4": 595.0, "5": 350.0, "6": 250.0, "7": 360.0, "8": 715.0, "9": 625.0, "10": 695.0, "11": 485.0, "15": 295.0, "16": 425.0, "17": 475.0, "18": 390.0, "19": 455.0, "20": 150.0, "22": 450.0, "23": 390.0, "24": 450.0, "25": 360.0, "26": 360.0, "27": 120.0, "29": 305.0, "30": 450.0}}, "5531122XST11A": {"5": {"4": 240.0, "5": 230.0, "6": 250.0, "7": 250.0, "8": 300.0, "11": 185.0, "12": 225.0, "13": 275.0, "14": 325.0, "15": 240.0, "16": 240.0, "18": 240.0, "19": 240.0, "20": 240.0, "21": 250.0, "22": 470.0, "23": 120.0, "25": 365.0, "26": 475.0, "27": 325.0, "28": 370.0, "29": 265.0, "30": 120.0}, "6": {"1": 495.0, "2": 685.0, "3": 470.0, "4": 595.0, "5": 350.0, "6": 250.0, "7": 360.0, "8": 715.0, "9": 625.0, "10": 695.0, "11": 485.0, "15": 295.0, "16": 425.0, "17": 475.0, "18": 390.0, "19": 455.0, "20": 150.0, "22": 450.0, "23": 390.0, "24": 450.0, "25": 360.0, "26": 360.0, "27": 120.0, "29": 305.0, "30": 450.0}}, "5531123XST11A": {"5": {"4": 240.0, "5": 230.0, "6": 250.0, "7": 250.0, "8": 300.0, "11": 185.0, "12": 225.0, "13": 275.0, "14": 325.0, "15": 240.0, "16": 240.0, "18": 240.0, "19": 240.0, "20": 240.0, "21": 250.0, "22": 470.0, "23": 120.0, "25": 365.0, "26": 475.0, "27": 325.0, "28": 370.0, "29": 265.0, "30": 120.0}, "6": {"1": 495.0, "2": 685.0, "3": 470.0, "4": 595.0, "5": 350.0, "6": 250.0, "7": 360.0, "8": 715.0, "9": 625.0, "10": 695.0, "11": 485.0, "15": 295.0, "16": 425.0, "17": 475.0, "18": 390.0, "19": 455.0, "20": 150.0, "22": 450.0, "23": 390.0, "24": 450.0, "25": 360.0, "26": 360.0, "27": 120.0, "29": 305.0, "30": 450.0}}, "5531131XST11A": {"5": {"4": 240.0, "5": 230.0, "6": 250.0, "7": 250.0, "8": 300.0, "11": 185.0, "12": 225.0, "13": 275.0, "14": 325.0, "15": 240.0, "16": 240.0, "18": 240.0, "19": 240.0, "20": 240.0, "21": 250.0, "22": 470.0, "23": 120.0, "25": 365.0, "26": 475.0, "27": 325.0, "28": 370.0, "29": 265.0, "30": 120.0}, "6": {"1": 495.0, "2": 685.0, "3": 470.0, "4": 595.0, "5": 350.0, "6": 250.0, "7": 360.0, "8": 715.0, "9": 625.0, "10": 695.0, "11": 485.0, "15": 295.0, "16": 425.0, "17": 475.0, "18": 390.0, "19": 455.0, "20": 150.0, "22": 450.0, "23": 390.0, "24": 450.0, "25": 360.0, "26": 360.0, "27": 120.0, "29": 305.0, "30": 450.0}}, "5531132XST11A": {"5": {"4": 240.0, "5": 230.0, "6": 250.0, "7": 250.0, "8": 300.0, "11": 185.0, "12": 225.0, "13": 275.0, "14": 325.0, "15": 240.0, "16": 240.0, "18": 240.0, "19": 240.0, "20": 240.0, "21": 250.0, "22": 470.0, "23": 120.0, "25": 365.0, "26": 475.0, "27": 325.0, "28": 370.0, "29": 265.0, "30": 120.0}, "6": {"1": 495.0, "2": 685.0, "3": 470.0, "4": 595.0, "5": 350.0, "6": 250.0, "7": 360.0, "8": 715.0, "9": 625.0, "10": 695.0, "11": 485.0, "15": 295.0, "16": 425.0, "17": 475.0, "18": 390.0, "19": 455.0, "20": 150.0, "22": 450.0, "23": 390.0, "24": 450.0, "25": 360.0, "26": 360.0, "27": 120.0, "29": 305.0, "30": 450.0}}, "5531133XST11A": {"5": {"4": 240.0, "5": 230.0, "6": 250.0, "7": 250.0, "8": 300.0, "11": 185.0, "12": 225.0, "13": 275.0, "14": 325.0, "15": 240.0, "16": 240.0, "18": 240.0, "19": 240.0, "20": 240.0, "21": 250.0, "22": 470.0, "23": 120.0, "25": 365.0, "26": 475.0, "27": 325.0, "28": 370.0, "29": 265.0, "30": 120.0}, "6": {"1": 495.0, "2": 685.0, "3": 470.0, "4": 595.0, "5": 350.0, "6": 250.0, "7": 360.0, "8": 715.0, "9": 625.0, "10": 695.0, "11": 485.0, "15": 295.0, "16": 425.0, "17": 475.0, "18": 390.0, "19": 455.0, "20": 150.0, "22": 450.0, "23": 390.0, "24": 450.0, "25": 360.0, "26": 360.0, "27": 120.0, "29": 305.0, "30": 450.0}}, "5531160XKN01B": {"5": {"4": 120.0, "5": 240.0, "7": 120.0, "8": 120.0, "11": 120.0, "12": 205.0, "13": 155.0, "14": 120.0, "15": 235.0, "16": 5.0, "18": 40.0, "19": 200.0, "20": 120.0, "21": 120.0, "22": 20.0, "23": 100.0, "25": 120.0, "27": 120.0, "29": 45.0, "30": 95.0}, "6": {"1": 325.0, "2": 135.0, "6": 325.0, "8": 155.0, "9": 120.0, "10": 120.0, "11": 120.0, "15": 120.0, "16": 120.0, "23": 120.0, "24": 120.0, "30": 120.0}}, "5531161XKN01B": {"5": {"4": 120.0, "5": 240.0, "7": 120.0, "8": 120.0, "11": 120.0, "12": 205.0, "13": 155.0, "14": 120.0, "15": 235.0, "16": 5.0, "18": 40.0, "19": 200.0, "20": 120.0, "21": 120.0, "22": 20.0, "23": 100.0, "25": 120.0, "27": 120.0, "29": 45.0, "30": 95.0}, "6": {"1": 325.0, "2": 135.0, "6": 325.0, "8": 155.0, "9": 120.0, "10": 120.0, "11": 120.0, "15": 120.0, "16": 120.0, "23": 120.0, "24": 120.0, "30": 120.0}}, "5531311XST11A": {"5": {"4": 240.0, "5": 230.0, "6": 250.0, "7": 250.0, "8": 300.0, "11": 185.0, "12": 225.0, "13": 275.0, "14": 325.0, "15": 240.0, "16": 240.0, "18": 240.0, "19": 240.0, "20": 240.0, "21": 250.0, "22": 470.0, "23": 120.0, "25": 365.0, "26": 475.0, "27": 325.0, "28": 370.0, "29": 265.0, "30": 120.0}, "6": {"1": 495.0, "2": 685.0, "3": 470.0, "4": 595.0, "5": 350.0, "6": 250.0, "7": 360.0, "8": 715.0, "9": 625.0, "10": 695.0, "11": 485.0, "15": 295.0, "16": 425.0, "17": 475.0, "18": 390.0, "19": 455.0, "20": 150.0, "22": 450.0, "23": 390.0, "24": 450.0, "25": 360.0, "26": 360.0, "27": 120.0, "29": 305.0, "30": 450.0}}, "5531313XST11A": {"5": {"4": 240.0, "5": 230.0, "6": 250.0, "7": 250.0, "8": 300.0, "11": 185.0, "12": 225.0, "13": 275.0, "14": 325.0, "15": 240.0, "16": 240.0, "18": 240.0, "19": 240.0, "20": 240.0, "21": 250.0, "22": 470.0, "23": 120.0, "25": 365.0, "26": 475.0, "27": 325.0, "28": 370.0, "29": 265.0, "30": 120.0}, "6": {"1": 495.0, "2": 685.0, "3": 470.0, "4": 595.0, "5": 350.0, "6": 250.0, "7": 360.0, "8": 715.0, "9": 625.0, "10": 695.0, "11": 485.0, "15": 295.0, "16": 425.0, "17": 475.0, "18": 390.0, "19": 455.0, "20": 150.0, "22": 450.0, "23": 390.0, "24": 450.0, "25": 360.0, "26": 360.0, "27": 120.0, "29": 305.0, "30": 450.0}}, "5601115xkn02a": {"5": {"4": 240.0, "5": 120.0, "7": 120.0, "8": 240.0, "11": 65.0, "12": 190.0, "13": 105.0, "14": 215.0, "15": 145.0, "18": 120.0, "19": 240.0, "21": 120.0, "22": 120.0, "25": 120.0, "27": 120.0, "29": 120.0, "30": 120.0}, "6": {"1": 240.0, "4": 120.0, "5": 160.0, "6": 80.0, "8": 165.0, "9": 75.0, "10": 165.0, "11": 75.0, "15": 120.0, "16": 240.0, "22": 110.0, "23": 30.0, "24": 100.0, "29": 120.0, "30": 120.0}}, "5601117xkn02a": {"5": {"4": 240.0, "5": 120.0, "7": 120.0, "8": 240.0, "11": 65.0, "12": 190.0, "13": 105.0, "14": 215.0, "15": 145.0, "18": 120.0, "19": 240.0, "21": 120.0, "22": 120.0, "25": 120.0, "27": 120.0, "29": 120.0, "30": 120.0}, "6": {"1": 240.0, "4": 120.0, "5": 160.0, "6": 80.0, "8": 165.0, "9": 75.0, "10": 165.0, "11": 75.0, "15": 120.0, "16": 240.0, "22": 110.0, "23": 30.0, "24": 100.0, "29": 120.0, "30": 120.0}}, "5601120XKN02A": {"5": {"4": 240.0, "5": 120.0, "7": 120.0, "8": 240.0, "11": 65.0, "12": 190.0, "13": 105.0, "14": 215.0, "15": 145.0, "18": 120.0, "19": 240.0, "21": 120.0, "22": 120.0, "25": 120.0, "27": 120.0, "29": 120.0, "30": 120.0}, "6": {"1": 240.0, "4": 120.0, "5": 160.0, "6": 80.0, "8": 165.0, "9": 75.0, "10": 165.0, "11": 75.0, "15": 120.0, "16": 240.0, "22": 110.0, "23": 30.0, "24": 100.0, "29": 120.0, "30": 120.0}}, "5601123XKN02A": {"5": {"4": 240.0, "5": 120.0, "7": 120.0, "8": 240.0, "11": 65.0, "12": 190.0, "13": 105.0, "14": 215.0, "15": 145.0, "18": 120.0, "19": 240.0, "21": 120.0, "22": 120.0, "25": 120.0, "27": 120.0, "29": 120.0, "30": 120.0}, "6": {"1": 240.0, "4": 120.0, "5": 160.0, "6": 80.0, "8": 165.0, "9": 75.0, "10": 165.0, "11": 75.0, "15": 120.0, "16": 240.0, "22": 110.0, "23": 30.0, "24": 100.0, "29": 120.0, "30": 120.0}}, "5604100XKN46A": {"5": {"4": 30.0, "5": 90.0, "6": 30.0, "7": 90.0, "8": 80.0, "11": 75.0, "12": 90.0, "13": 5.0, "14": 55.0, "15": 30.0, "16": 35.0, "18": 40.0, "19": 80.0, "21": 60.0, "22": 40.0, "23": 20.0, "26": 60.0, "27": 55.0, "28": 65.0, "29": 55.0, "30": 45.0}, "6": {"1": 115.0, "2": 5.0, "3": 110.0, "4": 45.0, "5": 10.0, "6": 125.0, "7": 20.0, "8": 200.0, "9": 200.0, "10": 30.0, "11": 125.0, "13": 85.0, "14": 15.0, "15": 200.0, "16": 90.0, "17": 85.0, "18": 90.0, "19": 65.0, "20": 10.0, "22": 90.0, "23": 90.0, "24": 50.0, "25": 70.0, "26": 90.0, "27": 45.0, "29": 35.0, "30": 85.0}}, "5604100XKN83A": {"5": {"4": 60.0, "6": 60.0, "13": 85.0, "14": 35.0, "15": 60.0, "16": 10.0, "18": 50.0, "19": 10.0, "20": 90.0, "21": 30.0, "22": 50.0, "23": 25.0, "25": 90.0, "26": 30.0, "27": 35.0, "28": 25.0, "29": 35.0}, "6": {"1": 85.0, "2": 195.0, "3": 90.0, "4": 90.0, "5": 80.0, "6": 30.0, "7": 90.0, "10": 170.0, "11": 10.0, "13": 25.0, "14": 95.0, "16": 110.0, "17": 65.0, "19": 25.0, "20": 35.0, "24": 40.0, "25": 20.0, "29": 55.0, "30": 5.0}}, "5604101AKN02A": {"5": {"4": 120.0, "5": 240.0, "6": 120.0, "7": 240.0, "8": 120.0, "11": 120.0, "12": 205.0, "13": 155.0, "14": 120.0, "15": 235.0, "16": 5.0, "18": 160.0, "19": 200.0, "20": 180.0, "21": 180.0, "22": 20.0, "23": 100.0, "25": 120.0, "27": 120.0, "28": 120.0, "29": 165.0, "30": 95.0}, "6": {"1": 445.0, "2": 255.0, "5": 120.0, "6": 325.0, "7": 120.0, "8": 165.0, "9": 315.0, "10": 275.0, "11": 265.0, "15": 215.0, "16": 205.0, "17": 35.0, "20": 120.0, "23": 120.0, "24": 120.0, "30": 120.0}}, "5604101XKN46A": {"5": {"4": 30.0, "5": 90.0, "6": 30.0, "7": 90.0, "8": 80.0, "11": 75.0, "12": 90.0, "13": 5.0, "14": 55.0, "15": 30.0, "16": 35.0, "18": 40.0, "19": 80.0, "21": 60.0, "22": 40.0, "23": 20.0, "26": 60.0, "27": 55.0, "28": 65.0, "29": 55.0, "30": 45.0}, "6": {"1": 115.0, "2": 5.0, "3": 110.0, "4": 45.0, "5": 10.0, "6": 125.0, "7": 20.0, "8": 200.0, "9": 200.0, "10": 30.0, "11": 125.0, "13": 85.0, "14": 15.0, "15": 200.0, "16": 90.0, "17": 85.0, "18": 90.0, "19": 65.0, "20": 10.0, "22": 90.0, "23": 90.0, "24": 50.0, "25": 70.0, "26": 90.0, "27": 45.0, "29": 35.0, "30": 85.0}}, "5604101XST11A": {"5": {"4": 240.0, "5": 220.0, "6": 260.0, "7": 240.0, "8": 300.0, "11": 180.0, "12": 240.0, "13": 240.0, "14": 360.0, "15": 240.0, "16": 120.0, "18": 360.0, "19": 240.0, "20": 230.0, "21": 250.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 365.0, "27": 435.0, "28": 285.0, "29": 275.0, "30": 205.0}, "6": {"1": 710.0, "2": 510.0, "3": 655.0, "4": 390.0, "5": 390.0, "6": 375.0, "7": 310.0, "8": 640.0, "9": 625.0, "10": 695.0, "11": 490.0, "13": 125.0, "15": 405.0, "16": 330.0, "17": 450.0, "18": 390.0, "19": 465.0, "20": 240.0, "22": 350.0, "23": 390.0, "24": 460.0, "25": 380.0, "26": 340.0, "27": 205.0, "29": 275.0, "30": 325.0}}, "5604102AKN02A": {"5": {"4": 120.0, "5": 240.0, "6": 120.0, "7": 240.0, "8": 120.0, "11": 120.0, "12": 205.0, "13": 155.0, "14": 120.0, "15": 235.0, "16": 5.0, "18": 160.0, "19": 200.0, "20": 180.0, "21": 180.0, "22": 20.0, "23": 100.0, "25": 120.0, "27": 120.0, "28": 120.0, "29": 165.0, "30": 95.0}, "6": {"1": 445.0, "2": 255.0, "5": 120.0, "6": 325.0, "7": 120.0, "8": 165.0, "9": 315.0, "10": 275.0, "11": 265.0, "15": 215.0, "16": 205.0, "17": 35.0, "20": 120.0, "23": 120.0, "24": 120.0, "30": 120.0}}, "5604102XKN46A": {"5": {"4": 30.0, "5": 90.0, "6": 30.0, "7": 90.0, "8": 80.0, "11": 75.0, "12": 90.0, "13": 5.0, "14": 55.0, "15": 30.0, "16": 35.0, "18": 40.0, "19": 80.0, "21": 60.0, "22": 40.0, "23": 20.0, "26": 60.0, "27": 55.0, "28": 65.0, "29": 55.0, "30": 45.0}, "6": {"1": 115.0, "2": 5.0, "3": 110.0, "4": 45.0, "5": 10.0, "6": 125.0, "7": 20.0, "8": 200.0, "9": 200.0, "10": 30.0, "11": 125.0, "13": 85.0, "14": 15.0, "15": 200.0, "16": 90.0, "17": 85.0, "18": 90.0, "19": 65.0, "20": 10.0, "22": 90.0, "23": 90.0, "24": 50.0, "25": 70.0, "26": 90.0, "27": 45.0, "29": 35.0, "30": 85.0}}, "5604105XST11A": {"5": {"4": 240.0, "5": 220.0, "6": 260.0, "7": 240.0, "8": 300.0, "11": 180.0, "12": 240.0, "13": 240.0, "14": 360.0, "15": 240.0, "16": 120.0, "18": 360.0, "19": 240.0, "20": 230.0, "21": 250.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 365.0, "27": 435.0, "28": 285.0, "29": 275.0, "30": 205.0}, "6": {"1": 710.0, "2": 510.0, "3": 655.0, "4": 390.0, "5": 390.0, "6": 375.0, "7": 310.0, "8": 640.0, "9": 625.0, "10": 695.0, "11": 490.0, "13": 125.0, "15": 405.0, "16": 330.0, "17": 450.0, "18": 390.0, "19": 465.0, "20": 240.0, "22": 350.0, "23": 390.0, "24": 460.0, "25": 380.0, "26": 340.0, "27": 205.0, "29": 275.0, "30": 325.0}}, "5604106XST11A": {"5": {"4": 240.0, "5": 220.0, "6": 260.0, "7": 240.0, "8": 300.0, "11": 180.0, "12": 240.0, "13": 240.0, "14": 360.0, "15": 240.0, "16": 120.0, "18": 360.0, "19": 240.0, "20": 230.0, "21": 250.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 365.0, "27": 435.0, "28": 285.0, "29": 275.0, "30": 205.0}, "6": {"1": 710.0, "2": 510.0, "3": 655.0, "4": 390.0, "5": 390.0, "6": 375.0, "7": 310.0, "8": 640.0, "9": 625.0, "10": 695.0, "11": 490.0, "13": 125.0, "15": 405.0, "16": 330.0, "17": 450.0, "18": 390.0, "19": 465.0, "20": 240.0, "22": 350.0, "23": 390.0, "24": 460.0, "25": 380.0, "26": 340.0, "27": 205.0, "29": 275.0, "30": 325.0}}, "5604107AKN02A": {"5": {"4": 120.0, "5": 240.0, "6": 120.0, "7": 240.0, "8": 120.0, "11": 120.0, "12": 205.0, "13": 155.0, "14": 120.0, "15": 235.0, "16": 5.0, "18": 160.0, "19": 200.0, "20": 180.0, "21": 180.0, "22": 20.0, "23": 100.0, "25": 120.0, "27": 120.0, "28": 120.0, "29": 165.0, "30": 95.0}, "6": {"1": 445.0, "2": 255.0, "5": 120.0, "6": 325.0, "7": 120.0, "8": 165.0, "9": 315.0, "10": 275.0, "11": 265.0, "15": 215.0, "16": 205.0, "17": 35.0, "20": 120.0, "23": 120.0, "24": 120.0, "30": 120.0}}, "5604108AKN02A": {"5": {"4": 120.0, "5": 240.0, "6": 120.0, "7": 240.0, "8": 120.0, "11": 120.0, "12": 205.0, "13": 155.0, "14": 120.0, "15": 235.0, "16": 5.0, "18": 160.0, "19": 200.0, "20": 180.0, "21": 180.0, "22": 20.0, "23": 100.0, "25": 120.0, "27": 120.0, "28": 120.0, "29": 165.0, "30": 95.0}, "6": {"1": 445.0, "2": 255.0, "5": 120.0, "6": 325.0, "7": 120.0, "8": 165.0, "9": 315.0, "10": 275.0, "11": 265.0, "15": 215.0, "16": 205.0, "17": 35.0, "20": 120.0, "23": 120.0, "24": 120.0, "30": 120.0}}, "5604300XKJ23A": {"5": {"4": 190.0, "5": 90.0, "6": 170.0, "7": 70.0, "8": 60.0, "11": 65.0, "12": 115.0, "13": 120.0, "14": 95.0, "15": 85.0, "16": 60.0, "18": 120.0, "19": 100.0, "20": 80.0, "21": 180.0, "22": 70.0, "23": 50.0, "25": 60.0, "26": 65.0, "27": 115.0, "28": 120.0, "29": 60.0, "30": 60.0}, "6": {"1": 60.0, "2": 120.0, "3": 60.0, "4": 60.0, "5": 60.0, "6": 60.0, "8": 60.0, "9": 120.0, "10": 60.0}}, "5604400XKJ23A": {"5": {"4": 190.0, "5": 90.0, "6": 170.0, "7": 70.0, "8": 60.0, "11": 65.0, "12": 115.0, "13": 120.0, "14": 95.0, "15": 85.0, "16": 60.0, "18": 120.0, "19": 100.0, "20": 80.0, "21": 180.0, "22": 70.0, "23": 50.0, "25": 60.0, "26": 65.0, "27": 115.0, "28": 120.0, "29": 60.0, "30": 60.0}, "6": {"1": 60.0, "2": 120.0, "3": 60.0, "4": 60.0, "5": 60.0, "6": 60.0, "8": 60.0, "9": 120.0, "10": 60.0}}, "5701102xkn02a": {"5": {"4": 240.0, "5": 120.0, "7": 120.0, "8": 240.0, "11": 65.0, "12": 190.0, "13": 105.0, "14": 215.0, "15": 145.0, "18": 120.0, "19": 240.0, "21": 120.0, "22": 120.0, "25": 120.0, "27": 120.0, "29": 120.0, "30": 120.0}, "6": {"1": 240.0, "4": 120.0, "5": 160.0, "6": 80.0, "8": 165.0, "9": 75.0, "10": 165.0, "11": 75.0, "15": 120.0, "16": 240.0, "22": 110.0, "23": 30.0, "24": 100.0, "29": 120.0, "30": 120.0}}, "5701103xkn02a": {"5": {"4": 240.0, "5": 120.0, "7": 120.0, "8": 240.0, "11": 65.0, "12": 190.0, "13": 105.0, "14": 215.0, "15": 145.0, "18": 120.0, "19": 240.0, "21": 120.0, "22": 120.0, "25": 120.0, "27": 120.0, "29": 120.0, "30": 120.0}, "6": {"1": 240.0, "4": 120.0, "5": 160.0, "6": 80.0, "8": 165.0, "9": 75.0, "10": 165.0, "11": 75.0, "15": 120.0, "16": 240.0, "22": 110.0, "23": 30.0, "24": 100.0, "29": 120.0, "30": 120.0}}, "5701109XKN06A": {"5": {"6": 240.0, "16": 20.0, "18": 100.0, "20": 120.0, "27": 15.0, "28": 225.0}, "6": {"1": 120.0, "2": 120.0, "5": 120.0, "6": 120.0, "8": 120.0, "9": 190.0, "10": 170.0, "11": 120.0, "16": 120.0, "19": 120.0}}, "5701151XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "6101135XKN02A": {"5": {"4": 240.0, "5": 120.0, "7": 120.0, "8": 240.0, "11": 65.0, "12": 190.0, "13": 105.0, "14": 215.0, "15": 145.0, "18": 120.0, "19": 240.0, "21": 120.0, "22": 120.0, "25": 120.0, "27": 120.0, "29": 120.0, "30": 120.0}, "6": {"1": 240.0, "4": 120.0, "5": 160.0, "6": 80.0, "8": 165.0, "9": 75.0, "10": 165.0, "11": 75.0, "15": 120.0, "16": 240.0, "22": 110.0, "23": 30.0, "24": 100.0, "29": 120.0, "30": 120.0}}, "6101136XKN02A": {"5": {"4": 240.0, "5": 120.0, "7": 120.0, "8": 240.0, "11": 65.0, "12": 190.0, "13": 105.0, "14": 215.0, "15": 145.0, "18": 120.0, "19": 240.0, "21": 120.0, "22": 120.0, "25": 120.0, "27": 120.0, "29": 120.0, "30": 120.0}, "6": {"1": 240.0, "4": 120.0, "5": 160.0, "6": 80.0, "8": 165.0, "9": 75.0, "10": 165.0, "11": 75.0, "15": 120.0, "16": 240.0, "22": 110.0, "23": 30.0, "24": 100.0, "29": 120.0, "30": 120.0}}, "6103100AKN02B": {"5": {"5": 240.0, "6": 120.0, "7": 240.0, "8": 120.0, "11": 120.0, "12": 195.0, "13": 190.0, "14": 95.0, "15": 225.0, "16": 95.0, "18": 70.0, "19": 210.0, "20": 240.0, "21": 120.0, "22": 120.0, "25": 120.0, "26": 120.0, "28": 145.0, "29": 215.0, "30": 10.0}, "6": {"1": 350.0, "2": 360.0, "3": 120.0, "6": 240.0, "7": 120.0, "8": 240.0, "9": 315.0, "10": 245.0, "11": 235.0, "13": 120.0, "15": 165.0, "16": 120.0, "17": 240.0, "22": 120.0, "23": 120.0, "24": 50.0, "25": 70.0, "30": 120.0}}, "6103100XKN46A": {"5": {"4": 30.0, "5": 90.0, "6": 30.0, "7": 90.0, "8": 80.0, "11": 75.0, "12": 90.0, "13": 5.0, "14": 55.0, "15": 30.0, "16": 35.0, "18": 40.0, "19": 80.0, "21": 60.0, "22": 40.0, "23": 20.0, "26": 60.0, "27": 55.0, "28": 65.0, "29": 55.0, "30": 45.0}, "6": {"1": 115.0, "2": 5.0, "3": 110.0, "4": 45.0, "5": 10.0, "6": 125.0, "7": 20.0, "8": 200.0, "9": 200.0, "10": 30.0, "11": 125.0, "13": 85.0, "14": 15.0, "15": 200.0, "16": 90.0, "17": 85.0, "18": 90.0, "19": 65.0, "20": 10.0, "22": 90.0, "23": 90.0, "24": 50.0, "25": 70.0, "26": 90.0, "27": 45.0, "29": 35.0, "30": 85.0}}, "6103101XKJ23A": {"5": {"4": 190.0, "5": 90.0, "6": 170.0, "7": 70.0, "8": 60.0, "11": 65.0, "12": 115.0, "13": 120.0, "14": 95.0, "15": 85.0, "16": 60.0, "18": 120.0, "19": 100.0, "20": 80.0, "21": 180.0, "22": 70.0, "23": 50.0, "25": 60.0, "26": 65.0, "27": 115.0, "28": 120.0, "29": 60.0, "30": 60.0}, "6": {"1": 60.0, "2": 120.0, "3": 60.0, "4": 60.0, "5": 60.0, "6": 60.0, "8": 60.0, "9": 120.0, "10": 60.0}}, "6103102XKJ23A": {"5": {"4": 190.0, "5": 90.0, "6": 170.0, "7": 70.0, "8": 60.0, "11": 65.0, "12": 115.0, "13": 120.0, "14": 95.0, "15": 85.0, "16": 60.0, "18": 120.0, "19": 100.0, "20": 80.0, "21": 180.0, "22": 70.0, "23": 50.0, "25": 60.0, "26": 65.0, "27": 115.0, "28": 120.0, "29": 60.0, "30": 60.0}, "6": {"1": 60.0, "2": 120.0, "3": 60.0, "4": 60.0, "5": 60.0, "6": 60.0, "8": 60.0, "9": 120.0, "10": 60.0}}, "6103200AKN02B": {"5": {"5": 240.0, "6": 120.0, "7": 240.0, "8": 120.0, "11": 120.0, "12": 195.0, "13": 190.0, "14": 95.0, "15": 225.0, "16": 95.0, "18": 70.0, "19": 210.0, "20": 240.0, "21": 120.0, "22": 120.0, "25": 120.0, "26": 120.0, "28": 145.0, "29": 215.0, "30": 10.0}, "6": {"1": 350.0, "2": 360.0, "3": 120.0, "6": 240.0, "7": 120.0, "8": 240.0, "9": 315.0, "10": 245.0, "11": 235.0, "13": 120.0, "15": 165.0, "16": 120.0, "17": 240.0, "22": 120.0, "23": 120.0, "24": 50.0, "25": 70.0, "30": 120.0}}, "6103200XKN46A": {"5": {"4": 30.0, "5": 90.0, "6": 30.0, "7": 90.0, "8": 80.0, "11": 75.0, "12": 90.0, "13": 5.0, "14": 55.0, "15": 30.0, "16": 35.0, "18": 40.0, "19": 80.0, "21": 60.0, "22": 40.0, "23": 20.0, "26": 60.0, "27": 55.0, "28": 65.0, "29": 55.0, "30": 45.0}, "6": {"1": 115.0, "2": 5.0, "3": 110.0, "4": 45.0, "5": 10.0, "6": 125.0, "7": 20.0, "8": 200.0, "9": 200.0, "10": 30.0, "11": 125.0, "13": 85.0, "14": 15.0, "15": 200.0, "16": 90.0, "17": 85.0, "18": 90.0, "19": 65.0, "20": 10.0, "22": 90.0, "23": 90.0, "24": 50.0, "25": 70.0, "26": 90.0, "27": 45.0, "29": 35.0, "30": 85.0}}, "6103300XST11A": {"5": {"4": 240.0, "5": 220.0, "6": 260.0, "7": 240.0, "8": 300.0, "11": 180.0, "12": 240.0, "13": 240.0, "14": 360.0, "15": 240.0, "16": 120.0, "18": 360.0, "19": 240.0, "20": 230.0, "21": 250.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 365.0, "27": 435.0, "28": 285.0, "29": 275.0, "30": 205.0}, "6": {"1": 710.0, "2": 510.0, "3": 655.0, "4": 390.0, "5": 390.0, "6": 375.0, "7": 310.0, "8": 640.0, "9": 625.0, "10": 695.0, "11": 490.0, "13": 125.0, "15": 405.0, "16": 330.0, "17": 450.0, "18": 390.0, "19": 465.0, "20": 240.0, "22": 350.0, "23": 390.0, "24": 460.0, "25": 380.0, "26": 340.0, "27": 205.0, "29": 275.0, "30": 325.0}}, "6103400XST11A": {"5": {"4": 240.0, "5": 220.0, "6": 260.0, "7": 240.0, "8": 300.0, "11": 180.0, "12": 240.0, "13": 240.0, "14": 360.0, "15": 240.0, "16": 120.0, "18": 360.0, "19": 240.0, "20": 230.0, "21": 250.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 365.0, "27": 435.0, "28": 285.0, "29": 275.0, "30": 205.0}, "6": {"1": 710.0, "2": 510.0, "3": 655.0, "4": 390.0, "5": 390.0, "6": 375.0, "7": 310.0, "8": 640.0, "9": 625.0, "10": 695.0, "11": 490.0, "13": 125.0, "15": 405.0, "16": 330.0, "17": 450.0, "18": 390.0, "19": 465.0, "20": 240.0, "22": 350.0, "23": 390.0, "24": 460.0, "25": 380.0, "26": 340.0, "27": 205.0, "29": 275.0, "30": 325.0}}, "6107100AKJ20A": {"5": {"4": 190.0, "5": 90.0, "6": 170.0, "7": 70.0, "8": 60.0, "11": 65.0, "12": 115.0, "13": 120.0, "14": 95.0, "15": 85.0, "16": 60.0, "18": 120.0, "19": 100.0, "20": 80.0, "21": 180.0, "22": 70.0, "23": 50.0, "25": 60.0, "26": 65.0, "27": 115.0, "28": 120.0, "29": 60.0, "30": 60.0}, "6": {"1": 60.0, "2": 120.0, "3": 60.0, "4": 60.0, "5": 60.0, "6": 60.0, "8": 60.0, "9": 120.0, "10": 60.0}}, "6107107AKJ20A": {"5": {"4": 190.0, "5": 90.0, "6": 170.0, "7": 70.0, "8": 60.0, "11": 65.0, "12": 115.0, "13": 120.0, "14": 95.0, "15": 85.0, "16": 60.0, "18": 120.0, "19": 100.0, "20": 80.0, "21": 180.0, "22": 70.0, "23": 50.0, "25": 60.0, "26": 65.0, "27": 115.0, "28": 120.0, "29": 60.0, "30": 60.0}, "6": {"1": 60.0, "2": 120.0, "3": 60.0, "4": 60.0, "5": 60.0, "6": 60.0, "8": 60.0, "9": 120.0, "10": 60.0}}, "6107107AST01A": {"5": {"4": 240.0, "5": 220.0, "6": 260.0, "7": 240.0, "8": 300.0, "11": 180.0, "12": 240.0, "13": 240.0, "14": 360.0, "15": 240.0, "16": 120.0, "18": 360.0, "19": 240.0, "20": 230.0, "21": 250.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 365.0, "27": 435.0, "28": 285.0, "29": 275.0, "30": 205.0}, "6": {"1": 710.0, "2": 510.0, "3": 655.0, "4": 390.0, "5": 390.0, "6": 375.0, "7": 310.0, "8": 640.0, "9": 625.0, "10": 695.0, "11": 490.0, "13": 125.0, "15": 405.0, "16": 330.0, "17": 450.0, "18": 390.0, "19": 465.0, "20": 240.0, "22": 350.0, "23": 390.0, "24": 460.0, "25": 380.0, "26": 340.0, "27": 205.0, "29": 275.0, "30": 325.0}}, "6107108AKJ20A": {"5": {"4": 190.0, "5": 90.0, "6": 170.0, "7": 70.0, "8": 60.0, "11": 65.0, "12": 115.0, "13": 120.0, "14": 95.0, "15": 85.0, "16": 60.0, "18": 120.0, "19": 100.0, "20": 80.0, "21": 180.0, "22": 70.0, "23": 50.0, "25": 60.0, "26": 65.0, "27": 115.0, "28": 120.0, "29": 60.0, "30": 60.0}, "6": {"1": 60.0, "2": 120.0, "3": 60.0, "4": 60.0, "5": 60.0, "6": 60.0, "8": 60.0, "9": 120.0, "10": 60.0}}, "6107108AST01A": {"5": {"4": 240.0, "5": 220.0, "6": 260.0, "7": 240.0, "8": 300.0, "11": 180.0, "12": 240.0, "13": 240.0, "14": 360.0, "15": 240.0, "16": 120.0, "18": 360.0, "19": 240.0, "20": 230.0, "21": 250.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 365.0, "27": 435.0, "28": 285.0, "29": 275.0, "30": 205.0}, "6": {"1": 710.0, "2": 510.0, "3": 655.0, "4": 390.0, "5": 390.0, "6": 375.0, "7": 310.0, "8": 640.0, "9": 625.0, "10": 695.0, "11": 490.0, "13": 125.0, "15": 405.0, "16": 330.0, "17": 450.0, "18": 390.0, "19": 465.0, "20": 240.0, "22": 350.0, "23": 390.0, "24": 460.0, "25": 380.0, "26": 340.0, "27": 205.0, "29": 275.0, "30": 325.0}}, "6107200AKJ20A": {"5": {"4": 190.0, "5": 90.0, "6": 170.0, "7": 70.0, "8": 60.0, "11": 65.0, "12": 115.0, "13": 120.0, "14": 95.0, "15": 85.0, "16": 60.0, "18": 120.0, "19": 100.0, "20": 80.0, "21": 180.0, "22": 70.0, "23": 50.0, "25": 60.0, "26": 65.0, "27": 115.0, "28": 120.0, "29": 60.0, "30": 60.0}, "6": {"1": 60.0, "2": 120.0, "3": 60.0, "4": 60.0, "5": 60.0, "6": 60.0, "8": 60.0, "9": 120.0, "10": 60.0}}, "6107300AST01A": {"5": {"4": 240.0, "5": 220.0, "6": 260.0, "7": 240.0, "8": 300.0, "11": 180.0, "12": 240.0, "13": 240.0, "14": 360.0, "15": 240.0, "16": 120.0, "18": 360.0, "19": 240.0, "20": 230.0, "21": 250.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 365.0, "27": 435.0, "28": 285.0, "29": 275.0, "30": 205.0}, "6": {"1": 710.0, "2": 510.0, "3": 655.0, "4": 390.0, "5": 390.0, "6": 375.0, "7": 310.0, "8": 640.0, "9": 625.0, "10": 695.0, "11": 490.0, "13": 125.0, "15": 405.0, "16": 330.0, "17": 450.0, "18": 390.0, "19": 465.0, "20": 240.0, "22": 350.0, "23": 390.0, "24": 460.0, "25": 380.0, "26": 340.0, "27": 205.0, "29": 275.0, "30": 325.0}}, "6107400AST01A": {"5": {"4": 240.0, "5": 220.0, "6": 260.0, "7": 240.0, "8": 300.0, "11": 180.0, "12": 240.0, "13": 240.0, "14": 360.0, "15": 240.0, "16": 120.0, "18": 360.0, "19": 240.0, "20": 230.0, "21": 250.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 365.0, "27": 435.0, "28": 285.0, "29": 275.0, "30": 205.0}, "6": {"1": 710.0, "2": 510.0, "3": 655.0, "4": 390.0, "5": 390.0, "6": 375.0, "7": 310.0, "8": 640.0, "9": 625.0, "10": 695.0, "11": 490.0, "13": 125.0, "15": 405.0, "16": 330.0, "17": 450.0, "18": 390.0, "19": 465.0, "20": 240.0, "22": 350.0, "23": 390.0, "24": 460.0, "25": 380.0, "26": 340.0, "27": 205.0, "29": 275.0, "30": 325.0}}, "6201103XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "6201104XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "6201121XKN02A": {"5": {"4": 240.0, "5": 120.0, "7": 120.0, "8": 240.0, "11": 65.0, "12": 190.0, "13": 105.0, "14": 215.0, "15": 145.0, "18": 120.0, "19": 240.0, "21": 120.0, "22": 120.0, "25": 120.0, "27": 120.0, "29": 120.0, "30": 120.0}, "6": {"1": 240.0, "4": 120.0, "5": 160.0, "6": 80.0, "8": 165.0, "9": 75.0, "10": 165.0, "11": 75.0, "15": 120.0, "16": 240.0, "22": 110.0, "23": 30.0, "24": 100.0, "29": 120.0, "30": 120.0}}, "6201122XKN02A": {"5": {"4": 240.0, "5": 120.0, "7": 120.0, "8": 240.0, "11": 65.0, "12": 190.0, "13": 105.0, "14": 215.0, "15": 145.0, "18": 120.0, "19": 240.0, "21": 120.0, "22": 120.0, "25": 120.0, "27": 120.0, "29": 120.0, "30": 120.0}, "6": {"1": 240.0, "4": 120.0, "5": 160.0, "6": 80.0, "8": 165.0, "9": 75.0, "10": 165.0, "11": 75.0, "15": 120.0, "16": 240.0, "22": 110.0, "23": 30.0, "24": 100.0, "29": 120.0, "30": 120.0}}, "6203100AKN02B": {"5": {"5": 240.0, "6": 120.0, "8": 120.0, "11": 120.0, "12": 195.0, "13": 190.0, "14": 95.0, "15": 225.0, "16": 95.0, "18": 40.0, "19": 120.0, "20": 240.0, "22": 120.0, "25": 120.0, "26": 120.0, "28": 120.0, "30": 10.0}, "6": {"1": 350.0, "2": 240.0, "6": 120.0, "8": 240.0, "9": 195.0, "10": 45.0, "11": 195.0, "15": 45.0, "16": 120.0, "17": 120.0, "23": 120.0, "24": 50.0, "25": 70.0, "30": 120.0}}, "6203100AKN61A": {"5": {"5": 240.0, "6": 120.0, "7": 240.0, "8": 120.0, "11": 120.0, "12": 195.0, "13": 190.0, "14": 95.0, "15": 225.0, "16": 95.0, "18": 70.0, "19": 210.0, "20": 240.0, "21": 120.0, "22": 120.0, "25": 120.0, "26": 120.0, "28": 145.0, "29": 215.0, "30": 10.0}, "6": {"1": 350.0, "2": 360.0, "3": 120.0, "6": 240.0, "7": 120.0, "8": 240.0, "9": 315.0, "10": 245.0, "11": 235.0, "13": 120.0, "15": 165.0, "16": 120.0, "17": 240.0, "22": 120.0, "23": 120.0, "24": 50.0, "25": 70.0, "30": 120.0}}, "6203100AST11A": {"5": {"5": 220.0, "6": 140.0, "7": 120.0, "8": 300.0, "11": 60.0, "12": 120.0, "13": 240.0, "14": 120.0, "15": 120.0, "27": 120.0}, "6": {"1": 40.0, "2": 200.0, "3": 160.0, "4": 200.0, "5": 100.0, "6": 95.0, "7": 230.0, "8": 415.0, "10": 225.0, "11": 250.0, "13": 125.0, "18": 120.0, "22": 120.0, "23": 20.0, "24": 460.0, "25": 260.0, "26": 340.0, "27": 120.0}}, "6203100XKJ23A": {"5": {"4": 190.0, "5": 90.0, "6": 170.0, "7": 70.0, "8": 60.0, "11": 65.0, "12": 115.0, "13": 120.0, "14": 95.0, "15": 85.0, "16": 60.0, "18": 120.0, "19": 100.0, "20": 80.0, "21": 180.0, "22": 70.0, "23": 50.0, "25": 60.0, "26": 65.0, "27": 115.0, "28": 120.0, "29": 60.0, "30": 60.0}, "6": {"1": 60.0, "2": 120.0, "3": 60.0, "4": 60.0, "5": 60.0, "6": 60.0, "8": 60.0, "9": 120.0, "10": 60.0}}, "6203100XKN46A": {"5": {"15": 25.0, "16": 35.0, "18": 40.0, "19": 20.0, "22": 40.0, "23": 20.0, "30": 40.0}, "6": {"1": 20.0, "14": 15.0, "15": 135.0, "16": 90.0, "17": 25.0, "18": 90.0, "19": 5.0}}, "6203100XKN83A": {"5": {"4": 640.0, "5": 640.0, "6": 640.0, "7": 640.0, "8": 560.0, "11": 440.0, "12": 640.0, "13": 640.0, "14": 640.0, "15": 640.0, "16": 320.0, "18": 640.0, "19": 640.0, "20": 640.0, "21": 640.0, "22": 640.0, "23": 320.0, "25": 640.0, "26": 640.0, "27": 640.0, "28": 640.0, "29": 640.0, "30": 320.0}, "6": {"1": 1320.0, "2": 1320.0, "3": 1320.0, "4": 980.0, "5": 660.0, "6": 980.0, "7": 660.0, "8": 1320.0, "9": 1320.0, "10": 1320.0, "11": 980.0, "13": 660.0, "14": 660.0, "15": 1320.0, "16": 1320.0, "17": 1025.0, "18": 660.0, "19": 660.0, "20": 320.0, "22": 660.0, "23": 660.0, "24": 660.0, "25": 660.0, "26": 660.0, "27": 320.0, "29": 660.0, "30": 660.0}}, "6203100XST11A": {"5": {"6": 120.0, "7": 120.0, "11": 120.0, "14": 120.0, "15": 120.0, "25": 115.0, "26": 5.0, "29": 240.0, "30": 120.0}, "6": {"1": 515.0, "2": 85.0, "3": 120.0, "5": 120.0, "6": 160.0, "7": 80.0, "9": 250.0, "10": 230.0, "29": 240.0}}, "6203101XKJ23A": {"5": {"4": 190.0, "5": 90.0, "6": 170.0, "7": 70.0, "8": 60.0, "11": 65.0, "12": 115.0, "13": 120.0, "14": 95.0, "15": 85.0, "16": 60.0, "18": 120.0, "19": 100.0, "20": 80.0, "21": 180.0, "22": 70.0, "23": 50.0, "25": 60.0, "26": 65.0, "27": 115.0, "28": 120.0, "29": 60.0, "30": 60.0}, "6": {"1": 60.0, "2": 120.0, "3": 60.0, "4": 60.0, "5": 60.0, "6": 60.0, "8": 60.0, "9": 120.0, "10": 60.0}}, "6203102XKJ23A": {"5": {"4": 190.0, "5": 90.0, "6": 170.0, "7": 70.0, "8": 60.0, "11": 65.0, "12": 115.0, "13": 120.0, "14": 95.0, "15": 85.0, "16": 60.0, "18": 120.0, "19": 100.0, "20": 80.0, "21": 180.0, "22": 70.0, "23": 50.0, "25": 60.0, "26": 65.0, "27": 115.0, "28": 120.0, "29": 60.0, "30": 60.0}, "6": {"1": 60.0, "2": 120.0, "3": 60.0, "4": 60.0, "5": 60.0, "6": 60.0, "8": 60.0, "9": 120.0, "10": 60.0}}, "6203200AKN02B": {"5": {"5": 240.0, "6": 120.0, "8": 120.0, "11": 120.0, "12": 195.0, "13": 190.0, "14": 95.0, "15": 225.0, "16": 95.0, "18": 40.0, "19": 120.0, "20": 240.0, "22": 120.0, "25": 120.0, "26": 120.0, "28": 120.0, "30": 10.0}, "6": {"1": 350.0, "2": 240.0, "6": 120.0, "8": 240.0, "9": 195.0, "10": 45.0, "11": 195.0, "15": 45.0, "16": 120.0, "17": 120.0, "23": 120.0, "24": 50.0, "25": 70.0, "30": 120.0}}, "6203200AKN61A": {"5": {"5": 240.0, "6": 120.0, "7": 240.0, "8": 120.0, "11": 120.0, "12": 195.0, "13": 190.0, "14": 95.0, "15": 225.0, "16": 95.0, "18": 70.0, "19": 210.0, "20": 240.0, "21": 120.0, "22": 120.0, "25": 120.0, "26": 120.0, "28": 145.0, "29": 215.0, "30": 10.0}, "6": {"1": 350.0, "2": 360.0, "3": 120.0, "6": 240.0, "7": 120.0, "8": 240.0, "9": 315.0, "10": 245.0, "11": 235.0, "13": 120.0, "15": 165.0, "16": 120.0, "17": 240.0, "22": 120.0, "23": 120.0, "24": 50.0, "25": 70.0, "30": 120.0}}, "6203200AST11A": {"5": {"5": 220.0, "6": 140.0, "7": 120.0, "8": 300.0, "11": 60.0, "12": 120.0, "13": 240.0, "14": 120.0, "15": 120.0, "27": 120.0}, "6": {"1": 40.0, "2": 200.0, "3": 160.0, "4": 200.0, "5": 100.0, "6": 95.0, "7": 230.0, "8": 415.0, "10": 225.0, "11": 250.0, "13": 125.0, "18": 120.0, "22": 120.0, "23": 20.0, "24": 460.0, "25": 260.0, "26": 340.0, "27": 120.0}}, "6203200XKJ23A": {"5": {"4": 190.0, "5": 90.0, "6": 170.0, "7": 70.0, "8": 60.0, "11": 65.0, "12": 115.0, "13": 120.0, "14": 95.0, "15": 85.0, "16": 60.0, "18": 120.0, "19": 100.0, "20": 80.0, "21": 180.0, "22": 70.0, "23": 50.0, "25": 60.0, "26": 65.0, "27": 115.0, "28": 120.0, "29": 60.0, "30": 60.0}, "6": {"1": 60.0, "2": 120.0, "3": 60.0, "4": 60.0, "5": 60.0, "6": 60.0, "8": 60.0, "9": 120.0, "10": 60.0}}, "6203200XKN46A": {"5": {"15": 25.0, "16": 35.0, "18": 40.0, "19": 20.0, "22": 40.0, "23": 20.0, "30": 40.0}, "6": {"1": 20.0, "14": 15.0, "15": 135.0, "16": 90.0, "17": 25.0, "18": 90.0, "19": 5.0}}, "6203200XKN83A": {"5": {"4": 640.0, "5": 640.0, "6": 640.0, "7": 640.0, "8": 560.0, "11": 440.0, "12": 640.0, "13": 640.0, "14": 640.0, "15": 640.0, "16": 320.0, "18": 640.0, "19": 640.0, "20": 640.0, "21": 640.0, "22": 640.0, "23": 320.0, "25": 640.0, "26": 640.0, "27": 640.0, "28": 640.0, "29": 640.0, "30": 320.0}, "6": {"1": 1320.0, "2": 1320.0, "3": 1320.0, "4": 980.0, "5": 660.0, "6": 980.0, "7": 660.0, "8": 1320.0, "9": 1320.0, "10": 1320.0, "11": 980.0, "13": 660.0, "14": 660.0, "15": 1320.0, "16": 1320.0, "17": 1025.0, "18": 660.0, "19": 660.0, "20": 320.0, "22": 660.0, "23": 660.0, "24": 660.0, "25": 660.0, "26": 660.0, "27": 320.0, "29": 660.0, "30": 660.0}}, "6203200XST11A": {"5": {"6": 120.0, "7": 120.0, "11": 120.0, "14": 120.0, "15": 120.0, "25": 115.0, "26": 5.0, "29": 240.0, "30": 120.0}, "6": {"1": 515.0, "2": 85.0, "3": 120.0, "5": 120.0, "6": 160.0, "7": 80.0, "9": 250.0, "10": 230.0, "29": 240.0}}, "6203300XKN46A": {"5": {"4": 120.0, "5": 640.0, "6": 400.0, "7": 520.0, "8": 500.0, "11": 255.0, "12": 405.0, "13": 580.0, "14": 400.0, "15": 270.0, "16": 70.0, "18": 380.0, "19": 160.0, "20": 390.0, "21": 330.0, "22": 300.0, "23": 75.0, "25": 210.0, "26": 510.0, "27": 450.0, "28": 535.0, "29": 365.0, "30": 65.0}, "6": {"1": 435.0, "2": 1105.0, "3": 850.0, "4": 680.0, "5": 540.0, "6": 760.0, "7": 580.0, "8": 1260.0, "9": 770.0, "10": 1090.0, "11": 620.0, "13": 660.0, "14": 510.0, "15": 350.0, "16": 485.0, "17": 750.0, "18": 240.0, "19": 190.0, "20": 80.0, "22": 430.0, "23": 230.0, "24": 600.0, "25": 420.0, "26": 490.0, "27": 310.0, "29": 245.0, "30": 535.0}}, "6203300XKN83A": {"5": {"4": 640.0, "5": 640.0, "6": 640.0, "7": 640.0, "8": 560.0, "11": 440.0, "12": 640.0, "13": 640.0, "14": 640.0, "15": 640.0, "16": 320.0, "18": 640.0, "19": 640.0, "20": 640.0, "21": 640.0, "22": 640.0, "23": 320.0, "25": 640.0, "26": 640.0, "27": 640.0, "28": 640.0, "29": 640.0, "30": 320.0}, "6": {"1": 1320.0, "2": 1320.0, "3": 1320.0, "4": 980.0, "5": 660.0, "6": 980.0, "7": 660.0, "8": 1320.0, "9": 1320.0, "10": 1320.0, "11": 980.0, "13": 660.0, "14": 660.0, "15": 1320.0, "16": 1320.0, "17": 1025.0, "18": 660.0, "19": 660.0, "20": 320.0, "22": 660.0, "23": 660.0, "24": 660.0, "25": 660.0, "26": 660.0, "27": 320.0, "29": 660.0, "30": 660.0}}, "6203400XKN46A": {"5": {"4": 120.0, "5": 640.0, "6": 400.0, "7": 520.0, "8": 500.0, "11": 255.0, "12": 405.0, "13": 580.0, "14": 400.0, "15": 270.0, "16": 70.0, "18": 380.0, "19": 160.0, "20": 390.0, "21": 330.0, "22": 300.0, "23": 75.0, "25": 210.0, "26": 510.0, "27": 450.0, "28": 535.0, "29": 365.0, "30": 65.0}, "6": {"1": 435.0, "2": 1105.0, "3": 850.0, "4": 680.0, "5": 540.0, "6": 760.0, "7": 580.0, "8": 1260.0, "9": 770.0, "10": 1090.0, "11": 620.0, "13": 660.0, "14": 510.0, "15": 350.0, "16": 485.0, "17": 750.0, "18": 240.0, "19": 190.0, "20": 80.0, "22": 430.0, "23": 230.0, "24": 600.0, "25": 420.0, "26": 490.0, "27": 310.0, "29": 245.0, "30": 535.0}}, "6203400XKN83A": {"5": {"4": 640.0, "5": 640.0, "6": 640.0, "7": 640.0, "8": 560.0, "11": 440.0, "12": 640.0, "13": 640.0, "14": 640.0, "15": 640.0, "16": 320.0, "18": 640.0, "19": 640.0, "20": 640.0, "21": 640.0, "22": 640.0, "23": 320.0, "25": 640.0, "26": 640.0, "27": 640.0, "28": 640.0, "29": 640.0, "30": 320.0}, "6": {"1": 1320.0, "2": 1320.0, "3": 1320.0, "4": 980.0, "5": 660.0, "6": 980.0, "7": 660.0, "8": 1320.0, "9": 1320.0, "10": 1320.0, "11": 980.0, "13": 660.0, "14": 660.0, "15": 1320.0, "16": 1320.0, "17": 1025.0, "18": 660.0, "19": 660.0, "20": 320.0, "22": 660.0, "23": 660.0, "24": 660.0, "25": 660.0, "26": 660.0, "27": 320.0, "29": 660.0, "30": 660.0}}, "6207100AKJ20A": {"5": {"4": 190.0, "5": 90.0, "6": 170.0, "7": 70.0, "8": 60.0, "11": 65.0, "12": 115.0, "13": 120.0, "14": 95.0, "15": 85.0, "16": 60.0, "18": 120.0, "19": 100.0, "20": 80.0, "21": 180.0, "22": 70.0, "23": 50.0, "25": 60.0, "26": 65.0, "27": 115.0, "28": 120.0, "29": 60.0, "30": 60.0}, "6": {"1": 60.0, "2": 120.0, "3": 60.0, "4": 60.0, "5": 60.0, "6": 60.0, "8": 60.0, "9": 120.0, "10": 60.0}}, "6207100AST01A": {"5": {"4": 240.0, "5": 220.0, "6": 260.0, "7": 240.0, "8": 300.0, "11": 180.0, "12": 240.0, "13": 240.0, "14": 360.0, "15": 240.0, "16": 120.0, "18": 360.0, "19": 240.0, "20": 230.0, "21": 250.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 365.0, "27": 435.0, "28": 285.0, "29": 275.0, "30": 205.0}, "6": {"1": 710.0, "2": 510.0, "3": 655.0, "4": 390.0, "5": 390.0, "6": 375.0, "7": 310.0, "8": 640.0, "9": 625.0, "10": 695.0, "11": 490.0, "13": 125.0, "15": 405.0, "16": 330.0, "17": 450.0, "18": 390.0, "19": 465.0, "20": 240.0, "22": 350.0, "23": 390.0, "24": 460.0, "25": 380.0, "26": 340.0, "27": 205.0, "29": 275.0, "30": 325.0}}, "6207107AKJ20A": {"5": {"4": 190.0, "5": 90.0, "6": 170.0, "7": 70.0, "8": 60.0, "11": 65.0, "12": 115.0, "13": 120.0, "14": 95.0, "15": 85.0, "16": 60.0, "18": 120.0, "19": 100.0, "20": 80.0, "21": 180.0, "22": 70.0, "23": 50.0, "25": 60.0, "26": 65.0, "27": 115.0, "28": 120.0, "29": 60.0, "30": 60.0}, "6": {"1": 60.0, "2": 120.0, "3": 60.0, "4": 60.0, "5": 60.0, "6": 60.0, "8": 60.0, "9": 120.0, "10": 60.0}}, "6207107AST01A": {"5": {"4": 240.0, "5": 220.0, "6": 260.0, "7": 240.0, "8": 300.0, "11": 180.0, "12": 240.0, "13": 240.0, "14": 360.0, "15": 240.0, "16": 120.0, "18": 360.0, "19": 240.0, "20": 230.0, "21": 250.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 365.0, "27": 435.0, "28": 285.0, "29": 275.0, "30": 205.0}, "6": {"1": 710.0, "2": 510.0, "3": 655.0, "4": 390.0, "5": 390.0, "6": 375.0, "7": 310.0, "8": 640.0, "9": 625.0, "10": 695.0, "11": 490.0, "13": 125.0, "15": 405.0, "16": 330.0, "17": 450.0, "18": 390.0, "19": 465.0, "20": 240.0, "22": 350.0, "23": 390.0, "24": 460.0, "25": 380.0, "26": 340.0, "27": 205.0, "29": 275.0, "30": 325.0}}, "6207108AKJ20A": {"5": {"4": 190.0, "5": 90.0, "6": 170.0, "7": 70.0, "8": 60.0, "11": 65.0, "12": 115.0, "13": 120.0, "14": 95.0, "15": 85.0, "16": 60.0, "18": 120.0, "19": 100.0, "20": 80.0, "21": 180.0, "22": 70.0, "23": 50.0, "25": 60.0, "26": 65.0, "27": 115.0, "28": 120.0, "29": 60.0, "30": 60.0}, "6": {"1": 60.0, "2": 120.0, "3": 60.0, "4": 60.0, "5": 60.0, "6": 60.0, "8": 60.0, "9": 120.0, "10": 60.0}}, "6207108AST01A": {"5": {"4": 240.0, "5": 220.0, "6": 260.0, "7": 240.0, "8": 300.0, "11": 180.0, "12": 240.0, "13": 240.0, "14": 360.0, "15": 240.0, "16": 120.0, "18": 360.0, "19": 240.0, "20": 230.0, "21": 250.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 365.0, "27": 435.0, "28": 285.0, "29": 275.0, "30": 205.0}, "6": {"1": 710.0, "2": 510.0, "3": 655.0, "4": 390.0, "5": 390.0, "6": 375.0, "7": 310.0, "8": 640.0, "9": 625.0, "10": 695.0, "11": 490.0, "13": 125.0, "15": 405.0, "16": 330.0, "17": 450.0, "18": 390.0, "19": 465.0, "20": 240.0, "22": 350.0, "23": 390.0, "24": 460.0, "25": 380.0, "26": 340.0, "27": 205.0, "29": 275.0, "30": 325.0}}, "6207200AKJ20A": {"5": {"4": 190.0, "5": 90.0, "6": 170.0, "7": 70.0, "8": 60.0, "11": 65.0, "12": 115.0, "13": 120.0, "14": 95.0, "15": 85.0, "16": 60.0, "18": 120.0, "19": 100.0, "20": 80.0, "21": 180.0, "22": 70.0, "23": 50.0, "25": 60.0, "26": 65.0, "27": 115.0, "28": 120.0, "29": 60.0, "30": 60.0}, "6": {"1": 60.0, "2": 120.0, "3": 60.0, "4": 60.0, "5": 60.0, "6": 60.0, "8": 60.0, "9": 120.0, "10": 60.0}}, "6207200AST01A": {"5": {"4": 240.0, "5": 220.0, "6": 260.0, "7": 240.0, "8": 300.0, "11": 180.0, "12": 240.0, "13": 240.0, "14": 360.0, "15": 240.0, "16": 120.0, "18": 360.0, "19": 240.0, "20": 230.0, "21": 250.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 365.0, "27": 435.0, "28": 285.0, "29": 275.0, "30": 205.0}, "6": {"1": 710.0, "2": 510.0, "3": 655.0, "4": 390.0, "5": 390.0, "6": 375.0, "7": 310.0, "8": 640.0, "9": 625.0, "10": 695.0, "11": 490.0, "13": 125.0, "15": 405.0, "16": 330.0, "17": 450.0, "18": 390.0, "19": 465.0, "20": 240.0, "22": 350.0, "23": 390.0, "24": 460.0, "25": 380.0, "26": 340.0, "27": 205.0, "29": 275.0, "30": 325.0}}, "6303100AKN04B": {"5": {"4": 30.0, "5": 90.0, "6": 30.0, "7": 90.0, "8": 80.0, "11": 75.0, "12": 90.0, "13": 5.0, "14": 55.0, "15": 30.0, "16": 35.0, "18": 40.0, "19": 80.0, "21": 60.0, "22": 40.0, "23": 20.0, "26": 60.0, "27": 55.0, "28": 65.0, "29": 55.0, "30": 45.0}, "6": {"1": 20.0, "17": 60.0, "19": 60.0, "20": 10.0, "22": 90.0, "23": 20.0}}, "6303100XKN02B": {"5": {"5": 240.0, "6": 120.0, "8": 120.0, "11": 120.0, "12": 195.0, "13": 190.0, "14": 95.0, "15": 225.0, "16": 95.0, "18": 40.0, "19": 120.0, "20": 240.0, "22": 120.0, "25": 120.0, "26": 120.0, "28": 120.0, "30": 10.0}, "6": {"1": 350.0, "2": 240.0, "6": 120.0, "8": 240.0, "9": 195.0, "10": 45.0, "11": 195.0, "15": 45.0, "16": 120.0, "17": 120.0, "23": 120.0, "24": 50.0, "25": 70.0, "30": 120.0}}, "6303100XKN83A": {"5": {"4": 60.0, "6": 60.0, "13": 85.0, "14": 35.0, "15": 60.0, "16": 10.0, "18": 50.0, "19": 10.0, "20": 90.0, "21": 30.0, "22": 50.0, "23": 25.0, "25": 90.0, "26": 30.0, "27": 35.0, "28": 25.0, "29": 35.0}, "6": {"1": 85.0, "2": 195.0, "3": 90.0, "4": 90.0, "5": 80.0, "6": 30.0, "7": 90.0, "10": 170.0, "11": 10.0, "13": 25.0, "14": 95.0, "16": 110.0, "17": 65.0, "19": 25.0, "20": 35.0, "24": 40.0, "25": 20.0, "29": 55.0, "30": 5.0}}, "6303100XST11A": {"5": {"6": 120.0, "7": 120.0, "11": 120.0, "14": 120.0, "15": 120.0, "25": 115.0, "26": 5.0, "29": 240.0, "30": 120.0}, "6": {"1": 515.0, "2": 85.0, "3": 120.0, "5": 120.0, "6": 160.0, "7": 80.0, "9": 250.0, "10": 230.0, "29": 240.0}}, "6303101XST11A": {"5": {"5": 220.0, "6": 140.0, "7": 120.0, "8": 300.0, "11": 60.0, "12": 120.0, "13": 240.0, "14": 120.0, "15": 120.0, "27": 120.0}, "6": {"1": 40.0, "2": 200.0, "3": 160.0, "4": 200.0, "5": 100.0, "6": 95.0, "7": 230.0, "8": 415.0, "10": 225.0, "11": 250.0, "13": 125.0, "18": 120.0, "22": 120.0, "23": 20.0, "24": 460.0, "25": 260.0, "26": 340.0, "27": 120.0}}, "6303200AKN04B": {"5": {"4": 30.0, "5": 90.0, "6": 30.0, "7": 90.0, "8": 80.0, "11": 75.0, "12": 90.0, "13": 5.0, "14": 55.0, "15": 30.0, "16": 35.0, "18": 40.0, "19": 80.0, "21": 60.0, "22": 40.0, "23": 20.0, "26": 60.0, "27": 55.0, "28": 65.0, "29": 55.0, "30": 45.0}, "6": {"1": 20.0, "17": 60.0, "19": 60.0, "20": 10.0, "22": 90.0, "23": 20.0}}, "6303200XKN61A": {"5": {"5": 240.0, "6": 120.0, "8": 120.0, "11": 120.0, "12": 195.0, "13": 190.0, "14": 95.0, "15": 225.0, "16": 95.0, "18": 40.0, "19": 120.0, "20": 240.0, "22": 120.0, "25": 120.0, "26": 120.0, "28": 120.0, "30": 10.0}, "6": {"1": 350.0, "2": 240.0, "6": 120.0, "8": 240.0, "9": 195.0, "10": 45.0, "11": 195.0, "15": 45.0, "16": 120.0, "17": 120.0, "23": 120.0, "24": 50.0, "25": 70.0, "30": 120.0}}, "6303200XKN62A": {"5": {"7": 240.0, "18": 30.0, "19": 90.0, "21": 120.0, "28": 25.0, "29": 215.0}, "6": {"2": 120.0, "3": 120.0, "6": 120.0, "7": 120.0, "9": 120.0, "10": 200.0, "11": 40.0, "13": 120.0, "15": 120.0, "17": 120.0, "22": 120.0}}, "6303200XKN83A": {"5": {"4": 60.0, "6": 60.0, "13": 85.0, "14": 35.0, "15": 60.0, "16": 10.0, "18": 50.0, "19": 10.0, "20": 90.0, "21": 30.0, "22": 50.0, "23": 25.0, "25": 90.0, "26": 30.0, "27": 35.0, "28": 25.0, "29": 35.0}, "6": {"1": 85.0, "2": 195.0, "3": 90.0, "4": 90.0, "5": 80.0, "6": 30.0, "7": 90.0, "10": 170.0, "11": 10.0, "13": 25.0, "14": 95.0, "16": 110.0, "17": 65.0, "19": 25.0, "20": 35.0, "24": 40.0, "25": 20.0, "29": 55.0, "30": 5.0}}, "6307100AST01A": {"5": {"4": 240.0, "5": 220.0, "6": 260.0, "7": 240.0, "8": 300.0, "11": 180.0, "12": 240.0, "13": 240.0, "14": 360.0, "15": 240.0, "16": 120.0, "18": 360.0, "19": 240.0, "20": 230.0, "21": 250.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 365.0, "27": 435.0, "28": 285.0, "29": 275.0, "30": 205.0}, "6": {"1": 710.0, "2": 510.0, "3": 655.0, "4": 390.0, "5": 390.0, "6": 375.0, "7": 310.0, "8": 640.0, "9": 625.0, "10": 695.0, "11": 490.0, "13": 125.0, "15": 405.0, "16": 330.0, "17": 450.0, "18": 390.0, "19": 465.0, "20": 240.0, "22": 350.0, "23": 390.0, "24": 460.0, "25": 380.0, "26": 340.0, "27": 205.0, "29": 275.0, "30": 325.0}}, "6307101AKJ20A": {"5": {"4": 190.0, "5": 90.0, "6": 170.0, "7": 70.0, "8": 60.0, "11": 65.0, "12": 115.0, "13": 120.0, "14": 95.0, "15": 85.0, "16": 60.0, "18": 120.0, "19": 100.0, "20": 80.0, "21": 180.0, "22": 70.0, "23": 50.0, "25": 60.0, "26": 65.0, "27": 115.0, "28": 120.0, "29": 60.0, "30": 60.0}, "6": {"1": 60.0, "2": 120.0, "3": 60.0, "4": 60.0, "5": 60.0, "6": 60.0, "8": 60.0, "9": 120.0, "10": 60.0}}, "6802101XKN61A8P": {"5": {"15": 225.0, "16": 95.0, "18": 40.0, "19": 120.0, "22": 120.0, "30": 10.0}, "6": {"1": 230.0}}, "6802102XKN61A8P": {"5": {"5": 240.0, "6": 120.0, "7": 240.0, "11": 120.0, "18": 30.0, "19": 90.0, "20": 240.0, "21": 120.0, "25": 120.0, "28": 25.0, "29": 215.0}, "6": {"1": 120.0, "2": 360.0, "3": 120.0, "6": 240.0, "7": 120.0, "8": 120.0, "9": 315.0, "10": 245.0, "11": 235.0, "13": 120.0, "15": 165.0, "17": 240.0, "22": 120.0, "23": 120.0, "24": 50.0, "25": 70.0}}, "6802105XKN61A8P": {"5": {"8": 120.0, "12": 195.0, "13": 190.0, "14": 95.0, "26": 120.0, "28": 120.0}, "6": {"8": 120.0, "16": 120.0, "30": 120.0}}, "6802107XKN61A8P": {"5": {"5": 240.0, "6": 120.0, "7": 240.0, "11": 120.0, "18": 30.0, "19": 90.0, "20": 240.0, "21": 120.0, "25": 120.0, "28": 25.0, "29": 215.0}, "6": {"1": 120.0, "2": 360.0, "3": 120.0, "6": 240.0, "7": 120.0, "8": 120.0, "9": 315.0, "10": 245.0, "11": 235.0, "13": 120.0, "15": 165.0, "17": 240.0, "22": 120.0, "23": 120.0, "24": 50.0, "25": 70.0}}, "6802110XST33A8P": {"5": {"6": 120.0, "7": 120.0, "11": 120.0, "14": 120.0, "15": 120.0, "25": 115.0, "26": 5.0, "29": 240.0, "30": 120.0}, "6": {"1": 515.0, "2": 85.0, "3": 120.0, "5": 120.0, "6": 160.0, "7": 80.0, "9": 250.0, "10": 230.0, "29": 240.0}}, "6802111XKN61A8P": {"5": {"15": 225.0, "16": 95.0, "18": 40.0, "19": 120.0, "22": 120.0, "30": 10.0}, "6": {"1": 230.0}}, "6802300XST33A8P": {"5": {"5": 220.0, "6": 140.0, "7": 120.0, "8": 300.0, "11": 60.0, "12": 120.0, "13": 240.0, "14": 120.0, "15": 120.0, "18": 240.0, "22": 240.0, "26": 240.0, "27": 360.0, "28": 240.0}, "6": {"1": 75.0, "2": 305.0, "3": 415.0, "4": 270.0, "5": 270.0, "6": 215.0, "7": 230.0, "8": 640.0, "9": 135.0, "10": 465.0, "11": 250.0, "13": 125.0, "15": 120.0, "16": 15.0, "17": 225.0, "18": 240.0, "22": 120.0, "23": 20.0, "24": 460.0, "25": 260.0, "26": 340.0, "27": 205.0, "29": 35.0, "30": 325.0}}, "6802300XST33AQT": {"5": {"5": 220.0, "6": 140.0, "7": 120.0, "8": 300.0, "11": 60.0, "12": 120.0, "13": 240.0, "14": 120.0, "15": 120.0, "27": 120.0}, "6": {"1": 40.0, "2": 200.0, "3": 160.0, "4": 200.0, "5": 100.0, "6": 95.0, "7": 230.0, "8": 415.0, "10": 225.0, "11": 250.0, "13": 125.0, "18": 120.0, "22": 120.0, "23": 20.0, "24": 460.0, "25": 260.0, "26": 340.0, "27": 120.0}}, "6802300XST33AZ6": {"5": {"5": 220.0, "6": 140.0, "7": 120.0, "8": 300.0, "11": 60.0, "12": 120.0, "13": 240.0, "14": 120.0, "15": 120.0, "18": 240.0, "22": 240.0, "26": 240.0, "27": 360.0, "28": 240.0}, "6": {"1": 75.0, "2": 305.0, "3": 415.0, "4": 270.0, "5": 270.0, "6": 215.0, "7": 230.0, "8": 640.0, "9": 135.0, "10": 465.0, "11": 250.0, "13": 125.0, "15": 120.0, "16": 15.0, "17": 225.0, "18": 240.0, "22": 120.0, "23": 20.0, "24": 460.0, "25": 260.0, "26": 340.0, "27": 205.0, "29": 35.0, "30": 325.0}}, "6802410XST33A8P": {"5": {"4": 240.0, "12": 120.0, "14": 120.0, "16": 120.0, "18": 120.0, "19": 240.0, "20": 230.0, "21": 250.0, "22": 120.0, "23": 225.0, "25": 255.0, "26": 120.0, "27": 75.0, "28": 45.0, "29": 35.0, "30": 85.0}, "6": {"1": 120.0, "2": 120.0, "3": 120.0, "4": 120.0, "9": 240.0, "11": 240.0, "15": 285.0, "16": 315.0, "17": 225.0, "18": 150.0, "19": 465.0, "20": 240.0, "22": 230.0, "23": 370.0, "25": 120.0}}, "6802500XST33A8P": {"5": {"4": 240.0, "6": 120.0, "7": 120.0, "11": 120.0, "12": 120.0, "14": 240.0, "15": 120.0, "16": 120.0, "18": 120.0, "19": 240.0, "20": 230.0, "21": 250.0, "22": 120.0, "23": 225.0, "25": 370.0, "26": 125.0, "27": 75.0, "28": 45.0, "29": 275.0, "30": 205.0}, "6": {"1": 635.0, "2": 205.0, "3": 240.0, "4": 120.0, "5": 120.0, "6": 160.0, "7": 80.0, "9": 490.0, "10": 230.0, "11": 240.0, "15": 285.0, "16": 315.0, "17": 225.0, "18": 150.0, "19": 465.0, "20": 240.0, "22": 230.0, "23": 370.0, "25": 120.0, "29": 240.0}}, "6802700XST33A8P": {"5": {"5": 220.0, "6": 140.0, "7": 120.0, "8": 300.0, "11": 60.0, "12": 120.0, "13": 240.0, "14": 120.0, "15": 120.0, "18": 240.0, "22": 240.0, "26": 240.0, "27": 360.0, "28": 240.0}, "6": {"1": 75.0, "2": 305.0, "3": 415.0, "4": 270.0, "5": 270.0, "6": 215.0, "7": 230.0, "8": 640.0, "9": 135.0, "10": 465.0, "11": 250.0, "13": 125.0, "15": 120.0, "16": 15.0, "17": 225.0, "18": 240.0, "22": 120.0, "23": 20.0, "24": 460.0, "25": 260.0, "26": 340.0, "27": 205.0, "29": 35.0, "30": 325.0}}, "6802700XST33AQT": {"5": {"5": 220.0, "6": 140.0, "7": 120.0, "8": 300.0, "11": 60.0, "12": 120.0, "13": 240.0, "14": 120.0, "15": 120.0, "27": 120.0}, "6": {"1": 40.0, "2": 200.0, "3": 160.0, "4": 200.0, "5": 100.0, "6": 95.0, "7": 230.0, "8": 415.0, "10": 225.0, "11": 250.0, "13": 125.0, "18": 120.0, "22": 120.0, "23": 20.0, "24": 460.0, "25": 260.0, "26": 340.0, "27": 120.0}}, "6802700XST33AZ6": {"5": {"5": 220.0, "6": 140.0, "7": 120.0, "8": 300.0, "11": 60.0, "12": 120.0, "13": 240.0, "14": 120.0, "15": 120.0, "18": 240.0, "22": 240.0, "26": 240.0, "27": 360.0, "28": 240.0}, "6": {"1": 75.0, "2": 305.0, "3": 415.0, "4": 270.0, "5": 270.0, "6": 215.0, "7": 230.0, "8": 640.0, "9": 135.0, "10": 465.0, "11": 250.0, "13": 125.0, "15": 120.0, "16": 15.0, "17": 225.0, "18": 240.0, "22": 120.0, "23": 20.0, "24": 460.0, "25": 260.0, "26": 340.0, "27": 205.0, "29": 35.0, "30": 325.0}}, "6803110XKJ20A": {"5": {"4": 190.0, "5": 90.0, "6": 170.0, "7": 70.0, "8": 60.0, "11": 65.0, "12": 115.0, "13": 120.0, "14": 95.0, "15": 85.0, "16": 60.0, "18": 120.0, "19": 100.0, "20": 80.0, "21": 180.0, "22": 70.0, "23": 50.0, "25": 60.0, "26": 65.0, "27": 115.0, "28": 120.0, "29": 60.0, "30": 60.0}, "6": {"1": 60.0, "2": 120.0, "3": 60.0, "4": 60.0, "5": 60.0, "6": 60.0, "8": 60.0, "9": 120.0, "10": 60.0}}, "6803111XKN02A": {"5": {"8": 120.0, "12": 195.0, "13": 190.0, "14": 95.0, "26": 120.0, "28": 120.0}, "6": {"8": 120.0, "16": 120.0, "30": 120.0}}, "6803112XKN02A": {"5": {"5": 240.0, "6": 120.0, "11": 120.0, "20": 240.0, "25": 120.0}, "6": {"1": 120.0, "2": 240.0, "6": 120.0, "8": 120.0, "9": 195.0, "10": 45.0, "11": 195.0, "15": 45.0, "17": 120.0, "23": 120.0, "24": 50.0, "25": 70.0}}, "6803113XKN01C": {"5": {"4": 240.0, "5": 220.0, "6": 260.0, "7": 240.0, "8": 300.0, "11": 180.0, "12": 240.0, "13": 240.0, "14": 360.0, "15": 240.0, "16": 120.0, "18": 360.0, "19": 240.0, "20": 230.0, "21": 250.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 365.0, "27": 435.0, "28": 285.0, "29": 275.0, "30": 205.0}, "6": {"1": 710.0, "2": 510.0, "3": 655.0, "4": 390.0, "5": 390.0, "6": 375.0, "7": 310.0, "8": 640.0, "9": 625.0, "10": 695.0, "11": 490.0, "13": 125.0, "15": 405.0, "16": 330.0, "17": 450.0, "18": 390.0, "19": 465.0, "20": 240.0, "22": 350.0, "23": 390.0, "24": 460.0, "25": 380.0, "26": 340.0, "27": 205.0, "29": 275.0, "30": 325.0}}, "6803113XST01A": {"5": {"18": 240.0, "22": 240.0, "26": 240.0, "27": 240.0, "28": 240.0}, "6": {"1": 35.0, "2": 105.0, "3": 255.0, "4": 70.0, "5": 170.0, "6": 120.0, "8": 225.0, "9": 135.0, "10": 240.0, "15": 120.0, "16": 15.0, "17": 225.0, "18": 120.0, "27": 85.0, "29": 35.0, "30": 325.0}}, "6803114XKN08A": {"5": {"4": 60.0, "6": 60.0, "13": 85.0, "14": 35.0, "15": 60.0, "16": 10.0, "18": 50.0, "19": 10.0, "20": 90.0, "21": 30.0, "22": 50.0, "23": 25.0, "25": 90.0, "26": 30.0, "27": 35.0, "28": 25.0, "29": 35.0}, "6": {"1": 85.0, "2": 195.0, "3": 90.0, "4": 90.0, "5": 80.0, "6": 30.0, "7": 90.0, "10": 170.0, "11": 10.0, "13": 25.0, "14": 95.0, "16": 110.0, "17": 65.0, "19": 25.0, "20": 35.0, "24": 40.0, "25": 20.0, "29": 55.0, "30": 5.0}}, "6805108XST33A": {"5": {"18": 240.0, "22": 240.0, "26": 240.0, "27": 240.0, "28": 240.0}, "6": {"1": 35.0, "2": 105.0, "3": 255.0, "4": 70.0, "5": 170.0, "6": 120.0, "8": 225.0, "9": 135.0, "10": 240.0, "15": 120.0, "16": 15.0, "17": 225.0, "18": 120.0, "27": 85.0, "29": 35.0, "30": 325.0}}, "6805117XKN61A": {"5": {"5": 240.0, "6": 120.0, "8": 120.0, "11": 120.0, "12": 195.0, "13": 190.0, "14": 95.0, "15": 225.0, "16": 95.0, "18": 40.0, "19": 120.0, "20": 240.0, "22": 120.0, "25": 120.0, "26": 120.0, "28": 120.0, "30": 10.0}, "6": {"1": 350.0, "2": 240.0, "6": 120.0, "8": 240.0, "9": 195.0, "10": 45.0, "11": 195.0, "15": 45.0, "16": 120.0, "17": 120.0, "23": 120.0, "24": 50.0, "25": 70.0, "30": 120.0}}, "6805210XKJ22A": {"5": {"4": 190.0, "5": 90.0, "6": 170.0, "7": 70.0, "8": 60.0, "11": 65.0, "12": 115.0, "13": 120.0, "14": 95.0, "15": 85.0, "16": 60.0, "18": 120.0, "19": 100.0, "20": 80.0, "21": 180.0, "22": 70.0, "23": 50.0, "25": 60.0, "26": 65.0, "27": 115.0, "28": 120.0, "29": 60.0, "30": 60.0}, "6": {"1": 60.0, "2": 120.0, "3": 60.0, "4": 60.0, "5": 60.0, "6": 60.0, "8": 60.0, "9": 120.0, "10": 60.0}}, "6805210XST01A": {"5": {"6": 120.0, "7": 120.0, "11": 120.0, "14": 120.0, "15": 120.0, "25": 115.0, "26": 5.0, "29": 240.0, "30": 120.0}, "6": {"1": 515.0, "2": 85.0, "3": 120.0, "5": 120.0, "6": 160.0, "7": 80.0, "9": 250.0, "10": 230.0, "29": 240.0}}, "6805268XKN81A": {"5": {"4": 60.0, "6": 60.0, "13": 85.0, "14": 35.0, "15": 60.0, "16": 10.0, "18": 50.0, "19": 10.0, "20": 90.0, "21": 30.0, "22": 50.0, "23": 25.0, "25": 90.0, "26": 30.0, "27": 35.0, "28": 25.0, "29": 35.0}, "6": {"1": 85.0, "2": 195.0, "3": 90.0, "4": 90.0, "5": 80.0, "6": 30.0, "7": 90.0, "10": 170.0, "11": 10.0, "13": 25.0, "14": 95.0, "16": 110.0, "17": 65.0, "19": 25.0, "20": 35.0, "24": 40.0, "25": 20.0, "29": 55.0, "30": 5.0}}, "6808101XKN61A8P": {"5": {"5": 240.0, "6": 120.0, "7": 240.0, "8": 120.0, "11": 120.0, "12": 195.0, "13": 190.0, "14": 95.0, "15": 225.0, "16": 95.0, "18": 70.0, "19": 210.0, "20": 240.0, "21": 120.0, "22": 120.0, "25": 120.0, "26": 120.0, "28": 145.0, "29": 215.0, "30": 10.0}, "6": {"1": 350.0, "2": 360.0, "3": 120.0, "6": 240.0, "7": 120.0, "8": 240.0, "9": 315.0, "10": 245.0, "11": 235.0, "13": 120.0, "15": 165.0, "16": 120.0, "17": 240.0, "22": 120.0, "23": 120.0, "24": 50.0, "25": 70.0, "30": 120.0}}, "6808201XKN61A8P": {"5": {"5": 240.0, "6": 120.0, "7": 240.0, "8": 120.0, "11": 120.0, "12": 195.0, "13": 190.0, "14": 95.0, "15": 225.0, "16": 95.0, "18": 70.0, "19": 210.0, "20": 240.0, "21": 120.0, "22": 120.0, "25": 120.0, "26": 120.0, "28": 145.0, "29": 215.0, "30": 10.0}, "6": {"1": 350.0, "2": 360.0, "3": 120.0, "6": 240.0, "7": 120.0, "8": 240.0, "9": 315.0, "10": 245.0, "11": 235.0, "13": 120.0, "15": 165.0, "16": 120.0, "17": 240.0, "22": 120.0, "23": 120.0, "24": 50.0, "25": 70.0, "30": 120.0}}, "6902101XKN61A8P": {"5": {"15": 225.0, "16": 95.0, "18": 40.0, "19": 120.0, "22": 120.0, "30": 10.0}, "6": {"1": 230.0}}, "6902102XKN61A8P": {"5": {"5": 240.0, "6": 120.0, "7": 240.0, "11": 120.0, "18": 30.0, "19": 90.0, "20": 240.0, "21": 120.0, "25": 120.0, "28": 25.0, "29": 215.0}, "6": {"1": 120.0, "2": 360.0, "3": 120.0, "6": 240.0, "7": 120.0, "8": 120.0, "9": 315.0, "10": 245.0, "11": 235.0, "13": 120.0, "15": 165.0, "17": 240.0, "22": 120.0, "23": 120.0, "24": 50.0, "25": 70.0}}, "6902110XST33A8P": {"5": {"6": 120.0, "7": 120.0, "11": 120.0, "14": 120.0, "15": 120.0, "25": 115.0, "26": 5.0, "29": 240.0, "30": 120.0}, "6": {"1": 515.0, "2": 85.0, "3": 120.0, "5": 120.0, "6": 160.0, "7": 80.0, "9": 250.0, "10": 230.0, "29": 240.0}}, "6902300XST33A8P": {"5": {"5": 220.0, "6": 140.0, "7": 120.0, "8": 300.0, "11": 60.0, "12": 120.0, "13": 240.0, "14": 120.0, "15": 120.0, "18": 240.0, "22": 240.0, "26": 240.0, "27": 360.0, "28": 240.0}, "6": {"1": 75.0, "2": 305.0, "3": 415.0, "4": 270.0, "5": 270.0, "6": 215.0, "7": 230.0, "8": 640.0, "9": 135.0, "10": 465.0, "11": 250.0, "13": 125.0, "15": 120.0, "16": 15.0, "17": 225.0, "18": 240.0, "22": 120.0, "23": 20.0, "24": 460.0, "25": 260.0, "26": 340.0, "27": 205.0, "29": 35.0, "30": 325.0}}, "6902300XST33AQT": {"5": {"5": 220.0, "6": 140.0, "7": 120.0, "8": 300.0, "11": 60.0, "12": 120.0, "13": 240.0, "14": 120.0, "15": 120.0, "27": 120.0}, "6": {"1": 40.0, "2": 200.0, "3": 160.0, "4": 200.0, "5": 100.0, "6": 95.0, "7": 230.0, "8": 415.0, "10": 225.0, "11": 250.0, "13": 125.0, "18": 120.0, "22": 120.0, "23": 20.0, "24": 460.0, "25": 260.0, "26": 340.0, "27": 120.0}}, "6902300XST33AZ6": {"5": {"5": 220.0, "6": 140.0, "7": 120.0, "8": 300.0, "11": 60.0, "12": 120.0, "13": 240.0, "14": 120.0, "15": 120.0, "18": 240.0, "22": 240.0, "26": 240.0, "27": 360.0, "28": 240.0}, "6": {"1": 75.0, "2": 305.0, "3": 415.0, "4": 270.0, "5": 270.0, "6": 215.0, "7": 230.0, "8": 640.0, "9": 135.0, "10": 465.0, "11": 250.0, "13": 125.0, "15": 120.0, "16": 15.0, "17": 225.0, "18": 240.0, "22": 120.0, "23": 20.0, "24": 460.0, "25": 260.0, "26": 340.0, "27": 205.0, "29": 35.0, "30": 325.0}}, "6902410XST33A8P": {"5": {"4": 240.0, "12": 120.0, "14": 120.0, "16": 120.0, "18": 120.0, "19": 240.0, "20": 230.0, "21": 250.0, "22": 120.0, "23": 225.0, "25": 255.0, "26": 120.0, "27": 75.0, "28": 45.0, "29": 35.0, "30": 85.0}, "6": {"1": 120.0, "2": 120.0, "3": 120.0, "4": 120.0, "9": 240.0, "11": 240.0, "15": 285.0, "16": 315.0, "17": 225.0, "18": 150.0, "19": 465.0, "20": 240.0, "22": 230.0, "23": 370.0, "25": 120.0}}, "6902500XST33A8P": {"5": {"4": 240.0, "6": 120.0, "7": 120.0, "11": 120.0, "12": 120.0, "14": 240.0, "15": 120.0, "16": 120.0, "18": 120.0, "19": 240.0, "20": 230.0, "21": 250.0, "22": 120.0, "23": 225.0, "25": 370.0, "26": 125.0, "27": 75.0, "28": 45.0, "29": 275.0, "30": 205.0}, "6": {"1": 635.0, "2": 205.0, "3": 240.0, "4": 120.0, "5": 120.0, "6": 160.0, "7": 80.0, "9": 490.0, "10": 230.0, "11": 240.0, "15": 285.0, "16": 315.0, "17": 225.0, "18": 150.0, "19": 465.0, "20": 240.0, "22": 230.0, "23": 370.0, "25": 120.0, "29": 240.0}}, "6902700XST33A8P": {"5": {"5": 220.0, "6": 140.0, "7": 120.0, "8": 300.0, "11": 60.0, "12": 120.0, "13": 240.0, "14": 120.0, "15": 120.0, "18": 240.0, "22": 240.0, "26": 240.0, "27": 360.0, "28": 240.0}, "6": {"1": 75.0, "2": 305.0, "3": 415.0, "4": 270.0, "5": 270.0, "6": 215.0, "7": 230.0, "8": 640.0, "9": 135.0, "10": 465.0, "11": 250.0, "13": 125.0, "15": 120.0, "16": 15.0, "17": 225.0, "18": 240.0, "22": 120.0, "23": 20.0, "24": 460.0, "25": 260.0, "26": 340.0, "27": 205.0, "29": 35.0, "30": 325.0}}, "6902700XST33AQT": {"5": {"5": 220.0, "6": 140.0, "7": 120.0, "8": 300.0, "11": 60.0, "12": 120.0, "13": 240.0, "14": 120.0, "15": 120.0, "27": 120.0}, "6": {"1": 40.0, "2": 200.0, "3": 160.0, "4": 200.0, "5": 100.0, "6": 95.0, "7": 230.0, "8": 415.0, "10": 225.0, "11": 250.0, "13": 125.0, "18": 120.0, "22": 120.0, "23": 20.0, "24": 460.0, "25": 260.0, "26": 340.0, "27": 120.0}}, "6902700XST33AZ6": {"5": {"5": 220.0, "6": 140.0, "7": 120.0, "8": 300.0, "11": 60.0, "12": 120.0, "13": 240.0, "14": 120.0, "15": 120.0, "18": 240.0, "22": 240.0, "26": 240.0, "27": 360.0, "28": 240.0}, "6": {"1": 75.0, "2": 305.0, "3": 415.0, "4": 270.0, "5": 270.0, "6": 215.0, "7": 230.0, "8": 640.0, "9": 135.0, "10": 465.0, "11": 250.0, "13": 125.0, "15": 120.0, "16": 15.0, "17": 225.0, "18": 240.0, "22": 120.0, "23": 20.0, "24": 460.0, "25": 260.0, "26": 340.0, "27": 205.0, "29": 35.0, "30": 325.0}}, "6903101XKN61A8P": {"5": {"15": 225.0, "16": 95.0, "18": 40.0, "19": 120.0, "22": 120.0, "30": 10.0}, "6": {"1": 230.0}}, "6903102XKN61A8P": {"5": {"8": 120.0, "12": 195.0, "13": 190.0, "14": 95.0, "26": 120.0, "28": 120.0}, "6": {"8": 120.0, "16": 120.0, "30": 120.0}}, "6903103XKN61A8P": {"5": {"5": 240.0, "6": 120.0, "7": 240.0, "11": 120.0, "18": 30.0, "19": 90.0, "20": 240.0, "21": 120.0, "25": 120.0, "28": 25.0, "29": 215.0}, "6": {"1": 120.0, "2": 360.0, "3": 120.0, "6": 240.0, "7": 120.0, "8": 120.0, "9": 315.0, "10": 245.0, "11": 235.0, "13": 120.0, "15": 165.0, "17": 240.0, "22": 120.0, "23": 120.0, "24": 50.0, "25": 70.0}}, "6903110XKJ20A": {"5": {"4": 190.0, "5": 90.0, "6": 170.0, "7": 70.0, "8": 60.0, "11": 65.0, "12": 115.0, "13": 120.0, "14": 95.0, "15": 85.0, "16": 60.0, "18": 120.0, "19": 100.0, "20": 80.0, "21": 180.0, "22": 70.0, "23": 50.0, "25": 60.0, "26": 65.0, "27": 115.0, "28": 120.0, "29": 60.0, "30": 60.0}, "6": {"1": 60.0, "2": 120.0, "3": 60.0, "4": 60.0, "5": 60.0, "6": 60.0, "8": 60.0, "9": 120.0, "10": 60.0}}, "6903112XKN02A": {"5": {"8": 120.0, "12": 195.0, "13": 190.0, "14": 95.0, "26": 120.0, "28": 120.0}, "6": {"8": 120.0, "16": 120.0, "30": 120.0}}, "6903113XKN02A": {"5": {"5": 240.0, "6": 120.0, "11": 120.0, "20": 240.0, "25": 120.0}, "6": {"1": 120.0, "2": 240.0, "6": 120.0, "8": 120.0, "9": 195.0, "10": 45.0, "11": 195.0, "15": 45.0, "17": 120.0, "23": 120.0, "24": 50.0, "25": 70.0}}, "6903113XKN08A": {"5": {"4": 60.0, "6": 60.0, "13": 85.0, "14": 35.0, "15": 60.0, "16": 10.0, "18": 50.0, "19": 10.0, "20": 90.0, "21": 30.0, "22": 50.0, "23": 25.0, "25": 90.0, "26": 30.0, "27": 35.0, "28": 25.0, "29": 35.0}, "6": {"1": 85.0, "2": 195.0, "3": 90.0, "4": 90.0, "5": 80.0, "6": 30.0, "7": 90.0, "10": 170.0, "11": 10.0, "13": 25.0, "14": 95.0, "16": 110.0, "17": 65.0, "19": 25.0, "20": 35.0, "24": 40.0, "25": 20.0, "29": 55.0, "30": 5.0}}, "6903116XKN01C": {"5": {"4": 240.0, "5": 220.0, "6": 260.0, "7": 240.0, "8": 300.0, "11": 180.0, "12": 240.0, "13": 240.0, "14": 360.0, "15": 240.0, "16": 120.0, "18": 360.0, "19": 240.0, "20": 230.0, "21": 250.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 365.0, "27": 435.0, "28": 285.0, "29": 275.0, "30": 205.0}, "6": {"1": 710.0, "2": 510.0, "3": 655.0, "4": 390.0, "5": 390.0, "6": 375.0, "7": 310.0, "8": 640.0, "9": 625.0, "10": 695.0, "11": 490.0, "13": 125.0, "15": 405.0, "16": 330.0, "17": 450.0, "18": 390.0, "19": 465.0, "20": 240.0, "22": 350.0, "23": 390.0, "24": 460.0, "25": 380.0, "26": 340.0, "27": 205.0, "29": 275.0, "30": 325.0}}, "6905107XST33A": {"5": {"5": 220.0, "6": 140.0, "7": 120.0, "8": 300.0, "11": 60.0, "12": 120.0, "13": 240.0, "14": 120.0, "15": 120.0, "27": 120.0}, "6": {"1": 40.0, "2": 200.0, "3": 160.0, "4": 200.0, "5": 100.0, "6": 95.0, "7": 230.0, "8": 415.0, "10": 225.0, "11": 250.0, "13": 125.0, "18": 120.0, "22": 120.0, "23": 20.0, "24": 460.0, "25": 260.0, "26": 340.0, "27": 120.0}}, "6905111XKN61A": {"5": {"5": 240.0, "6": 120.0, "8": 120.0, "11": 120.0, "12": 195.0, "13": 190.0, "14": 95.0, "15": 225.0, "16": 95.0, "18": 40.0, "19": 120.0, "20": 240.0, "22": 120.0, "25": 120.0, "26": 120.0, "28": 120.0, "30": 10.0}, "6": {"1": 350.0, "2": 240.0, "6": 120.0, "8": 240.0, "9": 195.0, "10": 45.0, "11": 195.0, "15": 45.0, "16": 120.0, "17": 120.0, "23": 120.0, "24": 50.0, "25": 70.0, "30": 120.0}}, "6905210XKJ22A": {"5": {"4": 190.0, "5": 90.0, "6": 170.0, "7": 70.0, "8": 60.0, "11": 65.0, "12": 115.0, "13": 120.0, "14": 95.0, "15": 85.0, "16": 60.0, "18": 120.0, "19": 100.0, "20": 80.0, "21": 180.0, "22": 70.0, "23": 50.0, "25": 60.0, "26": 65.0, "27": 115.0, "28": 120.0, "29": 60.0, "30": 60.0}, "6": {"1": 60.0, "2": 120.0, "3": 60.0, "4": 60.0, "5": 60.0, "6": 60.0, "8": 60.0, "9": 120.0, "10": 60.0}}, "6905210XST01A": {"5": {"6": 120.0, "7": 120.0, "11": 120.0, "14": 120.0, "15": 120.0, "25": 115.0, "26": 5.0, "29": 240.0, "30": 120.0}, "6": {"1": 515.0, "2": 85.0, "3": 120.0, "5": 120.0, "6": 160.0, "7": 80.0, "9": 250.0, "10": 230.0, "29": 240.0}}, "6905273XKN81A": {"5": {"4": 60.0, "6": 60.0, "13": 85.0, "14": 35.0, "15": 60.0, "16": 10.0, "18": 50.0, "19": 10.0, "20": 90.0, "21": 30.0, "22": 50.0, "23": 25.0, "25": 90.0, "26": 30.0, "27": 35.0, "28": 25.0, "29": 35.0}, "6": {"1": 85.0, "2": 195.0, "3": 90.0, "4": 90.0, "5": 80.0, "6": 30.0, "7": 90.0, "10": 170.0, "11": 10.0, "13": 25.0, "14": 95.0, "16": 110.0, "17": 65.0, "19": 25.0, "20": 35.0, "24": 40.0, "25": 20.0, "29": 55.0, "30": 5.0}}, "6908108XKN61A8P": {"5": {"5": 240.0, "6": 120.0, "7": 240.0, "8": 120.0, "11": 120.0, "12": 195.0, "13": 190.0, "14": 95.0, "15": 225.0, "16": 95.0, "18": 70.0, "19": 210.0, "20": 240.0, "21": 120.0, "22": 120.0, "25": 120.0, "26": 120.0, "28": 145.0, "29": 215.0, "30": 10.0}, "6": {"1": 350.0, "2": 360.0, "3": 120.0, "6": 240.0, "7": 120.0, "8": 240.0, "9": 315.0, "10": 245.0, "11": 235.0, "13": 120.0, "15": 165.0, "16": 120.0, "17": 240.0, "22": 120.0, "23": 120.0, "24": 50.0, "25": 70.0, "30": 120.0}}, "6908208XKN61A8P": {"5": {"5": 240.0, "6": 120.0, "7": 240.0, "8": 120.0, "11": 120.0, "12": 195.0, "13": 190.0, "14": 95.0, "15": 225.0, "16": 95.0, "18": 70.0, "19": 210.0, "20": 240.0, "21": 120.0, "22": 120.0, "25": 120.0, "26": 120.0, "28": 145.0, "29": 215.0, "30": 10.0}, "6": {"1": 350.0, "2": 360.0, "3": 120.0, "6": 240.0, "7": 120.0, "8": 240.0, "9": 315.0, "10": 245.0, "11": 235.0, "13": 120.0, "15": 165.0, "16": 120.0, "17": 240.0, "22": 120.0, "23": 120.0, "24": 50.0, "25": 70.0, "30": 120.0}}, "7002101XKN61A8P": {"5": {"15": 225.0, "16": 95.0, "18": 40.0, "19": 120.0, "22": 120.0, "30": 10.0}, "6": {"1": 230.0}}, "7002102XKN61A8P": {"5": {"5": 240.0, "6": 120.0, "7": 240.0, "11": 120.0, "18": 30.0, "19": 90.0, "20": 240.0, "21": 120.0, "25": 120.0, "28": 25.0, "29": 215.0}, "6": {"1": 120.0, "2": 360.0, "3": 120.0, "6": 240.0, "7": 120.0, "8": 120.0, "9": 315.0, "10": 245.0, "11": 235.0, "13": 120.0, "15": 165.0, "17": 240.0, "22": 120.0, "23": 120.0, "24": 50.0, "25": 70.0}}, "7002105XKN61A8P": {"5": {"15": 225.0, "16": 95.0, "18": 40.0, "19": 120.0, "22": 120.0, "30": 10.0}, "6": {"1": 230.0}}, "7002106XKN61A8P": {"5": {"5": 240.0, "6": 120.0, "7": 240.0, "11": 120.0, "18": 30.0, "19": 90.0, "20": 240.0, "21": 120.0, "25": 120.0, "28": 25.0, "29": 215.0}, "6": {"1": 120.0, "2": 360.0, "3": 120.0, "6": 240.0, "7": 120.0, "8": 120.0, "9": 315.0, "10": 245.0, "11": 235.0, "13": 120.0, "15": 165.0, "17": 240.0, "22": 120.0, "23": 120.0, "24": 50.0, "25": 70.0}}, "7002130XST33A8P": {"5": {"5": 220.0, "6": 140.0, "7": 120.0, "8": 300.0, "11": 60.0, "12": 120.0, "13": 240.0, "14": 120.0, "15": 120.0, "18": 240.0, "22": 240.0, "26": 240.0, "27": 360.0, "28": 240.0}, "6": {"1": 75.0, "2": 305.0, "3": 415.0, "4": 270.0, "5": 270.0, "6": 215.0, "7": 230.0, "8": 640.0, "9": 135.0, "10": 465.0, "11": 250.0, "13": 125.0, "15": 120.0, "16": 15.0, "17": 225.0, "18": 240.0, "22": 120.0, "23": 20.0, "24": 460.0, "25": 260.0, "26": 340.0, "27": 205.0, "29": 35.0, "30": 325.0}}, "7002130XST33AQT": {"5": {"5": 220.0, "6": 140.0, "7": 120.0, "8": 300.0, "11": 60.0, "12": 120.0, "13": 240.0, "14": 120.0, "15": 120.0, "27": 120.0}, "6": {"1": 40.0, "2": 200.0, "3": 160.0, "4": 200.0, "5": 100.0, "6": 95.0, "7": 230.0, "8": 415.0, "10": 225.0, "11": 250.0, "13": 125.0, "18": 120.0, "22": 120.0, "23": 20.0, "24": 460.0, "25": 260.0, "26": 340.0, "27": 120.0}}, "7002130XST33AZ6": {"5": {"5": 220.0, "6": 140.0, "7": 120.0, "8": 300.0, "11": 60.0, "12": 120.0, "13": 240.0, "14": 120.0, "15": 120.0, "18": 240.0, "22": 240.0, "26": 240.0, "27": 360.0, "28": 240.0}, "6": {"1": 75.0, "2": 305.0, "3": 415.0, "4": 270.0, "5": 270.0, "6": 215.0, "7": 230.0, "8": 640.0, "9": 135.0, "10": 465.0, "11": 250.0, "13": 125.0, "15": 120.0, "16": 15.0, "17": 225.0, "18": 240.0, "22": 120.0, "23": 20.0, "24": 460.0, "25": 260.0, "26": 340.0, "27": 205.0, "29": 35.0, "30": 325.0}}, "7002141XST33A8P": {"5": {"4": 240.0, "6": 120.0, "7": 120.0, "11": 120.0, "12": 120.0, "14": 240.0, "15": 120.0, "16": 120.0, "18": 120.0, "19": 240.0, "20": 230.0, "21": 250.0, "22": 120.0, "23": 225.0, "25": 370.0, "26": 125.0, "27": 75.0, "28": 45.0, "29": 275.0, "30": 205.0}, "6": {"1": 635.0, "2": 205.0, "3": 240.0, "4": 120.0, "5": 120.0, "6": 160.0, "7": 80.0, "9": 490.0, "10": 230.0, "11": 240.0, "15": 285.0, "16": 315.0, "17": 225.0, "18": 150.0, "19": 465.0, "20": 240.0, "22": 230.0, "23": 370.0, "25": 120.0, "29": 240.0}}, "7002211XST33A8P": {"5": {"4": 240.0, "6": 120.0, "7": 120.0, "11": 120.0, "12": 120.0, "14": 240.0, "15": 120.0, "16": 120.0, "18": 120.0, "19": 240.0, "20": 230.0, "21": 250.0, "22": 120.0, "23": 225.0, "25": 370.0, "26": 125.0, "27": 75.0, "28": 45.0, "29": 275.0, "30": 205.0}, "6": {"1": 635.0, "2": 205.0, "3": 240.0, "4": 120.0, "5": 120.0, "6": 160.0, "7": 80.0, "9": 490.0, "10": 230.0, "11": 240.0, "15": 285.0, "16": 315.0, "17": 225.0, "18": 150.0, "19": 465.0, "20": 240.0, "22": 230.0, "23": 370.0, "25": 120.0, "29": 240.0}}, "7002230XST33A8P": {"5": {"5": 220.0, "6": 140.0, "7": 120.0, "8": 300.0, "11": 60.0, "12": 120.0, "13": 240.0, "14": 120.0, "15": 120.0, "18": 240.0, "22": 240.0, "26": 240.0, "27": 360.0, "28": 240.0}, "6": {"1": 75.0, "2": 305.0, "3": 415.0, "4": 270.0, "5": 270.0, "6": 215.0, "7": 230.0, "8": 640.0, "9": 135.0, "10": 465.0, "11": 250.0, "13": 125.0, "15": 120.0, "16": 15.0, "17": 225.0, "18": 240.0, "22": 120.0, "23": 20.0, "24": 460.0, "25": 260.0, "26": 340.0, "27": 205.0, "29": 35.0, "30": 325.0}}, "7002230XST33AQT": {"5": {"5": 220.0, "6": 140.0, "7": 120.0, "8": 300.0, "11": 60.0, "12": 120.0, "13": 240.0, "14": 120.0, "15": 120.0, "27": 120.0}, "6": {"1": 40.0, "2": 200.0, "3": 160.0, "4": 200.0, "5": 100.0, "6": 95.0, "7": 230.0, "8": 415.0, "10": 225.0, "11": 250.0, "13": 125.0, "18": 120.0, "22": 120.0, "23": 20.0, "24": 460.0, "25": 260.0, "26": 340.0, "27": 120.0}}, "7002230XST33AZ6": {"5": {"5": 220.0, "6": 140.0, "7": 120.0, "8": 300.0, "11": 60.0, "12": 120.0, "13": 240.0, "14": 120.0, "15": 120.0, "18": 240.0, "22": 240.0, "26": 240.0, "27": 360.0, "28": 240.0}, "6": {"1": 75.0, "2": 305.0, "3": 415.0, "4": 270.0, "5": 270.0, "6": 215.0, "7": 230.0, "8": 640.0, "9": 135.0, "10": 465.0, "11": 250.0, "13": 125.0, "15": 120.0, "16": 15.0, "17": 225.0, "18": 240.0, "22": 120.0, "23": 20.0, "24": 460.0, "25": 260.0, "26": 340.0, "27": 205.0, "29": 35.0, "30": 325.0}}, "7002310XST33A8P": {"5": {"4": 240.0, "6": 120.0, "7": 120.0, "11": 120.0, "12": 120.0, "14": 240.0, "15": 120.0, "16": 120.0, "18": 120.0, "19": 240.0, "20": 230.0, "21": 250.0, "22": 120.0, "23": 225.0, "25": 370.0, "26": 125.0, "27": 75.0, "28": 45.0, "29": 275.0, "30": 205.0}, "6": {"1": 635.0, "2": 205.0, "3": 240.0, "4": 120.0, "5": 120.0, "6": 160.0, "7": 80.0, "9": 490.0, "10": 230.0, "11": 240.0, "15": 285.0, "16": 315.0, "17": 225.0, "18": 150.0, "19": 465.0, "20": 240.0, "22": 230.0, "23": 370.0, "25": 120.0, "29": 240.0}}, "7002440XST33A8P": {"5": {"5": 220.0, "6": 140.0, "7": 120.0, "8": 300.0, "11": 60.0, "12": 120.0, "13": 240.0, "14": 120.0, "15": 120.0, "18": 240.0, "22": 240.0, "26": 240.0, "27": 360.0, "28": 240.0}, "6": {"1": 75.0, "2": 305.0, "3": 415.0, "4": 270.0, "5": 270.0, "6": 215.0, "7": 230.0, "8": 640.0, "9": 135.0, "10": 465.0, "11": 250.0, "13": 125.0, "15": 120.0, "16": 15.0, "17": 225.0, "18": 240.0, "22": 120.0, "23": 20.0, "24": 460.0, "25": 260.0, "26": 340.0, "27": 205.0, "29": 35.0, "30": 325.0}}, "7002440XST33AQT": {"5": {"5": 220.0, "6": 140.0, "7": 120.0, "8": 300.0, "11": 60.0, "12": 120.0, "13": 240.0, "14": 120.0, "15": 120.0, "27": 120.0}, "6": {"1": 40.0, "2": 200.0, "3": 160.0, "4": 200.0, "5": 100.0, "6": 95.0, "7": 230.0, "8": 415.0, "10": 225.0, "11": 250.0, "13": 125.0, "18": 120.0, "22": 120.0, "23": 20.0, "24": 460.0, "25": 260.0, "26": 340.0, "27": 120.0}}, "7002440XST33AZ6": {"5": {"5": 220.0, "6": 140.0, "7": 120.0, "8": 300.0, "11": 60.0, "12": 120.0, "13": 240.0, "14": 120.0, "15": 120.0, "18": 240.0, "22": 240.0, "26": 240.0, "27": 360.0, "28": 240.0}, "6": {"1": 75.0, "2": 305.0, "3": 415.0, "4": 270.0, "5": 270.0, "6": 215.0, "7": 230.0, "8": 640.0, "9": 135.0, "10": 465.0, "11": 250.0, "13": 125.0, "15": 120.0, "16": 15.0, "17": 225.0, "18": 240.0, "22": 120.0, "23": 20.0, "24": 460.0, "25": 260.0, "26": 340.0, "27": 205.0, "29": 35.0, "30": 325.0}}, "7003100XKJ22A": {"5": {"4": 190.0, "5": 90.0, "6": 170.0, "7": 70.0, "8": 60.0, "11": 65.0, "12": 115.0, "13": 120.0, "14": 95.0, "15": 85.0, "16": 60.0, "18": 120.0, "19": 100.0, "20": 80.0, "21": 180.0, "22": 70.0, "23": 50.0, "25": 60.0, "26": 65.0, "27": 115.0, "28": 120.0, "29": 60.0, "30": 60.0}, "6": {"1": 60.0, "2": 120.0, "3": 60.0, "4": 60.0, "5": 60.0, "6": 60.0, "8": 60.0, "9": 120.0, "10": 60.0}}, "7003101XKN61A": {"5": {"5": 240.0, "6": 120.0, "8": 120.0, "11": 120.0, "12": 195.0, "13": 190.0, "14": 95.0, "15": 225.0, "16": 95.0, "18": 40.0, "19": 120.0, "20": 240.0, "22": 120.0, "25": 120.0, "26": 120.0, "28": 120.0, "30": 10.0}, "6": {"1": 350.0, "2": 240.0, "6": 120.0, "8": 240.0, "9": 195.0, "10": 45.0, "11": 195.0, "15": 45.0, "16": 120.0, "17": 120.0, "23": 120.0, "24": 50.0, "25": 70.0, "30": 120.0}}, "7003110XST01A": {"5": {"6": 120.0, "7": 120.0, "11": 120.0, "14": 120.0, "15": 120.0, "25": 115.0, "26": 5.0, "29": 240.0, "30": 120.0}, "6": {"1": 515.0, "2": 85.0, "3": 120.0, "5": 120.0, "6": 160.0, "7": 80.0, "9": 250.0, "10": 230.0, "29": 240.0}}, "7003110XST11A": {"5": {"5": 220.0, "6": 140.0, "7": 120.0, "8": 300.0, "11": 60.0, "12": 120.0, "13": 240.0, "14": 120.0, "15": 120.0, "27": 120.0}, "6": {"1": 40.0, "2": 200.0, "3": 160.0, "4": 200.0, "5": 100.0, "6": 95.0, "7": 230.0, "8": 415.0, "10": 225.0, "11": 250.0, "13": 125.0, "18": 120.0, "22": 120.0, "23": 20.0, "24": 460.0, "25": 260.0, "26": 340.0, "27": 120.0}}, "7003112XKN81A": {"5": {"4": 60.0, "6": 60.0, "13": 85.0, "14": 35.0, "15": 60.0, "16": 10.0, "18": 50.0, "19": 10.0, "20": 90.0, "21": 30.0, "22": 50.0, "23": 25.0, "25": 90.0, "26": 30.0, "27": 35.0, "28": 25.0, "29": 35.0}, "6": {"1": 85.0, "2": 195.0, "3": 90.0, "4": 90.0, "5": 80.0, "6": 30.0, "7": 90.0, "10": 170.0, "11": 10.0, "13": 25.0, "14": 95.0, "16": 110.0, "17": 65.0, "19": 25.0, "20": 35.0, "24": 40.0, "25": 20.0, "29": 55.0, "30": 5.0}}, "7005104XKN07A": {"5": {"5": 240.0, "6": 120.0, "8": 120.0, "11": 120.0, "12": 195.0, "13": 190.0, "14": 95.0, "15": 225.0, "16": 95.0, "18": 40.0, "19": 120.0, "20": 240.0, "22": 120.0, "25": 120.0, "26": 120.0, "28": 120.0, "30": 10.0}, "6": {"1": 350.0, "2": 240.0, "6": 120.0, "8": 240.0, "9": 195.0, "10": 45.0, "11": 195.0, "15": 45.0, "16": 120.0, "17": 120.0, "23": 120.0, "24": 50.0, "25": 70.0, "30": 120.0}}, "7005107XKN81A": {"5": {"4": 60.0, "6": 60.0, "13": 85.0, "14": 35.0, "15": 60.0, "16": 10.0, "18": 50.0, "19": 10.0, "20": 90.0, "21": 30.0, "22": 50.0, "23": 25.0, "25": 90.0, "26": 30.0, "27": 35.0, "28": 25.0, "29": 35.0}, "6": {"1": 85.0, "2": 195.0, "3": 90.0, "4": 90.0, "5": 80.0, "6": 30.0, "7": 90.0, "10": 170.0, "11": 10.0, "13": 25.0, "14": 95.0, "16": 110.0, "17": 65.0, "19": 25.0, "20": 35.0, "24": 40.0, "25": 20.0, "29": 55.0, "30": 5.0}}, "7005110XKJ22A": {"5": {"4": 190.0, "5": 90.0, "6": 170.0, "7": 70.0, "8": 60.0, "11": 65.0, "12": 115.0, "13": 120.0, "14": 95.0, "15": 85.0, "16": 60.0, "18": 120.0, "19": 100.0, "20": 80.0, "21": 180.0, "22": 70.0, "23": 50.0, "25": 60.0, "26": 65.0, "27": 115.0, "28": 120.0, "29": 60.0, "30": 60.0}, "6": {"1": 60.0, "2": 120.0, "3": 60.0, "4": 60.0, "5": 60.0, "6": 60.0, "8": 60.0, "9": 120.0, "10": 60.0}}, "7005111XKN07A": {"5": {"5": 240.0, "6": 120.0, "8": 120.0, "11": 120.0, "12": 195.0, "13": 190.0, "14": 95.0, "15": 225.0, "16": 95.0, "18": 40.0, "19": 120.0, "20": 240.0, "22": 120.0, "25": 120.0, "26": 120.0, "28": 120.0, "30": 10.0}, "6": {"1": 350.0, "2": 240.0, "6": 120.0, "8": 240.0, "9": 195.0, "10": 45.0, "11": 195.0, "15": 45.0, "16": 120.0, "17": 120.0, "23": 120.0, "24": 50.0, "25": 70.0, "30": 120.0}}, "7005113XKN81B": {"5": {"4": 60.0, "6": 60.0, "13": 85.0, "14": 35.0, "15": 60.0, "16": 10.0, "18": 50.0, "19": 10.0, "20": 90.0, "21": 30.0, "22": 50.0, "23": 25.0, "25": 90.0, "26": 30.0, "27": 35.0, "28": 25.0, "29": 35.0}, "6": {"1": 85.0, "2": 195.0, "3": 90.0, "4": 90.0, "5": 80.0, "6": 30.0, "7": 90.0, "10": 170.0, "11": 10.0, "13": 25.0, "14": 95.0, "16": 110.0, "17": 65.0, "19": 25.0, "20": 35.0, "24": 40.0, "25": 20.0, "29": 55.0, "30": 5.0}}, "7005127XKN61A8P": {"5": {"15": 225.0, "16": 95.0, "18": 40.0, "19": 120.0, "22": 120.0, "30": 10.0}, "6": {"1": 230.0}}, "7005128XKN61A8P": {"5": {"5": 240.0, "6": 120.0, "7": 240.0, "11": 120.0, "18": 30.0, "19": 90.0, "20": 240.0, "21": 120.0, "25": 120.0, "28": 25.0, "29": 215.0}, "6": {"1": 120.0, "2": 360.0, "3": 120.0, "6": 240.0, "7": 120.0, "8": 120.0, "9": 315.0, "10": 245.0, "11": 235.0, "13": 120.0, "15": 165.0, "17": 240.0, "22": 120.0, "23": 120.0, "24": 50.0, "25": 70.0}}, "7005310XKJ22A": {"5": {"4": 190.0, "5": 90.0, "6": 170.0, "7": 70.0, "8": 60.0, "11": 65.0, "12": 115.0, "13": 120.0, "14": 95.0, "15": 85.0, "16": 60.0, "18": 120.0, "19": 100.0, "20": 80.0, "21": 180.0, "22": 70.0, "23": 50.0, "25": 60.0, "26": 65.0, "27": 115.0, "28": 120.0, "29": 60.0, "30": 60.0}, "6": {"1": 60.0, "2": 120.0, "3": 60.0, "4": 60.0, "5": 60.0, "6": 60.0, "8": 60.0, "9": 120.0, "10": 60.0}}, "7005320XST01A": {"5": {"4": 240.0, "5": 220.0, "6": 260.0, "7": 240.0, "8": 300.0, "11": 180.0, "12": 240.0, "13": 240.0, "14": 360.0, "15": 240.0, "16": 120.0, "18": 360.0, "19": 240.0, "20": 230.0, "21": 250.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 365.0, "27": 435.0, "28": 285.0, "29": 275.0, "30": 205.0}, "6": {"1": 710.0, "2": 510.0, "3": 655.0, "4": 390.0, "5": 390.0, "6": 375.0, "7": 310.0, "8": 640.0, "9": 625.0, "10": 695.0, "11": 490.0, "13": 125.0, "15": 405.0, "16": 330.0, "17": 450.0, "18": 390.0, "19": 465.0, "20": 240.0, "22": 350.0, "23": 390.0, "24": 460.0, "25": 380.0, "26": 340.0, "27": 205.0, "29": 275.0, "30": 325.0}}, "7005410XST01A": {"5": {"4": 240.0, "5": 220.0, "6": 260.0, "7": 240.0, "8": 300.0, "11": 180.0, "12": 240.0, "13": 240.0, "14": 360.0, "15": 240.0, "16": 120.0, "18": 360.0, "19": 240.0, "20": 230.0, "21": 250.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 365.0, "27": 435.0, "28": 285.0, "29": 275.0, "30": 205.0}, "6": {"1": 710.0, "2": 510.0, "3": 655.0, "4": 390.0, "5": 390.0, "6": 375.0, "7": 310.0, "8": 640.0, "9": 625.0, "10": 695.0, "11": 490.0, "13": 125.0, "15": 405.0, "16": 330.0, "17": 450.0, "18": 390.0, "19": 465.0, "20": 240.0, "22": 350.0, "23": 390.0, "24": 460.0, "25": 380.0, "26": 340.0, "27": 205.0, "29": 275.0, "30": 325.0}}, "7008100XST11A8P": {"5": {"4": 240.0, "5": 220.0, "6": 260.0, "7": 240.0, "8": 300.0, "11": 180.0, "12": 240.0, "13": 240.0, "14": 360.0, "15": 240.0, "16": 120.0, "18": 360.0, "19": 240.0, "20": 230.0, "21": 250.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 365.0, "27": 435.0, "28": 285.0, "29": 275.0, "30": 205.0}, "6": {"1": 710.0, "2": 510.0, "3": 655.0, "4": 390.0, "5": 390.0, "6": 375.0, "7": 310.0, "8": 640.0, "9": 625.0, "10": 695.0, "11": 490.0, "13": 125.0, "15": 405.0, "16": 330.0, "17": 450.0, "18": 390.0, "19": 465.0, "20": 240.0, "22": 350.0, "23": 390.0, "24": 460.0, "25": 380.0, "26": 340.0, "27": 205.0, "29": 275.0, "30": 325.0}}, "7008101XKN61A8P": {"5": {"5": 240.0, "6": 120.0, "7": 240.0, "8": 120.0, "11": 120.0, "12": 195.0, "13": 190.0, "14": 95.0, "15": 225.0, "16": 95.0, "18": 70.0, "19": 210.0, "20": 240.0, "21": 120.0, "22": 120.0, "25": 120.0, "26": 120.0, "28": 145.0, "29": 215.0, "30": 10.0}, "6": {"1": 350.0, "2": 360.0, "3": 120.0, "6": 240.0, "7": 120.0, "8": 240.0, "9": 315.0, "10": 245.0, "11": 235.0, "13": 120.0, "15": 165.0, "16": 120.0, "17": 240.0, "22": 120.0, "23": 120.0, "24": 50.0, "25": 70.0, "30": 120.0}}, "7008102XKN61A8P": {"5": {"5": 240.0, "6": 120.0, "7": 240.0, "8": 120.0, "11": 120.0, "12": 195.0, "13": 190.0, "14": 95.0, "15": 225.0, "16": 95.0, "18": 70.0, "19": 210.0, "20": 240.0, "21": 120.0, "22": 120.0, "25": 120.0, "26": 120.0, "28": 145.0, "29": 215.0, "30": 10.0}, "6": {"1": 350.0, "2": 360.0, "3": 120.0, "6": 240.0, "7": 120.0, "8": 240.0, "9": 315.0, "10": 245.0, "11": 235.0, "13": 120.0, "15": 165.0, "16": 120.0, "17": 240.0, "22": 120.0, "23": 120.0, "24": 50.0, "25": 70.0, "30": 120.0}}, "7008110XST11A8P": {"5": {"4": 240.0, "5": 220.0, "6": 260.0, "7": 240.0, "8": 300.0, "11": 180.0, "12": 240.0, "13": 240.0, "14": 360.0, "15": 240.0, "16": 120.0, "18": 360.0, "19": 240.0, "20": 230.0, "21": 250.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 365.0, "27": 435.0, "28": 285.0, "29": 275.0, "30": 205.0}, "6": {"1": 710.0, "2": 510.0, "3": 655.0, "4": 390.0, "5": 390.0, "6": 375.0, "7": 310.0, "8": 640.0, "9": 625.0, "10": 695.0, "11": 490.0, "13": 125.0, "15": 405.0, "16": 330.0, "17": 450.0, "18": 390.0, "19": 465.0, "20": 240.0, "22": 350.0, "23": 390.0, "24": 460.0, "25": 380.0, "26": 340.0, "27": 205.0, "29": 275.0, "30": 325.0}}, "7008112XKN61A8P": {"5": {"5": 240.0, "6": 120.0, "7": 240.0, "8": 120.0, "11": 120.0, "12": 195.0, "13": 190.0, "14": 95.0, "15": 225.0, "16": 95.0, "18": 70.0, "19": 210.0, "20": 240.0, "21": 120.0, "22": 120.0, "25": 120.0, "26": 120.0, "28": 145.0, "29": 215.0, "30": 10.0}, "6": {"1": 350.0, "2": 360.0, "3": 120.0, "6": 240.0, "7": 120.0, "8": 240.0, "9": 315.0, "10": 245.0, "11": 235.0, "13": 120.0, "15": 165.0, "16": 120.0, "17": 240.0, "22": 120.0, "23": 120.0, "24": 50.0, "25": 70.0, "30": 120.0}}, "7008113XKN61A8P": {"5": {"5": 240.0, "6": 120.0, "7": 240.0, "8": 120.0, "11": 120.0, "12": 195.0, "13": 190.0, "14": 95.0, "15": 225.0, "16": 95.0, "18": 70.0, "19": 210.0, "20": 240.0, "21": 120.0, "22": 120.0, "25": 120.0, "26": 120.0, "28": 145.0, "29": 215.0, "30": 10.0}, "6": {"1": 350.0, "2": 360.0, "3": 120.0, "6": 240.0, "7": 120.0, "8": 240.0, "9": 315.0, "10": 245.0, "11": 235.0, "13": 120.0, "15": 165.0, "16": 120.0, "17": 240.0, "22": 120.0, "23": 120.0, "24": 50.0, "25": 70.0, "30": 120.0}}, "7008120XST11A8P": {"5": {"4": 240.0, "5": 220.0, "6": 260.0, "7": 240.0, "8": 300.0, "11": 180.0, "12": 240.0, "13": 240.0, "14": 360.0, "15": 240.0, "16": 120.0, "18": 360.0, "19": 240.0, "20": 230.0, "21": 250.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 365.0, "27": 435.0, "28": 285.0, "29": 275.0, "30": 205.0}, "6": {"1": 710.0, "2": 510.0, "3": 655.0, "4": 390.0, "5": 390.0, "6": 375.0, "7": 310.0, "8": 640.0, "9": 625.0, "10": 695.0, "11": 490.0, "13": 125.0, "15": 405.0, "16": 330.0, "17": 450.0, "18": 390.0, "19": 465.0, "20": 240.0, "22": 350.0, "23": 390.0, "24": 460.0, "25": 380.0, "26": 340.0, "27": 205.0, "29": 275.0, "30": 325.0}}, "7008120XST11AQT": {"5": {"4": 240.0, "5": 220.0, "6": 260.0, "7": 240.0, "8": 300.0, "11": 180.0, "12": 240.0, "13": 240.0, "14": 360.0, "15": 240.0, "16": 120.0, "18": 360.0, "19": 240.0, "20": 230.0, "21": 250.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 365.0, "27": 435.0, "28": 285.0, "29": 275.0, "30": 205.0}, "6": {"1": 710.0, "2": 510.0, "3": 655.0, "4": 390.0, "5": 390.0, "6": 375.0, "7": 310.0, "8": 640.0, "9": 625.0, "10": 695.0, "11": 490.0, "13": 125.0, "15": 405.0, "16": 330.0, "17": 450.0, "18": 390.0, "19": 465.0, "20": 240.0, "22": 350.0, "23": 390.0, "24": 460.0, "25": 380.0, "26": 340.0, "27": 205.0, "29": 275.0, "30": 325.0}}, "7008120XST11AZ6": {"5": {"4": 240.0, "5": 220.0, "6": 260.0, "7": 240.0, "8": 300.0, "11": 180.0, "12": 240.0, "13": 240.0, "14": 360.0, "15": 240.0, "16": 120.0, "18": 360.0, "19": 240.0, "20": 230.0, "21": 250.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 365.0, "27": 435.0, "28": 285.0, "29": 275.0, "30": 205.0}, "6": {"1": 710.0, "2": 510.0, "3": 655.0, "4": 390.0, "5": 390.0, "6": 375.0, "7": 310.0, "8": 640.0, "9": 625.0, "10": 695.0, "11": 490.0, "13": 125.0, "15": 405.0, "16": 330.0, "17": 450.0, "18": 390.0, "19": 465.0, "20": 240.0, "22": 350.0, "23": 390.0, "24": 460.0, "25": 380.0, "26": 340.0, "27": 205.0, "29": 275.0, "30": 325.0}}, "7008200XST11A8P": {"5": {"4": 240.0, "5": 220.0, "6": 260.0, "7": 240.0, "8": 300.0, "11": 180.0, "12": 240.0, "13": 240.0, "14": 360.0, "15": 240.0, "16": 120.0, "18": 360.0, "19": 240.0, "20": 230.0, "21": 250.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 365.0, "27": 435.0, "28": 285.0, "29": 275.0, "30": 205.0}, "6": {"1": 710.0, "2": 510.0, "3": 655.0, "4": 390.0, "5": 390.0, "6": 375.0, "7": 310.0, "8": 640.0, "9": 625.0, "10": 695.0, "11": 490.0, "13": 125.0, "15": 405.0, "16": 330.0, "17": 450.0, "18": 390.0, "19": 465.0, "20": 240.0, "22": 350.0, "23": 390.0, "24": 460.0, "25": 380.0, "26": 340.0, "27": 205.0, "29": 275.0, "30": 325.0}}, "7008200XST11AZ6": {"5": {"4": 240.0, "5": 220.0, "6": 260.0, "7": 240.0, "8": 300.0, "11": 180.0, "12": 240.0, "13": 240.0, "14": 360.0, "15": 240.0, "16": 120.0, "18": 360.0, "19": 240.0, "20": 230.0, "21": 250.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 365.0, "27": 435.0, "28": 285.0, "29": 275.0, "30": 205.0}, "6": {"1": 710.0, "2": 510.0, "3": 655.0, "4": 390.0, "5": 390.0, "6": 375.0, "7": 310.0, "8": 640.0, "9": 625.0, "10": 695.0, "11": 490.0, "13": 125.0, "15": 405.0, "16": 330.0, "17": 450.0, "18": 390.0, "19": 465.0, "20": 240.0, "22": 350.0, "23": 390.0, "24": 460.0, "25": 380.0, "26": 340.0, "27": 205.0, "29": 275.0, "30": 325.0}}, "7925100XKJ23A": {"5": {"4": 190.0, "5": 90.0, "6": 170.0, "7": 70.0, "8": 60.0, "11": 65.0, "12": 115.0, "13": 120.0, "14": 95.0, "15": 85.0, "16": 60.0, "18": 120.0, "19": 100.0, "20": 80.0, "21": 180.0, "22": 70.0, "23": 50.0, "25": 60.0, "26": 65.0, "27": 115.0, "28": 120.0, "29": 60.0, "30": 60.0}, "6": {"1": 60.0, "2": 120.0, "3": 60.0, "4": 60.0, "5": 60.0, "6": 60.0, "8": 60.0, "9": 120.0, "10": 60.0}}, "7925100XKN83A": {"5": {"4": 60.0, "6": 60.0, "13": 85.0, "14": 35.0, "15": 60.0, "16": 10.0, "18": 50.0, "19": 10.0, "20": 90.0, "21": 30.0, "22": 50.0, "23": 25.0, "25": 90.0, "26": 30.0, "27": 35.0, "28": 25.0, "29": 35.0}, "6": {"1": 85.0, "2": 195.0, "3": 90.0, "4": 90.0, "5": 80.0, "6": 30.0, "7": 90.0, "10": 170.0, "11": 10.0, "13": 25.0, "14": 95.0, "16": 110.0, "17": 65.0, "19": 25.0, "20": 35.0, "24": 40.0, "25": 20.0, "29": 55.0, "30": 5.0}}, "7925100XST33A": {"5": {"4": 240.0, "5": 220.0, "6": 260.0, "7": 240.0, "8": 300.0, "11": 180.0, "12": 240.0, "13": 240.0, "14": 360.0, "15": 240.0, "16": 120.0, "18": 360.0, "19": 240.0, "20": 230.0, "21": 250.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 365.0, "27": 435.0, "28": 285.0, "29": 275.0, "30": 205.0}, "6": {"1": 710.0, "2": 510.0, "3": 655.0, "4": 390.0, "5": 390.0, "6": 375.0, "7": 310.0, "8": 640.0, "9": 625.0, "10": 695.0, "11": 490.0, "13": 125.0, "15": 405.0, "16": 330.0, "17": 450.0, "18": 390.0, "19": 465.0, "20": 240.0, "22": 350.0, "23": 390.0, "24": 460.0, "25": 380.0, "26": 340.0, "27": 205.0, "29": 275.0, "30": 325.0}}, "7925101XKJ23A": {"5": {"4": 190.0, "5": 90.0, "6": 170.0, "7": 70.0, "8": 60.0, "11": 65.0, "12": 115.0, "13": 120.0, "14": 95.0, "15": 85.0, "16": 60.0, "18": 120.0, "19": 100.0, "20": 80.0, "21": 180.0, "22": 70.0, "23": 50.0, "25": 60.0, "26": 65.0, "27": 115.0, "28": 120.0, "29": 60.0, "30": 60.0}, "6": {"1": 60.0, "2": 120.0, "3": 60.0, "4": 60.0, "5": 60.0, "6": 60.0, "8": 60.0, "9": 120.0, "10": 60.0}}, "7925101XKN83A": {"5": {"4": 60.0, "6": 60.0, "13": 85.0, "14": 35.0, "15": 60.0, "16": 10.0, "18": 50.0, "19": 10.0, "20": 90.0, "21": 30.0, "22": 50.0, "23": 25.0, "25": 90.0, "26": 30.0, "27": 35.0, "28": 25.0, "29": 35.0}, "6": {"1": 85.0, "2": 195.0, "3": 90.0, "4": 90.0, "5": 80.0, "6": 30.0, "7": 90.0, "10": 170.0, "11": 10.0, "13": 25.0, "14": 95.0, "16": 110.0, "17": 65.0, "19": 25.0, "20": 35.0, "24": 40.0, "25": 20.0, "29": 55.0, "30": 5.0}}, "7925102XKN46A": {"5": {"4": 30.0, "5": 90.0, "6": 30.0, "7": 90.0, "8": 80.0, "11": 75.0, "12": 90.0, "13": 5.0, "14": 55.0, "15": 30.0, "16": 35.0, "18": 40.0, "19": 80.0, "21": 60.0, "22": 40.0, "23": 20.0, "26": 60.0, "27": 55.0, "28": 65.0, "29": 55.0, "30": 45.0}, "6": {"1": 20.0, "17": 60.0, "19": 60.0, "20": 10.0, "22": 90.0, "23": 20.0}}, "7925102XKN61A": {"5": {"5": 240.0, "6": 120.0, "7": 240.0, "8": 120.0, "11": 120.0, "12": 195.0, "13": 190.0, "14": 95.0, "15": 225.0, "16": 95.0, "18": 70.0, "19": 210.0, "20": 240.0, "21": 120.0, "22": 120.0, "25": 120.0, "26": 120.0, "28": 145.0, "29": 215.0, "30": 10.0}, "6": {"1": 350.0, "2": 360.0, "3": 120.0, "6": 240.0, "7": 120.0, "8": 240.0, "9": 315.0, "10": 245.0, "11": 235.0, "13": 120.0, "15": 165.0, "16": 120.0, "17": 240.0, "22": 120.0, "23": 120.0, "24": 50.0, "25": 70.0, "30": 120.0}}, "7925102XKN83A": {"5": {"4": 60.0, "6": 60.0, "13": 85.0, "14": 35.0, "15": 60.0, "16": 10.0, "18": 50.0, "19": 10.0, "20": 90.0, "21": 30.0, "22": 50.0, "23": 25.0, "25": 90.0, "26": 30.0, "27": 35.0, "28": 25.0, "29": 35.0}, "6": {"1": 85.0, "2": 195.0, "3": 90.0, "4": 90.0, "5": 80.0, "6": 30.0, "7": 90.0, "10": 170.0, "11": 10.0, "13": 25.0, "14": 95.0, "16": 110.0, "17": 65.0, "19": 25.0, "20": 35.0, "24": 40.0, "25": 20.0, "29": 55.0, "30": 5.0}}, "7925103XKN46A": {"5": {"4": 30.0, "5": 90.0, "6": 30.0, "7": 90.0, "8": 80.0, "11": 75.0, "12": 90.0, "13": 5.0, "14": 55.0, "15": 30.0, "16": 35.0, "18": 40.0, "19": 80.0, "21": 60.0, "22": 40.0, "23": 20.0, "26": 60.0, "27": 55.0, "28": 65.0, "29": 55.0, "30": 45.0}, "6": {"1": 20.0, "17": 60.0, "19": 60.0, "20": 10.0, "22": 90.0, "23": 20.0}}, "7925104XKQ41B": {"5": {"4": 520.0, "5": 640.0, "6": 640.0, "7": 640.0, "8": 560.0, "11": 440.0, "12": 640.0, "13": 640.0, "14": 640.0, "15": 640.0, "16": 320.0, "18": 640.0, "19": 640.0, "20": 640.0, "21": 640.0, "22": 640.0, "23": 320.0, "25": 640.0, "26": 640.0, "27": 640.0, "28": 640.0, "29": 640.0, "30": 320.0}, "6": {"1": 1225.0, "2": 1185.0, "3": 925.0, "4": 540.0, "5": 530.0, "6": 705.0, "7": 520.0, "8": 940.0, "9": 1060.0, "10": 1170.0, "11": 735.0, "13": 270.0, "14": 95.0, "15": 570.0, "16": 560.0, "17": 815.0, "18": 390.0, "19": 550.0, "20": 285.0, "22": 560.0, "23": 530.0, "24": 550.0, "25": 470.0, "26": 340.0, "27": 205.0, "29": 330.0, "30": 450.0}}, "7925105XKQ41A": {"5": {"4": 520.0, "5": 640.0, "6": 640.0, "7": 640.0, "8": 560.0, "11": 440.0, "12": 640.0, "13": 640.0, "14": 640.0, "15": 640.0, "16": 320.0, "18": 640.0, "19": 640.0, "20": 640.0, "21": 640.0, "22": 640.0, "23": 320.0, "25": 640.0, "26": 640.0, "27": 640.0, "28": 640.0, "29": 640.0, "30": 320.0}, "6": {"1": 1225.0, "2": 1185.0, "3": 925.0, "4": 540.0, "5": 530.0, "6": 705.0, "7": 520.0, "8": 940.0, "9": 1060.0, "10": 1170.0, "11": 735.0, "13": 270.0, "14": 95.0, "15": 570.0, "16": 560.0, "17": 815.0, "18": 390.0, "19": 550.0, "20": 285.0, "22": 560.0, "23": 530.0, "24": 550.0, "25": 470.0, "26": 340.0, "27": 205.0, "29": 330.0, "30": 450.0}}, "7925108XST11A": {"5": {"4": 240.0, "5": 220.0, "6": 260.0, "7": 240.0, "8": 300.0, "11": 180.0, "12": 240.0, "13": 240.0, "14": 360.0, "15": 240.0, "16": 120.0, "18": 360.0, "19": 240.0, "20": 230.0, "21": 250.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 365.0, "27": 435.0, "28": 285.0, "29": 275.0, "30": 205.0}, "6": {"1": 710.0, "2": 510.0, "3": 655.0, "4": 390.0, "5": 390.0, "6": 375.0, "7": 310.0, "8": 640.0, "9": 625.0, "10": 695.0, "11": 490.0, "13": 125.0, "15": 405.0, "16": 330.0, "17": 450.0, "18": 390.0, "19": 465.0, "20": 240.0, "22": 350.0, "23": 390.0, "24": 460.0, "25": 380.0, "26": 340.0, "27": 205.0, "29": 275.0, "30": 325.0}}, "7925109XKN61A": {"5": {"5": 240.0, "6": 120.0, "7": 240.0, "8": 120.0, "11": 120.0, "12": 195.0, "13": 190.0, "14": 95.0, "15": 225.0, "16": 95.0, "18": 70.0, "19": 210.0, "20": 240.0, "21": 120.0, "22": 120.0, "25": 120.0, "26": 120.0, "28": 145.0, "29": 215.0, "30": 10.0}, "6": {"1": 350.0, "2": 360.0, "3": 120.0, "6": 240.0, "7": 120.0, "8": 240.0, "9": 315.0, "10": 245.0, "11": 235.0, "13": 120.0, "15": 165.0, "16": 120.0, "17": 240.0, "22": 120.0, "23": 120.0, "24": 50.0, "25": 70.0, "30": 120.0}}, "7925109XST11A": {"5": {"6": 120.0, "7": 120.0, "11": 120.0, "14": 120.0, "15": 120.0, "25": 115.0, "26": 5.0, "29": 240.0, "30": 120.0}, "6": {"1": 515.0, "2": 85.0, "3": 120.0, "5": 120.0, "6": 160.0, "7": 80.0, "9": 250.0, "10": 230.0, "29": 240.0}}, "7925113XST11A": {"5": {"6": 120.0, "7": 120.0, "11": 120.0, "14": 120.0, "15": 120.0, "25": 115.0, "26": 5.0, "29": 240.0, "30": 120.0}, "6": {"1": 515.0, "2": 85.0, "3": 120.0, "5": 120.0, "6": 160.0, "7": 80.0, "9": 250.0, "10": 230.0, "29": 240.0}}, "8400131XGW01A": {"5": {"4": 240.0, "5": 120.0, "7": 120.0, "8": 240.0, "11": 65.0, "12": 190.0, "13": 105.0, "14": 215.0, "15": 145.0, "18": 120.0, "19": 240.0, "21": 120.0, "22": 120.0, "25": 120.0, "27": 120.0, "29": 120.0, "30": 120.0}, "6": {"1": 240.0, "4": 120.0, "5": 160.0, "6": 80.0, "8": 165.0, "9": 75.0, "10": 165.0, "11": 75.0, "15": 120.0, "16": 240.0, "22": 110.0, "23": 30.0, "24": 100.0, "29": 120.0, "30": 120.0}}, "8400138XGW01B": {"5": {"4": 240.0, "5": 120.0, "7": 120.0, "8": 240.0, "11": 65.0, "12": 190.0, "13": 105.0, "14": 215.0, "15": 145.0, "18": 120.0, "19": 240.0, "21": 120.0, "22": 120.0, "25": 120.0, "27": 120.0, "29": 120.0, "30": 120.0}, "6": {"1": 240.0, "4": 120.0, "5": 160.0, "6": 80.0, "8": 165.0, "9": 75.0, "10": 165.0, "11": 75.0, "15": 120.0, "16": 240.0, "22": 110.0, "23": 30.0, "24": 100.0, "29": 120.0, "30": 120.0}}, "8400141XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "8400142XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "8400156XGW01B": {"5": {"4": 240.0, "5": 120.0, "7": 120.0, "8": 240.0, "11": 65.0, "12": 190.0, "13": 105.0, "14": 215.0, "15": 145.0, "18": 120.0, "19": 240.0, "21": 120.0, "22": 120.0, "25": 120.0, "27": 120.0, "29": 120.0, "30": 120.0}, "6": {"1": 240.0, "4": 120.0, "5": 160.0, "6": 80.0, "8": 165.0, "9": 75.0, "10": 165.0, "11": 75.0, "15": 120.0, "16": 240.0, "22": 110.0, "23": 30.0, "24": 100.0, "29": 120.0, "30": 120.0}}, "8400157XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "8400158XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "8400171XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "8400172XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "8400173XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "8400174XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "8400181XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "8400195XGW01B": {"5": {"4": 240.0, "5": 120.0, "7": 120.0, "8": 240.0, "11": 65.0, "12": 190.0, "13": 105.0, "14": 215.0, "15": 145.0, "18": 120.0, "19": 240.0, "21": 120.0, "22": 120.0, "25": 120.0, "27": 120.0, "29": 120.0, "30": 120.0}, "6": {"1": 240.0, "4": 120.0, "5": 160.0, "6": 80.0, "8": 165.0, "9": 75.0, "10": 165.0, "11": 75.0, "15": 120.0, "16": 240.0, "22": 110.0, "23": 30.0, "24": 100.0, "29": 120.0, "30": 120.0}}, "8400202XGW01B": {"5": {"4": 240.0, "5": 120.0, "7": 120.0, "8": 240.0, "11": 65.0, "12": 190.0, "13": 105.0, "14": 215.0, "15": 145.0, "18": 120.0, "19": 240.0, "21": 120.0, "22": 120.0, "25": 120.0, "27": 120.0, "29": 120.0, "30": 120.0}, "6": {"1": 240.0, "4": 120.0, "5": 160.0, "6": 80.0, "8": 165.0, "9": 75.0, "10": 165.0, "11": 75.0, "15": 120.0, "16": 240.0, "22": 110.0, "23": 30.0, "24": 100.0, "29": 120.0, "30": 120.0}}, "8400220XGW01B": {"5": {"4": 240.0, "5": 120.0, "7": 120.0, "8": 240.0, "11": 65.0, "12": 190.0, "13": 105.0, "14": 215.0, "15": 145.0, "18": 120.0, "19": 240.0, "21": 120.0, "22": 120.0, "25": 120.0, "27": 120.0, "29": 120.0, "30": 120.0}, "6": {"1": 240.0, "4": 120.0, "5": 160.0, "6": 80.0, "8": 165.0, "9": 75.0, "10": 165.0, "11": 75.0, "15": 120.0, "16": 240.0, "22": 110.0, "23": 30.0, "24": 100.0, "29": 120.0, "30": 120.0}}, "8400323XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "8400324XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "8400331XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "8400332XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "8400343XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "8400344XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "8400371XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "8400372XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "8400411XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "8400412XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "8400413XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "8400414XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "8400421XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "8400422XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "8400451XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "8400452XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "8400473XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "8400474XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "8400751XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "8400752XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "8402105AKJ20A": {"5": {"4": 190.0, "5": 90.0, "6": 170.0, "7": 70.0, "8": 60.0, "11": 65.0, "12": 115.0, "13": 120.0, "14": 95.0, "15": 85.0, "16": 60.0, "18": 120.0, "19": 100.0, "20": 80.0, "21": 180.0, "22": 70.0, "23": 50.0, "25": 60.0, "26": 65.0, "27": 115.0, "28": 120.0, "29": 60.0, "30": 60.0}, "6": {"1": 60.0, "2": 120.0, "3": 60.0, "4": 60.0, "5": 60.0, "6": 60.0, "8": 60.0, "9": 120.0, "10": 60.0}}, "8402106AKJ20A": {"5": {"4": 190.0, "5": 90.0, "6": 170.0, "7": 70.0, "8": 60.0, "11": 65.0, "12": 115.0, "13": 120.0, "14": 95.0, "15": 85.0, "16": 60.0, "18": 120.0, "19": 100.0, "20": 80.0, "21": 180.0, "22": 70.0, "23": 50.0, "25": 60.0, "26": 65.0, "27": 115.0, "28": 120.0, "29": 60.0, "30": 60.0}, "6": {"1": 60.0, "2": 120.0, "3": 60.0, "4": 60.0, "5": 60.0, "6": 60.0, "8": 60.0, "9": 120.0, "10": 60.0}}, "8402110AST01A": {"5": {"4": 240.0, "5": 220.0, "6": 260.0, "7": 240.0, "8": 300.0, "11": 180.0, "12": 240.0, "13": 240.0, "14": 360.0, "15": 240.0, "16": 120.0, "18": 360.0, "19": 240.0, "20": 230.0, "21": 250.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 365.0, "27": 435.0, "28": 285.0, "29": 275.0, "30": 205.0}, "6": {"1": 710.0, "2": 510.0, "3": 655.0, "4": 390.0, "5": 390.0, "6": 375.0, "7": 310.0, "8": 640.0, "9": 625.0, "10": 695.0, "11": 490.0, "13": 125.0, "15": 405.0, "16": 330.0, "17": 450.0, "18": 390.0, "19": 465.0, "20": 240.0, "22": 350.0, "23": 390.0, "24": 460.0, "25": 380.0, "26": 340.0, "27": 205.0, "29": 275.0, "30": 325.0}}, "8402115XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "8402116XST11A": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "ALAA001234": {"5": {"4": 24.0, "5": 48.0, "6": 24.0, "8": 24.0, "11": 24.0, "12": 39.0, "13": 38.0, "14": 19.0, "15": 45.0, "16": 19.0, "18": 8.0, "19": 24.0, "20": 48.0, "22": 24.0, "25": 24.0, "26": 24.0, "28": 24.0, "30": 2.0}, "6": {"1": 70.0, "2": 50.0, "3": 46.0, "4": 66.0, "5": 24.0, "6": 30.0, "8": 72.0, "9": 39.0, "10": 9.0, "11": 39.0, "13": 13.0, "14": 11.0, "15": 9.0, "16": 48.0, "17": 24.0, "18": 24.0, "23": 24.0, "24": 10.0, "25": 38.0, "26": 22.0, "27": 2.0, "29": 24.0, "30": 48.0}}, "ALAA003255": {"5": {"4": 391.58, "5": 404.65, "6": 237.56, "7": 393.29, "8": 445.57, "11": 355.21, "12": 518.6, "13": 558.77, "14": 576.48, "15": 538.5, "16": 331.05, "18": 543.14, "19": 532.15, "20": 402.19, "21": 528.17, "22": 252.91, "23": 269.67, "25": 596.75, "26": 656.24, "27": 449.74, "28": 525.42, "29": 388.17, "30": 264.84}, "6": {"1": 946.84, "2": 1040.05, "3": 1021.86, "4": 580.84, "5": 402.95, "6": 755.6, "7": 600.16, "8": 849.37, "9": 829.77, "10": 785.44, "11": 853.83, "13": 271.1, "14": 229.7, "15": 857.9, "16": 1138.85, "17": 752.76, "18": 163.21, "19": 548.54, "20": 276.21, "22": 661.16, "23": 452.39, "24": 117.08, "25": 518.32, "26": 606.98, "27": 262.0, "29": 540.49, "30": 325.09}}, "ALAA003256": {"5": {"4": 325.09, "5": 335.12, "6": 508.58, "7": 346.92, "8": 198.24, "11": 150.45, "12": 216.82, "13": 175.13, "14": 156.74, "15": 196.18, "16": 33.92, "18": 191.36, "19": 202.76, "20": 337.68, "21": 206.89, "22": 492.65, "23": 97.64, "25": 135.7, "26": 73.95, "27": 288.31, "28": 209.74, "29": 352.23, "30": 102.66}, "6": {"1": 574.66, "2": 477.9, "3": 499.14, "4": 553.42, "5": 360.49, "6": 372.0, "7": 155.76, "8": 675.84, "9": 696.2, "10": 746.94, "11": 270.02, "13": 497.37, "14": 540.34, "15": 667.0, "16": 375.34, "17": 97.64, "18": 609.37, "19": 209.35, "20": 90.86, "22": 92.43, "23": 309.16, "24": 657.26, "25": 240.72, "26": 148.68, "27": 105.61, "29": 217.71, "30": 441.32}}, "ALAA003257": {"5": {"4": 322.88, "5": 321.3, "6": 324.45, "7": 718.2, "8": 989.1, "11": 417.38, "12": 652.05, "13": 296.63, "14": 247.54, "15": 718.99, "16": 72.45, "18": 725.03, "19": 678.3, "20": 326.55, "21": 687.75, "22": 144.9, "23": 155.93, "25": 237.83, "26": 870.71, "27": 269.59, "28": 587.48, "29": 326.81, "30": 150.41}, "6": {"1": 1004.06, "2": 1673.44, "3": 1589.18, "4": 837.9, "5": 255.15, "6": 1339.54, "7": 250.43, "8": 507.94, "9": 906.41, "10": 1437.19, "11": 1750.88, "13": 321.3, "14": 279.3, "15": 665.44, "16": 1941.45, "17": 383.51, "18": 230.48, "19": 1295.96, "20": 523.69, "22": 183.75, "23": 220.5, "24": 166.95, "25": 1195.43, "26": 628.43, "27": 512.66, "29": 285.86, "30": 557.55}}, "ALAA003286": {"5": {"4": 80.98, "5": 80.58, "6": 81.37, "7": 180.12, "8": 248.06, "11": 104.67, "12": 163.53, "13": 74.39, "14": 62.08, "15": 180.32, "16": 18.17, "18": 181.83, "19": 170.11, "20": 81.9, "21": 172.48, "22": 36.34, "23": 39.1, "25": 59.64, "26": 218.37, "27": 67.61, "28": 147.33, "29": 81.96, "30": 37.72}, "6": {"1": 251.81, "2": 419.69, "3": 398.56, "4": 210.14, "5": 63.99, "6": 335.95, "7": 62.8, "8": 127.39, "9": 227.32, "10": 360.44, "11": 439.11, "13": 80.58, "14": 70.05, "15": 166.89, "16": 486.9, "17": 96.18, "18": 57.8, "19": 325.02, "20": 131.34, "22": 46.08, "23": 55.3, "24": 41.87, "25": 299.8, "26": 157.6, "27": 128.57, "29": 71.69, "30": 139.83}}, "ALAA003287": {"5": {"4": 967.2, "5": 998.4, "6": 998.4, "7": 998.4, "8": 873.6, "11": 686.4, "12": 998.4, "13": 998.4, "14": 998.4, "15": 998.4, "16": 499.2, "18": 998.4, "19": 998.4, "20": 998.4, "21": 998.4, "22": 998.4, "23": 499.2, "25": 998.4, "26": 998.4, "27": 998.4, "28": 998.4, "29": 998.4, "30": 499.2}, "6": {"1": 2059.2, "2": 2059.2, "3": 2059.2, "4": 1528.8, "5": 1029.6, "6": 1528.8, "7": 1029.6, "8": 2059.2, "9": 2059.2, "10": 2059.2, "11": 1528.8, "13": 1029.6, "14": 1029.6, "15": 2059.2, "16": 2059.2, "17": 1162.2, "18": 1029.6, "19": 1029.6, "20": 499.2, "22": 1029.6, "23": 1029.6, "24": 1029.6, "25": 1029.6, "26": 1029.6, "27": 499.2, "29": 1029.6, "30": 1029.6}}, "ALAA003288": {"5": {"4": 380.27, "5": 392.53, "6": 392.53, "7": 392.53, "8": 343.47, "11": 269.87, "12": 392.53, "13": 392.53, "14": 392.53, "15": 392.53, "16": 196.27, "18": 392.53, "19": 392.53, "20": 392.53, "21": 392.53, "22": 392.53, "23": 196.27, "25": 392.53, "26": 392.53, "27": 392.53, "28": 392.53, "29": 392.53, "30": 196.27}, "6": {"1": 809.6, "2": 809.6, "3": 809.6, "4": 601.07, "5": 404.8, "6": 601.07, "7": 404.8, "8": 809.6, "9": 809.6, "10": 809.6, "11": 601.07, "13": 404.8, "14": 404.8, "15": 809.6, "16": 809.6, "17": 456.93, "18": 404.8, "19": 404.8, "20": 196.27, "22": 404.8, "23": 404.8, "24": 404.8, "25": 404.8, "26": 404.8, "27": 196.27, "29": 404.8, "30": 404.8}}, "ALAA003289": {"5": {"4": 49.6, "5": 51.2, "6": 51.2, "7": 51.2, "8": 44.8, "11": 35.2, "12": 51.2, "13": 51.2, "14": 51.2, "15": 51.2, "16": 25.6, "18": 51.2, "19": 51.2, "20": 51.2, "21": 51.2, "22": 51.2, "23": 25.6, "25": 51.2, "26": 51.2, "27": 51.2, "28": 51.2, "29": 51.2, "30": 25.6}, "6": {"1": 105.6, "2": 105.6, "3": 105.6, "4": 78.4, "5": 52.8, "6": 78.4, "7": 52.8, "8": 105.6, "9": 105.6, "10": 105.6, "11": 78.4, "13": 52.8, "14": 52.8, "15": 105.6, "16": 105.6, "17": 59.6, "18": 52.8, "19": 52.8, "20": 25.6, "22": 52.8, "23": 52.8, "24": 52.8, "25": 52.8, "26": 52.8, "27": 25.6, "29": 52.8, "30": 52.8}}, "ALAA003338": {"5": {"4": 5642.0, "5": 5824.0, "6": 5824.0, "7": 5824.0, "8": 5096.0, "11": 4004.0, "12": 5824.0, "13": 5824.0, "14": 5824.0, "15": 5824.0, "16": 2912.0, "18": 5824.0, "19": 5824.0, "20": 5824.0, "21": 5824.0, "22": 5824.0, "23": 2912.0, "25": 5824.0, "26": 5824.0, "27": 5824.0, "28": 5824.0, "29": 5824.0, "30": 2912.0}, "6": {"1": 12012.0, "2": 12012.0, "3": 12012.0, "4": 8918.0, "5": 6006.0, "6": 8918.0, "7": 6006.0, "8": 12012.0, "9": 12012.0, "10": 12012.0, "11": 8918.0, "13": 6006.0, "14": 6006.0, "15": 12012.0, "16": 12012.0, "17": 6779.5, "18": 6006.0, "19": 6006.0, "20": 2912.0, "22": 6006.0, "23": 6006.0, "24": 6006.0, "25": 6006.0, "26": 6006.0, "27": 2912.0, "29": 6006.0, "30": 6006.0}}, "ALAA003339": {"5": {"4": 930.0, "5": 960.0, "6": 960.0, "7": 960.0, "8": 840.0, "11": 660.0, "12": 960.0, "13": 960.0, "14": 960.0, "15": 960.0, "16": 480.0, "18": 960.0, "19": 960.0, "20": 960.0, "21": 960.0, "22": 960.0, "23": 480.0, "25": 960.0, "26": 960.0, "27": 960.0, "28": 960.0, "29": 960.0, "30": 480.0}, "6": {"1": 1980.0, "2": 1980.0, "3": 1980.0, "4": 1470.0, "5": 990.0, "6": 1470.0, "7": 990.0, "8": 1980.0, "9": 1980.0, "10": 1980.0, "11": 1470.0, "13": 990.0, "14": 990.0, "15": 1980.0, "16": 1980.0, "17": 1117.5, "18": 990.0, "19": 990.0, "20": 480.0, "22": 990.0, "23": 990.0, "24": 990.0, "25": 990.0, "26": 990.0, "27": 480.0, "29": 990.0, "30": 990.0}}, "ALAA003526": {"5": {"4": 550.0, "5": 640.0, "6": 640.0, "7": 640.0, "8": 560.0, "11": 440.0, "12": 640.0, "13": 640.0, "14": 640.0, "15": 640.0, "16": 320.0, "18": 640.0, "19": 640.0, "20": 640.0, "21": 640.0, "22": 640.0, "23": 320.0, "25": 640.0, "26": 640.0, "27": 640.0, "28": 640.0, "29": 640.0, "30": 320.0}, "6": {"1": 1175.0, "2": 1125.0, "3": 680.0, "4": 745.0, "5": 540.0, "6": 735.0, "7": 500.0, "8": 940.0, "9": 1190.0, "10": 1020.0, "11": 750.0, "13": 95.0, "14": 25.0, "15": 560.0, "16": 825.0, "17": 535.0, "18": 455.0, "19": 545.0, "20": 315.0, "22": 490.0, "23": 530.0, "24": 610.0, "25": 360.0, "26": 360.0, "27": 155.0, "29": 330.0, "30": 570.0}}, "ALAA004586": {"5": {"4": 640.15, "5": 630.51, "6": 1054.53, "7": 677.32, "8": 357.93, "11": 202.37, "12": 340.04, "13": 307.46, "14": 268.68, "15": 313.19, "16": 72.96, "18": 366.65, "19": 329.94, "20": 680.53, "21": 367.11, "22": 1116.48, "23": 165.2, "25": 229.22, "26": 157.4, "27": 597.01, "28": 386.16, "29": 702.1, "30": 166.58}, "6": {"1": 1032.5, "2": 861.79, "3": 859.04, "4": 1149.52, "5": 744.78, "6": 633.96, "7": 276.71, "8": 1289.25, "9": 1326.42, "10": 1495.75, "11": 391.89, "13": 1057.28, "14": 1233.95, "15": 1316.78, "16": 716.1, "17": 181.72, "18": 1273.19, "19": 457.51, "20": 161.76, "22": 193.65, "23": 630.51, "24": 1401.45, "25": 408.87, "26": 243.67, "27": 181.72, "29": 366.19, "30": 893.46}}, "ALAA004871": {"5": {"13": 414.0, "14": 207.0, "16": 207.0, "25": 414.0}, "6": {"3": 207.0, "7": 207.0, "8": 828.0, "17": 414.0, "26": 207.0}}, "ALAA004873": {"5": {"12": 414.0, "14": 207.0, "16": 207.0, "29": 414.0}, "6": {"3": 207.0, "7": 207.0, "8": 207.0, "16": 414.0, "17": 414.0, "26": 207.0}}, "ALAA005672": {"5": {"4": 20.7, "5": 28.98, "6": 124.2, "7": 24.84, "8": 57.96, "11": 124.2, "12": 79.69, "13": 23.8, "14": 46.57, "15": 22.77, "16": 5.17, "18": 12.42, "19": 49.68, "20": 24.84, "21": 37.26, "22": 33.12, "23": 11.38, "25": 32.08, "26": 22.77, "27": 14.49, "28": 80.73, "29": 68.31, "30": 31.05}, "6": {"1": 120.06, "2": 95.22, "3": 120.06, "4": 89.01, "5": 37.26, "6": 82.8, "7": 55.89, "8": 194.58, "9": 182.16, "10": 175.95, "11": 119.02, "13": 55.89, "14": 40.36, "15": 102.46, "16": 87.97, "17": 47.61, "18": 124.2, "19": 46.58, "20": 1.03, "22": 33.12, "23": 37.26, "24": 99.36, "25": 130.41, "26": 55.89, "27": 68.31, "29": 43.47, "30": 105.57}}, "ALAA006324": {"5": {"4": 74.4, "5": 76.8, "6": 76.8, "7": 76.8, "8": 67.2, "11": 52.8, "12": 76.8, "13": 76.8, "14": 76.8, "15": 76.8, "16": 38.4, "18": 76.8, "19": 76.8, "20": 76.8, "21": 76.8, "22": 76.8, "23": 38.4, "25": 76.8, "26": 76.8, "27": 76.8, "28": 76.8, "29": 76.8, "30": 38.4}, "6": {"1": 158.4, "2": 158.4, "3": 158.4, "4": 117.6, "5": 79.2, "6": 117.6, "7": 79.2, "8": 158.4, "9": 158.4, "10": 158.4, "11": 117.6, "13": 79.2, "14": 79.2, "15": 158.4, "16": 158.4, "17": 89.4, "18": 79.2, "19": 79.2, "20": 38.4, "22": 79.2, "23": 79.2, "24": 79.2, "25": 79.2, "26": 79.2, "27": 38.4, "29": 79.2, "30": 79.2}}, "ALAA006522": {"5": {"4": 161.2, "5": 166.4, "6": 166.4, "7": 166.4, "8": 145.6, "11": 114.4, "12": 166.4, "13": 166.4, "14": 166.4, "15": 166.4, "16": 83.2, "18": 166.4, "19": 166.4, "20": 166.4, "21": 166.4, "22": 166.4, "23": 83.2, "25": 166.4, "26": 166.4, "27": 166.4, "28": 166.4, "29": 166.4, "30": 83.2}, "6": {"1": 343.2, "2": 343.2, "3": 343.2, "4": 254.8, "5": 171.6, "6": 254.8, "7": 171.6, "8": 343.2, "9": 343.2, "10": 343.2, "11": 254.8, "13": 171.6, "14": 171.6, "15": 343.2, "16": 343.2, "17": 193.7, "18": 171.6, "19": 171.6, "20": 83.2, "22": 171.6, "23": 171.6, "24": 171.6, "25": 171.6, "26": 171.6, "27": 83.2, "29": 171.6, "30": 171.6}}, "ALAA006523": {"5": {"4": 550.0, "5": 640.0, "6": 640.0, "7": 640.0, "8": 560.0, "11": 440.0, "12": 640.0, "13": 640.0, "14": 640.0, "15": 640.0, "16": 320.0, "18": 640.0, "19": 640.0, "20": 640.0, "21": 640.0, "22": 640.0, "23": 320.0, "25": 640.0, "26": 640.0, "27": 640.0, "28": 640.0, "29": 640.0, "30": 320.0}, "6": {"1": 1175.0, "2": 1125.0, "3": 680.0, "4": 745.0, "5": 540.0, "6": 735.0, "7": 500.0, "8": 940.0, "9": 1190.0, "10": 1020.0, "11": 750.0, "13": 95.0, "14": 25.0, "15": 560.0, "16": 825.0, "17": 535.0, "18": 455.0, "19": 545.0, "20": 315.0, "22": 490.0, "23": 530.0, "24": 610.0, "25": 360.0, "26": 360.0, "27": 155.0, "29": 330.0, "30": 570.0}}, "ALAA006524": {"5": {"4": 4.34, "5": 4.48, "6": 4.48, "7": 4.48, "8": 3.92, "11": 3.08, "12": 4.48, "13": 4.48, "14": 4.48, "15": 4.48, "16": 2.24, "18": 4.48, "19": 4.48, "20": 4.48, "21": 4.48, "22": 4.48, "23": 2.24, "25": 4.48, "26": 4.48, "27": 4.48, "28": 4.48, "29": 4.48, "30": 2.24}, "6": {"1": 9.24, "2": 9.24, "3": 9.24, "4": 6.86, "5": 4.62, "6": 6.86, "7": 4.62, "8": 9.24, "9": 9.24, "10": 9.24, "11": 6.86, "13": 4.62, "14": 4.62, "15": 9.24, "16": 9.24, "17": 5.21, "18": 4.62, "19": 4.62, "20": 2.24, "22": 4.62, "23": 4.62, "24": 4.62, "25": 4.62, "26": 4.62, "27": 2.24, "29": 4.62, "30": 4.62}}, "ALAA006526": {"5": {"4": 173.6, "5": 179.2, "6": 179.2, "7": 179.2, "8": 156.8, "11": 123.2, "12": 179.2, "13": 179.2, "14": 179.2, "15": 179.2, "16": 89.6, "18": 179.2, "19": 179.2, "20": 179.2, "21": 179.2, "22": 179.2, "23": 89.6, "25": 179.2, "26": 179.2, "27": 179.2, "28": 179.2, "29": 179.2, "30": 89.6}, "6": {"1": 369.6, "2": 369.6, "3": 369.6, "4": 274.4, "5": 184.8, "6": 274.4, "7": 184.8, "8": 369.6, "9": 369.6, "10": 369.6, "11": 274.4, "13": 184.8, "14": 184.8, "15": 369.6, "16": 369.6, "17": 208.6, "18": 184.8, "19": 184.8, "20": 89.6, "22": 184.8, "23": 184.8, "24": 184.8, "25": 184.8, "26": 184.8, "27": 89.6, "29": 184.8, "30": 184.8}}, "ALAA007195": {"5": {"4": 700.6, "5": 723.2, "6": 723.2, "7": 723.2, "8": 632.8, "11": 497.2, "12": 723.2, "13": 723.2, "14": 723.2, "15": 723.2, "16": 361.6, "18": 723.2, "19": 723.2, "20": 723.2, "21": 723.2, "22": 723.2, "23": 361.6, "25": 723.2, "26": 723.2, "27": 723.2, "28": 723.2, "29": 723.2, "30": 361.6}, "6": {"1": 1491.6, "2": 1491.6, "3": 1491.6, "4": 1107.4, "5": 745.8, "6": 1107.4, "7": 745.8, "8": 1491.6, "9": 1491.6, "10": 1491.6, "11": 1107.4, "13": 745.8, "14": 745.8, "15": 1491.6, "16": 1491.6, "17": 841.85, "18": 745.8, "19": 745.8, "20": 361.6, "22": 745.8, "23": 745.8, "24": 745.8, "25": 745.8, "26": 745.8, "27": 361.6, "29": 745.8, "30": 745.8}}, "ALAA007196": {"5": {"4": 68.2, "5": 70.4, "6": 70.4, "7": 70.4, "8": 61.6, "11": 48.4, "12": 70.4, "13": 70.4, "14": 70.4, "15": 70.4, "16": 35.2, "18": 70.4, "19": 70.4, "20": 70.4, "21": 70.4, "22": 70.4, "23": 35.2, "25": 70.4, "26": 70.4, "27": 70.4, "28": 70.4, "29": 70.4, "30": 35.2}, "6": {"1": 145.2, "2": 145.2, "3": 145.2, "4": 107.8, "5": 72.6, "6": 107.8, "7": 72.6, "8": 145.2, "9": 145.2, "10": 145.2, "11": 107.8, "13": 72.6, "14": 72.6, "15": 145.2, "16": 145.2, "17": 81.95, "18": 72.6, "19": 72.6, "20": 35.2, "22": 72.6, "23": 72.6, "24": 72.6, "25": 72.6, "26": 72.6, "27": 35.2, "29": 72.6, "30": 72.6}}, "ALAA007197": {"5": {"4": 18.6, "5": 19.2, "6": 19.2, "7": 19.2, "8": 16.8, "11": 13.2, "12": 19.2, "13": 19.2, "14": 19.2, "15": 19.2, "16": 9.6, "18": 19.2, "19": 19.2, "20": 19.2, "21": 19.2, "22": 19.2, "23": 9.6, "25": 19.2, "26": 19.2, "27": 19.2, "28": 19.2, "29": 19.2, "30": 9.6}, "6": {"1": 39.6, "2": 39.6, "3": 39.6, "4": 29.4, "5": 19.8, "6": 29.4, "7": 19.8, "8": 39.6, "9": 39.6, "10": 39.6, "11": 29.4, "13": 19.8, "14": 19.8, "15": 39.6, "16": 39.6, "17": 22.35, "18": 19.8, "19": 19.8, "20": 9.6, "22": 19.8, "23": 19.8, "24": 19.8, "25": 19.8, "26": 19.8, "27": 9.6, "29": 19.8, "30": 19.8}}, "ALAA007198": {"5": {"4": 550.0, "5": 640.0, "6": 640.0, "7": 640.0, "8": 560.0, "11": 440.0, "12": 640.0, "13": 640.0, "14": 640.0, "15": 640.0, "16": 320.0, "18": 640.0, "19": 640.0, "20": 640.0, "21": 640.0, "22": 640.0, "23": 320.0, "25": 640.0, "26": 640.0, "27": 640.0, "28": 640.0, "29": 640.0, "30": 320.0}, "6": {"1": 1175.0, "2": 1125.0, "3": 680.0, "4": 745.0, "5": 540.0, "6": 735.0, "7": 500.0, "8": 940.0, "9": 1190.0, "10": 1020.0, "11": 750.0, "13": 95.0, "14": 25.0, "15": 560.0, "16": 825.0, "17": 535.0, "18": 455.0, "19": 545.0, "20": 315.0, "22": 490.0, "23": 530.0, "24": 610.0, "25": 360.0, "26": 360.0, "27": 155.0, "29": 330.0, "30": 570.0}}, "ALAA007199": {"5": {"4": 550.0, "5": 640.0, "6": 640.0, "7": 640.0, "8": 560.0, "11": 440.0, "12": 640.0, "13": 640.0, "14": 640.0, "15": 640.0, "16": 320.0, "18": 640.0, "19": 640.0, "20": 640.0, "21": 640.0, "22": 640.0, "23": 320.0, "25": 640.0, "26": 640.0, "27": 640.0, "28": 640.0, "29": 640.0, "30": 320.0}, "6": {"1": 1175.0, "2": 1125.0, "3": 680.0, "4": 745.0, "5": 540.0, "6": 735.0, "7": 500.0, "8": 940.0, "9": 1190.0, "10": 1020.0, "11": 750.0, "13": 95.0, "14": 25.0, "15": 560.0, "16": 825.0, "17": 535.0, "18": 455.0, "19": 545.0, "20": 315.0, "22": 490.0, "23": 530.0, "24": 610.0, "25": 360.0, "26": 360.0, "27": 155.0, "29": 330.0, "30": 570.0}}, "ALAA007200": {"5": {"4": 31.0, "5": 32.0, "6": 32.0, "7": 32.0, "8": 28.0, "11": 22.0, "12": 32.0, "13": 32.0, "14": 32.0, "15": 32.0, "16": 16.0, "18": 32.0, "19": 32.0, "20": 32.0, "21": 32.0, "22": 32.0, "23": 16.0, "25": 32.0, "26": 32.0, "27": 32.0, "28": 32.0, "29": 32.0, "30": 16.0}, "6": {"1": 66.0, "2": 66.0, "3": 66.0, "4": 49.0, "5": 33.0, "6": 49.0, "7": 33.0, "8": 66.0, "9": 66.0, "10": 66.0, "11": 49.0, "13": 33.0, "14": 33.0, "15": 66.0, "16": 66.0, "17": 37.25, "18": 33.0, "19": 33.0, "20": 16.0, "22": 33.0, "23": 33.0, "24": 33.0, "25": 33.0, "26": 33.0, "27": 16.0, "29": 33.0, "30": 33.0}}, "ALAA007201": {"5": {"4": 124.0, "5": 128.0, "6": 128.0, "7": 128.0, "8": 112.0, "11": 88.0, "12": 128.0, "13": 128.0, "14": 128.0, "15": 128.0, "16": 64.0, "18": 128.0, "19": 128.0, "20": 128.0, "21": 128.0, "22": 128.0, "23": 64.0, "25": 128.0, "26": 128.0, "27": 128.0, "28": 128.0, "29": 128.0, "30": 64.0}, "6": {"1": 264.0, "2": 264.0, "3": 264.0, "4": 196.0, "5": 132.0, "6": 196.0, "7": 132.0, "8": 264.0, "9": 264.0, "10": 264.0, "11": 196.0, "13": 132.0, "14": 132.0, "15": 264.0, "16": 264.0, "17": 149.0, "18": 132.0, "19": 132.0, "20": 64.0, "22": 132.0, "23": 132.0, "24": 132.0, "25": 132.0, "26": 132.0, "27": 64.0, "29": 132.0, "30": 132.0}}, "ALAA007202": {"5": {"4": 657.2, "5": 678.4, "6": 678.4, "7": 678.4, "8": 593.6, "11": 466.4, "12": 678.4, "13": 678.4, "14": 678.4, "15": 678.4, "16": 339.2, "18": 678.4, "19": 678.4, "20": 678.4, "21": 678.4, "22": 678.4, "23": 339.2, "25": 678.4, "26": 678.4, "27": 678.4, "28": 678.4, "29": 678.4, "30": 339.2}, "6": {"1": 1399.2, "2": 1399.2, "3": 1399.2, "4": 1038.8, "5": 699.6, "6": 1038.8, "7": 699.6, "8": 1399.2, "9": 1399.2, "10": 1399.2, "11": 1038.8, "13": 699.6, "14": 699.6, "15": 1399.2, "16": 1399.2, "17": 789.7, "18": 699.6, "19": 699.6, "20": 339.2, "22": 699.6, "23": 699.6, "24": 699.6, "25": 699.6, "26": 699.6, "27": 339.2, "29": 699.6, "30": 699.6}}, "ALAA007267": {"5": {"4": 1116.0, "5": 1152.0, "6": 1152.0, "7": 1152.0, "8": 1008.0, "11": 792.0, "12": 1152.0, "13": 1152.0, "14": 1152.0, "15": 1152.0, "16": 576.0, "18": 1152.0, "19": 1152.0, "20": 1152.0, "21": 1152.0, "22": 1152.0, "23": 576.0, "25": 1152.0, "26": 1152.0, "27": 1152.0, "28": 1152.0, "29": 1152.0, "30": 576.0}, "6": {"1": 2376.0, "2": 2376.0, "3": 2376.0, "4": 1764.0, "5": 1188.0, "6": 1764.0, "7": 1188.0, "8": 2376.0, "9": 2376.0, "10": 2376.0, "11": 1764.0, "13": 1188.0, "14": 1188.0, "15": 2376.0, "16": 2376.0, "17": 1341.0, "18": 1188.0, "19": 1188.0, "20": 576.0, "22": 1188.0, "23": 1188.0, "24": 1188.0, "25": 1188.0, "26": 1188.0, "27": 576.0, "29": 1188.0, "30": 1188.0}}, "ALAA007434": {"5": {"4": 550.0, "5": 640.0, "6": 640.0, "7": 640.0, "8": 560.0, "11": 440.0, "12": 640.0, "13": 640.0, "14": 640.0, "15": 640.0, "16": 320.0, "18": 640.0, "19": 640.0, "20": 640.0, "21": 640.0, "22": 640.0, "23": 320.0, "25": 640.0, "26": 640.0, "27": 640.0, "28": 640.0, "29": 640.0, "30": 320.0}, "6": {"1": 1175.0, "2": 1125.0, "3": 680.0, "4": 745.0, "5": 540.0, "6": 735.0, "7": 500.0, "8": 940.0, "9": 1190.0, "10": 1020.0, "11": 750.0, "13": 95.0, "14": 25.0, "15": 560.0, "16": 825.0, "17": 535.0, "18": 455.0, "19": 545.0, "20": 315.0, "22": 490.0, "23": 530.0, "24": 610.0, "25": 360.0, "26": 360.0, "27": 155.0, "29": 330.0, "30": 570.0}}, "ALAA007435": {"5": {"4": 550.0, "5": 640.0, "6": 640.0, "7": 640.0, "8": 560.0, "11": 440.0, "12": 640.0, "13": 640.0, "14": 640.0, "15": 640.0, "16": 320.0, "18": 640.0, "19": 640.0, "20": 640.0, "21": 640.0, "22": 640.0, "23": 320.0, "25": 640.0, "26": 640.0, "27": 640.0, "28": 640.0, "29": 640.0, "30": 320.0}, "6": {"1": 1175.0, "2": 1125.0, "3": 680.0, "4": 745.0, "5": 540.0, "6": 735.0, "7": 500.0, "8": 940.0, "9": 1190.0, "10": 1020.0, "11": 750.0, "13": 95.0, "14": 25.0, "15": 560.0, "16": 825.0, "17": 535.0, "18": 455.0, "19": 545.0, "20": 315.0, "22": 490.0, "23": 530.0, "24": 610.0, "25": 360.0, "26": 360.0, "27": 155.0, "29": 330.0, "30": 570.0}}, "ALAA007440": {"5": {"4": 550.0, "5": 640.0, "6": 640.0, "7": 640.0, "8": 560.0, "11": 440.0, "12": 640.0, "13": 640.0, "14": 640.0, "15": 640.0, "16": 320.0, "18": 640.0, "19": 640.0, "20": 640.0, "21": 640.0, "22": 640.0, "23": 320.0, "25": 640.0, "26": 640.0, "27": 640.0, "28": 640.0, "29": 640.0, "30": 320.0}, "6": {"1": 1175.0, "2": 1125.0, "3": 680.0, "4": 745.0, "5": 540.0, "6": 735.0, "7": 500.0, "8": 940.0, "9": 1190.0, "10": 1020.0, "11": 750.0, "13": 95.0, "14": 25.0, "15": 560.0, "16": 825.0, "17": 535.0, "18": 455.0, "19": 545.0, "20": 315.0, "22": 490.0, "23": 530.0, "24": 610.0, "25": 360.0, "26": 360.0, "27": 155.0, "29": 330.0, "30": 570.0}}, "ALAA007466": {"5": {"4": 498.55, "5": 599.95, "6": 84.5, "7": 185.9, "8": 101.4, "11": 439.4, "12": 237.3, "13": 621.78, "14": 737.97, "15": 571.08, "16": 366.87, "18": 456.3, "19": 585.87, "20": 521.08, "21": 540.8, "22": 404.19, "23": 466.86, "25": 806.98, "26": 730.22, "27": 683.75, "28": 709.1, "29": 162.66, "30": 462.64}, "6": {"1": 1249.9, "2": 958.37, "3": 435.17, "4": 553.48, "5": 635.16, "6": 373.21, "7": 780.22, "8": 577.42, "9": 1042.87, "10": 428.84, "11": 372.5, "13": 195.05, "14": 24.65, "15": 1149.9, "16": 477.42, "17": 711.21, "18": 120.41, "19": 54.22, "20": 169.0, "22": 1267.5, "23": 797.12, "24": 140.83, "25": 101.4, "26": 439.4, "27": 64.08, "29": 965.41, "30": 277.44}}, "ALAA007467": {"5": {"4": 101.84, "5": 128.64, "6": 48.24, "7": 112.56, "8": 64.32, "11": 64.32, "12": 109.88, "13": 83.08, "14": 64.32, "15": 125.96, "16": 2.68, "18": 69.68, "19": 107.2, "20": 88.44, "21": 88.44, "22": 10.72, "23": 53.6, "25": 64.32, "27": 64.32, "28": 48.24, "29": 72.36, "30": 50.92}, "6": {"1": 222.44, "2": 184.92, "3": 219.76, "4": 80.4, "5": 69.68, "6": 174.2, "7": 48.24, "8": 154.1, "9": 172.19, "10": 126.63, "11": 154.77, "13": 64.32, "15": 166.83, "16": 98.49, "17": 14.07, "18": 64.32, "20": 48.24, "23": 64.32, "24": 64.32, "25": 64.32, "26": 64.32, "27": 18.76, "29": 109.88, "30": 64.32}}, "ALAA007732": {"5": {"4": 141.4, "5": 125.24, "6": 266.64, "7": 96.96, "8": 96.96, "11": 96.96, "12": 181.8, "13": 181.8, "14": 148.47, "15": 106.05, "16": 48.48, "18": 230.28, "19": 167.66, "20": 195.94, "21": 181.8, "22": 84.84, "23": 88.88, "25": 92.92, "26": 121.2, "27": 157.56, "28": 96.96, "29": 169.68, "30": 84.84}, "6": {"1": 254.52, "2": 84.84, "3": 351.48, "4": 141.4, "5": 96.96, "6": 365.62, "7": 98.98, "8": 181.8, "9": 254.52, "10": 224.22, "11": 146.45, "13": 271.69, "14": 331.28, "15": 384.81, "16": 301.99, "17": 96.96, "18": 84.84, "19": 162.61, "20": 7.07, "22": 181.8, "23": 96.96, "25": 139.38, "26": 127.26, "27": 169.68, "29": 96.96}}, "ALAA007765": {"5": {"4": 186.0, "5": 192.0, "6": 192.0, "7": 192.0, "8": 168.0, "11": 132.0, "12": 192.0, "13": 192.0, "14": 192.0, "15": 192.0, "16": 96.0, "18": 192.0, "19": 192.0, "20": 192.0, "21": 192.0, "22": 192.0, "23": 96.0, "25": 192.0, "26": 192.0, "27": 192.0, "28": 192.0, "29": 192.0, "30": 96.0}, "6": {"1": 396.0, "2": 396.0, "3": 396.0, "4": 294.0, "5": 198.0, "6": 294.0, "7": 198.0, "8": 396.0, "9": 396.0, "10": 396.0, "11": 294.0, "13": 198.0, "14": 198.0, "15": 396.0, "16": 396.0, "17": 223.5, "18": 198.0, "19": 198.0, "20": 96.0, "22": 198.0, "23": 198.0, "24": 198.0, "25": 198.0, "26": 198.0, "27": 96.0, "29": 198.0, "30": 198.0}}, "ALAA007787": {"5": {"4": 30.38, "5": 31.36, "6": 31.36, "7": 31.36, "8": 27.44, "11": 21.56, "12": 31.36, "13": 31.36, "14": 31.36, "15": 31.36, "16": 15.68, "18": 31.36, "19": 31.36, "20": 31.36, "21": 31.36, "22": 31.36, "23": 15.68, "25": 31.36, "26": 31.36, "27": 31.36, "28": 31.36, "29": 31.36, "30": 15.68}, "6": {"1": 64.68, "2": 64.68, "3": 64.68, "4": 48.02, "5": 32.34, "6": 48.02, "7": 32.34, "8": 64.68, "9": 64.68, "10": 64.68, "11": 48.02, "13": 32.34, "14": 32.34, "15": 64.68, "16": 64.68, "17": 36.51, "18": 32.34, "19": 32.34, "20": 15.68, "22": 32.34, "23": 32.34, "24": 32.34, "25": 32.34, "26": 32.34, "27": 15.68, "29": 32.34, "30": 32.34}}, "ALAA008937": {"5": {"4": 967.2, "5": 998.4, "6": 998.4, "7": 998.4, "8": 873.6, "11": 686.4, "12": 998.4, "13": 998.4, "14": 998.4, "15": 998.4, "16": 499.2, "18": 998.4, "19": 998.4, "20": 998.4, "21": 998.4, "22": 998.4, "23": 499.2, "25": 998.4, "26": 998.4, "27": 998.4, "28": 998.4, "29": 998.4, "30": 499.2}, "6": {"1": 2059.2, "2": 2059.2, "3": 2059.2, "4": 1528.8, "5": 1029.6, "6": 1528.8, "7": 1029.6, "8": 2059.2, "9": 2059.2, "10": 2059.2, "11": 1528.8, "13": 1029.6, "14": 1029.6, "15": 2059.2, "16": 2059.2, "17": 1162.2, "18": 1029.6, "19": 1029.6, "20": 499.2, "22": 1029.6, "23": 1029.6, "24": 1029.6, "25": 1029.6, "26": 1029.6, "27": 499.2, "29": 1029.6, "30": 1029.6}}, "ALAB000025": {"5": {"4": 20.16, "5": 24.98, "6": 21.34, "7": 26.16, "8": 23.7, "11": 16.62, "12": 23.91, "13": 23.66, "14": 25.99, "15": 25.41, "16": 11.83, "18": 24.74, "19": 24.66, "20": 25.57, "21": 20.75, "22": 27.24, "23": 13.27, "25": 27.83, "26": 27.53, "27": 25.67, "28": 24.06, "29": 26.98, "30": 12.59}, "6": {"1": 59.39, "2": 48.59, "3": 56.14, "4": 39.51, "5": 29.01, "6": 35.62, "7": 24.29, "8": 55.76, "9": 52.62, "10": 53.25, "11": 40.66, "13": 23.7, "14": 23.99, "15": 51.32, "16": 46.22, "17": 42.38, "18": 29.01, "19": 27.43, "20": 14.16, "22": 26.65, "23": 29.01, "24": 29.64, "25": 31.92, "26": 25.56, "27": 12.59, "29": 22.23, "30": 31.18}}, "ALAB000028": {"5": {"4": 69.6, "5": 89.8, "6": 73.4, "7": 93.6, "8": 81.0, "11": 58.2, "12": 84.6, "13": 83.6, "14": 87.4, "15": 90.6, "16": 41.8, "18": 82.4, "19": 87.6, "20": 91.7, "21": 71.5, "22": 92.4, "23": 42.75, "25": 94.3, "26": 93.35, "27": 82.65, "28": 83.15, "29": 95.25, "30": 40.95}, "6": {"1": 204.9, "2": 170.9, "3": 194.45, "4": 140.1, "5": 98.1, "6": 125.25, "7": 82.9, "8": 193.6, "9": 181.75, "10": 181.05, "11": 140.1, "13": 83.55, "14": 79.4, "15": 171.7, "16": 158.2, "17": 145.85, "18": 98.1, "19": 88.35, "20": 45.6, "22": 90.5, "23": 98.1, "24": 97.4, "25": 110.2, "26": 86.6, "27": 40.95, "29": 76.25, "30": 109.75}}, "ALAB000046": {"5": {"4": 38141.24, "5": 43486.06, "6": 40661.0, "7": 46005.82, "8": 45695.51, "11": 30568.67, "12": 43056.6, "13": 42728.91, "14": 51621.37, "15": 45022.75, "16": 21364.46, "18": 49982.91, "19": 44039.68, "20": 44745.94, "21": 39401.12, "22": 53259.82, "23": 28367.09, "25": 54519.71, "26": 53889.77, "27": 54844.35, "28": 45449.16, "29": 48776.96, "30": 26502.71}, "6": {"1": 112470.41, "2": 88583.52, "3": 105541.05, "4": 70823.67, "5": 57040.91, "6": 65000.12, "7": 46960.41, "8": 104306.61, "9": 99467.57, "10": 103699.07, "11": 77196.36, "13": 43031.16, "14": 48999.84, "15": 102865.6, "16": 88882.73, "17": 80675.91, "18": 57040.91, "19": 58625.44, "20": 30256.91, "22": 52001.38, "23": 57040.91, "24": 61272.41, "25": 60368.7, "26": 50086.11, "27": 26502.71, "29": 42552.26, "30": 56716.26}}, "ALAB000049": {"5": {"4": 708.74, "5": 757.88, "6": 716.24, "7": 782.42, "8": 659.7, "11": 497.17, "12": 744.14, "13": 741.78, "14": 741.94, "15": 758.3, "16": 370.89, "18": 732.27, "19": 757.61, "20": 761.63, "21": 721.01, "22": 753.74, "23": 365.43, "25": 757.49, "26": 755.62, "27": 725.23, "28": 739.19, "29": 781.97, "30": 362.64}, "6": {"1": 1576.09, "2": 1519.05, "3": 1563.99, "4": 1165.6, "5": 785.43, "6": 1119.74, "7": 743.51, "8": 1554.56, "9": 1536.22, "10": 1535.11, "11": 1161.1, "13": 749.82, "14": 723.06, "15": 1504.79, "16": 1475.42, "17": 1209.18, "18": 785.43, "19": 756.92, "20": 371.05, "22": 778.95, "23": 785.43, "24": 778.64, "25": 814.72, "26": 761.96, "27": 362.64, "29": 742.31, "30": 817.7}}, "ALAB000052": {"5": {"4": 60.5, "5": 60.5, "6": 60.5, "7": 60.5, "8": 52.8, "11": 40.15, "12": 60.5, "13": 60.5, "14": 60.5, "15": 60.5, "16": 30.25, "18": 60.5, "19": 60.5, "20": 60.5, "21": 60.5, "22": 60.5, "23": 30.25, "25": 60.5, "26": 60.5, "27": 60.5, "28": 60.5, "29": 60.5, "30": 30.25}, "6": {"1": 123.2, "2": 123.2, "3": 123.2, "4": 92.95, "5": 62.7, "6": 90.75, "7": 60.5, "8": 123.2, "9": 123.2, "10": 123.2, "11": 92.95, "13": 60.5, "14": 60.5, "15": 123.2, "16": 123.2, "17": 96.25, "18": 62.7, "19": 62.7, "20": 30.25, "22": 62.7, "23": 62.7, "24": 62.7, "25": 62.7, "26": 62.7, "27": 30.25, "29": 62.7, "30": 62.7}}, "ALAB000055": {"5": {"4": 520.0, "5": 640.0, "6": 640.0, "7": 640.0, "8": 560.0, "11": 440.0, "12": 640.0, "13": 640.0, "14": 640.0, "15": 640.0, "16": 320.0, "18": 640.0, "19": 640.0, "20": 640.0, "21": 640.0, "22": 640.0, "23": 320.0, "25": 640.0, "26": 640.0, "27": 640.0, "28": 640.0, "29": 640.0, "30": 320.0}, "6": {"1": 1225.0, "2": 1185.0, "3": 925.0, "4": 540.0, "5": 530.0, "6": 705.0, "7": 520.0, "8": 940.0, "9": 1060.0, "10": 1170.0, "11": 735.0, "13": 270.0, "14": 95.0, "15": 570.0, "16": 560.0, "17": 815.0, "18": 390.0, "19": 550.0, "20": 285.0, "22": 560.0, "23": 530.0, "24": 550.0, "25": 470.0, "26": 340.0, "27": 205.0, "29": 330.0, "30": 450.0}}, "ALAB000091": {"5": {"4": 520.0, "5": 640.0, "6": 640.0, "7": 640.0, "8": 560.0, "11": 440.0, "12": 640.0, "13": 640.0, "14": 640.0, "15": 640.0, "16": 320.0, "18": 640.0, "19": 640.0, "20": 640.0, "21": 640.0, "22": 640.0, "23": 320.0, "25": 640.0, "26": 640.0, "27": 640.0, "28": 640.0, "29": 640.0, "30": 320.0}, "6": {"1": 1225.0, "2": 1185.0, "3": 925.0, "4": 540.0, "5": 530.0, "6": 705.0, "7": 520.0, "8": 940.0, "9": 1060.0, "10": 1170.0, "11": 735.0, "13": 270.0, "14": 95.0, "15": 570.0, "16": 560.0, "17": 815.0, "18": 390.0, "19": 550.0, "20": 285.0, "22": 560.0, "23": 530.0, "24": 550.0, "25": 470.0, "26": 340.0, "27": 205.0, "29": 330.0, "30": 450.0}}, "ALAB000208": {"5": {"4": 207.36, "5": 267.56, "6": 218.68, "7": 278.88, "8": 241.32, "11": 173.4, "12": 252.06, "13": 249.08, "14": 260.38, "15": 269.94, "16": 124.54, "18": 245.48, "19": 261.0, "20": 273.22, "21": 213.02, "22": 275.28, "23": 127.35, "25": 280.94, "26": 278.11, "27": 246.21, "28": 247.73, "29": 283.79, "30": 121.99}, "6": {"1": 610.46, "2": 509.18, "3": 579.33, "4": 417.42, "5": 292.26, "6": 373.17, "7": 246.98, "8": 576.8, "9": 541.49, "10": 539.39, "11": 417.4, "13": 248.93, "14": 236.54, "15": 511.52, "16": 471.32, "17": 434.53, "18": 292.26, "19": 263.19, "20": 135.84, "22": 269.62, "23": 292.26, "24": 290.16, "25": 328.32, "26": 258.0, "27": 121.99, "29": 227.17, "30": 326.99}}, "ALAB000321": {"5": {"4": 17.28, "5": 22.34, "6": 18.22, "7": 18.48, "8": 20.1, "11": 14.46, "12": 21.03, "13": 20.78, "14": 21.67, "15": 22.53, "16": 10.39, "18": 19.82, "19": 19.98, "20": 22.81, "21": 15.35, "22": 22.92, "23": 10.57, "25": 23.39, "26": 23.16, "27": 20.44, "28": 20.14, "29": 19.37, "30": 10.13}, "6": {"1": 50.87, "2": 40.07, "3": 45.88, "4": 34.83, "5": 24.33, "6": 28.72, "7": 18.17, "8": 48.08, "9": 42.73, "10": 40.91, "11": 33.98, "13": 18.36, "14": 19.67, "15": 40.16, "16": 39.26, "17": 33.8, "18": 24.33, "19": 21.86, "20": 11.28, "22": 20.05, "23": 24.33, "24": 24.12, "25": 27.36, "26": 21.48, "27": 10.13, "29": 18.93, "30": 27.28}}, "ALAB000663": {"5": {"4": 293.19, "5": 295.93, "6": 290.45, "7": 294.68, "8": 243.44, "11": 191.72, "12": 293.19, "13": 293.19, "14": 276.72, "15": 293.18, "16": 146.59, "18": 276.91, "19": 293.75, "20": 294.56, "21": 292.57, "22": 276.72, "23": 132.18, "25": 275.34, "26": 276.03, "27": 266.43, "28": 287.17, "29": 289.72, "30": 134.93}, "6": {"1": 566.66, "2": 594.87, "3": 574.96, "4": 447.52, "5": 284.46, "6": 438.47, "7": 284.33, "8": 576.27, "9": 579.08, "10": 569.98, "11": 434.05, "13": 293.25, "14": 276.72, "15": 564.68, "16": 584.51, "17": 448.9, "18": 284.46, "19": 274.17, "20": 130.13, "22": 290.7, "23": 284.46, "24": 274.85, "25": 285.83, "26": 291.32, "27": 134.93, "29": 300.25, "30": 293.37}}, "ALAB000664": {"5": {"4": 29.11, "5": 24.92, "6": 29.11, "7": 24.92, "8": 22.15, "11": 20.77, "12": 24.92, "13": 30.85, "14": 27.36, "15": 29.11, "16": 13.16, "18": 28.41, "19": 25.62, "20": 31.2, "21": 27.01, "22": 28.41, "23": 14.2, "25": 31.2, "26": 27.01, "27": 27.36, "28": 26.66, "29": 27.36, "30": 12.46}, "6": {"1": 61.31, "2": 68.98, "3": 61.66, "4": 43.66, "5": 30.5, "6": 45.01, "7": 36.74, "8": 55.38, "9": 55.38, "10": 67.24, "11": 38.08, "13": 32.2, "14": 37.09, "15": 55.38, "16": 63.05, "17": 46.07, "18": 24.92, "19": 26.66, "20": 14.9, "22": 24.92, "23": 24.92, "24": 27.71, "25": 26.32, "26": 24.92, "27": 12.46, "29": 28.76, "30": 25.27}}, "ALAB000709": {"5": {"4": 7.68, "5": 7.04, "6": 8.32, "7": 7.68, "8": 9.6, "11": 5.76, "12": 7.68, "13": 7.68, "14": 11.52, "15": 7.68, "16": 3.84, "18": 11.52, "19": 7.68, "20": 7.36, "21": 8.0, "22": 11.52, "23": 7.2, "25": 11.84, "26": 11.68, "27": 13.92, "28": 9.12, "29": 8.8, "30": 6.56}, "6": {"1": 22.72, "2": 16.32, "3": 20.96, "4": 12.48, "5": 12.48, "6": 12.0, "7": 9.92, "8": 20.48, "9": 20.0, "10": 22.24, "11": 15.68, "13": 7.84, "14": 11.52, "15": 23.36, "16": 18.56, "17": 16.48, "18": 12.48, "19": 14.88, "20": 7.68, "22": 11.2, "23": 12.48, "24": 14.72, "25": 12.16, "26": 10.88, "27": 6.56, "29": 8.8, "30": 10.4}}, "ALAB000986": {"5": {"4": 23.84, "5": 18.16, "6": 22.35, "7": 16.93, "8": 13.13, "11": 11.86, "12": 19.35, "13": 19.65, "14": 16.37, "15": 17.56, "16": 9.83, "18": 17.9, "19": 18.56, "20": 17.41, "21": 23.23, "22": 14.88, "23": 7.67, "25": 14.13, "26": 14.5, "27": 16.45, "28": 19.01, "29": 15.78, "30": 8.56}, "6": {"1": 25.7, "2": 39.58, "3": 29.94, "4": 26.33, "5": 14.41, "6": 29.39, "7": 18.75, "8": 30.33, "9": 34.28, "10": 33.32, "11": 24.58, "13": 19.71, "14": 18.76, "15": 35.41, "16": 41.99, "17": 25.18, "18": 14.41, "19": 15.99, "20": 6.54, "22": 17.54, "23": 14.41, "24": 13.37, "25": 10.97, "26": 18.75, "27": 8.56, "29": 23.01, "30": 12.09}}, "ALAB001037": {"5": {"4": 28611.8, "5": 25463.2, "6": 28026.0, "7": 24685.4, "8": 21877.8, "11": 17490.7, "12": 26277.8, "13": 26433.4, "14": 25875.0, "15": 25344.2, "16": 13216.7, "18": 26629.0, "19": 25739.0, "20": 25170.3, "21": 28222.9, "22": 25097.0, "23": 13097.65, "25": 24804.1, "26": 24950.55, "27": 26634.65, "28": 26495.75, "29": 24458.25, "30": 13372.25}, "6": {"1": 48896.5, "2": 54036.1, "3": 50411.45, "4": 38724.4, "5": 25571.7, "6": 40060.95, "7": 26944.9, "8": 50635.6, "9": 52379.35, "10": 52443.45, "11": 38719.8, "13": 26825.95, "14": 27599.2, "15": 53971.9, "16": 56127.4, "17": 40157.6, "18": 25571.7, "19": 27109.35, "20": 12658.3, "22": 26647.3, "23": 25571.7, "24": 25699.8, "25": 23686.2, "26": 27347.4, "27": 13372.25, "29": 28940.05, "30": 23741.15}}, "ALAB001069": {"5": {"4": 1650.0, "5": 4950.0, "6": 1650.0, "7": 4950.0, "8": 4400.0, "11": 4125.0, "12": 4950.0, "13": 275.0, "14": 3025.0, "15": 1650.0, "16": 1925.0, "18": 2200.0, "19": 4400.0, "21": 3300.0, "22": 2200.0, "23": 1100.0, "26": 3300.0, "27": 3025.0, "28": 3575.0, "29": 3025.0, "30": 2475.0}, "6": {"1": 6325.0, "2": 275.0, "3": 6050.0, "4": 2475.0, "5": 550.0, "6": 6875.0, "7": 1100.0, "8": 11000.0, "9": 11000.0, "10": 1650.0, "11": 6875.0, "13": 4675.0, "14": 825.0, "15": 11000.0, "16": 4950.0, "17": 4675.0, "18": 4950.0, "19": 3575.0, "20": 550.0, "22": 4950.0, "23": 4950.0, "24": 2750.0, "25": 3850.0, "26": 4950.0, "27": 2475.0, "29": 1925.0, "30": 4675.0}}, "ALAB001078": {"5": {"4": 520.0, "5": 640.0, "6": 640.0, "7": 640.0, "8": 560.0, "11": 440.0, "12": 640.0, "13": 640.0, "14": 640.0, "15": 640.0, "16": 320.0, "18": 640.0, "19": 640.0, "20": 640.0, "21": 640.0, "22": 640.0, "23": 320.0, "25": 640.0, "26": 640.0, "27": 640.0, "28": 640.0, "29": 640.0, "30": 320.0}, "6": {"1": 1225.0, "2": 1185.0, "3": 925.0, "4": 540.0, "5": 530.0, "6": 705.0, "7": 520.0, "8": 940.0, "9": 1060.0, "10": 1170.0, "11": 735.0, "13": 270.0, "14": 95.0, "15": 570.0, "16": 560.0, "17": 815.0, "18": 390.0, "19": 550.0, "20": 285.0, "22": 560.0, "23": 530.0, "24": 550.0, "25": 470.0, "26": 340.0, "27": 205.0, "29": 330.0, "30": 450.0}}, "ALAB001079": {"5": {"4": 520.0, "5": 640.0, "6": 640.0, "7": 640.0, "8": 560.0, "11": 440.0, "12": 640.0, "13": 640.0, "14": 640.0, "15": 640.0, "16": 320.0, "18": 640.0, "19": 640.0, "20": 640.0, "21": 640.0, "22": 640.0, "23": 320.0, "25": 640.0, "26": 640.0, "27": 640.0, "28": 640.0, "29": 640.0, "30": 320.0}, "6": {"1": 1225.0, "2": 1185.0, "3": 925.0, "4": 540.0, "5": 530.0, "6": 705.0, "7": 520.0, "8": 940.0, "9": 1060.0, "10": 1170.0, "11": 735.0, "13": 270.0, "14": 95.0, "15": 570.0, "16": 560.0, "17": 815.0, "18": 390.0, "19": 550.0, "20": 285.0, "22": 560.0, "23": 530.0, "24": 550.0, "25": 470.0, "26": 340.0, "27": 205.0, "29": 330.0, "30": 450.0}}, "ALAB001080": {"5": {"4": 520.0, "5": 640.0, "6": 640.0, "7": 640.0, "8": 560.0, "11": 440.0, "12": 640.0, "13": 640.0, "14": 640.0, "15": 640.0, "16": 320.0, "18": 640.0, "19": 640.0, "20": 640.0, "21": 640.0, "22": 640.0, "23": 320.0, "25": 640.0, "26": 640.0, "27": 640.0, "28": 640.0, "29": 640.0, "30": 320.0}, "6": {"1": 1225.0, "2": 1185.0, "3": 925.0, "4": 540.0, "5": 530.0, "6": 705.0, "7": 520.0, "8": 940.0, "9": 1060.0, "10": 1170.0, "11": 735.0, "13": 270.0, "14": 95.0, "15": 570.0, "16": 560.0, "17": 815.0, "18": 390.0, "19": 550.0, "20": 285.0, "22": 560.0, "23": 530.0, "24": 550.0, "25": 470.0, "26": 340.0, "27": 205.0, "29": 330.0, "30": 450.0}}, "ALAB001507": {"5": {"4": 520.0, "5": 640.0, "6": 640.0, "7": 640.0, "8": 560.0, "11": 440.0, "12": 640.0, "13": 640.0, "14": 640.0, "15": 640.0, "16": 320.0, "18": 640.0, "19": 640.0, "20": 640.0, "21": 640.0, "22": 640.0, "23": 320.0, "25": 640.0, "26": 640.0, "27": 640.0, "28": 640.0, "29": 640.0, "30": 320.0}, "6": {"1": 1225.0, "2": 1185.0, "3": 925.0, "4": 540.0, "5": 530.0, "6": 705.0, "7": 520.0, "8": 940.0, "9": 1060.0, "10": 1170.0, "11": 735.0, "13": 270.0, "14": 95.0, "15": 570.0, "16": 560.0, "17": 815.0, "18": 390.0, "19": 550.0, "20": 285.0, "22": 560.0, "23": 530.0, "24": 550.0, "25": 470.0, "26": 340.0, "27": 205.0, "29": 330.0, "30": 450.0}}, "ALAB001571": {"5": {"4": 5208000.0, "5": 5376000.0, "6": 5376000.0, "7": 5376000.0, "8": 4704000.0, "11": 3696000.0, "12": 5376000.0, "13": 5376000.0, "14": 5376000.0, "15": 5376000.0, "16": 2688000.0, "18": 5376000.0, "19": 5376000.0, "20": 5376000.0, "21": 5376000.0, "22": 5376000.0, "23": 2688000.0, "25": 5376000.0, "26": 5376000.0, "27": 5376000.0, "28": 5376000.0, "29": 5376000.0, "30": 2688000.0}, "6": {"1": 11088000.0, "2": 11088000.0, "3": 11088000.0, "4": 8232000.0, "5": 5544000.0, "6": 8232000.0, "7": 5544000.0, "8": 11088000.0, "9": 11088000.0, "10": 11088000.0, "11": 8232000.0, "13": 5544000.0, "14": 5544000.0, "15": 11088000.0, "16": 11088000.0, "17": 6258000.0, "18": 5544000.0, "19": 5544000.0, "20": 2688000.0, "22": 5544000.0, "23": 5544000.0, "24": 5544000.0, "25": 5544000.0, "26": 5544000.0, "27": 2688000.0, "29": 5544000.0, "30": 5544000.0}}, "ALAB001572": {"5": {"4": 3720000.0, "5": 3840000.0, "6": 3840000.0, "7": 3840000.0, "8": 3360000.0, "11": 2640000.0, "12": 3840000.0, "13": 3840000.0, "14": 3840000.0, "15": 3840000.0, "16": 1920000.0, "18": 3840000.0, "19": 3840000.0, "20": 3840000.0, "21": 3840000.0, "22": 3840000.0, "23": 1920000.0, "25": 3840000.0, "26": 3840000.0, "27": 3840000.0, "28": 3840000.0, "29": 3840000.0, "30": 1920000.0}, "6": {"1": 7920000.0, "2": 7920000.0, "3": 7920000.0, "4": 5880000.0, "5": 3960000.0, "6": 5880000.0, "7": 3960000.0, "8": 7920000.0, "9": 7920000.0, "10": 7920000.0, "11": 5880000.0, "13": 3960000.0, "14": 3960000.0, "15": 7920000.0, "16": 7920000.0, "17": 4470000.0, "18": 3960000.0, "19": 3960000.0, "20": 1920000.0, "22": 3960000.0, "23": 3960000.0, "24": 3960000.0, "25": 3960000.0, "26": 3960000.0, "27": 1920000.0, "29": 3960000.0, "30": 3960000.0}}, "ALAB001573": {"5": {"4": 434000.0, "5": 448000.0, "6": 448000.0, "7": 448000.0, "8": 392000.0, "11": 308000.0, "12": 448000.0, "13": 448000.0, "14": 448000.0, "15": 448000.0, "16": 224000.0, "18": 448000.0, "19": 448000.0, "20": 448000.0, "21": 448000.0, "22": 448000.0, "23": 224000.0, "25": 448000.0, "26": 448000.0, "27": 448000.0, "28": 448000.0, "29": 448000.0, "30": 224000.0}, "6": {"1": 924000.0, "2": 924000.0, "3": 924000.0, "4": 686000.0, "5": 462000.0, "6": 686000.0, "7": 462000.0, "8": 924000.0, "9": 924000.0, "10": 924000.0, "11": 686000.0, "13": 462000.0, "14": 462000.0, "15": 924000.0, "16": 924000.0, "17": 521500.0, "18": 462000.0, "19": 462000.0, "20": 224000.0, "22": 462000.0, "23": 462000.0, "24": 462000.0, "25": 462000.0, "26": 462000.0, "27": 224000.0, "29": 462000.0, "30": 462000.0}}, "ALAC000077": {"5": {"4": 520.0, "5": 640.0, "6": 640.0, "7": 640.0, "8": 560.0, "11": 440.0, "12": 640.0, "13": 640.0, "14": 640.0, "15": 640.0, "16": 320.0, "18": 640.0, "19": 640.0, "20": 640.0, "21": 640.0, "22": 640.0, "23": 320.0, "25": 640.0, "26": 640.0, "27": 640.0, "28": 640.0, "29": 640.0, "30": 320.0}, "6": {"1": 1225.0, "2": 1185.0, "3": 925.0, "4": 540.0, "5": 530.0, "6": 705.0, "7": 520.0, "8": 940.0, "9": 1060.0, "10": 1170.0, "11": 735.0, "13": 270.0, "14": 95.0, "15": 570.0, "16": 560.0, "17": 815.0, "18": 390.0, "19": 550.0, "20": 285.0, "22": 560.0, "23": 530.0, "24": 550.0, "25": 470.0, "26": 340.0, "27": 205.0, "29": 330.0, "30": 450.0}}, "ALAC000386": {"5": {"4": 5440.0, "5": 5440.0, "6": 5440.0, "7": 5440.0, "8": 4760.0, "11": 3740.0, "12": 5440.0, "13": 5440.0, "14": 5440.0, "15": 5440.0, "16": 2720.0, "18": 5440.0, "19": 5440.0, "20": 5440.0, "21": 5440.0, "22": 5440.0, "23": 2720.0, "25": 5440.0, "26": 5440.0, "27": 5440.0, "28": 5440.0, "29": 5440.0, "30": 2720.0}, "6": {"1": 11220.0, "2": 11220.0, "3": 11220.0, "4": 8330.0, "5": 5610.0, "6": 8330.0, "7": 5610.0, "8": 11220.0, "9": 11220.0, "10": 11220.0, "11": 8330.0, "13": 5610.0, "14": 5610.0, "15": 11220.0, "16": 11220.0, "17": 8712.5, "18": 5610.0, "19": 5610.0, "20": 2720.0, "22": 5610.0, "23": 5610.0, "24": 5610.0, "25": 5610.0, "26": 5610.0, "27": 2720.0, "29": 5610.0, "30": 5610.0}}, "ALAC011940": {"5": {"4": 186000.0, "5": 192000.0, "6": 192000.0, "7": 192000.0, "8": 168000.0, "11": 132000.0, "12": 192000.0, "13": 192000.0, "14": 192000.0, "15": 192000.0, "16": 96000.0, "18": 192000.0, "19": 192000.0, "20": 192000.0, "21": 192000.0, "22": 192000.0, "23": 96000.0, "25": 192000.0, "26": 192000.0, "27": 192000.0, "28": 192000.0, "29": 192000.0, "30": 96000.0}, "6": {"1": 396000.0, "2": 396000.0, "3": 396000.0, "4": 294000.0, "5": 198000.0, "6": 294000.0, "7": 198000.0, "8": 396000.0, "9": 396000.0, "10": 396000.0, "11": 294000.0, "13": 198000.0, "14": 198000.0, "15": 396000.0, "16": 396000.0, "17": 223500.0, "18": 198000.0, "19": 198000.0, "20": 96000.0, "22": 198000.0, "23": 198000.0, "24": 198000.0, "25": 198000.0, "26": 198000.0, "27": 96000.0, "29": 198000.0, "30": 198000.0}}, "ALAC011941": {"5": {"4": 279000.0, "5": 288000.0, "6": 288000.0, "7": 288000.0, "8": 252000.0, "11": 198000.0, "12": 288000.0, "13": 288000.0, "14": 288000.0, "15": 288000.0, "16": 144000.0, "18": 288000.0, "19": 288000.0, "20": 288000.0, "21": 288000.0, "22": 288000.0, "23": 144000.0, "25": 288000.0, "26": 288000.0, "27": 288000.0, "28": 288000.0, "29": 288000.0, "30": 144000.0}, "6": {"1": 594000.0, "2": 594000.0, "3": 594000.0, "4": 441000.0, "5": 297000.0, "6": 441000.0, "7": 297000.0, "8": 594000.0, "9": 594000.0, "10": 594000.0, "11": 441000.0, "13": 297000.0, "14": 297000.0, "15": 594000.0, "16": 594000.0, "17": 335250.0, "18": 297000.0, "19": 297000.0, "20": 144000.0, "22": 297000.0, "23": 297000.0, "24": 297000.0, "25": 297000.0, "26": 297000.0, "27": 144000.0, "29": 297000.0, "30": 297000.0}}, "ALAC011971": {"5": {"4": 520.0, "5": 640.0, "6": 640.0, "7": 640.0, "8": 560.0, "11": 440.0, "12": 640.0, "13": 640.0, "14": 640.0, "15": 640.0, "16": 320.0, "18": 640.0, "19": 640.0, "20": 640.0, "21": 640.0, "22": 640.0, "23": 320.0, "25": 640.0, "26": 640.0, "27": 640.0, "28": 640.0, "29": 640.0, "30": 320.0}, "6": {"1": 1225.0, "2": 1185.0, "3": 925.0, "4": 540.0, "5": 530.0, "6": 705.0, "7": 520.0, "8": 940.0, "9": 1060.0, "10": 1170.0, "11": 735.0, "13": 270.0, "14": 95.0, "15": 570.0, "16": 560.0, "17": 815.0, "18": 390.0, "19": 550.0, "20": 285.0, "22": 560.0, "23": 530.0, "24": 550.0, "25": 470.0, "26": 340.0, "27": 205.0, "29": 330.0, "30": 450.0}}, "ALAC011977": {"5": {"4": 416.0, "5": 416.0, "6": 416.0, "7": 416.0, "8": 364.0, "11": 286.0, "12": 416.0, "13": 416.0, "14": 416.0, "15": 416.0, "16": 208.0, "18": 416.0, "19": 416.0, "20": 416.0, "21": 416.0, "22": 416.0, "23": 208.0, "25": 416.0, "26": 416.0, "27": 416.0, "28": 416.0, "29": 416.0, "30": 208.0}, "6": {"1": 858.0, "2": 858.0, "3": 858.0, "4": 637.0, "5": 429.0, "6": 637.0, "7": 429.0, "8": 858.0, "9": 858.0, "10": 858.0, "11": 637.0, "13": 429.0, "14": 429.0, "15": 858.0, "16": 858.0, "17": 666.25, "18": 429.0, "19": 429.0, "20": 208.0, "22": 429.0, "23": 429.0, "24": 429.0, "25": 429.0, "26": 429.0, "27": 208.0, "29": 429.0, "30": 429.0}}, "ALAC012158": {"5": {"4": 520.0, "5": 640.0, "6": 640.0, "7": 640.0, "8": 560.0, "11": 440.0, "12": 640.0, "13": 640.0, "14": 640.0, "15": 640.0, "16": 320.0, "18": 640.0, "19": 640.0, "20": 640.0, "21": 640.0, "22": 640.0, "23": 320.0, "25": 640.0, "26": 640.0, "27": 640.0, "28": 640.0, "29": 640.0, "30": 320.0}, "6": {"1": 1225.0, "2": 1185.0, "3": 925.0, "4": 540.0, "5": 530.0, "6": 705.0, "7": 520.0, "8": 940.0, "9": 1060.0, "10": 1170.0, "11": 735.0, "13": 270.0, "14": 95.0, "15": 570.0, "16": 560.0, "17": 815.0, "18": 390.0, "19": 550.0, "20": 285.0, "22": 560.0, "23": 530.0, "24": 550.0, "25": 470.0, "26": 340.0, "27": 205.0, "29": 330.0, "30": 450.0}}, "ALAC012327": {"5": {"6": 120.0, "7": 120.0, "11": 120.0, "14": 120.0, "15": 120.0, "25": 115.0, "26": 5.0, "29": 240.0, "30": 120.0}, "6": {"1": 515.0, "2": 85.0, "3": 120.0, "5": 120.0, "6": 160.0, "7": 80.0, "9": 250.0, "10": 230.0, "29": 240.0}}, "ALAC012335": {"5": {"4": 512.0, "5": 512.0, "6": 512.0, "7": 512.0, "8": 448.0, "11": 352.0, "12": 512.0, "13": 512.0, "14": 512.0, "15": 512.0, "16": 256.0, "18": 512.0, "19": 512.0, "20": 512.0, "21": 512.0, "22": 512.0, "23": 256.0, "25": 512.0, "26": 512.0, "27": 512.0, "28": 512.0, "29": 512.0, "30": 256.0}, "6": {"1": 1056.0, "2": 1056.0, "3": 1056.0, "4": 784.0, "5": 528.0, "6": 784.0, "7": 528.0, "8": 1056.0, "9": 1056.0, "10": 1056.0, "11": 784.0, "13": 528.0, "14": 528.0, "15": 1056.0, "16": 1056.0, "17": 820.0, "18": 528.0, "19": 528.0, "20": 256.0, "22": 528.0, "23": 528.0, "24": 528.0, "25": 528.0, "26": 528.0, "27": 256.0, "29": 528.0, "30": 528.0}}, "ALAC012350": {"5": {"4": 520.0, "5": 640.0, "6": 640.0, "7": 640.0, "8": 560.0, "11": 440.0, "12": 640.0, "13": 640.0, "14": 640.0, "15": 640.0, "16": 320.0, "18": 640.0, "19": 640.0, "20": 640.0, "21": 640.0, "22": 640.0, "23": 320.0, "25": 640.0, "26": 640.0, "27": 640.0, "28": 640.0, "29": 640.0, "30": 320.0}, "6": {"1": 1225.0, "2": 1185.0, "3": 925.0, "4": 540.0, "5": 530.0, "6": 705.0, "7": 520.0, "8": 940.0, "9": 1060.0, "10": 1170.0, "11": 735.0, "13": 270.0, "14": 95.0, "15": 570.0, "16": 560.0, "17": 815.0, "18": 390.0, "19": 550.0, "20": 285.0, "22": 560.0, "23": 530.0, "24": 550.0, "25": 470.0, "26": 340.0, "27": 205.0, "29": 330.0, "30": 450.0}}, "ALAC013461": {"5": {"4": 5120.0, "5": 5120.0, "6": 5120.0, "7": 5120.0, "8": 4480.0, "11": 3520.0, "12": 5120.0, "13": 5120.0, "14": 5120.0, "15": 5120.0, "16": 2560.0, "18": 5120.0, "19": 5120.0, "20": 5120.0, "21": 5120.0, "22": 5120.0, "23": 2560.0, "25": 5120.0, "26": 5120.0, "27": 5120.0, "28": 5120.0, "29": 5120.0, "30": 2560.0}, "6": {"1": 10560.0, "2": 10560.0, "3": 10560.0, "4": 7840.0, "5": 5280.0, "6": 7840.0, "7": 5280.0, "8": 10560.0, "9": 10560.0, "10": 10560.0, "11": 7840.0, "13": 5280.0, "14": 5280.0, "15": 10560.0, "16": 10560.0, "17": 8200.0, "18": 5280.0, "19": 5280.0, "20": 2560.0, "22": 5280.0, "23": 5280.0, "24": 5280.0, "25": 5280.0, "26": 5280.0, "27": 2560.0, "29": 5280.0, "30": 5280.0}}, "ALAC013503": {"5": {"4": 490.0, "5": 330.0, "6": 330.0, "7": 390.0, "8": 140.0, "11": 195.0, "12": 210.0, "13": 150.0, "14": 305.0, "15": 400.0, "16": 260.0, "18": 580.0, "19": 600.0, "20": 620.0, "21": 520.0, "22": 580.0, "23": 200.0, "25": 405.0, "26": 390.0, "27": 510.0, "28": 475.0, "29": 400.0, "30": 200.0}, "6": {"1": 590.0, "2": 980.0, "3": 685.0, "4": 300.0, "5": 310.0, "6": 465.0, "7": 440.0, "8": 820.0, "9": 570.0, "10": 880.0, "11": 495.0, "13": 270.0, "14": 95.0, "15": 405.0, "16": 125.0, "17": 590.0, "18": 120.0, "19": 85.0, "20": 45.0, "22": 210.0, "23": 140.0, "24": 90.0, "25": 90.0, "27": 85.0, "29": 90.0, "30": 330.0}}, "ALAC013836": {"5": {"4": 520.0, "5": 640.0, "6": 640.0, "7": 640.0, "8": 560.0, "11": 440.0, "12": 640.0, "13": 640.0, "14": 640.0, "15": 640.0, "16": 320.0, "18": 640.0, "19": 640.0, "20": 640.0, "21": 640.0, "22": 640.0, "23": 320.0, "25": 640.0, "26": 640.0, "27": 640.0, "28": 640.0, "29": 640.0, "30": 320.0}, "6": {"1": 1225.0, "2": 1185.0, "3": 925.0, "4": 540.0, "5": 530.0, "6": 705.0, "7": 520.0, "8": 940.0, "9": 1060.0, "10": 1170.0, "11": 735.0, "13": 270.0, "14": 95.0, "15": 570.0, "16": 560.0, "17": 815.0, "18": 390.0, "19": 550.0, "20": 285.0, "22": 560.0, "23": 530.0, "24": 550.0, "25": 470.0, "26": 340.0, "27": 205.0, "29": 330.0, "30": 450.0}}, "LK012145": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "LK012146": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "LK012147": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "LK012155": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "LK012156": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "LK012714": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "LK013036": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "LK014388": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "LK014389": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "LK014390": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "LK014391": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "LK014393": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "LK014396": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "LK014399": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "LK014400": {"5": {"4": 210.0, "5": 270.0, "6": 230.0, "7": 370.0, "8": 120.0, "11": 240.0, "12": 240.0, "13": 360.0, "14": 240.0, "15": 345.0, "16": 135.0, "18": 240.0, "19": 220.0, "20": 260.0, "21": 360.0, "22": 360.0, "23": 225.0, "25": 370.0, "26": 430.0, "27": 295.0, "28": 265.0, "29": 335.0, "30": 130.0}, "6": {"1": 540.0, "2": 635.0, "3": 625.0, "4": 425.0, "5": 170.0, "6": 505.0, "7": 370.0, "8": 655.0, "9": 675.0, "10": 605.0, "11": 415.0, "15": 345.0, "16": 435.0, "17": 390.0, "18": 495.0, "19": 345.0, "20": 215.0, "22": 400.0, "23": 480.0, "24": 350.0, "25": 370.0, "26": 360.0, "27": 120.0, "29": 295.0, "30": 450.0}}, "LK015519": {"5": {"4": 240.0, "5": 120.0, "7": 120.0, "8": 240.0, "11": 65.0, "12": 190.0, "13": 105.0, "14": 215.0, "15": 145.0, "18": 120.0, "19": 240.0, "21": 120.0, "22": 120.0, "25": 120.0, "27": 120.0, "29": 120.0, "30": 120.0}, "6": {"1": 240.0, "4": 120.0, "5": 160.0, "6": 80.0, "8": 165.0, "9": 75.0, "10": 165.0, "11": 75.0, "15": 120.0, "16": 240.0, "22": 110.0, "23": 30.0, "24": 100.0, "29": 120.0, "30": 120.0}}, "LK015520": {"5": {"4": 240.0, "5": 120.0, "7": 120.0, "8": 240.0, "11": 65.0, "12": 190.0, "13": 105.0, "14": 215.0, "15": 145.0, "18": 120.0, "19": 240.0, "21": 120.0, "22": 120.0, "25": 120.0, "27": 120.0, "29": 120.0, "30": 120.0}, "6": {"1": 240.0, "4": 120.0, "5": 160.0, "6": 80.0, "8": 165.0, "9": 75.0, "10": 165.0, "11": 75.0, "15": 120.0, "16": 240.0, "22": 110.0, "23": 30.0, "24": 100.0, "29": 120.0, "30": 120.0}}, "LK015521": {"5": {"4": 240.0, "5": 120.0, "7": 120.0, "8": 240.0, "11": 65.0, "12": 190.0, "13": 105.0, "14": 215.0, "15": 145.0, "18": 120.0, "19": 240.0, "21": 120.0, "22": 120.0, "25": 120.0, "27": 120.0, "29": 120.0, "30": 120.0}, "6": {"1": 240.0, "4": 120.0, "5": 160.0, "6": 80.0, "8": 165.0, "9": 75.0, "10": 165.0, "11": 75.0, "15": 120.0, "16": 240.0, "22": 110.0, "23": 30.0, "24": 100.0, "29": 120.0, "30": 120.0}}, "LK015522": {"5": {"4": 240.0, "5": 120.0, "7": 120.0, "8": 240.0, "11": 65.0, "12": 190.0, "13": 105.0, "14": 215.0, "15": 145.0, "18": 120.0, "19": 240.0, "21": 120.0, "22": 120.0, "25": 120.0, "27": 120.0, "29": 120.0, "30": 120.0}, "6": {"1": 240.0, "4": 120.0, "5": 160.0, "6": 80.0, "8": 165.0, "9": 75.0, "10": 165.0, "11": 75.0, "15": 120.0, "16": 240.0, "22": 110.0, "23": 30.0, "24": 100.0, "29": 120.0, "30": 120.0}}, "LK015524": {"5": {"4": 240.0, "5": 120.0, "7": 120.0, "8": 240.0, "11": 65.0, "12": 190.0, "13": 105.0, "14": 215.0, "15": 145.0, "18": 120.0, "19": 240.0, "21": 120.0, "22": 120.0, "25": 120.0, "27": 120.0, "29": 120.0, "30": 120.0}, "6": {"1": 240.0, "4": 120.0, "5": 160.0, "6": 80.0, "8": 165.0, "9": 75.0, "10": 165.0, "11": 75.0, "15": 120.0, "16": 240.0, "22": 110.0, "23": 30.0, "24": 100.0, "29": 120.0, "30": 120.0}}, "LK015525": {"5": {"4": 240.0, "5": 120.0, "7": 120.0, "8": 240.0, "11": 65.0, "12": 190.0, "13": 105.0, "14": 215.0, "15": 145.0, "18": 120.0, "19": 240.0, "21": 120.0, "22": 120.0, "25": 120.0, "27": 120.0, "29": 120.0, "30": 120.0}, "6": {"1": 240.0, "4": 120.0, "5": 160.0, "6": 80.0, "8": 165.0, "9": 75.0, "10": 165.0, "11": 75.0, "15": 120.0, "16": 240.0, "22": 110.0, "23": 30.0, "24": 100.0, "29": 120.0, "30": 120.0}}, "LK015526": {"5": {"4": 240.0, "5": 120.0, "7": 120.0, "8": 240.0, "11": 65.0, "12": 190.0, "13": 105.0, "14": 215.0, "15": 145.0, "18": 120.0, "19": 240.0, "21": 120.0, "22": 120.0, "25": 120.0, "27": 120.0, "29": 120.0, "30": 120.0}, "6": {"1": 240.0, "4": 120.0, "5": 160.0, "6": 80.0, "8": 165.0, "9": 75.0, "10": 165.0, "11": 75.0, "15": 120.0, "16": 240.0, "22": 110.0, "23": 30.0, "24": 100.0, "29": 120.0, "30": 120.0}}, "LK015528": {"5": {"4": 240.0, "5": 120.0, "7": 120.0, "8": 240.0, "11": 65.0, "12": 190.0, "13": 105.0, "14": 215.0, "15": 145.0, "18": 120.0, "19": 240.0, "21": 120.0, "22": 120.0, "25": 120.0, "27": 120.0, "29": 120.0, "30": 120.0}, "6": {"1": 240.0, "4": 120.0, "5": 160.0, "6": 80.0, "8": 165.0, "9": 75.0, "10": 165.0, "11": 75.0, "15": 120.0, "16": 240.0, "22": 110.0, "23": 30.0, "24": 100.0, "29": 120.0, "30": 120.0}}, "LK015529": {"5": {"4": 240.0, "5": 120.0, "7": 120.0, "8": 240.0, "11": 65.0, "12": 190.0, "13": 105.0, "14": 215.0, "15": 145.0, "18": 120.0, "19": 240.0, "21": 120.0, "22": 120.0, "25": 120.0, "27": 120.0, "29": 120.0, "30": 120.0}, "6": {"1": 240.0, "4": 120.0, "5": 160.0, "6": 80.0, "8": 165.0, "9": 75.0, "10": 165.0, "11": 75.0, "15": 120.0, "16": 240.0, "22": 110.0, "23": 30.0, "24": 100.0, "29": 120.0, "30": 120.0}}, "LK015530": {"5": {"4": 240.0, "5": 120.0, "7": 120.0, "8": 240.0, "11": 65.0, "12": 190.0, "13": 105.0, "14": 215.0, "15": 145.0, "18": 120.0, "19": 240.0, "21": 120.0, "22": 120.0, "25": 120.0, "27": 120.0, "29": 120.0, "30": 120.0}, "6": {"1": 240.0, "4": 120.0, "5": 160.0, "6": 80.0, "8": 165.0, "9": 75.0, "10": 165.0, "11": 75.0, "15": 120.0, "16": 240.0, "22": 110.0, "23": 30.0, "24": 100.0, "29": 120.0, "30": 120.0}}, "LK015560": {"5": {"4": 240.0, "5": 120.0, "7": 120.0, "8": 240.0, "11": 65.0, "12": 190.0, "13": 105.0, "14": 215.0, "15": 145.0, "18": 120.0, "19": 240.0, "21": 120.0, "22": 120.0, "25": 120.0, "27": 120.0, "29": 120.0, "30": 120.0}, "6": {"1": 240.0, "4": 120.0, "5": 160.0, "6": 80.0, "8": 165.0, "9": 75.0, "10": 165.0, "11": 75.0, "15": 120.0, "16": 240.0, "22": 110.0, "23": 30.0, "24": 100.0, "29": 120.0, "30": 120.0}}, "LK015561": {"5": {"4": 240.0, "5": 120.0, "7": 120.0, "8": 240.0, "11": 65.0, "12": 190.0, "13": 105.0, "14": 215.0, "15": 145.0, "18": 120.0, "19": 240.0, "21": 120.0, "22": 120.0, "25": 120.0, "27": 120.0, "29": 120.0, "30": 120.0}, "6": {"1": 240.0, "4": 120.0, "5": 160.0, "6": 80.0, "8": 165.0, "9": 75.0, "10": 165.0, "11": 75.0, "15": 120.0, "16": 240.0, "22": 110.0, "23": 30.0, "24": 100.0, "29": 120.0, "30": 120.0}}, "LK015567": {"5": {"4": 240.0, "5": 120.0, "7": 120.0, "8": 240.0, "11": 65.0, "12": 190.0, "13": 105.0, "14": 215.0, "15": 145.0, "18": 120.0, "19": 240.0, "21": 120.0, "22": 120.0, "25": 120.0, "27": 120.0, "29": 120.0, "30": 120.0}, "6": {"1": 240.0, "4": 120.0, "5": 160.0, "6": 80.0, "8": 165.0, "9": 75.0, "10": 165.0, "11": 75.0, "15": 120.0, "16": 240.0, "22": 110.0, "23": 30.0, "24": 100.0, "29": 120.0, "30": 120.0}}, "LK015569": {"5": {"4": 240.0, "5": 120.0, "7": 120.0, "8": 240.0, "11": 65.0, "12": 190.0, "13": 105.0, "14": 215.0, "15": 145.0, "18": 120.0, "19": 240.0, "21": 120.0, "22": 120.0, "25": 120.0, "27": 120.0, "29": 120.0, "30": 120.0}, "6": {"1": 240.0, "4": 120.0, "5": 160.0, "6": 80.0, "8": 165.0, "9": 75.0, "10": 165.0, "11": 75.0, "15": 120.0, "16": 240.0, "22": 110.0, "23": 30.0, "24": 100.0, "29": 120.0, "30": 120.0}}, "LK015570": {"5": {"4": 240.0, "5": 120.0, "7": 120.0, "8": 240.0, "11": 65.0, "12": 190.0, "13": 105.0, "14": 215.0, "15": 145.0, "18": 120.0, "19": 240.0, "21": 120.0, "22": 120.0, "25": 120.0, "27": 120.0, "29": 120.0, "30": 120.0}, "6": {"1": 240.0, "4": 120.0, "5": 160.0, "6": 80.0, "8": 165.0, "9": 75.0, "10": 165.0, "11": 75.0, "15": 120.0, "16": 240.0, "22": 110.0, "23": 30.0, "24": 100.0, "29": 120.0, "30": 120.0}}, "LK015571": {"5": {"4": 240.0, "5": 120.0, "7": 120.0, "8": 240.0, "11": 65.0, "12": 190.0, "13": 105.0, "14": 215.0, "15": 145.0, "18": 120.0, "19": 240.0, "21": 120.0, "22": 120.0, "25": 120.0, "27": 120.0, "29": 120.0, "30": 120.0}, "6": {"1": 240.0, "4": 120.0, "5": 160.0, "6": 80.0, "8": 165.0, "9": 75.0, "10": 165.0, "11": 75.0, "15": 120.0, "16": 240.0, "22": 110.0, "23": 30.0, "24": 100.0, "29": 120.0, "30": 120.0}}, "LK015572": {"5": {"4": 240.0, "5": 120.0, "7": 120.0, "8": 240.0, "11": 65.0, "12": 190.0, "13": 105.0, "14": 215.0, "15": 145.0, "18": 120.0, "19": 240.0, "21": 120.0, "22": 120.0, "25": 120.0, "27": 120.0, "29": 120.0, "30": 120.0}, "6": {"1": 240.0, "4": 120.0, "5": 160.0, "6": 80.0, "8": 165.0, "9": 75.0, "10": 165.0, "11": 75.0, "15": 120.0, "16": 240.0, "22": 110.0, "23": 30.0, "24": 100.0, "29": 120.0, "30": 120.0}}, "LP000293": {"5": {"4": 240.0, "5": 120.0, "7": 120.0, "8": 240.0, "11": 65.0, "12": 190.0, "13": 105.0, "14": 215.0, "15": 145.0, "18": 120.0, "19": 240.0, "21": 120.0, "22": 120.0, "25": 120.0, "27": 120.0, "29": 120.0, "30": 120.0}, "6": {"1": 240.0, "4": 120.0, "5": 160.0, "6": 80.0, "8": 165.0, "9": 75.0, "10": 165.0, "11": 75.0, "15": 120.0, "16": 240.0, "22": 110.0, "23": 30.0, "24": 100.0, "29": 120.0, "30": 120.0}}, "LP000294": {"5": {"4": 240.0, "5": 120.0, "7": 120.0, "8": 240.0, "11": 65.0, "12": 190.0, "13": 105.0, "14": 215.0, "15": 145.0, "18": 120.0, "19": 240.0, "21": 120.0, "22": 120.0, "25": 120.0, "27": 120.0, "29": 120.0, "30": 120.0}, "6": {"1": 240.0, "4": 120.0, "5": 160.0, "6": 80.0, "8": 165.0, "9": 75.0, "10": 165.0, "11": 75.0, "15": 120.0, "16": 240.0, "22": 110.0, "23": 30.0, "24": 100.0, "29": 120.0, "30": 120.0}}, "LP000295": {"5": {"6": 240.0, "16": 20.0, "18": 100.0, "20": 120.0, "27": 15.0, "28": 225.0}, "6": {"1": 120.0, "2": 120.0, "5": 120.0, "6": 120.0, "8": 120.0, "9": 190.0, "10": 170.0, "11": 120.0, "16": 120.0, "19": 120.0}}, "LP000339": {"5": {"4": 100.0, "5": 160.0, "6": 80.0, "7": 60.0, "8": 120.0, "11": 60.0, "12": 120.0, "13": 85.0, "14": 95.0, "15": 60.0, "16": 120.0, "18": 90.0, "19": 90.0, "20": 170.0, "21": 70.0, "22": 70.0, "23": 50.0, "25": 60.0, "26": 120.0, "27": 120.0, "28": 60.0, "29": 95.0, "30": 25.0}, "6": {"1": 100.0, "2": 80.0, "3": 60.0, "4": 60.0, "5": 120.0, "8": 120.0, "9": 60.0}}, "LP000340": {"5": {"4": 100.0, "5": 160.0, "6": 80.0, "7": 60.0, "8": 120.0, "11": 60.0, "12": 120.0, "13": 85.0, "14": 95.0, "15": 60.0, "16": 120.0, "18": 90.0, "19": 90.0, "20": 170.0, "21": 70.0, "22": 70.0, "23": 50.0, "25": 60.0, "26": 120.0, "27": 120.0, "28": 60.0, "29": 95.0, "30": 25.0}, "6": {"1": 100.0, "2": 80.0, "3": 60.0, "4": 60.0, "5": 120.0, "8": 120.0, "9": 60.0}}}
# Convert string keys to int for day numbers
v7_daily = {code: {int(mnum): {int(d): q for d,q in days.items()}
                   for mnum,days in months.items()}
            for code, months in v7_daily_raw.items()}

# ── Исправления единиц измерения (согласовано) ───────────────
# ALAB000046: норма A08 была введена в кг (0.0718) при нормах в граммах — фикс.
if 'ALAB000046' in chem_norms and 0 < chem_norms['ALAB000046'].get('A08', 0) < 1:
    chem_norms['ALAB000046']['A08'] = round(chem_norms['ALAB000046']['A08'] * 1000, 1)
    print("[INIT] Норма ALAB000046 A08 переведена из кг в граммы (71.8 г/авто)")
# Эфтек: эти материалы ведём в КИЛОГРАММАХ (нормы были в граммах; остатки
# warehouse 16 и упаковки-бочки 250/218/246 — уже в кг).
EFTEC_KG_CODES = {'ALAB001571', 'ALAB001572', 'ALAB001573', 'ALAC011940', 'ALAC011941'}
for _c_kg in EFTEC_KG_CODES:
    if _c_kg in chem_norms:
        chem_norms[_c_kg] = {m: round(v / 1000.0, 4) for m, v in chem_norms[_c_kg].items()}
print(f"[INIT] Эфтек в кг: {', '.join(sorted(EFTEC_KG_CODES))} (нормы г → кг)")
# Color-filtered paints need V7 demand (requires Order Statistics not uploaded)
# ALAA005669 excluded — FIX#2 already corrected its plan tab
color_filtered_paints = {k for k,v in paint_colors.items() if isinstance(v, list) and k != 'ALAA005669'}

# ── Грунты: потребность = Σ(кузова цветов связанных красок) × норма грунта по модели ──
# Цвета берутся из paint_colors связанных красок; кузова фильтруются по BOM-применяемости краски.
PRIMER_LINKED_PAINTS = {
    'ALAA003255': ['ALAA004871', 'ALAA004873', 'ALAA003257', 'ALAA007466', 'ALAA007732'],  # светлый грунт
    'ALAA003256': ['ALAA004586', 'ALAA005669', 'ALAA005672', 'ALAA007467'],                 # тёмный грунт
}

# ── Замена кодов красок YIERTE → Litum (полная смена SKU с месяца from_month) ──
PAINT_REPLACEMENT_SUPPLIER = 'Litum'
PAINT_CODE_REPLACEMENTS = [
    {'old': 'ALAA004586', 'new': 'ALAA009173', 'from_month': 8},   # Golden Black — с августа
    {'old': 'ALAA003257', 'new': 'ALAA009165', 'from_month': 9},   # Hamilton white C1 — с сентября
    # Запланировано (from_month=None — код в BOM, потребность не переключается)
    {'old': 'ALAA003286', 'new': 'ALAA009166', 'from_month': None},
]
# ── Частичный перенос грунта на Litum: только доля, связанная с переходящей краской ──
# YIERTE-грунт остаётся для остальных связанных красок.
PRIMER_SPLIT_TRANSITIONS = [
    {
        'yierte_primer': 'ALAA003256', 'litum_primer': 'ALAA009164', 'from_month': 8,
        'paint_old': 'ALAA004586', 'paint_new': 'ALAA009173',
    },
    {
        'yierte_primer': 'ALAA003255', 'litum_primer': 'ALAA009163', 'from_month': 9,
        'paint_old': 'ALAA003257', 'paint_new': 'ALAA009165',
    },
]
REPLACEMENT_FROM_MONTH = {}   # old → (new, from_month) — только полная замена краски
REPLACEMENT_NEW_CODES = {}    # new → (old, from_month)
LITUM_PRIMER_CODES = {}       # litum_primer → from_month


def _linked_paint_color_keys(paint_code):
    cols = paint_colors.get(paint_code)
    if paint_code == 'ALAA005669' and not isinstance(cols, list):
        return ['FU_GREY']
    return list(cols) if isinstance(cols, list) else []


def _primer_color_keys(primer_code):
    keys = []
    for paint_code in PRIMER_LINKED_PAINTS.get(primer_code, []):
        for ck in _linked_paint_color_keys(paint_code):
            if ck not in keys:
                keys.append(ck)
    return keys


paint_colors['ALAA005669'] = ['FU_GREY']
for _primer_code in PRIMER_LINKED_PAINTS:
    paint_colors[_primer_code] = _primer_color_keys(_primer_code)
print(f"  tab_map={len(part_tab_map)}, chem={len(chem_norms)}, bumper={len(bumper_clr_map)}, v7_daily={len(v7_daily)}, color_paints={len(color_filtered_paints)}")

# ═══ STEP 1: Stock from STOCK_FILES (multi-format parser) ═══
print("\nLoading stock from files...")
IS_PART_CODE = re.compile(r'^[A-Za-z0-9А-ЯЁа-яё0-9]{8,20}$')  # включает кириллицу для обнаружения

# Замена визуально идентичных кириллических букв на латинские
_CYR_LAT = {
    'А':'A','В':'B','С':'C','Е':'E','Н':'H','К':'K','М':'M','О':'O','Р':'P','Т':'T','Х':'X','У':'Y',
    'а':'a','в':'b','с':'c','е':'e','н':'h','к':'k','м':'m','о':'o','р':'p','т':'t','х':'x','у':'y',
}

def normalize_code(code):
    """Нормализует партномер: заменяет кириллические омоглифы на латинские,
    приводит к верхнему регистру для сравнения.
    Пример: '2803104XKN61A8Т' (кир Т) → '2803104XKN61A8T'"""
    if not code: return code
    return ''.join(_CYR_LAT.get(ch, ch) for ch in code.strip())

def _normalize_stock_date(value):
    """Нормализует дату пересчёта остатков (заголовок файла / имя / Excel serial)."""
    if value is None or value == '':
        return None
    if isinstance(value, datetime.datetime):
        dt = value.date()
    elif isinstance(value, datetime.date):
        dt = value
    elif isinstance(value, (int, float)) and value > 0:
        try:
            dt = from_excel(value).date()
        except Exception:
            return None
    else:
        s = str(value).strip()
        if not s:
            return None
        dt = None
        for fmt in ('%d.%m.%Y', '%d.%m.%y', '%Y-%m-%d', '%d/%m/%Y', '%d/%m/%y', '%d-%m-%Y', '%d-%m-%y'):
            try:
                dt = datetime.datetime.strptime(s, fmt).date()
                break
            except ValueError:
                pass
        if dt is None:
            return None
    if dt.year < 2010 or dt.year > 2100:
        return None
    return dt

def _header_col_date(header_row, col_idx, fallback_date):
    if col_idx is None or col_idx >= len(header_row):
        return fallback_date
    return _normalize_stock_date(header_row[col_idx]) or fallback_date

def _add_stock_snapshot(snapshots, code, snap_date, qty):
    if qty is None or qty <= 0 or not snap_date:
        return
    snapshots.setdefault(code, {})
    snapshots[code][snap_date] = snapshots[code].get(snap_date, 0) + qty

def _update_stock_date(dates, code, snap_date):
    if snap_date and (code not in dates or snap_date > dates[code]):
        dates[code] = snap_date

def _filter_stock_to_calc(stock, stock_src, stock_date_map, calc_codes, stock_by_date=None):
    """Оставляет только остатки деталей, входящих в расчёт MRP."""
    calc_norm = {normalize_code(c) for c in calc_codes}
    calc_norm.update(str(c).strip() for c in calc_codes)
    filtered, src_f, dates_f, by_date_f = {}, {}, {}, {}
    dropped = 0
    for code, qty in stock.items():
        nc = normalize_code(code)
        if nc in calc_norm or code in calc_norm:
            filtered[code] = qty
            if code in stock_src:
                src_f[code] = stock_src[code]
            if code in stock_date_map:
                dates_f[code] = stock_date_map[code]
        else:
            dropped += 1
    if stock_by_date:
        for code, dt_map in stock_by_date.items():
            nc = normalize_code(code)
            if nc in calc_norm or code in calc_norm:
                by_date_f[code] = dict(dt_map)
    if dropped:
        print(f"  ℹ️  Отфильтровано {dropped} кодов вне расчёта MRP")
    return filtered, src_f, dates_f, by_date_f
import datetime as _dt

def _wh16_qty_norm(raw_qty, raw_weight, unit_str):
    """WH16: qty_result = M(шт упаковок) × F(вес единицы) в БАЗОВОЙ единице кода.
    Коды с весом в граммах ведутся в MRP в ГРАММАХ (исправлено: раньше
    переводилось в кг и остаток занижался в 1000 раз — ALAB001037 и т.п.)."""
    try:
        q = float(raw_qty) if raw_qty is not None else 0
        w = float(str(raw_weight).replace(',','.')) if raw_weight is not None else 1
        unit = str(unit_str).lower().strip().rstrip('.')
        total = q * w
        if unit in ('g','г','г.'): return total          # граммы — как есть (код в г)
        if unit in ('ml','мл'):    return total/1000     # мл → л
        return total  # kg, l, шт и прочее — как есть
    except: return 0

def _parse_wh16_sheet(ws_rows):
    """Парсинг 'warehouse 16 list': SUM(col_M * col_F) per code."""
    result = {}
    h = ws_rows[0]
    # Verify columns: A=code, F=weight, G=unit, M=qty_remaining
    # col indices 0-based: A=0, F=5, G=6, M=12
    for row in ws_rows[1:]:
        code = normalize_code(str(row[0]).strip()) if row[0] else ''
        if not code or code=='nan' or not re.match(r'^[A-Za-z0-9]{8,20}$', code): continue
        weight_f = row[5] if len(row)>5 else None
        unit_g   = str(row[6]).strip() if len(row)>6 and row[6] else ''
        qty_m    = row[12] if len(row)>12 else None
        qty = _wh16_qty_norm(qty_m, weight_f, unit_g)
        result[code] = result.get(code, 0) + qty
    return result

def _safe_num(v):
    """Безопасное преобразование в число: '(143 на отправку) +92' → 143."""
    if v is None: return 0
    try: return float(v)
    except:
        m = re.match(r'^\(?([0-9]+(?:\.[0-9]+)?)', str(v).strip())
        return float(m.group(1)) if m else 0

def _parse_local_accounting(ws_rows):
    """accounting_for_local_materials: Склад+Линия.
    Для КАЖДОЙ строки берёт последнюю дату где есть данные (per-row подход).
    Это важно: разные детали могут иметь данные на разные последние даты.
    Структура: row0=даты, row1=Склад/Линия, col0=Партномер.
    Также обрабатывает скрытые фильтром строки (читает ВСЕ строки без исключения).
    """
    result = {}
    result_dates = {}
    if len(ws_rows) < 3: return result, result_dates
    header_dates = ws_rows[0]
    header_subs  = ws_rows[1]
    # Строим все тройки (дата, col_склад, col_линия)
    triplets = []
    i = 0
    while i < len(header_dates):
        if header_dates[i] is not None and i > 2:
            sub_i  = str(header_subs[i]).strip().lower() if i < len(header_subs) and header_subs[i] else ''
            sub_i1 = str(header_subs[i+1]).strip().lower() if i+1 < len(header_subs) and header_subs[i+1] else ''
            if 'склад' in sub_i and 'линия' in sub_i1:
                triplets.append((header_dates[i], i, i+1))
            elif 'линия' in sub_i and 'склад' in sub_i1:
                triplets.append((header_dates[i], i+1, i))
        i += 1
    if not triplets: return result, result_dates
    # PER-ROW: для каждой детали ищем последнюю дату с данными в её строке
    for row in ws_rows[2:]:
        code = normalize_code(str(row[0]).strip()) if row[0] else ''
        if not code or code == 'nan' or not re.match(r'^[A-Za-z0-9]{8,20}$', code): continue
        # Идём с конца по датам, ищем первую где есть непустые данные
        for date_val, sc, lc in reversed(triplets):
            s_raw = row[sc] if sc < len(row) else None
            l_raw = row[lc] if lc < len(row) else None
            if s_raw is not None or l_raw is not None:
                total = _safe_num(s_raw) + _safe_num(l_raw)
                recalc_dt = _normalize_stock_date(date_val)
                if total > 0:
                    result[code] = total
                    if recalc_dt:
                        result_dates[code] = recalc_dt
                break  # нашли последнюю дату с данными для этой строки
    return result, result_dates

def _sheet_local_recount_format(rows):
    """Формат листа «пересчет локал»: Склад/Линия или DL/WL/COMP."""
    if not rows or len(rows) < 2:
        return None
    subs = ' '.join(str(x or '').upper() for x in (rows[1] or [])[:25])
    if 'DL' in subs and 'WL' in subs and 'COMP' in subs:
        return 'dl_wl_comp'
    subs_l = subs.lower()
    if 'склад' in subs_l and 'линия' in subs_l:
        return 'sklad_linia'
    return None


def _parse_dl_wl_comp_recount(ws_rows):
    """Жуйчuan «пересчет локал»: row0=даты, row1=DL/WL/COMP, col0=код."""
    result = {}
    result_dates = {}
    if len(ws_rows) < 3:
        return result, result_dates
    header_dates = ws_rows[0]
    header_subs = ws_rows[1]
    blocks = []
    i = 0
    while i < len(header_subs):
        sub = str(header_subs[i] or '').strip().upper()
        if sub == 'DL' and i + 2 < len(header_subs):
            sub_w = str(header_subs[i + 1] or '').strip().upper()
            sub_c = str(header_subs[i + 2] or '').strip().upper()
            if sub_w == 'WL' and sub_c == 'COMP':
                dt = _normalize_stock_date(header_dates[i] if i < len(header_dates) else None)
                if not dt:
                    for j in range(i, max(-1, i - 4), -1):
                        if j < len(header_dates):
                            dt = _normalize_stock_date(header_dates[j])
                            if dt:
                                break
                blocks.append((dt, i, i + 1, i + 2))
                i += 3
                continue
        i += 1
    if not blocks:
        return result, result_dates
    fallback_dt = next((b[0] for b in reversed(blocks) if b[0]), None)
    if not fallback_dt:
        fallback_dt = _find_recalc_date_in_sheet(ws_rows, None)
    for row in ws_rows[2:]:
        code = normalize_code(str(row[0]).strip()) if row[0] else ''
        if not code or code == 'nan' or not _re.match(r'^[A-Za-z0-9]{8,20}$', code):
            continue
        for dt, dl, wl, comp in reversed(blocks):
            vals = []
            for c in (dl, wl, comp):
                vals.append(row[c] if c < len(row) else None)
            if not any(v is not None for v in vals):
                continue
            total = sum(_safe_num(v) for v in vals)
            recalc_dt = dt or fallback_dt
            if total > 0:
                result[code] = total
                if recalc_dt:
                    result_dates[code] = recalc_dt
            break
    return result, result_dates

def _parse_bumper_recount(rows):
    """Парсер Пересчет бамперов(3).xlsx — суммирует улица + линия + буфер."""
    out = {}
    if not rows or len(rows) < 2: return out
    h = rows[0]
    col_code = None; cols_qty = []
    for j, v in enumerate(h):
        s = str(v or '').lower()
        if 'партномер' in s or 'sap' in s and 'old' not in s or 'коддетали' in s.replace(' ',''):
            if col_code is None: col_code = j
        # Stock-bearing columns: улица / линия / буфер
        if 'улиц' in s or 'линия' in s or 'линии' in s or 'буфер' in s or '线边' in str(v or '') or '缓存' in str(v or ''):
            cols_qty.append(j)
    if col_code is None: col_code = 0
    # Если по заголовкам ничего не нашли — ищем по ключевым словам в первых 3 строках
    if not cols_qty:
        for r in range(min(3, len(rows))):
            for j, v in enumerate(rows[r]):
                s = str(v or '').lower()
                if any(k in s for k in ['улиц','линия','линии','буфер','线边','缓存']):
                    if j not in cols_qty: cols_qty.append(j)
    if not cols_qty: return out
    for row in rows[1:]:
        if col_code >= len(row): continue
        code = normalize_code(str(row[col_code]).strip()) if row[col_code] else ''
        if not code or not re.match(r'^[A-Za-z0-9]{8,20}$', code): continue
        total = 0.0
        for c in cols_qty:
            if c < len(row):
                try:
                    v = row[c]
                    if v is None or v == '': continue
                    total += float(v)
                except: pass
        if total > 0:
            out[code] = out.get(code, 0) + total
    return out


def _filename_stock_date(fpath):
    """Дата пересчёта ИЗ ИМЕНИ файла или None.
    Форматы: «…_08.08.26.xlsx», «17.08.2026 ГСК.xlsx», «…пересчет 08.08.xlsx» (ДД.ММ
    без года — год берётся из mtime, с откатом на год назад если дата «из будущего»)."""
    bn = os.path.basename(fpath)
    mm = _re.search(r'[._](\d{2})[._](\d{2})[._](\d{2})\.(xlsx|xls)$', bn, _re.IGNORECASE)
    if mm:
        try:
            d, m, y = int(mm.group(1)), int(mm.group(2)), 2000 + int(mm.group(3))
            if 1 <= d <= 31 and 1 <= m <= 12:
                return datetime.date(y, m, d)
        except Exception:
            pass
    # «ДД.ММ.ГГГГ …» в начале имени (напр. «11.06.2026 ГСК.xlsx»)
    mm2 = _re.match(r'^(\d{2})[._](\d{2})[._](\d{4})\b', bn)
    if mm2:
        try:
            d, m, y = int(mm2.group(1)), int(mm2.group(2)), int(mm2.group(3))
            if 1 <= d <= 31 and 1 <= m <= 12 and 2020 <= y <= 2035:
                return datetime.date(y, m, d)
        except Exception:
            pass
    # «… пересчет 08.08.xlsx» — ДД.ММ без года (год из mtime файла)
    stem = os.path.splitext(bn)[0]
    mm3 = None
    # пара должна стоять отдельным «словом»: «пересчет 08.08», но НЕ «остатки_v1.2»
    for cand in _re.finditer(r'(?:^|(?<=[\s(]))(\d{1,2})[._\-](\d{1,2})(?=$|[\s)])', stem):
        mm3 = cand          # берём ПОСЛЕДНЮЮ пару ДД.ММ в имени
    if mm3:
        try:
            d, m = int(mm3.group(1)), int(mm3.group(2))
            if 1 <= d <= 31 and 1 <= m <= 12:
                try:
                    base = datetime.date.fromtimestamp(os.path.getmtime(fpath))
                except Exception:
                    base = datetime.date.today()
                cand_dt = datetime.date(base.year, m, d)
                if (cand_dt - base).days > 45:       # «08.08» в файле от января → прошлый год
                    cand_dt = datetime.date(base.year - 1, m, d)
                return cand_dt
        except Exception:
            pass
    return None


def _parse_file_date(fpath):
    """Дата файла: из имени (см. _filename_stock_date), иначе — дата изменения."""
    named = _filename_stock_date(fpath)
    if named:
        return named
    try:
        return datetime.date.fromtimestamp(os.path.getmtime(fpath))
    except Exception:
        return datetime.date.today()


def _load_all_stock_sheets(fpath):
    """Все вкладки xlsx/xls → {имя_листа: [строки]}."""
    ext = os.path.splitext(fpath)[1].lower()
    if ext == '.xls':
        try:
            xls = pd.ExcelFile(fpath)
            out = {}
            for sname in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name=sname, header=None)
                out[sname] = [
                    tuple(None if (v is None or (isinstance(v, float) and pd.isna(v))) else v
                          for v in row)
                    for row in df.values.tolist()
                ]
            return out
        except Exception as e:
            print(f"  ❌ Ошибка чтения xls {os.path.basename(fpath)}: {e}")
            return {}
    try:
        wb = load_workbook(fpath, read_only=True, data_only=True)
        out = {sname: list(wb[sname].iter_rows(values_only=True)) for sname in wb.sheetnames}
        wb.close()
        return out
    except Exception as e:
        print(f"  ❌ Ошибка чтения {os.path.basename(fpath)}: {e}")
        return {}


def _find_recalc_date_in_sheet(rows, fallback, prefer_fallback=False):
    """Дата пересчёта из шапки листа (или fallback — имя файла / mtime).

    Приоритет:
      1) «сильная» дата — ячейка-дата или ячейка, текст которой и есть дата
         («08.08.2026», «на 08.08.2026»);
      2) fallback, если он взят из ИМЕНИ файла (prefer_fallback=True);
      3) «слабая» дата — дата ВНУТРИ длинной подписи
         («КОЛИЧЕСТВО (улица)21.06.2026\\n数量» — прошлый пересчёт, шапку не обновили);
      4) fallback (mtime).
    Пункт 2 важен: имя файла («…пересчет 08.08.xlsx») операторы правят всегда,
    а подпись колонки — нет; раньше вкладка МСА получала дату 21.06 из подписи и
    её остатки выпадали из расчётного окна.
    """
    def _sane(dt):
        # отсечка мусора: числа остатков (400, 720…) Excel считает датами 1901 г.
        return dt is not None and datetime.date(2024, 1, 1) <= dt <= datetime.date(2030, 12, 31)
    weak = None
    for row in rows[:10]:
        if not row:
            continue
        for cell in row[:18]:
            dt = _normalize_stock_date(cell)
            if _sane(dt):
                return dt
            s = str(cell or '')
            m = _re.search(r'(\d{1,2})[./\-](\d{1,2})[./\-](\d{2,4})', s)
            if m:
                d, mo, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
                if y < 100:
                    y += 2000
                try:
                    cand = datetime.date(y, mo, d)
                except ValueError:
                    continue
                if not _sane(cand):
                    continue
                # дата занимает почти всю ячейку → сильная; внутри подписи → слабая
                if len(s.strip()) - len(m.group(0)) <= 6:
                    return cand
                if weak is None:
                    weak = cand
    if prefer_fallback and fallback:
        return fallback
    if weak is not None:
        return weak
    return fallback


def _sheet_is_bumper_recount(rows, sname=''):
    """Лист бамперов: улица + линия + буфер (не любой «пересчёт»)."""
    sn = str(sname or '').lower()
    if 'бампер' in sn:
        return True
    if not rows:
        return False
    joined = ' '.join(str(x or '').lower() for row in rows[:3] for x in (row or [])[:25])
    if 'партномер' not in joined and 'sap' not in joined and 'коддетали' not in joined.replace(' ', ''):
        return False
    return any(k in joined for k in ('улиц', 'линия', 'линии', 'буфер', '线边', '缓存'))


def _parse_stock_table_rows(rows, default_date, prefer_default=False):
    """Универсальный парсер таблицы остатков (FORMAT A/B/C, AAT). → (dict, snap_date)."""
    out = {}
    if not rows or len(rows) < 2:
        return out, default_date
    h0 = rows[0]
    h1 = rows[1] if len(rows) > 1 else []
    code_col = qty_col = data_start = None
    snap_date = default_date
    # AAT / Purem: «Код материала» + «Количество при инвентаризации» (2 или 3 колонки)
    if rows[0] and len(rows[0]) >= 2:
        h_str = ' '.join(str(x or '').lower() for x in rows[0][:6])
        if 'код' in h_str and 'количеств' in h_str and 'инвентар' in h_str:
            qty_col = 1
            for j, v in enumerate(rows[0]):
                vs = str(v or '').lower()
                if 'количеств' in vs and 'инвентар' in vs:
                    qty_col = j
                    break
            for r in rows[1:]:
                if not r or not r[0]:
                    continue
                code = normalize_code(str(r[0]).strip())
                if not _re.match(r'^[A-Za-z0-9]{8,20}$', code):
                    continue
                if qty_col >= len(r):
                    continue
                try:
                    q = float(r[qty_col]) if r[qty_col] is not None else 0
                except Exception:
                    q = 0
                if q > 0:
                    out[code] = out.get(code, 0) + q
            return out, _find_recalc_date_in_sheet(rows, default_date, prefer_default)
    # FORMAT A: ГСК — date headers + «склад»
    has_sklad = any(i < len(h1) and h1[i] and 'склад' in str(h1[i]).lower()
                    for i in range(1, min(len(h1), 10)))
    if has_sklad:
        date_sklad = [i for i in range(1, len(h0))
                      if i < len(h1) and h1[i] and 'склад' in str(h1[i]).lower()]
        target_col = next((c for c in reversed(date_sklad)
            if any(rows[r][c] is not None for r in range(2, min(8, len(rows)))
                   if c < len(rows[r]))), date_sklad[-1] if date_sklad else None)
        if target_col is None:
            return out, default_date
        code_col, qty_col, data_start = 0, target_col, 2
        snap_date = _header_col_date(h0, target_col, default_date)
    # FORMAT B: даты в заголовке, последняя колонка с данными
    elif any(isinstance(h0[i], (_dt.datetime, _dt.date)) or
             (isinstance(h0[i], str) and _re.search(
                 r'\d{1,2}[.,/]\d{2}[.,/]\d{4}|\d{4}-\d{2}-\d{2}|\d+[а-яА-Яa-zA-Z]+\s*\d{4}', str(h0[i])))
             for i in range(1, min(len(h0), 5)) if h0[i] is not None):
        date_cols = [i for i, v in enumerate(h0)
                     if i > 0 and v is not None and (
                         isinstance(v, (_dt.datetime, _dt.date)) or
                         (isinstance(v, str) and _re.search(r'\d', v)))]
        target_col = next((c for c in reversed(date_cols)
            if any(rows[r][c] is not None for r in range(1, min(6, len(rows)))
                   if c < len(rows[r]))), date_cols[-1] if date_cols else None)
        if target_col is None:
            return out, default_date
        code_col, qty_col, data_start = 0, target_col, 1
        snap_date = _header_col_date(h0, target_col, default_date)
    else:
        snap_date = _find_recalc_date_in_sheet(rows, default_date, prefer_default)
        # FORMAT C: именованные колонки (код + остаток / всего / 合计)
        for hi, hrow in enumerate(rows[:5]):
            for j, v in enumerate(hrow):
                vs = str(v).lower().replace('\n', ' ') if v else ''
                _vs_ns = vs.replace(' ', '')
                if (('коддетали' in _vs_ns or 'кодматериала' in _vs_ns
                        or 'партномер' in vs or 'компонент' in vs
                        or vs.strip() == 'код' or vs.strip().startswith('код '))
                        and code_col is None):
                    code_col = j
                if (('остаток' in vs and 'выдач' in vs) or 'количеств' in vs
                        or 'остатки' in vs or 'qty' in vs) and qty_col is None:
                    qty_col = j
                if 'всего' in vs or '合计' in str(v or ''):
                    if code_col == 0 or code_col is None:
                        qty_col = j
            if code_col is not None and qty_col is not None:
                data_start = hi + 1
                break
        if code_col is None or qty_col is None:
            return out, snap_date
    for row in rows[data_start:]:
        if code_col >= len(row):
            continue
        code = normalize_code(str(row[code_col]).strip()) if row[code_col] else ''
        if not code or code in ('nan', 'None', ''):
            continue
        if not _re.match(r'^[A-Za-z0-9]{8,20}$', code):
            continue
        if qty_col >= len(row):
            continue
        try:
            q = float(row[qty_col]) if row[qty_col] is not None else 0
            if q >= 0:
                out[code] = out.get(code, 0) + q
        except Exception:
            pass
    return out, snap_date


def load_all_stock(file_list):
    stock = {}; sources = {}; dates = {}; stock_by_date = {}
    for fpath in file_list:
        if not os.path.exists(fpath) or os.path.getsize(fpath) == 0:
            print(f"  ⚠️  Файл не найден: {os.path.basename(fpath)}"); continue
        fname = os.path.basename(fpath)
        fdate = _parse_file_date(fpath)
        # дата из ИМЕНИ файла главнее даты, выдранной из подписи колонки
        _fdate_named = _filename_stock_date(fpath) is not None
        sheets = _load_all_stock_sheets(fpath)
        if not sheets:
            continue
        file_stock = {}
        file_snapshots = {}
        sheet_count = 0
        empty_sheets = []      # вкладки без строк — норма
        skipped_sheets = []    # вкладки-справочники (служебный список)
        unparsed_sheets = []   # вкладки с данными, но формат не распознан
        for sname, rows in sheets.items():
            if len(rows) < 2:
                empty_sheets.append(sname)
                continue
            if sname in ('Low Stock  Warnng', '3 months not used', 'expired',
                         'approaching expiration', 'information', 'warehouse used table'):
                skipped_sheets.append(sname)
                continue
            sheet_snap = _find_recalc_date_in_sheet(rows, fdate, _fdate_named)
            # Пересчёт бамперов — только на листах с улица/линия/буфер
            if _sheet_is_bumper_recount(rows, sname):
                partial = _parse_bumper_recount(rows)
                if partial:
                    for code, qty in partial.items():
                        file_stock[code] = file_stock.get(code, 0) + qty
                        _add_stock_snapshot(file_snapshots, code, sheet_snap, qty)
                    sheet_count += 1
                    print(f"    ↳ «{sname}» (бамперы): {len(partial)} кодов, дата {sheet_snap.strftime('%d.%m.%Y')}")
                else:
                    unparsed_sheets.append(f"{sname} (бамперы: нет ненулевых остатков)")
                continue
            _local_fmt = _sheet_local_recount_format(rows)
            if _local_fmt == 'dl_wl_comp':
                partial, partial_dates = _parse_dl_wl_comp_recount(rows)
                for code, qty in partial.items():
                    file_stock[code] = file_stock.get(code, 0) + qty
                    dt = partial_dates.get(code, sheet_snap)
                    _add_stock_snapshot(file_snapshots, code, dt, qty)
                if partial:
                    sheet_count += 1
                    _ld = max(partial_dates.values()) if partial_dates else sheet_snap
                    print(f"    ↳ «{sname}» (жуйчuan DL/WL/COMP): {len(partial)} кодов, дата {_ld.strftime('%d.%m.%Y')}")
                else:
                    unparsed_sheets.append(f"{sname} (DL/WL/COMP: нет ненулевых остатков)")
                continue
            if _local_fmt == 'sklad_linia' or sname == 'пересчет локал':
                partial, partial_dates = _parse_local_accounting(rows)
                for code, qty in partial.items():
                    file_stock[code] = file_stock.get(code, 0) + qty
                    dt = partial_dates.get(code, sheet_snap)
                    _add_stock_snapshot(file_snapshots, code, dt, qty)
                if partial:
                    sheet_count += 1
                    _ld = max(partial_dates.values()) if partial_dates else sheet_snap
                    print(f"    ↳ «{sname}» (локал): {len(partial)} кодов, дата {_ld.strftime('%d.%m.%Y')}")
                else:
                    unparsed_sheets.append(f"{sname} (Склад/Линия: нет ненулевых остатков)")
                continue
            if sname == 'warehouse 16 list':
                partial = _parse_wh16_sheet(rows)
                for code, qty in partial.items():
                    file_stock[code] = file_stock.get(code, 0) + qty
                    _add_stock_snapshot(file_snapshots, code, sheet_snap, qty)
                if partial:
                    sheet_count += 1
                    print(f"    ↳ «{sname}» (WH16): {len(partial)} кодов")
                else:
                    unparsed_sheets.append(f"{sname} (WH16: нет ненулевых остатков)")
                continue
            partial, snap_date = _parse_stock_table_rows(rows, sheet_snap, _fdate_named)
            if partial:
                for code, qty in partial.items():
                    file_stock[code] = file_stock.get(code, 0) + qty
                    _add_stock_snapshot(file_snapshots, code, snap_date, qty)
                sheet_count += 1
                print(f"    ↳ «{sname}»: {len(partial)} кодов, дата {snap_date.strftime('%d.%m.%Y')}")
            else:
                unparsed_sheets.append(f"{sname} ({len(rows)} стр.: формат не распознан / нет остатков)")
        if unparsed_sheets:
            print(f"    ⚠️  {fname}: вкладки БЕЗ загруженных остатков ({len(unparsed_sheets)}):")
            for _us in unparsed_sheets:
                print(f"         ⏭ {_us}")
        if file_stock:
            _snap_dates = sorted({d for dm in file_snapshots.values() for d in dm})
            _date_hint = _snap_dates[-1].strftime('%d.%m.%Y') if _snap_dates else fdate.strftime('%d.%m.%Y')
            print(f"  ✅ {fname} [{sheet_count} вкл., пересчёт {_date_hint}]: {len(file_stock)} деталей")
        else:
            print(f"  ⚠️  {fname}: нет распознанных остатков (пропуск)")
            continue
        for code, qty in file_stock.items():
            stock[code] = stock.get(code, 0) + qty
            sources[code] = sources.get(code, set()); sources[code].add(fname[:25])
            for dt, q in file_snapshots.get(code, {fdate: qty}).items():
                _add_stock_snapshot(stock_by_date, code, dt, q)
                _update_stock_date(dates, code, dt)
    # ── Слияние снимков: суммируем ТОЛЬКО в пределах одной даты ──────────
    # Внутри одной даты слагаемые — разные зоны хранения (склад/линия/буфер,
    # вкладки одного файла, разные поставщики) → сумма корректна.
    # Разные даты — это ПОВТОРНЫЕ замеры одного и того же остатка: суммировать
    # нельзя, берём самый свежий снимок (старые остаются видны в stock_by_date
    # как история и рисуются жёлтыми колонками Ввод_Остатков).
    _merged = 0
    _merge_examples = []
    for code, dt_map in stock_by_date.items():
        if not dt_map:
            continue
        fresh_dt = max(dt_map)
        fresh_qty = dt_map[fresh_dt]
        old_qty = stock.get(code)
        if old_qty is not None and abs(old_qty - fresh_qty) > 1e-6:
            _merged += 1
            if len(_merge_examples) < 5:
                _stale = sorted(d for d in dt_map if d != fresh_dt)
                _merge_examples.append(
                    f"{code}: {old_qty:g} → {fresh_qty:g} "
                    f"(свежий {fresh_dt.strftime('%d.%m')}, отброшены "
                    f"{', '.join(d.strftime('%d.%m') for d in _stale)})")
        stock[code] = fresh_qty
        dates[code] = fresh_dt
    if _merged:
        print(f"  🔁 Снимки на разные даты объединены по свежести: {_merged} кодов "
              f"(раньше складывались — остаток задваивался)")
        for _ex in _merge_examples:
            print(f"       • {_ex}")
    return stock, sources, dates, stock_by_date

stock, stock_src, stock_date_map, file_stock_by_date = load_all_stock(STOCK_FILES)
stock_sources = [(list(s)[0] if s else '?', qty) for code, (qty, s)
                 in {c: (stock.get(c,0), stock_src.get(c,set())) for c in stock}.items()]
print(f"  Итого: {len(stock)} деталей с ненулевым остатком из {len(STOCK_FILES)} файла(ов)")
if not stock:
    print("  ℹ️  Добавьте xlsx/xls с остатками в папку MRP (все вкладки сканируются автоматически)")

# ═══ STEP 1a: Теоретические остатки ═══════════════════════════
# Приоритет: обычные файлы (AAT, GSK…) главные.
# Теоретические остатки:
#   • заполняют пропуски — коды, которых нет в обычных файлах
#   • колонка «Остатки цеха» прибавляется ко всем кодам независимо
# Формат файла: Код | Наименование | Поставщик | Ед. | Тeoр.остаток | Остатки цеха
_theor_filled   = 0   # кодов заполнено из теоретических
_workshop_added = 0   # кодов с добавкой цеха
if THEOR_STOCK_FILE and os.path.exists(THEOR_STOCK_FILE):
    _theor_fdate = _parse_file_date(THEOR_STOCK_FILE)
    print(f"\nLoading теоретические остатки: {os.path.basename(THEOR_STOCK_FILE)} [{_theor_fdate}]")
    try:
        _twb = load_workbook(THEOR_STOCK_FILE, read_only=True, data_only=True)
        _tws = _twb.active
        _theor_rows = list(_tws.iter_rows(values_only=True))
        _twb.close()

        # Определяем колонки по заголовку строки 1
        _t_hdr = [str(v).strip().lower() if v else '' for v in (_theor_rows[0] if _theor_rows else [])]
        def _tcol(keywords):
            for kw in keywords:
                for i, h in enumerate(_t_hdr):
                    if kw in h: return i
            return None

        _col_code    = _tcol(['код', 'part', 'номер'])
        _col_theor   = _tcol(['теор', 'теоретич'])       # Теор.остаток
        _col_workshop= _tcol(['цех', 'workshop', 'остатки цеха'])  # Остатки цеха

        if _col_code is None:
            print("  ⚠️  Не найдена колонка с кодом детали в файле теоретических остатков")
        else:
            for _tr in _theor_rows[1:]:
                if not _tr or _col_code >= len(_tr): continue
                _code = normalize_code(str(_tr[_col_code]).strip()) if _tr[_col_code] else ''
                if not _code or _code in ('nan','None',''): continue

                # Теоретический остаток — только для кодов без обычных остатков
                if _col_theor is not None and _col_theor < len(_tr):
                    try:
                        _tval = float(_tr[_col_theor]) if _tr[_col_theor] is not None else 0.0
                        if _tval > 0 and _code not in stock:
                            stock[_code] = _tval
                            stock_src.setdefault(_code, set()).add(os.path.basename(THEOR_STOCK_FILE)[:25])
                            stock_date_map[_code] = _theor_fdate
                            _add_stock_snapshot(file_stock_by_date, _code, _theor_fdate, _tval)
                            _theor_filled += 1
                    except: pass

                # Остатки цеха — прибавляются всегда; дата обновляется если теор.файл свежее
                if _col_workshop is not None and _col_workshop < len(_tr):
                    try:
                        _wval = float(_tr[_col_workshop]) if _tr[_col_workshop] is not None else 0.0
                        if _wval > 0:
                            stock[_code] = stock.get(_code, 0) + _wval
                            stock_src.setdefault(_code, set()).add(os.path.basename(THEOR_STOCK_FILE)[:25] + ' [цех]')
                            _add_stock_snapshot(file_stock_by_date, _code, _theor_fdate, _wval)
                            # Обновляем дату: теор.файл свежее основного → применяем его дату
                            if _code not in stock_date_map or _theor_fdate > stock_date_map[_code]:
                                stock_date_map[_code] = _theor_fdate
                            _workshop_added += 1
                    except: pass

        print(f"  ✅ Теор.остаток заполнил {_theor_filled} кодов (не было в др. файлах)")
        print(f"  ✅ Остатки цеха добавлены к {_workshop_added} кодам")
    except Exception as _te:
        print(f"  ⚠️  Ошибка чтения теоретических остатков: {_te}")
else:
    print("\n[INIT] Теоретические остатки: файл не найден (необязательно)")

# ═══ STEP 1b: Per-part safety stock days (Stock in days.xlsx) ════
print("\nLoading safety stock days...")
safety_by_code   = {}   # code -> days (int)
safety_by_supp   = {}   # normalized_supplier -> days (int, mode)
if os.path.exists(SAFETY_DAYS_FILE):
    try:
        _wb_s = load_workbook(SAFETY_DAYS_FILE, read_only=True, data_only=True)
        _ws_s = _wb_s.active
        _supp_acc = {}
        for _row in _ws_s.iter_rows(min_row=2, values_only=True):
            _code = str(_row[0]).strip() if _row[0] else ''
            _supp = str(_row[1]).strip() if _row[1] else ''
            _days = _row[2]
            if not _code or _code == 'nan': continue
            try:
                _days = int(float(_days))
            except (TypeError, ValueError):
                continue
            if _days < 0: continue
            safety_by_code[_code.upper()] = _days
            _sn = normalize_supplier(_supp)
            _supp_acc.setdefault(_sn, []).append(_days)
        # Build per-supplier mode as fallback
        from statistics import mode as _mode
        for _sn, _vals in _supp_acc.items():
            try: safety_by_supp[_sn] = _mode(_vals)
            except: safety_by_supp[_sn] = _vals[0]
        _wb_s.close()
        print(f"  Safety days: {len(safety_by_code)} деталей, {len(safety_by_supp)} поставщиков")
    except Exception as _e:
        print(f"  ⚠️  Не удалось загрузить {SAFETY_DAYS_FILE}: {_e}")
else:
    print(f"  ℹ️  Файл {SAFETY_DAYS_FILE} не найден — используется SAFETY_DAYS={SAFETY_DAYS}")

def get_safety_days(code, supplier=''):
    """Return safety stock days for a given code, with fallbacks."""
    v = safety_by_code.get(code.upper())
    if v is not None: return v
    sn = normalize_supplier(supplier)
    v = safety_by_supp.get(sn)
    if v is not None: return v
    return SAFETY_DAYS

def _parse_manual_stock_date(value):
    """Parse stock date from Ввод_Остатков col F; blank means keep current source date."""
    return _normalize_stock_date(value)

def _same_stock_date(left, right):
    left_n = _normalize_stock_date(left)
    right_n = _normalize_stock_date(right)
    return left_n == right_n

def _baseline_stock_date(code, stock_date_baseline):
    return stock_date_baseline.get(code, _stock_date)

def get_stock_date(code):
    """Дата актуальности остатков для данного кода.
    Если код есть в stock_date_map — берём её, иначе глобальная _stock_date."""
    return stock_date_map.get(code, _stock_date)

def _calc_horizon_dates():
    return [datetime.date(MONTH_YEAR[mn], mn, d) for mn, _, nd in MONTHS for d in range(1, nd + 1)]

def _parse_stock_input_header_date(value):
    """Parse dd.mm header on Ввод_Остатков into a date inside the calc horizon."""
    norm = _normalize_stock_date(value)
    if norm is not None:
        return norm
    s = str(value or '').strip()
    if not s:
        return None
    for fmt in ('%d.%m', '%d.%m.%y'):
        try:
            parsed = datetime.datetime.strptime(s, fmt)
            for dt in _calc_horizon_dates():
                if dt.day == parsed.day and dt.month == parsed.month:
                    return dt
            return None
        except ValueError:
            pass
    return None

_MAN_STOCK_HDR_ROW = 3
_MAN_STOCK_DATA_START = 4
_MAN_STOCK_OPEN_COL = 5
_MAN_STOCK_FIRST_DATE_COL = 6

# ═══ STEP 2: Master BOM ══════════════════════════════════════
# Priority: 1) Live BOM_Детальный из MRP_System_v9.xlsx (ручные правки),
#           2) Master_BOM_актуальный.xlsx

def _parse_bom_cfg_cell(v):
    """Ячейка применяемости: ✓ / число / пусто."""
    if v is None:
        return None
    vs = str(v).strip()
    if not vs or vs.lower() in ('nan', 'none'):
        return None
    if vs in ('✓', 'v', 'V', 'x', 'X', 'да', 'Да', 'yes', 'Yes'):
        return 1
    try:
        q = float(vs)
        if q > 0:
            return q
    except (TypeError, ValueError):
        pass
    return None

def load_bom_from_live_output(path, apply_auto_marks=False):
    """
    Читает лист BOM_Детальный из MRP_System_v9.xlsx после ручного редактирования.
    apply_auto_marks=False — не добавлять B02_2WD_elite (сохранить ручные ✓).
    """
    wb = load_workbook(path, data_only=True)
    if 'BOM_Детальный' not in wb.sheetnames:
        wb.close()
        return None
    ws = wb['BOM_Детальный']

    header_row = 4
    cfg_hdr_row = None
    for r in range(1, 8):
        row = next(ws.iter_rows(min_row=r, max_row=r, values_only=True), None)
        if not row:
            continue
        joined = ' '.join(str(x or '').lower() for x in row[:6])
        if 'код' in joined:
            header_row = r
        row_joined = ' '.join(str(x or '').lower() for x in row)
        if 'a01' in row_joined and 'comfort' in row_joined:
            cfg_hdr_row = list(row)
    cfg_start_col = 5
    if cfg_hdr_row:
        for i, v in enumerate(cfg_hdr_row):
            vs = str(v or '').lower()
            if 'a01' in vs and 'comfort' in vs:
                cfg_start_col = i
                break
    cfg_keys = detect_cfg_keys_from_header(cfg_hdr_row)

    bom_out = {}
    sections_out = {}
    ordered_codes = []
    live_packages = {}
    current_section = 'Прочее'

    for row in ws.iter_rows(min_row=header_row + 1, values_only=True):
        row = list(row) if row else []
        while len(row) < 29:
            row.append(None)

        code_raw = row[1]
        code = str(code_raw).strip() if code_raw else ''
        if not code or code.lower() in ('nan', 'none', 'код'):
            sec = str(row[0]).strip() if row[0] else ''
            if not sec and row[2]:
                sec = str(row[2]).strip()
            if sec and len(sec) > 2 and not sec.startswith('№'):
                current_section = sec[:80]
            continue

        if code in OBSOLETE_CODES:
            continue

        name = str(row[2]).strip() if row[2] else ''
        unit = str(row[3]).strip() if row[3] else 'шт'
        supp = normalize_supplier(str(row[4]).strip() if row[4] else '')

        cfgs = {}
        for i, cfg_key in enumerate(cfg_keys):
            col = cfg_start_col + i
            v = _parse_bom_cfg_cell(row[col] if col < len(row) else None)
            if v is not None:
                cfgs[cfg_key] = v
        if code in B16_ELITE:
            cfgs.setdefault('B16_4WD_elite', 1)
        if apply_auto_marks and code not in BOM_NO_DEMAND_CODES_DEFAULT and not _is_prerestyle_bumper_code(code):
            apply_b02_2wd_elite_marks(cfgs, code)

        pkg = 1
        if row[27] is not None:
            try:
                p = float(row[27])
                if p >= 1:
                    pkg = p
                    live_packages[code] = p
            except (TypeError, ValueError):
                pass

        notes = str(row[28]).strip() if row[28] else ''
        bom_out[code] = {
            'name': name,
            'unit': unit or 'шт',
            'supplier': supp,
            'configs': cfgs,
            'model_qty': {},
            'section': current_section,
            'package': pkg,
            'notes': notes,
        }
        sections_out.setdefault(current_section, [])
        if code not in sections_out[current_section]:
            sections_out[current_section].append(code)
        if code not in ordered_codes:
            ordered_codes.append(code)

    wb.close()
    if not bom_out:
        return None
    return bom_out, sections_out, ordered_codes, live_packages


def _snapshot_live_bom_configs():
    """Запоминает ✓ из загруженного LIVE BOM (до авто-правок скрипта)."""
    global live_bom_configs_snapshot
    live_bom_configs_snapshot = {}
    for code, info in bom.items():
        cfgs = info.get('configs') or {}
        if isinstance(cfgs, set):
            cfgs = {k: 1 for k in cfgs}
        live_bom_configs_snapshot[code] = dict(cfgs)


def _finalize_live_bom_configs():
    """LIVE: вернуть применяемость из файла; затем только «Исключить_применяемость»."""
    if not BOM_LIVE_ACTIVE or not live_bom_configs_snapshot:
        return 0
    for code, cfgs in live_bom_configs_snapshot.items():
        if code in bom:
            bom[code]['configs'] = dict(cfgs)
    _apply_no_demand_exclusions(bom)
    return len(live_bom_configs_snapshot)

def _merge_bom_summary_model_qty(bom_dict, summary_path):
    """Дополняет model_qty / notes из BOM_Сводный Master BOM."""
    if not os.path.exists(summary_path):
        return
    BOM_MDL = {5: 'A01', 6: 'A08', 7: 'B02', 8: 'B04', 9: 'B06', 10: 'B16'}
    wb = load_workbook(summary_path, read_only=True, data_only=True)
    if 'BOM_Сводный' not in wb.sheetnames:
        wb.close()
        return
    for row in wb['BOM_Сводный'].iter_rows(min_row=3, values_only=True):
        code = str(row[1]).strip() if row[1] else ''
        if not code or code in ('nan', 'None', 'Код'):
            continue
        if code not in bom_dict:
            continue
        mq = {}
        for col, model in BOM_MDL.items():
            v = row[col] if col < len(row) else None
            if v not in (None, '') and str(v).strip() not in ('', 'None', 'nan'):
                try:
                    mq[model] = float(v)
                except (TypeError, ValueError):
                    pass
        if mq:
            bom_dict[code]['model_qty'] = mq
        notes = str(row[13]).strip() if len(row) > 13 and row[13] else ''
        if notes and not bom_dict[code].get('notes'):
            bom_dict[code]['notes'] = notes
    wb.close()

def _cfg_to_bom_cell(qty):
    try:
        q = float(qty)
        return '✓' if q == 1 else q
    except (TypeError, ValueError):
        return '✓'

def _overlay_live_bom_configs(bom_dict, live_bom_dict, full_replace=True):
    """Накладывает применяемость из LIVE BOM_Детальный (MRP output) на расчётный bom."""
    updated = 0
    for code, info in live_bom_dict.items():
        if code not in bom_dict:
            continue
        src = info.get('configs')
        if src is None:
            continue
        if isinstance(src, set):
            src = {k: 1 for k in src}
        else:
            src = dict(src or {})
        if full_replace:
            bom_dict[code]['configs'] = src
            updated += 1
        elif src:
            dst = bom_dict[code].get('configs', {})
            if isinstance(dst, set):
                dst = {k: 1 for k in dst}
            else:
                dst = dict(dst or {})
            added = False
            for k, v in src.items():
                if k not in dst:
                    dst[k] = v
                    added = True
            if added:
                bom_dict[code]['configs'] = dst
                updated += 1
    return updated


def _apply_live_bom_meta(bom_dict, bom_sections, bom_ordered_codes, live_packages,
                         live_tuple, use_live_structure=False):
    """Упаковки и (опционально) порядок разделов из LIVE BOM."""
    live_bom, live_sections, live_ordered, live_pkgs = live_tuple
    live_packages.clear()
    live_packages.update(live_pkgs or {})
    for code, pkg in live_pkgs.items():
        if code in bom_dict and pkg:
            bom_dict[code]['package'] = pkg
    if use_live_structure and live_sections:
        bom_sections.clear()
        bom_sections.update(live_sections)
        bom_ordered_codes[:] = list(live_ordered)


def _merge_out_bom_configs_into_bom(bom_dict):
    """Дополняет расчётный bom отметками из BOM_Детальный вывода MRP (если Master BOM устарел)."""
    if not os.path.exists(OUT):
        return 0
    _live = load_bom_from_live_output(OUT)
    if not _live:
        return 0
    return _overlay_live_bom_configs(bom_dict, _live[0], full_replace=False)

def _find_master_det_layout(ws):
    """Строка заголовка, начало колонок конфигураций и первая строка данных на BOM_Детальный."""
    header_row_idx = None
    cfg_hdr_row = None
    cfg_hdr = []
    data_start = 5
    for r in range(1, 8):
        row_vals = [c.value for c in ws[r]]
        joined = ' '.join(str(x or '').lower() for x in row_vals[:8])
        if 'код' in joined and ('наименован' in joined or 'поставщик' in joined or 'ед.' in joined):
            header_row_idx = r
            data_start = r + 1
        rj = ' '.join(str(x or '').lower() for x in row_vals)
        if 'a01' in rj and 'comfort' in rj:
            cfg_hdr_row = r
            cfg_hdr = row_vals
    cfg_start = 5
    if cfg_hdr:
        for i, v in enumerate(cfg_hdr):
            vs = str(v or '').lower()
            if 'a01' in vs and 'comfort' in vs:
                cfg_start = i
                break
    return header_row_idx, cfg_hdr_row, cfg_start, data_start, cfg_hdr

def _copy_file_with_retry(src, dst, *, pauses=(0, 1, 2, 3, 5, 8, 12, 15)):
    """Копирует файл с паузами (OneDrive / антивирус часто блокируют перезапись)."""
    import time as _time
    import shutil as _shutil
    last_err = None
    for pause in pauses:
        if pause:
            _time.sleep(pause)
        try:
            _shutil.copy2(src, dst)
            return True, None
        except (PermissionError, OSError) as e:
            last_err = e
    return False, last_err

def _master_bom_autosave_path(master_path):
    return os.path.join(
        os.path.dirname(master_path),
        os.path.splitext(os.path.basename(master_path))[0] + ' (autosave).xlsx',
    )

def _promote_pending_master_bom(master_path):
    """Заменяет master из autosave (OneDrive блокирует copy поверх, но rename часто проходит)."""
    if not master_path or not os.path.isfile(master_path):
        return False
    alt = _master_bom_autosave_path(master_path)
    if not os.path.isfile(alt):
        return False
    try:
        if os.path.getmtime(alt) <= os.path.getmtime(master_path):
            return False
    except OSError:
        pass
    try:
        os.remove(master_path)
        os.replace(alt, master_path)
        print(f"  ♻️  Master BOM обновлён из «{os.path.basename(alt)}»")
        return True
    except OSError:
        return False

def _sync_master_bom_detailed(bom_dict, master_path):
    """Копия ✓ из расчётного bom в Master_BOM (архив). Не влияет на сохранение MRP_System_v9.xlsx."""
    if not master_path or not os.path.exists(master_path):
        return 0
    if os.environ.get('MRP_SKIP_MASTER_BOM_SYNC', '').strip().lower() in ('1', 'true', 'yes'):
        print('  ℹ️  Master BOM sync пропущен (MRP_SKIP_MASTER_BOM_SYNC=1)')
        return 0

    import tempfile as _tempfile
    import shutil as _shutil
    import uuid as _uuid

    _alt = _master_bom_autosave_path(master_path)
    work_dir = os.path.join(ROOT, 'Логи', '_bom_sync')
    try:
        os.makedirs(work_dir, exist_ok=True)
    except OSError:
        work_dir = os.path.join(_tempfile.gettempdir(), 'mrp_master_bom_sync')
        os.makedirs(work_dir, exist_ok=True)
    work_path = os.path.join(work_dir, f'sync_{_uuid.uuid4().hex}.xlsx')

    try:
        _shutil.copy2(master_path, work_path)
    except OSError as e:
        print(f"  ⚠️  Master BOM sync: не удалось скопировать {os.path.basename(master_path)}: {e}")
        return 0

    wb = None
    try:
        wb = load_workbook(work_path)
    except Exception as e:
        print(f"  ⚠️  Master BOM sync: не удалось открыть копию: {e}")
        return 0
    if 'BOM_Детальный' not in wb.sheetnames:
        wb.close()
        print("  ⚠️  Master BOM sync: лист BOM_Детальный не найден")
        return 0
    ws = wb['BOM_Детальный']
    _, cfg_hdr_row, cfg_start, data_start, cfg_hdr = _find_master_det_layout(ws)
    if not _bom_header_has_b02_2wd_elite(cfg_hdr):
        ins_col = cfg_start + CFG_KEYS.index('B02_2WD_elite') + 1
        ws.insert_cols(ins_col)
        lbl = 'B02\n2WD\nelite'
        if cfg_hdr_row:
            ws.cell(cfg_hdr_row, ins_col).value = lbl
            ws.cell(cfg_hdr_row, ins_col).alignment = Alignment(
                horizontal='center', vertical='center', wrap_text=True)
        print(f"  Master BOM sync: добавлена колонка B02_2WD_elite (col {get_column_letter(ins_col)})")
    code_to_row = {}
    for r in range(data_start, ws.max_row + 1):
        raw = ws.cell(r, 2).value
        if raw in (None, ''):
            continue
        code = normalize_code(str(raw).strip())
        if not code or code.lower() in ('nan', 'none', 'код'):
            continue
        code_to_row[code] = r
    updated = 0
    cleared = 0
    b02_elite_written = 0
    for code, info in bom_dict.items():
        nc = normalize_code(code)
        row = code_to_row.get(nc) or code_to_row.get(code)
        if not row:
            continue
        cfgs = info.get('configs') or {}
        if isinstance(cfgs, set):
            cfgs = {k: 1 for k in cfgs}
        else:
            cfgs = dict(cfgs or {})
        for i, cfg_key in enumerate(CFG_KEYS):
            col = cfg_start + i + 1
            if cfg_key in cfgs:
                ws.cell(row, col).value = _cfg_to_bom_cell(cfgs[cfg_key])
                updated += 1
                if cfg_key == 'B02_2WD_elite':
                    b02_elite_written += 1
            elif ws.cell(row, col).value not in (None, ''):
                ws.cell(row, col).value = None
                cleared += 1
    try:
        wb.save(work_path)
    except Exception as e:
        _fallback = work_path + '.retry.xlsx'
        try:
            wb.save(_fallback)
            work_path = _fallback
        except Exception:
            if wb:
                wb.close()
            print(f"  ⚠️  Master BOM sync: не удалось сохранить рабочую копию — {e}")
            return 0
    if wb:
        wb.close()

    _saved = False
    _last_err = None
    _ok, _last_err = _copy_file_with_retry(work_path, master_path)
    if _ok:
        _saved = True
    else:
        _new_path = master_path + '.mrp_new.xlsx'
        try:
            _shutil.copy2(work_path, _new_path)
            try:
                os.remove(master_path)
            except OSError:
                pass
            os.replace(_new_path, master_path)
            _saved = True
        except OSError as e:
            _last_err = e
            for _p in (_new_path, master_path + '.mrp_sync.tmp.xlsx'):
                try:
                    if os.path.isfile(_p):
                        os.remove(_p)
                except OSError:
                    pass

    if _saved:
        print(f"  Master BOM sync: {updated} отметок → {os.path.basename(master_path)}"
              f" (B02_2WD_elite: {b02_elite_written}, снято: {cleared})")
        if os.path.isfile(_alt):
            try:
                os.remove(_alt)
            except OSError:
                pass
    else:
        try:
            _shutil.copy2(work_path, _alt)
            if _promote_pending_master_bom(master_path):
                print(f"  Master BOM sync: {updated} отметок → {os.path.basename(master_path)}"
                      f" (B02_2WD_elite: {b02_elite_written}, снято: {cleared})")
            else:
                print(f"  ⚠️  Master BOM sync: OneDrive блокирует перезапись "
                      f"«{os.path.basename(master_path)}».")
                if 'onedrive' in master_path.lower():
                    print("      При следующем запуске MRP попробует заменить master из autosave.")
                    print("      Либо: ПКМ на файле → «Всегда хранить на этом устройстве».")
                print(f"      Отметки сохранены в «{os.path.basename(_alt)}».")
        except Exception:
            print(f"  ⚠️  Master BOM sync: не удалось сохранить {os.path.basename(master_path)} — {_last_err}")
            updated = 0
    for _cleanup in (work_path, work_path + '.retry.xlsx'):
        try:
            if _cleanup and os.path.isfile(_cleanup):
                os.remove(_cleanup)
        except OSError:
            pass
    return updated

BOM_SOURCE = None
BOM_LIVE_ACTIVE = False
live_packages = {}
bom = {}
bom_sections = {}
bom_ordered_codes = []
BOM_COL_OFFSET = 0

# OneDrive: autosave с прошлого запуска → master (до чтения BOM)
if BOM_NEW and os.path.exists(BOM_NEW):
    _promote_pending_master_bom(BOM_NEW)
elif BOM_FILE and os.path.exists(BOM_FILE):
    _promote_pending_master_bom(BOM_FILE)

# LIVE MODE — приоритет BOM_Детальный из MRP_System_v9.xlsx (ручные ✓ в выводе)
if os.path.exists(OUT_LIVE_INPUT):
    try:
        _live = load_bom_from_live_output(OUT_LIVE_INPUT, apply_auto_marks=False)
        if _live:
            bom, bom_sections, bom_ordered_codes, live_packages = _live
            BOM_SOURCE = OUT_LIVE_INPUT
            BOM_LIVE_ACTIVE = True
            _snapshot_live_bom_configs()
            _master_for_mq = BOM_NEW if (BOM_NEW and os.path.exists(BOM_NEW)) else BOM_FILE
            if _master_for_mq and os.path.exists(_master_for_mq):
                _merge_bom_summary_model_qty(bom, _master_for_mq)
            print(f"\nLoading BOM from MRP output (LIVE MODE — BOM_Детальный из {os.path.basename(OUT_LIVE_INPUT)})")
            print(f"  LIVE: {len(bom)} деталей | упаковок в BOM: {len(live_packages)} | "
                  f"снимок ✓: {len(live_bom_configs_snapshot)} кодов")
    except Exception as _live_e:
        print(f"  ⚠️  Ошибка чтения LIVE BOM_Детальный: {_live_e}")

if not BOM_LIVE_ACTIVE:
    BOM_SOURCE = BOM_NEW if (BOM_NEW and os.path.exists(BOM_NEW)) else BOM_FILE
    print(f"\nLoading Master BOM from: {os.path.basename(BOM_SOURCE)}")

    wb_bom = load_workbook(BOM_SOURCE, data_only=True)
    det_sheet = 'BOM_Детальный'
    min_row = 5

    _header_row_idx = None
    _hdr = []
    for r_idx in range(1, 7):
        r_rows = list(wb_bom[det_sheet].iter_rows(min_row=r_idx, max_row=r_idx, values_only=True))
        if not r_rows:
            continue
        r = r_rows[0]
        joined = ' '.join(str(x or '').lower() for x in r[:8])
        if 'код' in joined and ('наименован' in joined or 'имя' in joined or 'ед.' in joined or 'поставщик' in joined):
            _header_row_idx = r_idx
            _hdr = r
            break
    if not _hdr:
        _hdr_rows = list(wb_bom[det_sheet].iter_rows(min_row=min_row, max_row=min_row, values_only=True))
        _hdr = _hdr_rows[0] if _hdr_rows else []

    _first_cfg_col = None
    for i, v in enumerate(_hdr):
        if v is None:
            continue
        vs = str(v).strip()
        if any(x in vs for x in ('A01', 'A08', 'B02', 'B04', 'B06', 'B16')) and any(
                x in vs for x in ('comfort', 'elite', 'premium', 'Tech', '2WD', '4WD')):
            _first_cfg_col = i
            break
    if _first_cfg_col == 5:
        BOM_COL_OFFSET = 0
    elif _first_cfg_col == 6:
        BOM_COL_OFFSET = 1
    else:
        BOM_COL_OFFSET = 0
    if _header_row_idx:
        min_row = _header_row_idx + 1

    _cfg_read_keys = detect_cfg_keys_from_header(_hdr)
    _cfg_start = _first_cfg_col if _first_cfg_col is not None else (5 + BOM_COL_OFFSET)
    _NAME_COL = 2 + BOM_COL_OFFSET
    _UNIT_COL = 3 + BOM_COL_OFFSET
    _SUPP_COL = 4 + BOM_COL_OFFSET
    print(f"  BOM col offset: {BOM_COL_OFFSET} | конфигов: {len(_cfg_read_keys)}")

    current_section = 'Прочее'
    for row in wb_bom[det_sheet].iter_rows(min_row=min_row + 1, values_only=True):
        code = str(row[1]).strip() if row[1] else ''
        if not code or code == 'nan':
            txt = str(row[0]).strip() if row[0] else ''
            if not txt:
                txt = str(row[2]).strip() if len(row) > 2 and row[2] else ''
            if txt and len(txt) > 4 and not txt.startswith('№') and not txt.startswith('ДЕТАЛЬНЫЙ'):
                current_section = txt[:80]
            continue
        if code in OBSOLETE_CODES:
            print(f"  ⊘ Исключён устаревший код: {code}")
            continue
        name = str(row[_NAME_COL]).strip() if len(row) > _NAME_COL and row[_NAME_COL] else ''
        unit = str(row[_UNIT_COL]).strip() if len(row) > _UNIT_COL and row[_UNIT_COL] else 'шт'
        supp = normalize_supplier(str(row[_SUPP_COL]).strip() if len(row) > _SUPP_COL and row[_SUPP_COL] else '')
        cfgs = {}
        for i, cfg in enumerate(_cfg_read_keys):
            col = _cfg_start + i
            v = row[col] if col < len(row) else None
            qty = _parse_bom_cfg_cell(v)
            if qty is not None:
                cfgs[cfg] = qty
        if code in B16_ELITE:
            cfgs.setdefault('B16_4WD_elite', 1)
        if code not in BOM_NO_DEMAND_CODES_DEFAULT and not _is_prerestyle_bumper_code(code):
            apply_b02_2wd_elite_marks(cfgs, code)
        pkg_from_bom = None
        _pkg_col = 27
        if _pkg_col < len(row) and row[_pkg_col] is not None:
            try:
                p = float(row[_pkg_col])
                if p >= 1:
                    pkg_from_bom = p
            except (TypeError, ValueError):
                pass
        if code not in bom:
            bom[code] = {'name': name, 'unit': unit, 'supplier': supp, 'configs': cfgs,
                         'model_qty': {}, 'section': current_section, 'package': pkg_from_bom or 1, 'notes': ''}
        else:
            bom[code]['configs'].update(cfgs)
            bom[code]['section'] = current_section
            if pkg_from_bom:
                bom[code]['package'] = pkg_from_bom
        bom_sections.setdefault(current_section, [])
        if code not in bom_sections[current_section]:
            bom_sections[current_section].append(code)

    BOM_MDL = {5: 'A01', 6: 'A08', 7: 'B02', 8: 'B04', 9: 'B06', 10: 'B16'}
    if 'BOM_Сводный' in wb_bom.sheetnames:
        sec = 'General'
        for row in wb_bom['BOM_Сводный'].iter_rows(min_row=3, values_only=True):
            col0 = str(row[0]).strip() if row[0] else ''
            code = str(row[1]).strip() if row[1] else ''
            if code in ('', 'nan', 'None', 'Код') and col0 not in ('', 'nan', 'None', '№'):
                sec = col0[:80]
                continue
            if not code or code in ('nan', 'None', 'Код'):
                continue
            mq = {}
            for col, model in BOM_MDL.items():
                v = row[col] if col < len(row) else None
                if v not in (None, '') and str(v).strip() not in ('', 'None', 'nan'):
                    try:
                        mq[model] = float(v)
                    except (TypeError, ValueError):
                        pass
            notes = str(row[13]).strip() if len(row) > 13 and row[13] else ''
            if code in bom:
                bom[code]['model_qty'] = mq
                if not bom[code].get('section') or bom[code]['section'] == 'Прочее':
                    bom[code]['section'] = sec
                if notes:
                    bom[code]['notes'] = notes
            else:
                name = str(row[2]).strip() if row[2] else ''
                unit = str(row[3]).strip() if row[3] else 'шт'
                supp = str(row[4]).strip() if row[4] else ''
                bom[code] = {'name': name, 'unit': unit, 'supplier': supp, 'model_qty': mq,
                             'configs': {}, 'section': sec, 'package': 1, 'notes': notes}

    bom_sections = {}
    bom_ordered_codes = []
    wb_sec = load_workbook(BOM_NEW if os.path.exists(BOM_NEW) else BOM_FILE, data_only=True)
    cur_sec = 'Прочее'
    if 'BOM_Сводный' in wb_sec.sheetnames:
        for row in wb_sec['BOM_Сводный'].iter_rows(min_row=3, values_only=True):
            col0 = str(row[0]).strip() if row[0] else ''
            code = str(row[1]).strip() if row[1] else ''
            if code in ('', 'nan', 'None', 'Код') and col0 not in ('', 'nan', 'None', '№'):
                cur_sec = col0[:80]
                continue
            if not code or code in ('nan', 'None', 'Код'):
                continue
            bom_sections.setdefault(cur_sec, [])
            if code not in bom_sections[cur_sec]:
                bom_sections[cur_sec].append(code)
            if code not in bom_ordered_codes:
                bom_ordered_codes.append(code)
    wb_sec.close()
    wb_bom.close()

bom_mode = 'LIVE (BOM_Детальный из MRP)' if BOM_LIVE_ACTIVE else f'Источник: {os.path.basename(BOM_SOURCE or "?")}'
print(f"  {len(bom)} деталей в BOM | {len(bom_sections)} разделов | Режим: {bom_mode}")

# ── Разделы BOM: если структура выродилась (все коды в одном разделе, напр.
#    «Краски Litum» — заголовки разделов потерялись в LIVE-файле), перегруппируем
#    по разделам Master_BOM_актуальный.xlsx (BOM_Детальный). ──
def _rebuild_sections_from_master():
    global bom_sections, bom_ordered_codes
    real = {s: c for s, c in bom_sections.items() if c}
    _max_sec = max((len(c) for c in real.values()), default=0)
    degenerate = (len(bom) > 50 and (len(real) <= 2 or _max_sec > 0.7 * len(bom)))
    if not degenerate:
        return False
    src = BOM_NEW if (BOM_NEW and os.path.exists(BOM_NEW)) else BOM_FILE
    if not src or not os.path.exists(src):
        return False
    try:
        wb_m = load_workbook(src, read_only=True, data_only=True)
    except Exception as _e:
        print(f"  ⚠️  Разделы BOM: не удалось прочитать {os.path.basename(str(src))}: {_e}")
        return False
    if 'BOM_Детальный' not in wb_m.sheetnames:
        wb_m.close()
        return False
    master_pairs = []   # [(code, section)] в порядке мастера
    cur = 'Прочее'
    seen_secs = []
    for row in wb_m['BOM_Детальный'].iter_rows(min_row=2, values_only=True):
        code = str(row[1]).strip() if len(row) > 1 and row[1] else ''
        if not code or code.lower() in ('nan', 'none', 'код'):
            txt = str(row[0]).strip() if row[0] else ''
            if not txt and len(row) > 2 and row[2]:
                txt = str(row[2]).strip()
            if (txt and len(txt) > 4 and not txt.startswith('№')
                    and not txt.upper().startswith('BOM')):
                cur = txt[:80]
                if cur not in seen_secs:
                    seen_secs.append(cur)
            continue
        master_pairs.append((code, cur))
    wb_m.close()
    if len(seen_secs) < 3:
        return False
    sec_of = {c: s for c, s in master_pairs}
    new_sections = {}
    new_order = []
    for c, s in master_pairs:
        if c not in bom:
            continue
        bom[c]['section'] = s
        new_sections.setdefault(s, [])
        if c not in new_sections[s]:
            new_sections[s].append(c)
        if c not in new_order:
            new_order.append(c)
    # Коды вне мастера — группируем по поставщику (Litum-замены добавятся позже
    # в раздел «Краски Litum» через apply_paint_code_replacements)
    for c in bom:
        if c in sec_of:
            continue
        supp = str(bom[c].get('supplier', '') or '').strip() or 'Прочее'
        s = f'Прочее — {supp}'
        bom[c]['section'] = s
        new_sections.setdefault(s, [])
        if c not in new_sections[s]:
            new_sections[s].append(c)
        if c not in new_order:
            new_order.append(c)
    bom_sections = new_sections
    bom_ordered_codes = new_order
    print(f"  ♻️  Разделы BOM перегруппированы по {os.path.basename(str(src))}: "
          f"{len(new_sections)} разделов")
    return True

try:
    _rebuild_sections_from_master()
except Exception as _sec_e:
    print(f"  ⚠️  Перегруппировка разделов BOM не выполнена: {_sec_e}")

# Оставляем только остатки деталей из BOM / красок / бамперов (расчётные коды)
_pre_calc_codes = sorted(set(list(bom.keys()) + list(chem_norms.keys()) + list(bumper_clr_map.keys())))
stock, stock_src, stock_date_map, file_stock_by_date = _filter_stock_to_calc(
    stock, stock_src, stock_date_map, _pre_calc_codes, file_stock_by_date)
print(f"  Остатки по расчётным деталям: {len(stock)} кодов")

# ═══ STEP 3: Packaging ════════════════════════════════════════
print("Loading packaging...")
# Приоритет упаковок:
# 1) «Упаковка локала.xlsx» (лист Haval_Stock) — основной источник
# 2) PACKAGE_OVERRIDES — точечные правки
# 3) DEFAULT_BUMPER_PKG — если после (1)-(2) осталось 1
# Лист «Упаковка» в LIVE MRP — только примечания (Lear), не размер упаковки.

def _load_packages_from_local_file():
    """Загрузка упаковок из «Упаковка локала» — перезаписывает bom[].package."""
    loaded = 0
    if not PKG_FILE or not os.path.isfile(PKG_FILE):
        print(f"  ⚠️  Файл не найден: Упаковка локала ({PKG_FILE or '—'})")
        return loaded
    try:
        wb = load_workbook(PKG_FILE, read_only=True, data_only=True)
    except Exception as e:
        print(f"  ⚠️  Ошибка чтения {os.path.basename(PKG_FILE)}: {e}")
        return loaded
    sheet = 'Haval_Stock' if 'Haval_Stock' in wb.sheetnames else wb.sheetnames[0]
    for row in wb[sheet].iter_rows(min_row=2, values_only=True):
        code = str(row[0]).strip() if row[0] else ''
        if not code or code.lower() == 'nan':
            continue
        try:
            raw = row[3] if len(row) > 3 else None
            if raw in (None, '', 0):
                continue
            pkg = max(1, int(float(raw)))
        except (TypeError, ValueError):
            continue
        if code not in bom:
            bom[code] = {
                'name': str(row[1]).strip() if len(row) > 1 and row[1] else '',
                'unit': str(row[4]).strip() if len(row) > 4 and row[4] else 'шт',
                'supplier': str(row[2]).strip() if len(row) > 2 and row[2] else '',
                'model_qty': {}, 'configs': set(), 'section': 'Упаковка', 'package': pkg,
            }
        else:
            bom[code]['package'] = pkg
        loaded += 1
    wb.close()
    print(f"  Упаковок из {os.path.basename(PKG_FILE)} ({sheet}): {loaded}")
    return loaded

def _apply_bumper_package_defaults():
    """Упаковка бамперов, если не задана в локальном файле: ПЕРЕДНИЕ (2803…) —
    10 шт/тара, ЗАДНИЕ (2804…) — 8 шт/тара (согласовано по эталону MSA)."""
    n = 0
    for code in list(bom.keys()):
        info = bom[code]
        is_bumper = code in bumper_meta or code.startswith('2803') or code.startswith('2804')
        if is_bumper and info.get('package', 1) == 1:
            info['package'] = 10 if code.startswith('2803') else DEFAULT_BUMPER_PKG
            n += 1
    if n:
        print(f"  Бамперы без упаковки в локале → default (перед 10 / зад {DEFAULT_BUMPER_PKG}): {n}")

# Ручные остатки: Ввод_Остатков col E (остаток на нач.мес.) + даты в колонках F+

live_opening_stock = {}  # code -> qty (col E из LIVE — остаток на 1-е число месяца)
manual_stock_by_date = {}  # code -> {date: qty} (снимки остатка по дням для График_Поставок)
manual_stock_date_override = {}  # code -> date (legacy col F — дата актуальности для спроса)
manual_pkg_notes = {}  # code -> примечание из листа Упаковка col F (Пена / Подголовник)
live_sheet_packages = {}  # code -> размер упаковки из листа Упаковка col C (ручные правки LIVE)
live_sheet_pkg_auto = {}  # code -> авто-значение col G листа Упаковка (для детекта ручных правок)
manual_deliveries     = {}  # не используется, сохраняется для совместимости simulate_deliveries
stock_baseline = dict(stock)
stock_date_baseline = dict(stock_date_map)
_calc_last_date = datetime.date(MONTH_YEAR[MONTHS[-1][0]], MONTHS[-1][0], MONTHS[-1][2])

if os.path.exists(OUT):
    try:
        wb_live = load_workbook(OUT, read_only=True, data_only=True)
        # A. Лист 'Упаковка' (LIVE): col C — РУЧНЫЕ размеры упаковки (сохраняются
        #    между запусками, перекрывают «Упаковка локала»); col F — примечания (Lear)
        pkg_notes_loaded = 0
        if 'Упаковка' in wb_live.sheetnames:
            for row in wb_live['Упаковка'].iter_rows(min_row=3, values_only=True):
                code = str(row[0]).strip() if row[0] else ''
                if not code or code.lower() == 'nan':
                    continue
                # col C = размер упаковки (значение 1 считаем «не задано»);
                # col G = авто-значение прошлого запуска (для детекта ручной правки)
                pkg_v = row[2] if len(row) > 2 else None
                try:
                    p = float(pkg_v) if pkg_v not in (None, '') else 0
                    if p > 1:
                        live_sheet_packages[code] = int(p)
                except (TypeError, ValueError):
                    pass
                auto_v = row[6] if len(row) > 6 else None
                try:
                    pa = float(auto_v) if auto_v not in (None, '') else 0
                    if pa >= 1:
                        live_sheet_pkg_auto[code] = int(pa)
                except (TypeError, ValueError):
                    pass
                note_v = row[5] if len(row) > 5 else None
                if note_v not in (None, ''):
                    note_s = str(note_v).strip()
                    if note_s:
                        manual_pkg_notes[code] = note_s
                        if code in bom:
                            bom[code]['pkg_note'] = note_s
                        pkg_notes_loaded += 1
            if pkg_notes_loaded:
                print(f"  Примечаний упаковки (LIVE col F): {pkg_notes_loaded}")
            if live_sheet_packages:
                print(f"  Размеров упаковок из листа Упаковка (LIVE col C): {len(live_sheet_packages)}")

        # B. Ручные остатки из Ввод_Остатков
        manual_count = 0
        manual_snap_count = 0
        manual_date_count = 0
        ignored_date_count = 0
        if 'Ввод_Остатков' in wb_live.sheetnames:
            ws_man_live = wb_live['Ввод_Остатков']
            hdr_row = next(ws_man_live.iter_rows(
                min_row=_MAN_STOCK_HDR_ROW, max_row=_MAN_STOCK_HDR_ROW, values_only=True), ())
            hdr_joined = ' '.join(str(x or '').lower() for x in hdr_row[:8])
            is_date_grid = (
                'дата остатка' not in hdr_joined
                and any(
                    _parse_stock_input_header_date(hdr_row[i])
                    for i in range(_MAN_STOCK_FIRST_DATE_COL - 1, len(hdr_row))
                    if i < len(hdr_row) and hdr_row[i] not in (None, '')
                )
            )
            date_col_map = {}
            if is_date_grid:
                for ci in range(_MAN_STOCK_FIRST_DATE_COL - 1, len(hdr_row)):
                    dt = _parse_stock_input_header_date(hdr_row[ci])
                    if dt:
                        date_col_map[ci] = dt
            data_start = _MAN_STOCK_DATA_START if is_date_grid else 3
            for row in ws_man_live.iter_rows(min_row=data_start, values_only=True):
                code = str(row[0]).strip() if row[0] else ''
                if not code or code == 'nan': continue
                val = row[4] if len(row) > 4 else None  # col E = остаток на нач.мес.
                if val is not None:
                    try:
                        q = float(val)
                        live_opening_stock[code] = q
                        baseline_q = stock_baseline.get(code, 0)
                        if abs(q - baseline_q) > 1e-6:
                            manual_count += 1
                    except: pass
                if is_date_grid:
                    for ci, dt in date_col_map.items():
                        if ci >= len(row) or row[ci] in (None, ''):
                            continue
                        try:
                            snap_q = float(row[ci])
                        except: 
                            continue
                        manual_stock_by_date.setdefault(code, {})[dt] = snap_q
                        manual_snap_count += 1
                else:
                    date_val = row[5] if len(row) > 5 else None  # legacy col F
                    manual_date = _parse_manual_stock_date(date_val)
                    if manual_date is not None:
                        baseline_date = _baseline_stock_date(code, stock_date_baseline)
                        if _same_stock_date(manual_date, baseline_date):
                            continue
                        if manual_date > _calc_last_date:
                            ignored_date_count += 1
                            continue
                        manual_stock_date_override[code] = manual_date
                        manual_date_count += 1
        print(f"  Ручных остатков (нач.мес.) из Ввод_Остатков: {manual_count}")
        print(f"  Ручных снимков остатков по дням: {manual_snap_count}")
        if manual_date_count:
            print(f"  Ручных дат актуальности (legacy F): {manual_date_count}")
        if ignored_date_count:
            print(f"  ℹ️  Игнорировано дат остатков позже расчётного горизонта: {ignored_date_count}")
        wb_live.close()
    except Exception as e:
        print(f"  ⚠️  Ошибка чтения LIVE файла: {e}")

# Файлы остатков → колонки Ввод_Остатков (не перезаписываем ручной ввод из LIVE)
_file_snap_applied = 0
for code, dt_map in file_stock_by_date.items():
    for dt, qty in dt_map.items():
        dst = manual_stock_by_date.setdefault(code, {})
        if dt not in dst:
            dst[dt] = qty
            _file_snap_applied += 1
if _file_snap_applied:
    print(f"  Снимков из файлов (дата пересчёта) → Ввод_Остатков: {_file_snap_applied}")

def _first_calc_month_open_date():
    """1-е число первого расчётного месяца (остаток col E / Graph G)."""
    return datetime.date(MONTH_YEAR[STOCK_AS_OF_MONTH], STOCK_AS_OF_MONTH, 1)

def _finalize_live_opening_stock():
    """Col E не должен дублировать жёлтый снимок на дату пересчёта (если это не 1-е)."""
    open_dt = _first_calc_month_open_date()
    dropped = 0
    for code in list(live_opening_stock.keys()):
        sdate = get_stock_date(code)
        if sdate == open_dt:
            continue
        recalc_q = manual_stock_by_date.get(code, {}).get(sdate)
        if recalc_q is not None and abs(live_opening_stock[code] - recalc_q) < 1e-6:
            del live_opening_stock[code]
            dropped += 1
        elif recalc_q is None and abs(live_opening_stock[code] - stock_baseline.get(code, 0)) < 1e-6:
            del live_opening_stock[code]
            dropped += 1
    if dropped:
        print(f"  Col E: сброшено {dropped} знач. (= снимок на дату пересчёта, не нач.мес.)", flush=True)

_finalize_live_opening_stock()

def _apply_stock_snapshots_to_opening_balance():
    """Остаток на дату пересчёта (get_stock_date) → stock[] для заказов и риска.
    Col E (нач.мес.) сюда не подмешивается — см. _opening_stock()."""
    updated = 0
    codes = set(stock.keys()) | set(manual_stock_by_date.keys())
    for code in codes:
        sdate = get_stock_date(code)
        snap = manual_stock_by_date.get(code, {}).get(sdate)
        if snap is not None:
            stock[code] = snap
            updated += 1
    return updated

_snap_stock_sync = _apply_stock_snapshots_to_opening_balance()
if _snap_stock_sync:
    print(f"  Остаток на дату пересчёта (снимок/ручной) → расчёт: {_snap_stock_sync} кодов")

def _opening_stock(code):
    """Остаток на 1-е число первого расчётного месяца (Ввод_Остатков E, График G).
    Снимки по другим датам (жёлтые колонки) сюда не подтягиваются."""
    open_dt = _first_calc_month_open_date()
    if code in live_opening_stock:
        return live_opening_stock[code]
    snaps = manual_stock_by_date.get(code, {})
    if open_dt in snaps:
        return snaps[open_dt]
    by = file_stock_by_date.get(code, {})
    if open_dt in by:
        return by[open_dt]
    sdate = get_stock_date(code)
    if sdate == open_dt:
        return stock_baseline.get(code, 0)
    return 0.0

def _opening_stock_for_ss(code):
    """G / Ss[0]: не ниже potr первого дня спроса (Del[-1] в формуле нет)."""
    o = float(_opening_stock(code))
    for d in _delivery_series_demand(code):
        if d > 0:
            return max(o, float(d))
    return o

for code, stock_dt in manual_stock_date_override.items():
    stock_date_map[code] = stock_dt  # ручная дата меняет горизонт спроса/поставок

# ── Упаковки: «Упаковка локала» → PACKAGE_OVERRIDES → default бампер ──
_load_packages_from_local_file()
for code in bom:
    if 'package' not in bom[code]:
        bom[code]['package'] = 1

pkg_override_count = 0
for code, pkg_val in PACKAGE_OVERRIDES.items():
    if code in bom:
        bom[code]['package'] = pkg_val
        pkg_override_count += 1
    else:
        bom[code] = {'name': code, 'unit': 'pcs', 'supplier': '',
                     'model_qty': {}, 'configs': {}, 'section': 'Упаковка',
                     'package': pkg_val, 'notes': ''}
        pkg_override_count += 1
if pkg_override_count:
    print(f"  Переопределено упаковок (PACKAGE_OVERRIDES): {pkg_override_count}")

# ── Ручные исключения из расчёта (лист Исключить_применяемость в MRP_System_v9.xlsx) ──
BOM_NO_DEMAND_CODES, _no_demand_sheet = _load_no_demand_codes_from_excel()
if _no_demand_sheet:
    print(f"  Исключить_применяемость: {len(BOM_NO_DEMAND_CODES)} кодов без потребности")
elif BOM_NO_DEMAND_CODES:
    print(f"  Без потребности (встроенный список): {sorted(BOM_NO_DEMAND_CODES)}")

# ── Применяем переопределения применяемости BOM (BOM_APPLICABILITY_OVERRIDES) ──
# В LIVE MODE не трогаем — приоритет ручным ✓ из BOM_Детальный
bom_appl_count = 0
if not BOM_LIVE_ACTIVE:
    for code, allowed_cfgs in BOM_APPLICABILITY_OVERRIDES.items():
        if code in bom:
            bom[code]['configs'] = {k: 1 for k in allowed_cfgs}
            bom_appl_count += 1
            print(f"  BOM применяемость {code}: {sorted(allowed_cfgs)}")
else:
    print("  BOM применяемость: LIVE — только ✓ из BOM_Детальный (overrides пропущены)")
print(f"  Переопределено применяемостей BOM: {bom_appl_count}")

# B02_2WD_elite: только если не LIVE (в LIVE ✓ задаёт пользователь)
_b02_elite_added = 0
if not BOM_LIVE_ACTIVE:
    for _code, _info in bom.items():
        if _is_no_demand_code(_code):
            continue
        if _is_prerestyle_bumper_code(_code):
            continue
        _before = 'B02_2WD_elite' in _info.get('configs', {})
        _info['configs'] = apply_b02_2wd_elite_marks(_info.get('configs', {}), _code)
        if not _before and 'B02_2WD_elite' in _info.get('configs', {}):
            _b02_elite_added += 1
    for code, allowed_cfgs in BOM_APPLICABILITY_OVERRIDES.items():
        if code in bom:
            bom[code]['configs'] = {k: 1 for k in allowed_cfgs}
    if _b02_elite_added:
        print(f"  B02_2WD_elite: добавлена применяемость для {_b02_elite_added} деталей")

# Принудительно обнуляем применяемость исключённых кодов (после всех дополнений)
_no_demand_cleared = _apply_no_demand_exclusions(bom)
if _no_demand_cleared:
    print(f"  Исключено из потребности (применяемость очищена): {_no_demand_cleared} кодов")

_eco_n_codes, _eco_n_rules = load_ecoalliance_use(ECOALLIANCE_USE_FILE)
if ECOALLIANCE_USE_FILE:
    if _eco_n_codes:
        print(f"  EcoAlliance use ({os.path.basename(ECOALLIANCE_USE_FILE)}): "
              f"{_eco_n_codes} кодов, {_eco_n_rules} правил колонок")
        if eco_superseded_codes:
            print(f"    Замена 2.0: {', '.join(sorted(eco_superseded_codes))} → 1205279XGW01A")
    else:
        print(f"  ⚠️  EcoAlliance use: не распознан {os.path.basename(ECOALLIANCE_USE_FILE)}")

_eco_appl, _eco_cleared = 0, 0
if not BOM_LIVE_ACTIVE:
    _eco_appl, _eco_cleared = apply_ecoalliance_applicability(bom)
elif eco_applicability:
    print("  EcoAlliance: потребность из файла use.xlsx (✓ в LIVE BOM не перезаписываются)")

# LIVE: восстановить ✓ из файла перед расчётом и записью в Excel
if BOM_LIVE_ACTIVE:
    _n_live_rest = _finalize_live_bom_configs()
    if _n_live_rest:
        print(f"  LIVE BOM: восстановлено {_n_live_rest} кодов из снимка файла")

# Колёса B02 elite: применяемость B02_2WD_elite = как B02_4WD_elite (тот же привод elite,
# оба исполнения). Исключение XKN61 сделано для бамперов-рестайла и к колёсам (3101…) не
# относится. Работает в обоих режимах (LIVE и Master BOM).
_wheel_b02_elite = 0
for _wcode, _winfo in bom.items():
    if not str(_wcode).startswith('3101'):
        continue
    _wcf = _winfo.get('configs')
    if not isinstance(_wcf, dict):
        continue
    if 'B02_4WD_elite' in _wcf and 'B02_2WD_elite' not in _wcf:
        _wcf['B02_2WD_elite'] = _wcf['B02_4WD_elite']
        _wheel_b02_elite += 1
if _wheel_b02_elite:
    print(f"  Колёса B02 elite → добавлен B02_2WD_elite: {_wheel_b02_elite} код(ов)")

# Master BOM sync перенесён в конец скрипта (после сохранения MRP) — см. OUT_SAVED
_MASTER_BOM_SYNC_PATH = BOM_NEW if (BOM_NEW and os.path.exists(BOM_NEW)) else BOM_FILE


def paint_material_active(code, month_num):
    """Активен ли SKU в данном месяце (до/после замены кода поставщика)."""
    if code in REPLACEMENT_FROM_MONTH:
        _, from_m = REPLACEMENT_FROM_MONTH[code]
        if from_m is None:
            return True
        return month_num < from_m
    if code in REPLACEMENT_NEW_CODES:
        _, from_m = REPLACEMENT_NEW_CODES[code]
        if from_m is None:
            return False
        return month_num >= from_m
    return True


def _register_litum_sku(old, new, from_m, note=''):
    """Регистрирует новый код Litum (копия old), без отключения old."""
    if old in chem_norms:
        chem_norms[new] = dict(chem_norms[old])
    if old in paint_colors:
        paint_colors[new] = paint_colors[old]
    if old in part_tab_map and new not in part_tab_map:
        part_tab_map[new] = part_tab_map[old]
    if old in bom:
        bom[new] = dict(bom[old])
        bom[new]['supplier'] = PAINT_REPLACEMENT_SUPPLIER
        if note:
            prev = bom[new].get('notes', '') or ''
            bom[new]['notes'] = (prev + ' ' + note).strip()[:120]
    else:
        bom[new] = {
            'name': new, 'unit': 'кг', 'supplier': PAINT_REPLACEMENT_SUPPLIER,
            'configs': {}, 'model_qty': {}, 'section': 'Краски Litum',
            'package': 1, 'notes': note or f'Litum {new}',
        }
    stock.setdefault(new, 0.0)
    for sec, codes in bom_sections.items():
        if old in codes and new not in codes:
            codes.append(new)
            return
    bom_sections.setdefault('Краски Litum', []).append(new)


def _primer_split_for_litum(primer_code):
    for sp in PRIMER_SPLIT_TRANSITIONS:
        if sp['litum_primer'] == primer_code:
            return sp
    return None


def _primer_excluded_paints(yierte_primer, month_num):
    """Краски, чья доля грунта перешла на Litum — исключить из YIERTE с from_month."""
    out = []
    for sp in PRIMER_SPLIT_TRANSITIONS:
        if sp['yierte_primer'] == yierte_primer and month_num >= sp['from_month']:
            out.append(sp['paint_old'])
    return out


def apply_paint_code_replacements():
    """Полная замена красок + регистрация Litum-грунтов (частичный перенос)."""
    global color_filtered_paints
    for item in PAINT_CODE_REPLACEMENTS:
        old, new, from_m = item['old'], item['new'], item.get('from_month')
        if from_m is not None:
            REPLACEMENT_FROM_MONTH[old] = (new, from_m)
            REPLACEMENT_NEW_CODES[new] = (old, from_m)
        note = ''
        if from_m is not None:
            note = f"→ {new} с {_MONTH_EN_SHORT.get(from_m, from_m)} (Litum)"
        _register_litum_sku(old, new, from_m, note)
    for sp in PRIMER_SPLIT_TRANSITIONS:
        yp, lp, fm = sp['yierte_primer'], sp['litum_primer'], sp['from_month']
        po, pn = sp['paint_old'], sp['paint_new']
        LITUM_PRIMER_CODES[lp] = fm
        note = (f"грунт для {pn} с {_MONTH_EN_SHORT.get(fm, fm)}; "
                f"остальное на {yp} (YIERTE)")
        _register_litum_sku(yp, lp, fm, note)
        PRIMER_LINKED_PAINTS[lp] = [pn]
        paint_colors[lp] = _primer_color_keys(lp)
    color_filtered_paints = {
        k for k, v in paint_colors.items()
        if isinstance(v, list) and k != 'ALAA005669'
    }
    active = [f"{o}→{n} с {_MONTH_EN_SHORT.get(fm, '?')}"
              for o, (n, fm) in REPLACEMENT_FROM_MONTH.items() if fm]
    splits = [f"{sp['yierte_primer']}/{sp['litum_primer']} ({sp['paint_old']}) с "
              f"{_MONTH_EN_SHORT.get(sp['from_month'], '?')}"
              for sp in PRIMER_SPLIT_TRANSITIONS]
    planned = [f"{o}→{n}" for o, (n, fm) in REPLACEMENT_FROM_MONTH.items() if fm is None]
    if active:
        print(f"  Замена красок (Litum): {', '.join(active)}")
    if splits:
        print(f"  Частичный грунт (Litum): {', '.join(splits)}")
    if planned:
        print(f"  Замена красок (запланировано): {', '.join(planned)}")


apply_paint_code_replacements()

# ═══ STEP 4: Additional.xlsx ══════════════════════════════════
print("Loading Additional.xlsx...")
ADDL={}
try:
    df_add=pd.read_excel(ADD_FILE,sheet_name='Лист1',header=0)
    for _,row in df_add.iterrows():
        code=str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ''
        if not code: continue
        name=str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ''
        supp=str(row.iloc[2]).strip() if pd.notna(row.iloc[2]) else ''
        unit=str(row.iloc[3]).strip() if pd.notna(row.iloc[3]) else ''
        plan=str(row.get('Plan','')).strip() if pd.notna(row.get('Plan','')) else ''
        model=str(row.get('Model','')).strip() if pd.notna(row.get('Model','')) else ''
        cfg=str(row.get('Configuration','')).strip() if pd.notna(row.get('Configuration','')) else ''
        models=[m for m in ['A01','A08','B02','B04','B06','B16'] if m in model.upper()]
        if not models: models=['A01','A08','B02','B04','B06','B16']
        cfgs=set()
        cl=cfg.lower()
        if 'all' in cl or not cfg: cfgs={'comfort','elite','premium','TechPlus'}
        else:
            if 'comfort' in cl: cfgs.add('comfort')
            if 'elite' in cl: cfgs.add('elite')
            if 'premium' in cl: cfgs.add('premium')
            if 'tech' in cl: cfgs.add('TechPlus')
        ADDL[code]={'name':name,'supplier':supp,'unit':unit,'plan_tab':plan,'models':models,'configs':cfgs}
        if plan: part_tab_map[code]=plan
        if code not in bom:
            bom[code]={'name':name,'unit':unit,'supplier':supp,'model_qty':{},'configs':set(),'section':'Additional','package':1}
    print(f"  {len(ADDL)} additional parts")
except Exception as e:
    print(f"  Additional error: {e}")

# Build bumper_meta from bumper_clr_map + BOM model_qty
bumper_meta={}
for code,color_key in bumper_clr_map.items():
    mq=bom.get(code,{}).get('model_qty',{})
    models=list(mq.keys()) if mq else ['A01','A08','B02','B04','B06','B16']
    if code in ADDL: models=ADDL[code]['models']
    bumper_meta[code]={'color_key':color_key,'models':models}

_apply_bumper_package_defaults()

# ── NDKT: упаковки (3101100XST33A = бокс 30 шт; прочие колёса без кратности) ──
_ndkt_pkg_fixed = 0
for _c, _p in NDKT_PACKAGE_FIX.items():
    if _c in bom and int(bom[_c].get('package', 1) or 1) != int(_p):
        bom[_c]['package'] = int(_p)
        _ndkt_pkg_fixed += 1
# у остальных колёс NDKT базовая упаковка (стопка) — 24 шт; значения 1/384/540
# (старый формат «фура как упаковка») заменяем на 24
for _c in list(bom.keys()):
    if (normalize_supplier(bom[_c].get('supplier', '')) == NDKT_SUPPLIER
            and _c not in NDKT_PACKAGE_FIX
            and int(bom[_c].get('package', 1) or 1) in (1, 384, 540)):
        bom[_c]['package'] = NDKT_WHEEL_PKG_DEFAULT
        _ndkt_pkg_fixed += 1
if _ndkt_pkg_fixed:
    print(f"  NDKT: упаковки приведены к схеме (XST33A=30, прочие колёса=24): {_ndkt_pkg_fixed} кодов")

# ── Lear: вместимость тары и страх.запасы из файла «страх.запасы ЛИР.xlsx» ──
_load_lear_safety_file()
_lear_pkg_applied = 0
for _c, _p in LEAR_PKG_BY_CODE.items():
    if _c in bom and int(_p) >= 1 and int(bom[_c].get('package', 1) or 1) != int(_p):
        bom[_c]['package'] = int(_p)
        _lear_pkg_applied += 1
if _lear_pkg_applied:
    print(f"  Lear: вместимость тары из файла страх.запасов: {_lear_pkg_applied} кодов")

# ── Авто-база упаковок (для листа Упаковка col G — детект ручных правок) ──
pkg_base_by_code = {c: int(bom[c].get('package', 1) or 1) for c in bom}

# ── Ручные размеры упаковок из LIVE листа «Упаковка» (col C) — высший приоритет.
#    Правка считается ручной, если col C ≠ col G (авто-значение прошлого запуска);
#    для старых файлов без col G — если col C ≠ текущей авто-базе. ──
_live_pkg_applied = []
for _c, _p in live_sheet_packages.items():
    if _c not in bom or _p < 1:
        continue
    _auto_prev = live_sheet_pkg_auto.get(_c)
    _base_now = pkg_base_by_code.get(_c, 1)
    _is_manual = (_p != _auto_prev) if _auto_prev is not None else (_p != _base_now)
    if _is_manual and _p != int(bom[_c].get('package', 1) or 1):
        bom[_c]['package'] = int(_p)
        _live_pkg_applied.append(_c)
if _live_pkg_applied:
    print(f"  Ручные упаковки из листа Упаковка (LIVE): {len(_live_pkg_applied)} кодов "
          f"(напр. {', '.join(_live_pkg_applied[:5])})")

# ── NDKT: правило упаковок: 3101100XST33A — ВСЕГДА 30 (бокс);
#    остальные колёса — 24 (стопка) по умолчанию; ручные правки листа
#    «Упаковка» (col C ≠ col G) для остальных колёс сохраняются. ──
_ndkt_pkg_forced = 0
for _c in list(bom.keys()):
    if normalize_supplier(bom[_c].get('supplier', '')) != NDKT_SUPPLIER:
        continue
    _want = int(NDKT_PACKAGE_FIX.get(_c, NDKT_WHEEL_PKG_DEFAULT))
    pkg_base_by_code[_c] = _want
    if _c in NDKT_PACKAGE_FIX:
        if int(bom[_c].get('package', 1) or 1) != _want:
            bom[_c]['package'] = _want
            _ndkt_pkg_forced += 1
        continue
    if _c in _live_pkg_applied:
        continue   # ручная правка пользователя в листе Упаковка — не трогаем
    if int(bom[_c].get('package', 1) or 1) in (1, 384, 540):
        bom[_c]['package'] = _want
        _ndkt_pkg_forced += 1
if _ndkt_pkg_forced:
    print(f"  NDKT: упаковки приведены к правилу (XST33A=30, прочие колёса=24): {_ndkt_pkg_forced} кодов")

# ── Эфтек: единица измерения этих материалов — КИЛОГРАММЫ (согласовано);
#    упаковка = бочка (250 / 218 / 246 кг), остатки warehouse 16 — в кг ──
for _c_kg in EFTEC_KG_CODES:
    if _c_kg in bom and str(bom[_c_kg].get('unit', '')).lower() not in ('kg', 'кг'):
        bom[_c_kg]['unit'] = 'kg'

# For parts NOT in v7 tab_map, assign AS_in_F_A (most common)
for code in bom:
    if code not in part_tab_map:
        part_tab_map[code]='AS_in_F_A'

all_codes=sorted(set(list(bom.keys())+list(chem_norms.keys())+
                     list(bumper_meta.keys())+list(ADDL.keys())))
# Исключаем устаревшие/дублированные коды
all_codes = [c for c in all_codes if c not in OBSOLETE_CODES]
for c in list(OBSOLETE_CODES):
    bom.pop(c, None); part_tab_map.pop(c, None); bumper_meta.pop(c, None)

# Доп. коды из Additional — повторная фильтрация остатков
stock, stock_src, stock_date_map, file_stock_by_date = _filter_stock_to_calc(
    stock, stock_src, stock_date_map, all_codes, file_stock_by_date)

# Бамперы — только вкладка AS_in_F_A (A01/B02/B04 собираются на F_A; H_B пуст).
# Применяемость бамперов берётся НАПРЯМУЮ из BOM (BOM_Детальный), без хардкода:
#   2803120/2804104 → A01 comfort/elite/premium; 2803130/2804105 → A01 Tech Plus;
#   2803104/2804KN260004 → B02 elite; 2803105/2804KN260005 → B02/B04 premium+TechPlus.
# Сопоставление с планом — по (модель, комплектация), привод не важен (см. _bumper_cfg_match).
_bumper_tab_set = 0
for _bcode in bumper_clr_map:
    CODE_TAB_OVERRIDES[_bcode] = 'AS_in_F_A'
    _bumper_tab_set += 1
if _bumper_tab_set:
    print(f"  Бамперы: вкладка AS_in_F_A, применяемость из BOM → {_bumper_tab_set} кодов")

# ── ПЕРЕОПРЕДЕЛЕНИЕ ВКЛАДОК ПЛАНА ПО ПОСТАВЩИКАМ ──
# Главное правило: tab определяется по поставщику из карты SUPPLIER_TAB_MAP.
# Если поставщик не в карте, оставляем V7-маппинг part_tab_map.
# Точечные CODE_TAB_OVERRIDES имеют наивысший приоритет.
overrides_supp = 0
overrides_code = 0
for code in all_codes:
    if code in CODE_TAB_OVERRIDES:
        part_tab_map[code] = CODE_TAB_OVERRIDES[code]
        overrides_code += 1
        continue
    supp = bom.get(code, {}).get('supplier', '') or ADDL.get(code, {}).get('supplier', '')
    new_tab = supplier_tab(supp)
    if new_tab and part_tab_map.get(code) != new_tab:
        part_tab_map[code] = new_tab
        overrides_supp += 1

# Доп. правила по эвристике для B06 колёс (NDKT) — на AS_in_H_B
TAB_OVERRIDES_PREFIX = [
    ('XKN4A', 'AS_in_H_B'),  # 3101100XKN4AA, 3101101XKN4AA — B06 wheels
]
overrides_pref = 0
for code in all_codes:
    cur_tab = part_tab_map.get(code)
    for sub, tab in TAB_OVERRIDES_PREFIX:
        if sub in code and cur_tab != tab:
            if code not in {'3101100XKN46A','3101102XKN46A','3101105XKN46A','3101101XKN61A'}:
                part_tab_map[code] = tab
                overrides_pref += 1
                break

# NDKT: не дублировать потребность на F_A + H_B — оставляем H_B (B06)
for code in all_codes:
    if code in CODE_TAB_OVERRIDES:
        continue
    supp = normalize_supplier(bom.get(code, {}).get('supplier', ''))
    if supp != NDKT_SUPPLIER:
        continue
    if len(resolve_plan_tabs(part_tab_map.get(code, ''))) > 1:
        part_tab_map[code] = 'AS_in_H_B'
        overrides_pref += 1

print(f"\nTotal parts in MRP: {len(all_codes)} (исключено устаревших: {len(OBSOLETE_CODES)})")
print(f"  Переопределено вкладок по поставщикам: {overrides_supp}")
print(f"  Переопределено вкладок по префиксам:   {overrides_pref}")
print(f"  Переопределено вкладок по точечным кодам: {overrides_code}")

def _configs_nonempty(code):
    """Есть ли хотя бы одна отметка применяемости (✓) в BOM_Детальный."""
    if code not in bom:
        return False
    cfgs = bom[code].get('configs', {})
    if isinstance(cfgs, set):
        return len(cfgs) > 0
    return bool(cfgs)

def has_mrp_applicability(code):
    """Деталь участвует в потребности, заказах, поставках и риске дефицита."""
    if code in BOM_NO_DEMAND_CODES:
        return False  # ручное/встроенное исключение — без расчёта потребности
    if code in REFERENCE_ONLY_CODES:
        return False  # справочные коды — без расчёта потребности/заказов/графика
    if code in chem_norms or code in color_filtered_paints:
        return True
    if code in ADDL:
        return True
    if code in B16_ELITE:
        return True
    if code in bom:
        return _configs_nonempty(code)
    if code in bumper_meta:
        return True
    return False

# ── Применяемость (модель + конфигурация) для колонки B потребностей/заказов/графиков ──
_CFG_ORDER = {'comfort': 0, 'elite': 1, 'premium': 2, 'TechPlus': 3}
def _applicability_str(code):
    """Строка применяемости вида 'B02: elite, premium, TechPlus; B04: premium, TechPlus'.
    Берётся из BOM_Детальный (configs); привод (2WD/4WD) опускается — только модель+конфиг."""
    cfgs = bom.get(code, {}).get('configs', {})
    keys = list(cfgs.keys()) if isinstance(cfgs, dict) else list(cfgs or [])
    if not keys and code in ADDL:
        _a = ADDL[code]
        _m = _a.get('models', []) or []
        _c = sorted(_a.get('configs', set()) or [])
        if _m and _c:
            return '; '.join(f"{m}: {', '.join(_c)}" for m in _m)
        return ', '.join(_m) if _m else ''
    if not keys:
        return ''
    by_model = {}
    for k in keys:
        parts = str(k).split('_')
        if len(parts) >= 3:
            model, cfg = parts[0], parts[2]
        elif len(parts) == 2:
            model, cfg = parts[0], parts[1]
        else:
            model, cfg = k, ''
        by_model.setdefault(model, set()).add(cfg)
    out = []
    for m in sorted(by_model):
        cs = sorted((c for c in by_model[m] if c), key=lambda c: _CFG_ORDER.get(c, 9))
        out.append(f"{m}: {', '.join(cs)}" if cs else m)
    return '; '.join(out)

_COLOR_DISPLAY = {
    'GOLDEN_BLACK': 'Golden Black', 'WHITE_C1': 'White C1', 'C3_GREY': 'C3 Grey',
    'ATLANTIS': 'Atlantis', 'BLUE_5B': 'Blue 5B', 'GN_RED': 'GN Red',
    'AYERS_GREY': 'Ayers Grey', 'KU_GREY': 'KU Grey', 'ORANGE': 'Orange',
    'FU_GREY': 'FU Grey', 'CRYSTAL_BLACK': 'Crystal Black', 'SWAROVSKI': 'Swarovski',
    '9E_WHITE': '9E White',
}
def _color_str(code):
    """Цвет детали: бамперы — из bumper_clr_map; краски — из paint_colors."""
    if code in bumper_meta:
        ck = bumper_meta[code].get('color_key', '') or ''
        return _COLOR_DISPLAY.get(ck, ck.replace('_', ' ')) if ck else ''
    pc = paint_colors.get(code)
    if isinstance(pc, list) and pc:
        return ', '.join(_COLOR_DISPLAY.get(c, c.replace('_', ' ')) for c in pc)
    return ''

def _name_with_appl(code, name):
    """Колонка B: 'Наименование | модель: конфигурация | цвет' (без сдвига колонок)."""
    appl = _applicability_str(code)
    color = _color_str(code)
    tail = []
    if appl:
        tail.append(appl)
    if color:
        tail.append(f"цвет: {color}")
    if tail:
        return f"{str(name or '')[:38]} | {' | '.join(tail)}"[:150]
    return str(name or '')[:50]

mrp_codes = [c for c in all_codes if has_mrp_applicability(c)]
_bom_no_appl = [c for c in all_codes if c in bom and not _configs_nonempty(c)]
print(f"  В расчётах (потребность/заказы/поставки/риск): {len(mrp_codes)} из {len(all_codes)}")
if _bom_no_appl:
    print(f"  Исключено без применяемости в BOM: {len(_bom_no_appl)}"
          + (f" (напр. {_bom_no_appl[0]})" if len(_bom_no_appl) == 1 else ""))

# ═══ STEP 5: Parse plan tabs ══════════════════════════════════
# Парсер полностью авто-определяет позиции колонок из заголовочной строки.
# Не зависит от фиксированных смещений — устойчив к сдвигам колонок при обновлении файла.
print("Parsing plan tabs...")

# Альтернативные ключевые слова месяцев (на случай изменения формата заголовка)
# Динамически строится из MONTHS — не зависит от конкретных месяцев
_MONTH_EN_FULL_MAP = {1:'January',2:'February',3:'March',4:'April',5:'May',6:'June',
                      7:'July',8:'August',9:'September',10:'October',11:'November',12:'December'}
_MONTH_RU_FULL_MAP = {1:'Январь',2:'Февраль',3:'Март',4:'Апрель',5:'Май',6:'Июнь',
                      7:'Июль',8:'Август',9:'Сентябрь',10:'Октябрь',11:'Ноябрь',12:'Декабрь'}
_MONTH_RU_LC_MAP  = {k: v.lower() for k, v in _MONTH_RU_FULL_MAP.items()}
_MONTH_ZH_MAP     = {1:'1月',2:'2月',3:'3月',4:'4月',5:'5月',6:'6月',
                     7:'7月',8:'8月',9:'9月',10:'10月',11:'11月',12:'12月'}
_MONTH_ALT = {}
for _mn, _mlabel, _ in MONTHS:
    _yr = int(_mlabel.split()[1])
    _en_short = _mlabel[:3]      # e.g. "May"
    _en_full  = _MONTH_EN_FULL_MAP[_mn]  # e.g. "May"
    _ru_full  = _MONTH_RU_FULL_MAP[_mn]
    _ru_lc    = _MONTH_RU_LC_MAP[_mn]
    _zh       = _MONTH_ZH_MAP[_mn]
    _MONTH_ALT[_mn] = [
        MONTH_KW[_mn],
        f"{_mn:02d} {_en_full}",
        f"{_en_full} {_yr}",
        f"{_mn:02d} {_en_full} {_yr}",
        _ru_full, _ru_lc, _zh,
    ]

def _find_month_row(df, mnum):
    """Находит строку с заголовком месяца, сканируя все колонки."""
    keywords = _MONTH_ALT.get(mnum, [MONTH_KW[mnum]])
    for i in range(df.shape[0]):
        row_str = ' '.join(str(v) for v in df.iloc[i] if v is not None and v == v)
        if any(kw in row_str for kw in keywords):
            return i
    return None

def _find_header_row(df, start_row):
    """Находит строку заголовка (содержит 'Batch'), начиная с start_row."""
    for i in range(start_row, min(start_row + 8, df.shape[0])):
        if any('Batch' in str(v) for v in df.iloc[i]):
            return i
    return None

def _detect_cols(df, hrow, base_cols):
    """
    Авто-определяет позиции колонок из строки заголовка.
    base_cols используется как fallback.
    Возвращает (bc, mc, dc, cc, d1).
    """
    h = df.iloc[hrow]
    # Находим 'Batch no.'
    bc_det = next((i for i, v in enumerate(h) if 'Batch' in str(v)), None)
    if bc_det is None:
        bc_det = base_cols['batch']
        offset = 0
    else:
        offset = bc_det - base_cols['batch']

    # Находим 'Model' колонку (или применяем смещение)
    mc_det = next((i for i, v in enumerate(h)
                   if str(v).strip().lower() in ('model','модель')), None)
    mc_use = mc_det if mc_det is not None else base_cols['model'] + offset

    # Находим 'configuration' колонку
    cc_det = next((i for i, v in enumerate(h)
                   if 'config' in str(v).lower()), None)
    cc_use = cc_det if cc_det is not None else base_cols['config'] + offset

    dc_det = next((i for i, v in enumerate(h)
                   if any(k in str(v).lower() for k in
                          ('drive', '驱动', 'драйв', '4x2', '4x4', '两驱', '四驱'))), None)
    dc_use = dc_det if dc_det is not None else base_cols['drive'] + offset
    if dc_use == cc_use:
        dc_use = base_cols['drive'] + offset

    # day1: ищем первый '1' после конфигурации (числовой заголовок дня)
    d1_det = None
    search_start = max(cc_use, mc_use) + 2
    for i in range(search_start, min(search_start + 10, len(h))):
        try:
            if float(str(h.iloc[i]).strip()) == 1.0:
                d1_det = i
                break
        except: pass
    d1_use = d1_det if d1_det is not None else base_cols['day1'] + offset

    return bc_det if bc_det is not None else base_cols['batch'], mc_use, dc_use, cc_use, d1_use

plan_batches={}; batch_info={}; batch_info_by_tab={}
for tab,cols in TAB_COLS.items():
    try:
        df=pd.read_excel(PF_FILE,sheet_name=tab,header=None)
    except PermissionError:
        print(f"  Cannot read {tab}: файл открыт в Excel. Закройте {os.path.basename(PF_FILE)} и запустите скрипт снова.")
        continue
    except Exception as e:
        print(f"  Cannot read {tab}: {e}"); continue
    plan_batches[tab]={}
    for mnum,mlabel,n_days in MONTHS:
        may_row = _find_month_row(df, mnum)
        if may_row is None: plan_batches[tab][mnum]={}; continue

        hrow = _find_header_row(df, may_row)
        if hrow is None: plan_batches[tab][mnum]={}; continue

        # Полностью авто-определяем колонки из заголовочной строки
        _bc,_mc,_dc,_cc,_d1 = _detect_cols(df, hrow, cols)

        # Конечная строка секции: начало следующего месяца
        end_row=df.shape[0]
        for mn2,_,_ in MONTHS:
            if mn2<=mnum: continue
            next_row = _find_month_row(df, mn2)
            if next_row is not None and next_row > hrow:
                end_row = min(end_row, next_row)

        month_data={}
        for i in range(hrow+1,end_row):
            row=df.iloc[i]
            if _bc>=len(row) or not pd.notna(row.iloc[_bc]): continue
            bv = _normalize_batch_id(row.iloc[_bc])
            if not bv.startswith('R'): continue
            try:
                mr=MODEL_MAP.get(str(row.iloc[_mc]).strip() if _mc<len(row) and pd.notna(row.iloc[_mc]) else '','')
                dr=DRIVE_MAP.get(str(row.iloc[_dc]).strip() if _dc<len(row) and pd.notna(row.iloc[_dc]) else '','')
                cr=_map_plan_config(row.iloc[_cc] if _cc<len(row) else '')
                tab_bi = batch_info_by_tab.setdefault(tab, {})
                if bv not in tab_bi:
                    tab_bi[bv] = {'model': mr, 'drive': dr, 'config': cr, 'total': 0}
                else:
                    _bi = tab_bi[bv]
                    if mr and not _bi.get('model'):
                        _bi['model'] = mr
                    if dr and not _bi.get('drive'):
                        _bi['drive'] = dr
                    if cr and not _bi.get('config'):
                        _bi['config'] = cr
                if bv not in batch_info:
                    batch_info[bv] = dict(tab_bi[bv])
            except:
                tab_bi = batch_info_by_tab.setdefault(tab, {})
                if bv not in tab_bi:
                    tab_bi[bv] = {'model': '', 'drive': '', 'config': '', 'total': 0}
                if bv not in batch_info:
                    batch_info[bv] = dict(tab_bi[bv])
            for d in range(1,n_days+1):
                col_idx=_d1+d-1
                if col_idx<len(row) and pd.notna(row.iloc[col_idx]):
                    try:
                        q=float(row.iloc[col_idx])
                        if q>0:
                            month_data.setdefault(bv,{})[d]=month_data.get(bv,{}).get(d,0)+q
                            batch_info_by_tab[tab][bv]['total'] += q
                            batch_info[bv]['total'] += q
                    except: pass
        plan_batches[tab][mnum]=month_data
        n_cars=sum(sum(d.values()) for d in month_data.values())
        print(f"  {tab}/{mnum}: {len(month_data)} batches, {n_cars:.0f} cars")

# ═══ STEP 5b: Paint Statistics (batch → color → cars) ════════
# Связывает партии RXXxxxx из плана с распределением по цветам кузовов.
# Используется для (а) бамперов: остатки по цветам, (б) красок: фильтр по цвету.
print("\nLoading paint statistics (batch → color)...")

# Маппинг внутренних цветовых ключей → колонки Order_calculation_statistics
PAINT_COL_MAP = {
    'GN_RED':        'GN红GN Red Met B2',
    'BLUE_5B':       '5B蓝5B BLUE Met B2',
    'ORANGE':        '2C橙\norange',
    'FU_GREY':       'FU灰\ngrey',
    'WHITE_C1':      '汉密尔顿白hamilton white solid .bc(C1)',
    'GOLDEN_BLACK':  '太阳金乌黑 Golden Blak Met BC',
    'CRYSTAL_BLACK': '炫晶黑Crystal Black B2 ',
    'AYERS_GREY':    '艾尔斯灰Ayers Grey met. BC',
    'ATLANTIS':      '亚特兰蒂斯蓝\nAtlantis blue',
    'C3_GREY':       'C3灰\nC3 grey',
    'KU_GREY':       'KU灰\ngrey',
    'SWAROVSKI':     '施华洛蓝Swarovski Blue met.BC',
    '9E_WHITE':      '9E白',
}

batch_color = {}   # {'RAR2281': {'GOLDEN_BLACK': 120, ...}}
if PAINT_STATS and _is_valid_xlsx(PAINT_STATS):
    try:
        wb_ps = load_workbook(PAINT_STATS, read_only=True, data_only=True)
        ws_ps = wb_ps['DAP All batches'] if 'DAP All batches' in wb_ps.sheetnames else wb_ps.worksheets[0]
        # Заголовок: строка 2 содержит названия колонок
        hdr = {}
        row2 = next(ws_ps.iter_rows(min_row=2, max_row=2, values_only=True))
        for ci, v in enumerate(row2, 1):
            if v:
                vs = str(v)
                hdr[vs] = ci
        # Reverse PAINT_COL_MAP: column_name → internal_key
        col_to_key = {}
        for k, colname in PAINT_COL_MAP.items():
            if colname in hdr:
                col_to_key[hdr[colname]] = k
        batch_col = hdr.get('批次号\nBatch number', 1)
        # Read data
        for row in ws_ps.iter_rows(min_row=3, values_only=True):
            batch = row[batch_col-1] if batch_col-1 < len(row) else None
            if not batch: continue
            batch = _normalize_batch_id(batch)
            cd = {}
            for ci, k in col_to_key.items():
                if ci-1 < len(row):
                    try:
                        q = float(row[ci-1]) if row[ci-1] not in (None, '') else 0
                    except: q = 0
                    if q > 0:
                        cd[k] = q
            if cd:
                batch_color[batch] = cd
        wb_ps.close()
        print(f"  ✅ Paint stats loaded: {len(batch_color)} batches with colour split")
    except Exception as e:
        print(f"  ⚠️  Paint stats не загружен: {e}")
elif PAINT_STATS:
    print(f"  ⚠️  Paint stats повреждён (не xlsx): {os.path.basename(PAINT_STATS)}")
else:
    print(f"  ⚠️  Файл paint statistics не найден")

# ── НАДЁЖНОСТЬ: если файл статистики не загрузился (повреждён/занят Excel/
#    OneDrive) или оказался пуст — строим цветовую карту из GWM-файла. ──
if not batch_color:
    _gwm_fb = _GWM_INPUT or _find_gwm_batch_file()
    if _gwm_fb and _try_regenerate_paint_stats(_gwm_fb, _PAINT_OUT):
        print("  ♻️  Paint stats пересобран из GWM, повторная загрузка...")
        PAINT_STATS = _PAINT_OUT
        try:
            wb_ps = load_workbook(PAINT_STATS, read_only=True, data_only=True)
            ws_ps = wb_ps['DAP All batches'] if 'DAP All batches' in wb_ps.sheetnames else wb_ps.worksheets[0]
            hdr = {}
            row2 = next(ws_ps.iter_rows(min_row=2, max_row=2, values_only=True))
            for ci, v in enumerate(row2, 1):
                if v:
                    hdr[str(v)] = ci
            col_to_key = {hdr[colname]: k for k, colname in PAINT_COL_MAP.items() if colname in hdr}
            batch_col = hdr.get('批次号\nBatch number', 1)
            for row in ws_ps.iter_rows(min_row=3, values_only=True):
                batch = row[batch_col-1] if batch_col-1 < len(row) else None
                if not batch:
                    continue
                batch = _normalize_batch_id(batch)
                cd = {}
                for ci, k in col_to_key.items():
                    if ci - 1 < len(row):
                        try:
                            q = float(row[ci - 1]) if row[ci - 1] not in (None, '') else 0
                        except Exception:
                            q = 0
                        if q > 0:
                            cd[k] = q
                if cd:
                    batch_color[batch] = cd
            wb_ps.close()
            print(f"  ✅ Paint stats loaded: {len(batch_color)} batches with colour split")
        except Exception as _e_reload:
            print(f"  ⚠️  Повторная загрузка paint stats не удалась: {_e_reload}")
    if not batch_color and _gwm_fb and _is_valid_xlsx(_gwm_fb):
        print("  ⚠️  batch_color пуст — восстанавливаю цвета напрямую из GWM-файла...")
        try:
            _bp_fb = _find_build_order_calc_module()
            if not _bp_fb:
                raise FileNotFoundError('build_order_calc_v2.py не найден')
            import importlib.util as _ilu_fb
            _sp_fb = _ilu_fb.spec_from_file_location('build_order_calc_v2_fb', _bp_fb)
            _bm_fb = _ilu_fb.module_from_spec(_sp_fb)
            _sp_fb.loader.exec_module(_bm_fb)
            _raw_fb = _bm_fb.build_batch_color_map(_gwm_fb)
            _col2key_fb = {v: k for k, v in PAINT_COL_MAP.items()}
            for _b_fb, _cd_fb in _raw_fb.items():
                _norm_fb = {_col2key_fb[_c]: _q for _c, _q in _cd_fb.items() if _c in _col2key_fb}
                if _norm_fb:
                    batch_color[_normalize_batch_id(_b_fb)] = _norm_fb
            print(f"  ♻️  Цвета восстановлены из GWM напрямую: {len(batch_color)} партий")
        except Exception as _e_fb:
            print(f"  ❌  Не удалось восстановить цвета из GWM: {_e_fb}")
            print("      ВНИМАНИЕ: потребность по бамперам/цветным краскам будет НЕВЕРНОЙ (0).")
    elif not batch_color:
        print("  ❌  GWM-файл 各车型成套批次统计表 не найден — цвета партий недоступны.")

# ── ДИАГНОСТИКА: партии плана без цветовых данных. Такие партии молча выпадают
#    из потребности по бамперам и цветным краскам (как сообщалось по RAS2212).
#    Печатаем список, чтобы ошибки были видны для ВСЕХ расчётов. ──
def _audit_plan_batches_without_color():
    _seen = set(); _miss = []
    for _tab in plan_batches:
        for _mn in plan_batches[_tab]:
            for _bv, _dq in plan_batches[_tab][_mn].items():
                if _bv in _seen:
                    continue
                _seen.add(_bv)
                _cars = sum(_dq.values())
                if _cars <= 0:
                    continue
                if not _batch_color_lookup(_bv):
                    _bi = _get_batch_info(_tab, _bv)
                    _miss.append((_bv, _bi.get('model', ''), _bi.get('config', ''), _cars))
    if _miss:
        print(f"  ⚠️  Партии плана БЕЗ цветовых данных (выпадают из бамперов/красок): {len(_miss)}")
        for _x in sorted(_miss, key=lambda t: -t[3])[:30]:
            print(f"       {_x[0]:14s} {_x[1]:4s} {_x[2]:10s} cars={_x[3]:.0f}")
        print("      → проверьте, что эти партии есть в GWM-файле 各车型成套批次统计表.")
    else:
        print("  ✅ Все партии плана имеют цветовые данные (ничего не выпадает).")
# (вызов _audit_plan_batches_without_color() — ниже, после _batch_color_lookup)

# ═══ STEP 6: Demand (V7 logic) ════════════════════════════════
def _unique_batches_for_tabs(month_num, tabs):
    """Уникальные (вкладка, партия) с выбранных листов — метаданные не смешиваются."""
    seen = set()
    out = []
    for tab in tabs:
        if tab not in plan_batches or month_num not in plan_batches[tab]:
            continue
        for bv, day_qty in plan_batches[tab][month_num].items():
            key = (tab, bv)
            if key in seen:
                continue
            seen.add(key)
            out.append((tab, bv, day_qty))
    return out

def cars_by_color(month_num, color_key, models=None, configs=None, applicable=None, tabs=None):
    """Возвращает {day: cars} — кузова данного цвета в данном месяце.
    Метаданные и цвета партии берутся с той же вкладки плана.
    tabs — вкладки Plan-Fact из part_tab_map; если None, берутся все вкладки.
    """
    out = {}
    seen = set()
    tab_list = tabs if tabs is not None else list(plan_batches.keys())
    for tab in tab_list:
        if tab not in plan_batches or month_num not in plan_batches[tab]:
            continue
        months_data = plan_batches[tab]
        for bv, day_qty in months_data[month_num].items():
            key = (tab, bv)
            if key in seen:
                continue
            seen.add(key)
            bi = _get_batch_info(tab, bv)
            b_model = bi.get('model','')
            b_config = bi.get('config','')
            if models and b_model not in models: continue
            if configs and b_config not in configs: continue
            if applicable:
                cfg_key = f"{b_model}_{bi.get('drive','')}_{b_config}"
                matched, _ = cfg_match(cfg_key, applicable)
                if not matched:
                    continue
            bc_split = _batch_color_lookup(bv)
            if not bc_split: continue
            total_in_batch = sum(bc_split.values())
            if total_in_batch <= 0: continue
            color_cars = bc_split.get(color_key, 0)
            if color_cars <= 0: continue
            share = color_cars / total_in_batch
            for day, cars in day_qty.items():
                out[day] = out.get(day, 0) + cars * share
    return out


def _primer_demand_for_paints(primer_code, month_num, paint_codes, tabs=None):
    """Σ(кузова цветов paint_codes × BOM краски) × норма грунта primer_code."""
    norms = chem_norms.get(primer_code, {})
    if not norms or not paint_codes:
        return {}
    daily = {}
    for paint_code in paint_codes:
        paint_cfgs = bom.get(paint_code, {}).get('configs', {})
        if isinstance(paint_cfgs, set):
            paint_cfgs = {k: 1 for k in paint_cfgs}
        applicable = paint_cfgs if paint_cfgs else None
        for color_key in _linked_paint_color_keys(paint_code):
            for model_key, norm in norms.items():
                if norm <= 0:
                    continue
                cars_d = cars_by_color(
                    month_num, color_key, models={model_key}, applicable=applicable, tabs=tabs)
                for day, cars in cars_d.items():
                    daily[day] = daily.get(day, 0) + cars * norm
    return daily


def get_primer_demand(primer_code, month_num, tabs=None):
    """YIERTE-грунт: все связанные краски минус перешедшие на Litum. Litum-грунт: одна краска."""
    sp_litum = _primer_split_for_litum(primer_code)
    if sp_litum:
        if month_num < sp_litum['from_month']:
            return {}
        return _primer_demand_for_paints(
            primer_code, month_num, [sp_litum['paint_new']], tabs=tabs)
    if primer_code not in PRIMER_LINKED_PAINTS:
        return {}
    linked = [p for p in PRIMER_LINKED_PAINTS[primer_code]
              if p not in _primer_excluded_paints(primer_code, month_num)]
    return _primer_demand_for_paints(primer_code, month_num, linked, tabs=tabs)

def cars_filter(month_num, models=None, configs=None, predicate=None, tabs=None):
    """Возвращает {day: cars} с произвольным фильтром по партиям."""
    out = {}
    tab_list = tabs if tabs is not None else list(plan_batches.keys())
    for tab in tab_list:
        if tab not in plan_batches or month_num not in plan_batches[tab]:
            continue
        months_data = plan_batches[tab]
        for bv, day_qty in months_data[month_num].items():
            bi = _get_batch_info(tab, bv)
            if models and bi.get('model','') not in models: continue
            if configs and bi.get('config','') not in configs: continue
            if predicate and not predicate(bv, bi): continue
            for day, cars in day_qty.items():
                out[day] = out.get(day, 0) + cars
    return out

def _get_batch_info(tab, bv):
    """Метаданные партии только с указанной вкладки (без fallback на другие листы)."""
    return batch_info_by_tab.get(tab, {}).get(_normalize_batch_id(bv), {})

def _batch_id_prefix(bv):
    m = re.match(r'^([A-Z]+)', _normalize_batch_id(bv))
    return m.group(1) if m else ''


def _batch_color_lookup(bv):
    """Цвета партии из paint stats. RAW ≠ RAR (разные префиксы не смешиваются)."""
    bv = _normalize_batch_id(bv)
    if bv in batch_color:
        return batch_color[bv]
    tail = re.sub(r'^[A-Z]+', '', bv)
    if not tail:
        return {}
    bp = _batch_id_prefix(bv)
    matches = [k for k in batch_color
               if re.sub(r'^[A-Z]+', '', k) == tail and _batch_id_prefix(k) == bp]
    if len(matches) == 1:
        return batch_color[matches[0]]
    return {}

def _batch_dominant_color(bv):
    """Преобладающий цвет кузова партии (из paint stats). None если данных нет."""
    bc = _batch_color_lookup(bv)
    if not bc:
        return None
    return max(bc.items(), key=lambda kv: kv[1])[0]

def _is_prerestyle_bumper(code):
    return code in bumper_meta and _is_prerestyle_bumper_code(code)


def _prerestyle_bumper_applicable():
    return {PRERESTYLE_PREMIUM_CFG: 1}


def _bumper_batch_cfg_key(bi, code=''):
    """Ключ конфигурации партии.
    Для дорестайл B02 premium принимаем 2WD ТОЛЬКО если привод в плане не указан
    (пустой). Если привод задан (4x4 → 4WD), уважаем его, иначе B02 4WD premium
    ошибочно матчился бы как 2WD premium и завышал потребность бампера."""
    model = bi.get('model', '') or ''
    drive = bi.get('drive', '') or ''
    config = bi.get('config', '') or ''
    if code and _is_prerestyle_bumper_code(code):
        if model == PRERESTYLE_PREMIUM_MODEL and config in PRERESTYLE_B02_CONFIGS and not drive:
            drive = '2WD'
    return f"{model}_{drive}_{config}"


def _bumper_color_share(bv, color_key):
    """Доля color_key в партии; 0.0 если цвета нет, None если paint stats для партии нет."""
    bc_split = _batch_color_lookup(bv)
    if not bc_split:
        return None
    total_in_batch = sum(bc_split.values())
    if total_in_batch <= 0:
        return None
    color_cars = bc_split.get(color_key, 0)
    if color_cars <= 0:
        return 0.0
    return color_cars / total_in_batch


def _bumper_add_batch_demand(daily, day_qty, share):
    if share is None or share <= 0:
        return
    for day, cars in day_qty.items():
        daily[day] = daily.get(day, 0) + cars * share

# (HEADREST_X2_CODES, OBSOLETE_CODES, CFG_FUZZY_ALT, cfg_match определены в начале файла)


def get_daily_demand_raw_bumper(code, month_num, b_models, tabs):
    """Helper: compute bumper demand without V7 scaling (for ratio calculation)"""
    d={}
    for tab in tabs:
        if tab not in plan_batches or month_num not in plan_batches[tab]: continue
        for bv,day_qty in plan_batches[tab][month_num].items():
            if b_models and _get_batch_info(tab, bv).get('model','') not in b_models: continue
            for day,cars in day_qty.items():
                d[day]=d.get(day,0)+cars
    return d

def get_daily_demand(code, month_num):
    if not paint_material_active(code, month_num):
        return {}
    plan_tab=part_tab_map.get(code,'')
    tabs=resolve_plan_tabs(plan_tab)
    if not tabs: return {}
    daily={}

    # 1. Chemistry / Paint
    if code in chem_norms:
        # ─ Грунты: цвета и применяемость от связанных красок ─
        if code in PRIMER_LINKED_PAINTS:
            return get_primer_demand(code, month_num, tabs=tabs)
        norms=chem_norms[code]
        color_filter=paint_colors.get(code,False)
        # ─ Цветная краска: norm × кузова нужного(их) цвета(ов) ─
        if isinstance(color_filter, list) and color_filter:
            for color_key in color_filter:
                # для каждого цвета — соберём кузова по моделям с подходящей нормой
                for model_key, norm in norms.items():
                    if norm <= 0: continue
                    cars_d = cars_by_color(month_num, color_key, models={model_key}, tabs=tabs)
                    for day, cars in cars_d.items():
                        daily[day] = daily.get(day, 0) + cars * norm
            return daily
        # ─ Краска без цветового фильтра: norm × все кузова ─
        for tab in tabs:
            if tab not in plan_batches or month_num not in plan_batches[tab]: continue
            for bv,day_qty in plan_batches[tab][month_num].items():
                bi = _get_batch_info(tab, bv)
                norm=norms.get(bi.get('model',''),0)
                if norm==0: continue
                for day,cars in day_qty.items():
                    daily[day]=daily.get(day,0)+cars*norm
        return daily

    # ─ Краски ALAA004871/ALAA004873/ALAA007732 без записи в chem_norms ─
    # Базовая эмаль типового расхода ~4 кг/кузов нужного цвета (по аналогии с ALAA005672=4.14)
    PAINT_FALLBACK_NORM = 4.0
    if code in color_filtered_paints and code not in chem_norms:
        color_keys = paint_colors.get(code, []) or []
        for color_key in color_keys:
            cars_d = cars_by_color(month_num, color_key, tabs=tabs)
            for day, cars in cars_d.items():
                daily[day] = daily.get(day, 0) + cars * PAINT_FALLBACK_NORM
        # fallback: если paint stats нет — берём V7
        if not daily and month_num in (5,6) and code in v7_daily and month_num in v7_daily[code]:
            return dict(v7_daily[code][month_num])
        return daily

    # 2. Bumpers — цветной split через batch_color, фильтр по applicable из BOM
    if code in bumper_meta:
        bm = bumper_meta[code]
        color_key = bm.get('color_key','')
        # Применяемость бампера — из BOM (BOM_Детальный). Сопоставление по
        # (модель, комплектация) без учёта привода (_bumper_cfg_match).
        applicable_b = bom.get(code, {}).get('configs', {})
        if not isinstance(applicable_b, dict):
            applicable_b = {k: 1 for k in applicable_b}
        # Запасная эвристика только если бампера НЕТ в BOM вообще.
        if not applicable_b and code not in bom:
            if any(p in code for p in ('XKN61', 'KN260004', 'KN260005')):
                applicable_b = {'B02_4WD_elite': 1, 'B02_4WD_TechPlus': 1}
            elif any(p in code for p in ('XST33', 'AST33', 'AKN02')):
                applicable_b = {'A01_2WD_premium': 1}
        if not applicable_b:
            return {}
        # ── Дедуплицируем партии только с вкладок из part_tab_map ──
        # Потребность цвета C = Σ по партиям (машины партии за день × доля цвета C).
        # Доля цвета берётся из paint statistics (batch → цвет → машины).
        unique_batches = _unique_batches_for_tabs(month_num, tabs)
        use_color_split = bool(color_key and batch_color)
        for tab, bv, day_qty in unique_batches:
            bi = _get_batch_info(tab, bv)
            cfg_key = _bumper_batch_cfg_key(bi, code)
            matched, qty = _bumper_cfg_match(code, cfg_key, applicable_b)
            if not matched:
                continue
            if use_color_split:
                share = _bumper_color_share(bv, color_key)
                if share is None or share <= 0:
                    continue
                _bumper_add_batch_demand(daily, day_qty, share * qty)
            else:
                _bumper_add_batch_demand(daily, day_qty, qty)
        return {d: max(0, int(round(v))) for d, v in daily.items() if round(v) > 0}

    # 3. BOM general — Ecoal'yance: матрица «EcoAlliance use.xlsx» (партии RAS/RBU/RBA)
    if code in eco_demand_rules and code not in eco_superseded_codes:
        eco_d = get_ecoalliance_daily_demand(code, month_num)
        if eco_d is not None:
            return eco_d

    is_b16 = (code in B16_ELITE)
    add_info = ADDL.get(code)
    applicable = bom.get(code, {}).get('configs', {})
    if not add_info and not is_b16:
        if isinstance(applicable, set):
            if not applicable:
                return {}
        elif not applicable:
            return {}
    for tab in tabs:
        if tab not in plan_batches or month_num not in plan_batches[tab]: continue
        for bv,day_qty in plan_batches[tab][month_num].items():
            bi = _get_batch_info(tab, bv)
            b_model=bi.get('model',''); b_drive=bi.get('drive',''); b_config=bi.get('config','')
            qty_per_car = 1  # default: 1 unit per car
            if is_b16:
                # B16 elite headrest: ТОЛЬКО elite-конфиг, не Tech+/premium
                if b_model!='B16': continue
                if b_config != 'elite': continue
            elif add_info:
                if b_model not in add_info['models']: continue
                if add_info['configs'] and b_config not in add_info['configs']: continue
            elif applicable:
                cfg_key = f"{b_model}_{b_drive}_{b_config}"
                matched, q = cfg_match(cfg_key, applicable)
                if not matched:
                    continue
                qty_per_car = q
            else:
                continue
            # Применяемость x2 для задних подголовников
            if code in HEADREST_X2_CODES:
                qty_per_car *= 2
            # Применяемость, заданная вручную (BOM хранит только ✓)
            if code in QTY_PER_CAR_OVERRIDE:
                qty_per_car = QTY_PER_CAR_OVERRIDE[code]
            for day,cars in day_qty.items():
                daily[day]=daily.get(day,0)+cars*qty_per_car
    return daily

# ═══ STEP 7: Pre-compute demand ═══════════════════════════════
try:
    _audit_plan_batches_without_color()
except Exception as _e_audit:
    print(f"  ⚠️  Аудит партий не выполнен: {_e_audit}")

print("\nPre-computing demand...")
demand={}
for code in mrp_codes:
    demand[code]={}
    for mnum,_,_ in MONTHS:
        demand[code][mnum]=get_daily_demand(code,mnum)

# ── SGK: дневной план ориентируем на Stamping plan actual (согласовано) ──
# Последний файл SGK* в папке «Графики поставок», лист «Stamp plan*»:
# блоки по 4 колонки с датой ddmm в 1-й строке и парами (код LK/LP, кол-во).
def _sgk_sheet_plan(rows):
    """{(code, date): qty} из листа SGK: дата в строке-заголовке (datetime или
    ddmm, напр. «1603»), под ней пары (код LK/LP в той же колонке, кол-во справа)."""
    col_date = {}
    for r in rows[:3]:
        for ci, v in enumerate(r or ()):
            dt = None
            if isinstance(v, (datetime.datetime, datetime.date)):
                dt = datetime.date(v.year, v.month, v.day)
            else:
                s = re.sub(r'\D', '', str(v or ''))
                if len(s) == 4:
                    dd, mm = int(s[:2]), int(s[2:])
                    yy = MONTH_YEAR.get(mm)
                    if yy and 1 <= dd <= 31:
                        try:
                            dt = datetime.date(yy, mm, dd)
                        except ValueError:
                            dt = None
            if dt is not None and ci not in col_date:
                col_date[ci] = dt
        if col_date:
            break
    plan = {}
    if not col_date:
        return plan
    for r in rows[1:]:
        if not r:
            continue
        for ci, dt in col_date.items():
            # два варианта раскладки: дата над колонкой КОДА (Stamp plan:
            # код в ci, кол-во в ci+1) или над колонкой КОЛИЧЕСТВА
            # (лист «план»: код в ci-1, кол-во в ci)
            for code_ci, qty_ci in ((ci, ci + 1), (ci - 1, ci)):
                if code_ci < 0:
                    continue
                code = (str(r[code_ci]).strip().upper()
                        if code_ci < len(r) and r[code_ci] else '')
                if not re.match(r'^L[KP]\d', code):
                    continue
                try:
                    q = (float(r[qty_ci])
                         if qty_ci < len(r) and r[qty_ci] is not None else 0.0)
                except (TypeError, ValueError):
                    continue
                if q > 0:
                    plan[(code, dt)] = plan.get((code, dt), 0.0) + q
                break
    return plan

def _apply_sgk_stamping_plan():
    folder = os.path.join(ROOT, 'Графики поставок')
    cands = []
    for pat_dir in (folder, ROOT, os.path.join(ROOT, 'Входные данные')):
        cands += glob.glob(os.path.join(pat_dir, '*SGK*.xls*'))
        cands += glob.glob(os.path.join(pat_dir, '*Stamping*plan*.xls*'))
        cands += glob.glob(os.path.join(pat_dir, '*Stamping*.xls*'))
    cands = [p for p in set(cands) if not _is_excel_lock_file(p)]
    if not cands:
        return
    path = max(cands, key=os.path.getmtime)
    try:
        wb_sgk = load_workbook(path, read_only=True, data_only=True)
    except Exception as _e:
        print(f"  ⚠️  SGK Stamping plan: не открыть {os.path.basename(path)}: {_e}")
        return
    # Приоритет (поздние перекрывают ранние по совпадающим датам):
    # Stamp plan* → план → «актуально …» (самые свежие корректировки)
    tiers = [
        [s for s in wb_sgk.sheetnames if s.lower().replace(' ', '').startswith('stamp')],
        [s for s in wb_sgk.sheetnames if s.strip().lower() == 'план'],
        [s for s in wb_sgk.sheetnames if 'актуал' in s.lower()],
    ]
    merged = {}
    used_sheets = []
    for tier in tiers:
        tier_plan = {}
        for sn in tier:
            rows = list(wb_sgk[sn].iter_rows(values_only=True))
            p = _sgk_sheet_plan(rows)
            if p:
                used_sheets.append(sn)
                for k, v in p.items():
                    tier_plan[k] = tier_plan.get(k, 0.0) + v
        if tier_plan:
            # перезапись дат, которые есть в этом уровне (более свежем)
            tier_dates = {dt for (_, dt) in tier_plan}
            merged = {k: v for k, v in merged.items() if k[1] not in tier_dates}
            merged.update(tier_plan)
    wb_sgk.close()
    n_over = 0
    skipped = 0
    for (code, dt), q in merged.items():
        if code not in demand:
            continue
        if dt.month not in MONTH_YEAR or MONTH_YEAR.get(dt.month) != dt.year:
            skipped += 1
            continue
        demand[code].setdefault(dt.month, {})[dt.day] = q
        n_over += 1
    if n_over:
        print(f"  SGK: дневной план переопределён по «{os.path.basename(path)}» "
              f"(листы: {', '.join(used_sheets)}): {n_over} код-дней"
              + (f"; вне горизонта пропущено: {skipped}" if skipped else ""))
    elif used_sheets:
        print(f"  ℹ️  SGK план прочитан ({', '.join(used_sheets)}), "
              f"но все даты вне расчётного горизонта")

try:
    _apply_sgk_stamping_plan()
except Exception as _sgk_e:
    print(f"  ⚠️  SGK Stamping plan не применён: {_sgk_e}")

with_demand=sum(1 for c in mrp_codes if any(demand[c].get(m) for m,_,_ in MONTHS))
print(f"  Parts with demand > 0: {with_demand} / {len(mrp_codes)}")
for code in ['6803112XKN08A','ALAA005669','1101100AGW01A','2803104XKN61A8T']:
    t=sum(demand.get(code,{}).get(MONTHS[0][0],{}).values())
    print(f"  {code}: May={t:.1f}  tab={part_tab_map.get(code,'?')}")

def _debug_bumper_day6(code):
    if code not in demand:
        return
    _cfgs = sorted(bom.get(code, {}).get('configs', {}).keys())
    print(f"  {code}: tab={part_tab_map.get(code)} BOM={_cfgs}")
    for mnum, mlabel, _ in MONTHS:
        _d6 = demand[code].get(mnum, {}).get(6, 0)
        _tabs = resolve_plan_tabs(part_tab_map.get(code, ''))
        _parts = []
        for _tab in _tabs:
            if _tab not in plan_batches or mnum not in plan_batches[_tab]:
                continue
            for _bv, _dq in plan_batches[_tab][mnum].items():
                _cars6 = _dq.get(6, 0)
                if _cars6 <= 0:
                    continue
                _bi = _get_batch_info(_tab, _bv)
                _ck = _bumper_batch_cfg_key(_bi, code)
                _ok, _ = _bumper_cfg_match(code, _ck, bom.get(code, {}).get('configs', {}))
                if not _ok:
                    continue
                _clr = bumper_meta.get(code, {}).get('color_key', '')
                _exact = 'EXACT' if _normalize_batch_id(_bv) in batch_color else 'fuzzy'
                _bc = _batch_color_lookup(_bv)
                _tot = sum(_bc.values()) if _bc else 0
                _share = _bumper_color_share(_bv, _clr)
                _share_str = 'no_paint' if _share is None else f"{_share:.3f}"
                _bc_str = ','.join(f"{k}:{int(v)}" for k, v in
                                   sorted(_bc.items(), key=lambda kv: -kv[1])[:5]) if _bc else 'no_paint'
                _parts.append(
                    f"{_tab}/{_bv}({_ck}) plan_cars={_cars6} paint={_exact} "
                    f"paint_total={_tot:.0f} share[{_clr}]={_share_str} -> {_cars6*(_share or 0):.0f} | [{_bc_str}]")
        print(f"    {mlabel} day6={_d6:.0f} clr={bumper_meta.get(code,{}).get('color_key','')}")
        for _p in _parts:
            print(f"        {_p}")


_PRERESTYLE_BUMPER_DEBUG = (
    '2803120XST33A8T', '2804104AST33A8T',
    '2803120XST33A9C', '2804104AST33A9C',
    '2803120XST33AC3', '2804104AST33AC3',
    '2803130XST33AC3', '2804105AST33AC3',
)
for _dbc in _PRERESTYLE_BUMPER_DEBUG:
    _debug_bumper_day6(_dbc)

# Сводка потребности по поставщикам (все месяцы расчёта, day-6 отдельно)
print("  Потребность по поставщикам:")
for _supp_key in sorted(set(normalize_supplier(bom.get(c, {}).get('supplier', ''))
                            for c in mrp_codes)):
    if not _supp_key:
        continue
    _parts = sum(1 for _c in mrp_codes
                 if normalize_supplier(bom.get(_c, {}).get('supplier', '')) == _supp_key
                 and any(demand[_c].get(_mn) for _mn, _, _ in MONTHS))
    _month_totals = []
    _day6_total = 0.0
    for _mn, _ml, _nd in MONTHS:
        _mt = sum(sum(demand[_c].get(_mn, {}).values())
                  for _c in mrp_codes
                  if normalize_supplier(bom.get(_c, {}).get('supplier', '')) == _supp_key)
        _month_totals.append(f"{_ml[:3]}={_mt:.0f}")
        if _nd >= 6:
            _day6_total += sum(demand[_c].get(_mn, {}).get(6, 0)
                               for _c in mrp_codes
                               if normalize_supplier(bom.get(_c, {}).get('supplier', '')) == _supp_key)
    print(f"    {_supp_key}: {_parts} поз. | {', '.join(_month_totals)} | day6={_day6_total:.0f}")

def get_info(code):
    if code in bom:
        b=bom[code]; return b.get('name',''),b.get('supplier',''),b.get('unit','шт')
    if code in ADDL:
        a=ADDL[code]; return a['name'],a['supplier'],a['unit']
    return '','','шт'

def get_norm_str(code):
    if code in chem_norms:
        ns='/'.join(f"{m}:{v:.3g}" for m,v in chem_norms[code].items() if v>0)
        pc=paint_colors.get(code,False)
        tag='' if pc is False else ' [все цвета]' if pc is None else f' [{",".join(str(x) for x in pc[:2])}...]'
        return ns+tag
    if code in bumper_meta:
        return f"1 шт [цвет: {bumper_meta[code].get('color_key','')}]"
    if code in bom:
        cfgs=bom[code].get('configs',set())
        return f"✓ ({len(cfgs)} конф.)" if cfgs else '— без применяемости'
    if code in ADDL:
        ai=ADDL[code]; return f"Доп: {','.join(ai['models'][:3])}"
    return '—'

# ── Страховой запас MSA-бамперов: 1 полная партия, разбитая по цветам кузовов ──
MSA_BUMPER_BATCH = 120  # размер 1 производственной партии (кузовов)
_ndkt_first_batch_cache = {}
_ndkt_transition_cache = {}

def _is_ndkt_supplier(supplier):
    return normalize_supplier(supplier) == NDKT_SUPPLIER

def _is_ndkt_part(code):
    return _is_ndkt_supplier(bom.get(code, {}).get('supplier', ''))

def _ndkt_batch_day_qty(tab, month_num, bv_norm):
    tab_data = plan_batches.get(tab, {}).get(month_num, {})
    for bvk, day_qty in tab_data.items():
        if _normalize_batch_id(bvk) == bv_norm:
            return day_qty, bvk
    return None, None

def _ndkt_part_tabs(code):
    """Вкладки плана для NDKT-детали (обычно AS_in_H_B для B06)."""
    tabs = [t for t in resolve_plan_tabs(part_tab_map.get(code, '')) if t in NDKT_AS_TABS]
    return tabs or ['AS_in_H_B']

def _ndkt_batch_days(month_num, bv_norm, tabs=None):
    """Календарные дни 1-й партии на выбранных вкладках (объединение без дублей)."""
    tabs = tabs or NDKT_AS_TABS
    days = set()
    for tab in tabs:
        day_qty, _ = _ndkt_batch_day_qty(tab, month_num, bv_norm)
        if day_qty:
            days.update(int(d) for d in day_qty.keys())
    return days

def _ndkt_plan_first_batch_id(month_num):
    """(first_day, batch_id_norm) — самая ранняя партия на AS_in_F_A / AS_in_H_B."""
    if month_num in _ndkt_first_batch_cache:
        return _ndkt_first_batch_cache[month_num]
    best = None
    for tab in NDKT_AS_TABS:
        for bv, day_qty in plan_batches.get(tab, {}).get(month_num, {}).items():
            if not day_qty:
                continue
            try:
                first_day = min(int(d) for d in day_qty.keys())
            except (TypeError, ValueError):
                continue
            key = (first_day, _normalize_batch_id(bv))
            if best is None or key < best:
                best = key
    _ndkt_first_batch_cache[month_num] = best
    return best

def _ndkt_first_batch_id(month_num):
    return _ndkt_plan_first_batch_id(month_num)

def _ndkt_part_demand_on_batch_days(code, month_num, batch_days):
    """Потребность code за дни партии — из уже рассчитанной demand (без двойного AS)."""
    if not batch_days:
        return 0.0
    daily = demand.get(code, {}).get(month_num, {})
    return sum(float(daily.get(d, 0) or 0) for d in batch_days)

def _ndkt_part_demand_from_first_batch(code, month_num, batch_id):
    """Потребность code на 1-ю производственную партию (дни партии × demand, одна вкладка детали)."""
    if not batch_id:
        return 0.0
    _, bv_norm = batch_id
    tabs = _ndkt_part_tabs(code)
    batch_days = _ndkt_batch_days(month_num, bv_norm, tabs)
    return _ndkt_part_demand_on_batch_days(code, month_num, batch_days)

def _ndkt_first_batch_for_part(code, month_num):
    """1-я партия плана, где у code есть потребность (на вкладках детали)."""
    plan_fb = _ndkt_plan_first_batch_id(month_num)
    tabs = _ndkt_part_tabs(code)
    best = None  # ((first_day, bv_norm), demand)
    if plan_fb:
        _, bv_norm = plan_fb
        days = _ndkt_batch_days(month_num, bv_norm, tabs)
        dem = _ndkt_part_demand_on_batch_days(code, month_num, days)
        if dem > 0:
            return plan_fb, dem
    for tab in tabs:
        for bv, day_qty in plan_batches.get(tab, {}).get(month_num, {}).items():
            if not day_qty:
                continue
            try:
                first_day = min(int(d) for d in day_qty.keys())
            except (TypeError, ValueError):
                continue
            bv_norm = _normalize_batch_id(bv)
            days = _ndkt_batch_days(month_num, bv_norm, [tab])
            dem = _ndkt_part_demand_on_batch_days(code, month_num, days)
            if dem <= 0:
                continue
            key = (first_day, bv_norm)
            if best is None or key < best[0]:
                best = (key, dem)
    if best:
        return best[0], best[1]
    return plan_fb, 0.0

def _ndkt_tab_first_batch(tab, month_num):
    """(first_day, batch_id_norm) — самая ранняя партия на конкретной вкладке."""
    best = None
    for bv, day_qty in plan_batches.get(tab, {}).get(month_num, {}).items():
        if not day_qty:
            continue
        try:
            first_day = min(int(d) for d in day_qty.keys())
        except (TypeError, ValueError):
            continue
        key = (first_day, _normalize_batch_id(bv))
        if best is None or key < best:
            best = key
    return best

def _ndkt_transition_qty(code):
    """Переходящий остаток NDKT (колонка F / расчёт заказа): ежедневный
    страховой запас по коду (NDKT_SAFETY_BY_CODE; default 80 шт).
    Потребность дня закрывается поставкой в день потребности."""
    if code in _ndkt_transition_cache:
        return _ndkt_transition_cache[code]
    if not _is_ndkt_part(code) or _code_3m_demand(code) <= 0:
        _ndkt_transition_cache[code] = 0.0
        return 0.0
    _ndkt_transition_cache[code] = _ndkt_safety_base(code)
    return _ndkt_transition_cache[code]

_BUMPER_COLOR_SUFFIX = ('A8T', 'A9C', 'AC3', 'AH4', 'A5B', 'AGN')
def _bumper_type_key(code):
    """Группа цветовых вариантов одного физического бампера (без цветового суффикса)."""
    for _suf in _BUMPER_COLOR_SUFFIX:
        if code.endswith(_suf):
            return code[:-len(_suf)]
    return code
def _is_msa_bumper(code):
    return (code in bumper_meta and code not in REFERENCE_ONLY_CODES
            and normalize_supplier(bom.get(code, {}).get('supplier', '')) == 'MSA')
_msa_type_demand_cache = {}
def _code_3m_demand(code):
    return sum(sum(demand.get(code, {}).get(mn, {}).values()) for mn, _, _ in MONTHS)
def _safety_qty(code, supplier=''):
    """Страховой запас (шт).
    Lear: из файла «страх.запасы ЛИР.xlsx» (по коду).
    Накладки MSA: без накрутки (партия 480/180 — сама буфер, запасы не раздуваем).
    MSA-бамперы: согласованные уровни (A01: ходовые 90 / неходовые 45;
    B02/B04: Premium+TechPlus 90 / Elite 60; база — полпартии 100%).
    NDKT: ежедневный страховой запас по коду (NDKT_SAFETY_BY_CODE).
    Остальные детали: avg_daily × дни страх.запаса."""
    if code in LEAR_SAFETY_BY_CODE:
        # Lear: страховой из файла ЛИР, но не ниже 1,5 средних дневных
        # потребностей (по эталонам планировщик держит 1,5–3 дня) — согласовано.
        _days_l = [q for _mn, _, _ in MONTHS
                   for q in demand.get(code, {}).get(_mn, {}).values() if q > 0]
        _avg_l = (sum(_days_l) / len(_days_l)) if _days_l else 0.0
        return round(max(float(LEAR_SAFETY_BY_CODE[code]), 1.5 * _avg_l), 1)
    if code in MSA_FRONT_TRIM_CODES or code in MSA_REAR_TRIM_CODES:
        return float(MSA_TRIM_SAFETY_QTY)
    if _is_msa_bumper(code):
        if _code_3m_demand(code) <= 0:
            return 0.0
        return _msa_bumper_safety_qty(code)
    if _is_ndkt_part(code):
        return _ndkt_transition_qty(code)
    _sd = get_safety_days(code, supplier)
    _tot = _code_3m_demand(code)
    _n3 = sum(nd for _, _, nd in MONTHS)
    _avg = _tot / _n3 if _n3 else 0
    return round(_avg * _sd, 2)

def _first_calc_month_num():
    return MONTHS[0][0] if MONTHS else CALC_RUN_MONTH

def _snapshot_stock_date(code):
    """Дата снимка остатка, введённого пользователем для кода.

    Когда остаток вводится в СЕТКУ листа Ввод_Остатков (жёлтый столбец с датой,
    напр. 15.06) или приходит снимком из файла, эта дата и есть «остатки на дату».
    Раньше окно потребности 1-го месяца бралось только из get_stock_date() —
    глобальной/файловой даты, не связанной со снимком, поэтому потребность
    июня считалась не с даты ввода остатков (с 15.06 → с 17.06), а с более
    поздней даты, и часть месяца терялась. Берём САМУЮ ПОЗДНЮЮ дату снимка
    в пределах 1-го расчётного месяца — остаток актуален на неё, а потребность
    считается с этого дня (выходные/дни без партий перед первой партией дают 0,
    поэтому фактический старт совпадает с первой производственной партией)."""
    first_m = _first_calc_month_num()
    y = MONTH_YEAR.get(first_m, CALC_RUN_YEAR)
    best = None
    for _src in (manual_stock_by_date, file_stock_by_date):
        for dt in _src.get(code, {}):
            if dt.year == y and dt.month == first_m:
                if best is None or dt > best:
                    best = dt
    return best

def _calc_from_date(code):
    """Дата, С КОТОРОЙ считается оставшаяся потребность 1-го месяца —
    дата предоставления остатков по коду. Приоритет — у введённого снимка
    (сетка Ввод_Остатков / файл), иначе per-code дата из stock_date_map."""
    snap = _snapshot_stock_date(code)
    if snap is not None:
        return snap
    return get_stock_date(code) or _stock_date

def _remaining_demand_for_month(code, month_num):
    """Оставшаяся потребность: 1-й расчётный месяц — с ДАТЫ ПРЕДОСТАВЛЕНИЯ
    ОСТАТКОВ кода (включительно) до конца месяца. Будущие месяцы — полностью."""
    daily = demand[code].get(month_num, {})
    if not daily:
        return 0.0
    if month_num == _first_calc_month_num():
        sd = _calc_from_date(code)
        y = MONTH_YEAR.get(month_num, CALC_RUN_YEAR)
        if sd.year == y and sd.month == month_num:
            return sum(q for d, q in daily.items() if d >= sd.day)
        if (sd.year, sd.month) > (y, month_num):
            return 0.0
        return sum(daily.values())
    return sum(daily.values())

def calc_order(code, month_num, stock_override=None):
    """Расчёт заказа на месяц.
    1-й расчётный месяц: потребность с даты запуска (CALC_RUN_DATE), не с 1-го числа.
    Будущие месяцы окна — полная потребность. Прошлые — 0.
    """
    d_full = sum(demand[code].get(month_num, {}).values())
    d_remaining = _remaining_demand_for_month(code, month_num)
    first_m = _first_calc_month_num()
    if month_num == first_m:
        d_future = d_remaining
        d_display = d_remaining
    elif month_num < first_m:
        d_future = 0.0
        d_display = 0.0
    else:
        d_future = d_full
        d_display = d_full
    safety = _safety_qty(code, bom.get(code, {}).get('supplier', ''))
    stk = stock_override if stock_override is not None else stock.get(code, 0)
    net = max(0, d_future + safety - stk)
    pkg = max(1, bom.get(code, {}).get('package', 1))
    pkgs = math.ceil(net / pkg) if net > 0 else 0
    return {'demand': round(d_display, 2), 'demand_full': round(d_full, 2),
            'demand_remaining': round(d_remaining, 2), 'stock': stk,
            'safety': round(safety, 2), 'net': round(net, 2), 'pkg_size': pkg,
            'packages': pkgs, 'order_qty': round(pkgs * pkg, 2)}

def weekly_demand_map(code):
    wd={}
    for mnum,_,n_days in MONTHS:
        for day,qty in demand[code].get(mnum,{}).items():
            wi=date_to_week(datetime.date(MONTH_YEAR.get(mnum,_today.year),mnum,day))
            if wi>=0: wd[wi]=wd.get(wi,0)+qty
    return wd

# ═══ Прогнозный остаток на 1 июня и 1 июля ════════════════════
# Исправление по замечанию: "После расчёта майского заказа не учитываются
# остатки на 1 июня — они берутся как остатки на 11-13 мая во всех вкладках."
#
# Алгоритм:
#   stock_jun1[code] = stock_may13 - demand_may(14..31) + may_order
#   stock_jul1[code] = stock_jun1  - demand_jun(1..30)  + jun_order
#
# projected_stock[code][month_num] = открывающий остаток на 1-е число месяца.
# Для мая (STOCK_AS_OF_MONTH) используется физический остаток напрямую.
print(f"\nПрогнозный расчёт остатков на начало каждого месяца...")
projected_stock = {}
for code in mrp_codes:
    s = _opening_stock(code)
    ps = {}
    for i, (mn, _, _) in enumerate(MONTHS):
        ps[mn] = max(0.0, s)
        if i == 0:
            remaining = _remaining_demand_for_month(code, mn)
            ord_qty = calc_order(code, mn)['order_qty']
        else:
            remaining = sum(demand[code].get(mn, {}).values())
            ord_qty = calc_order(code, mn, stock_override=s)['order_qty']
        s = max(0.0, s - remaining + ord_qty)
    projected_stock[code] = ps
print(f"  Прогнозный остаток рассчитан для {len(projected_stock)} позиций")

# ═══ STEP 8: Build Excel ══════════════════════════════════════
print("\nBuilding Excel...")
wb_out=Workbook(); wb_out.remove(wb_out.active)
all_dates=[(datetime.date(MONTH_YEAR[mn],mn,d),mn,d) for mn,_,nd in MONTHS for d in range(1,nd+1)]
_GLOBAL_DI = {(mn, d): di for di, (_dt, mn, d) in enumerate(all_dates)}  # (месяц, день) → индекс дня горизонта
# mc_start: динамически — стартовая колонка каждого месяца в График_Поставок
# col A=Код, B-F=info (5 cols), затем попарно (Del|Ss) на каждый день месяца
mc_start = {}
_col = 6  # первая колонка данных (A=1..F=6 — info)
for _mn, _, _nd in MONTHS:
    mc_start[_mn] = _col
    _col += _nd * 2   # каждый день = 2 колонки

# ═══ Расписание поставок (правила по поставщикам) ═══════════════
def _ceiling_pkg_qty(qty, pkg):
    pkg = max(1, int(pkg) if pkg else 1)
    if qty <= 0:
        return 0
    return math.ceil(qty / pkg) * pkg

def _floor_pkg_qty(qty, pkg):
    """Округление вниз до целых упаковок (при урезании фуры/лимита)."""
    pkg = max(1, int(pkg) if pkg else 1)
    if qty <= 0:
        return 0
    return int(qty // pkg) * pkg

def _pkg_size(code):
    return max(1, int(bom.get(code, {}).get('package', 1) or 1))

def _is_pkg_multiple(qty, pkg):
    if qty <= 0:
        return True
    pkg = max(1, int(pkg))
    rem = qty % pkg
    return rem < 1e-6 or abs(rem - pkg) < 1e-6

def _round_delivery_qty(code, qty, pkg=None):
    """Отгрузка = целое число упаковок одного кода (+ правила накладок MSA,
    фур 5304100XKN02A и кратности 3 уп. подголовников Lear)."""
    if qty <= 0:
        return 0
    if code == MSA_FULL_TRUCK_CODE:
        return _msa_full_truck_qty(qty)
    if pkg is None:
        pkg = _pkg_size(code)
    qty = _ceiling_pkg_qty(qty, pkg)
    if code in EFTEC_DRUM_MULTIPLE_CODES:
        # Кратно 4 бочкам: 1 бочка = 1 упаковка кода
        step = pkg * EFTEC_DRUMS_PER_SHIPMENT
        return math.ceil(qty / step) * step
    if code in VM_ROW_QTY:
        # ВМ Авто: отгрузка целыми рядами кузова
        step = VM_ROW_QTY[code]
        return math.ceil(qty / step) * step
    if code in MTS_BOX_QTY:
        # МТС-авто: отгрузка целыми паллетами (16 коробов)
        step = MTS_BOX_QTY[code] * MTS_BOXES_PER_PALLET
        return math.ceil(qty / step) * step
    if code in MSA_FRONT_TRIM_CODES:
        m = MSA_FRONT_TRIM_MULTIPLE
        qty = max(math.ceil(qty / m) * m, MSA_FRONT_TRIM_MIN_QTY)
    elif code in MSA_REAR_TRIM_CODES:
        # кратно таре 60, минимальный комплект 180 (по эталону MSA)
        m = MSA_REAR_TRIM_MULTIPLE
        qty = max(math.ceil(qty / m) * m, MSA_REAR_TRIM_MIN_QTY)
    # Подголовники Lear: кратность 3 применяется к СУММЕ упаковок дня
    # (таблица микса машины), НЕ к каждой позиции — по эталону планировщика.
    return _ceiling_pkg_qty(qty, pkg)

def _enforce_delivery_pkg_multiples(all_dels, codes=None):
    """Каждая 📦 — целое число упаковок одного кода (1 уп = 1 вид, без дробления)."""
    codes = codes or list(all_dels.keys())
    fixed = 0
    for code in codes:
        if code not in all_dels:
            continue
        if code in URAL_ALL_CODES:
            continue   # у Урала своя норма загрузки фуры (480 не кратно 360)
        pkg = _pkg_size(code)
        dels = all_dels[code]
        for di in range(len(dels)):
            q = float(dels[di] or 0)
            if q <= 0:
                dels[di] = 0.0
                continue
            if not _is_pkg_multiple(q, pkg):
                new_q = float(_round_delivery_qty(code, q, pkg))
                if abs(new_q - q) > 1e-6:
                    fixed += 1
                dels[di] = new_q
    if fixed:
        print(f"  Упаковки: приведено отгрузок к кратности упаковки: {fixed}", flush=True)

def _xlsheet_ref(sheet_name, cell_addr):
    """Ссылка на ячейку другого листа — всегда в кавычках (кириллица ломает Excel без них)."""
    sn = str(sheet_name).replace("'", "''")
    return f"'{sn}'!{cell_addr}"

def _sanitize_excel_formula(formula):
    """Отбрасывает битые формулы из старого файла (Excel repair / sheet10.xml)."""
    if not isinstance(formula, str):
        return None
    f = formula.strip()
    if not f:
        return None
    if not f.startswith('='):
        f = '=' + f
    body = f[1:]
    if len(f) > 8192:
        return None
    if re.search(r'[\x00-\x08\x0b\x0c\x0e-\x1f]', f):
        return None
    # VLOOKUP(...,41,) — лишняя запятая перед ')': Excel удаляет формулу при открытии
    if re.search(r',\s*\)', body):
        return None
    # Ссылки [1]DelJune! / [2]Лист2! — внешние книги, без них формула невалидна
    if re.search(r'\[\d+\]', body):
        return None
    return f

def _graph_stock_formula(ri, prev_col_l, prev_col_n, pkg, fallback):
    """Переходящий остаток из График_Поставок, округление вверх до упаковки."""
    pkg = max(1, int(pkg) if pkg else 1)
    fb = int(fallback) if fallback else 0
    tbl = _xlsheet_ref('График_Поставок', f'$A:${prev_col_l}')
    return (
        f"=IFERROR(CEILING(VLOOKUP(A{ri},{tbl},"
        f"{prev_col_n},0)/{pkg},1)*{pkg},{fb})"
    )

_delivery_series_cache = {}


def _delivery_series_demand(code):
    """Спрос для графика поставок — как на вкладке Потребность (все дни с потребностью).

    Кешируется: ряд не меняется за прогон, а функция вызывается десятки тысяч раз
    (из _compute_ss_path в цикле донабора). Без кеша пересборка 92 значений с
    round() на каждый вызов и была причиной «долгой финальной сверки Ss».
    Возвращаемый список НЕ мутируют — проверено по всем 20 местам использования.
    """
    s = _delivery_series_cache.get(code)
    if s is None:
        s = [round(demand[code].get(mn, {}).get(d, 0), 4)
             for (dt, mn, d) in all_dates]
        _delivery_series_cache[code] = s
    return s

def _delivery_safety_qty(code, supplier):
    return _safety_qty(code, supplier)

def _delivery_mode(supplier):
    s = normalize_supplier(supplier)
    if s == ECOALYANCE_SUPPLIER:
        return 'monthly_early'
    if s == ECOTEXIS_SUPPLIER:
        return 'monthly_lag'
    if s == SMC_SUPPLIER:
        return 'smc_weekly'
    if s == PUREM_SUPPLIER:
        return 'weekly'
    if s == MTS_SUPPLIER:
        # Согласовано: МТС-авто — 2 отгрузки в неделю (было — ежедневно)
        return 'twice_weekly'
    _su = s.upper()
    if _su.startswith('ZGM') or 'ЗГМ' in _su:
        # Согласовано: ЗГМ — 1 поставка в месяц (замечание планировщика)
        return 'monthly_early'
    if _su.startswith('AVTOKOM') or 'АВТОКОМ' in _su:
        # Согласовано (Кинешма): недельная консолидация, отгрузки по вторникам
        return 'avtokom_weekly'
    if _su.startswith('ITELMA') or 'ИТЕЛМА' in _su:
        # Согласовано: Ителма — 2 отгрузки в месяц (≈1-е и 16-е)
        return 'semimonthly'
    return 'lookahead'

def _lookahead_demand(dems, di):
    for _lk in range(1, 8):
        if di + _lk >= len(dems):
            break
        if all_dates[di + _lk][0].weekday() == 6:
            continue
        if dems[di + _lk] > 0:
            return dems[di + _lk]
    return 0.0

def _monthly_early_delivery_indices():
    """Ecoal'yance: 1 поставка в месяц, в 1–5 числах."""
    dis, seen = [], set()
    for di, (dt, mn, d) in enumerate(all_dates):
        if mn in seen:
            continue
        if 1 <= d <= 5 and dt.weekday() != 6:
            dis.append(di)
            seen.add(mn)
    return dis

def _ecotexis_delivery_indices():
    """Ecotexis: поставка через 35 дней от отправки заказа (1-е число месяца)."""
    dis = []
    for mn, _, _ in MONTHS:
        y = MONTH_YEAR.get(mn, _today.year)
        order_sent = datetime.date(y, mn, 1)
        if mn == _first_calc_month_num():
            order_sent = max(order_sent, CALC_RUN_DATE)
        del_dt = order_sent + datetime.timedelta(days=ECOTEXIS_ORDER_LAG_DAYS)
        for di, (dt, _, _) in enumerate(all_dates):
            if dt == del_dt:
                dis.append(di)
                break
    return dis

def _weekday_delivery_indices(weekday):
    return [di for di, (dt, _, _) in enumerate(all_dates) if dt.weekday() == weekday]

def _period_delivery_qty(ss_before, dems, del_di, next_del_di, safety, pkg):
    end_di = next_del_di if next_del_di is not None else len(dems)
    period = sum(dems[del_di:end_di])
    need = safety + period - ss_before
    return _ceiling_pkg_qty(max(0, need), pkg)

def _smc_units_per_pallet(code, pkg):
    return SMC_UNITS_PER_PALLET.get(code) or max(1, pkg)

def _smc_pallet_count(code, qty, pkg):
    upp = _smc_units_per_pallet(code, pkg)
    return math.ceil(qty / upp) if qty > 0 else 0

def _pkg_note_for_code(code):
    return (manual_pkg_notes.get(code)
            or bom.get(code, {}).get('pkg_note', '')
            or '')

_lear_kind_cache = {}

def _lear_product_kind(code):
    """Пена / Подголовник — из примечания листа Упаковка, имени, кода HEADREST_X2;
    для Lear-кодов без пометки — эвристика по коду (70xx — подголовники, 68/69 — пены)."""
    if code in _lear_kind_cache:
        return _lear_kind_cache[code]
    note = _pkg_note_for_code(code)
    name = bom.get(code, {}).get('name', '')
    text = f"{note} {name}".lower()
    kind = None
    if 'подголов' in text or 'headrest' in text or code in HEADREST_X2_CODES:
        kind = 'headrest'
    elif 'пен' in text or 'foam' in text:
        kind = 'foam'
    elif normalize_supplier(bom.get(code, {}).get('supplier', '')) == LEAR_SUPPLIER:
        # эвристика по коду: 70xxxxx — подголовники, 68/69xxxxx — пены сидений
        if code.startswith('70'):
            kind = 'headrest'
        elif code.startswith('68') or code.startswith('69'):
            kind = 'foam'
    _lear_kind_cache[code] = kind
    return kind

def _lear_pkg_slot_cost(code, kind=None):
    """Паллетоместа на 1 ШТУКУ: шт → упаковки (тара из файла страх.запасов ЛИР) →
    места (2 уп. пены = 1 место; 3 уп. подголовников = 1 место).
    Ранее считалось от штук без деления на упаковку — фура «переполнялась» в сотни раз
    и лимит 36 уп. пен / машину фактически не работал."""
    kind = kind or _lear_product_kind(code)
    pkg = max(1, bom.get(code, {}).get('package', 1))
    if kind == 'foam':
        return 1.0 / (pkg * LEAR_FOAM_PKGS_PER_SLOT)
    if kind == 'headrest':
        return 1.0 / (pkg * LEAR_HEADREST_PKGS_PER_SLOT)
    return 1.0 / pkg

def _lear_day_pallet_slots(raw_qty):
    """Сумма паллетомест по дню (Lear): 2 пены или 3 подголовника на место."""
    if not raw_qty:
        return 0.0
    total = 0.0
    for code, qty in raw_qty.items():
        if qty <= 0:
            continue
        total += qty * _lear_pkg_slot_cost(code)
    return total

def _cap_lear_day_deliveries(raw_qty):
    """Ограничение фуры Lear: ≤ 18 паллетомест (36 уп. пены или 54 подголовника)."""
    if not raw_qty:
        return raw_qty
    slots = _lear_day_pallet_slots(raw_qty)
    if slots <= LEAR_MAX_PALLET_SLOTS + 1e-9:
        return raw_qty
    scale = LEAR_MAX_PALLET_SLOTS / slots
    out = {}
    for code, qty in raw_qty.items():
        pkg = max(1, bom.get(code, {}).get('package', 1))
        new_q = qty * scale
        new_q = (int(new_q // pkg)) * pkg if pkg else 0
        if new_q > 0:
            out[code] = new_q
    return out

def _smooth_lear_trucks(all_dels, lear_codes):
    """Lear ≤ 18 паллетомест/день. Избыток переносим на более ранние дни с потребностью."""
    n = len(all_dates)
    CAP = LEAR_MAX_PALLET_SLOTS
    SOFT = _soft_truck_cap(CAP)
    kinds = {c: _lear_product_kind(c) for c in lear_codes}
    pkgs = {c: max(1, bom.get(c, {}).get('package', 1)) for c in lear_codes}
    demday = {c: [d > 0 for d in _delivery_series_demand(c)] for c in lear_codes}
    load = [_lear_day_pallet_slots({c: all_dels[c][di] for c in lear_codes if all_dels[c][di] > 0})
            for di in range(n)]
    for di in range(n - 1, -1, -1):
        guard = 0
        while load[di] > SOFT + 1e-9 and guard < 100000:
            guard += 1
            moved = False
            for c in lear_codes:
                pkg = pkgs[c]
                if all_dels[c][di] < pkg:
                    continue
                slot_cost = pkg * _lear_pkg_slot_cost(c, kinds[c])
                for k in range(di - 1, -1, -1):
                    if not demday[c][k]:
                        continue
                    if load[k] + slot_cost <= SOFT + 1e-9:
                        all_dels[c][di] -= pkg
                        all_dels[c][k] += pkg
                        load[di] -= slot_cost
                        load[k] += slot_cost
                        moved = True
                        break
                if moved:
                    break
            if not moved:
                capped = _cap_lear_day_deliveries(
                    {c: all_dels[c][di] for c in lear_codes if all_dels[c][di] > 0})
                for c in lear_codes:
                    all_dels[c][di] = capped.get(c, 0.0)
                load[di] = _lear_day_pallet_slots(
                    {c: all_dels[c][di] for c in lear_codes if all_dels[c][di] > 0})
                break

def _smooth_eftec_weight(all_dels):
    """Эфтек: суммарная дневная отгрузка ≤ 18 т. Излишек переносится на более
    ранние дни с потребностью (кратно упаковке)."""
    ef = [c for c in mrp_codes
          if normalize_supplier(get_info(c)[1]) == EFTEC_SUPPLIER and c in all_dels]
    if not ef:
        return
    n = len(all_dates)
    pkgs = {c: _pkg_size(c) for c in ef}
    w = {c: _eftec_unit_weight_g(c) for c in ef}
    demday = {c: [d > 0 for d in _delivery_series_demand(c)] for c in ef}
    # Запрещённые дни отгрузки Эфтек (пн/вс) — ни как источник, ни как цель
    ok_day = [all_dates[k][0].weekday() not in EFTEC_FORBIDDEN_WEEKDAYS
              for k in range(n)]

    def _day_w(di):
        return sum(float(all_dels[c][di] or 0) * w[c] for c in ef)

    # 1. Убрать отгрузки с запрещённых дней на ближайший разрешённый день
    #    со спросом ДО него (позже нельзя — материал нужен к дню потребности)
    _moved_off = 0
    for di in range(n):
        if ok_day[di]:
            continue
        for c in ef:
            q = float(all_dels[c][di] or 0)
            if q <= 0:
                continue
            pkg = pkgs[c]
            step_w = pkg * w[c]
            # Раскладываем по РАЗРЕШЁННЫМ дням со спросом, соблюдая лимит 18 т,
            # а не сваливаем всё на один день (иначе развесовка потом разгребает
            # это пакетами и прогон встаёт).
            for k in range(di - 1, -1, -1):
                if q <= 1e-9:
                    break
                if not (ok_day[k] and demday[c][k]):
                    continue
                room = EFTEC_MAX_TRUCK_WEIGHT_G - _day_w(k)
                if room < step_w:
                    continue
                take = min(q, (room // step_w) * pkg)
                if take <= 0:
                    continue
                all_dels[c][k] = float(all_dels[c][k] or 0) + take
                all_dels[c][di] -= take
                q -= take
            if q > 1e-9:
                # назад не влезло — ближайший разрешённый день со спросом вперёд
                for k in range(di + 1, n):
                    if ok_day[k] and demday[c][k]:
                        all_dels[c][k] = float(all_dels[c][k] or 0) + q
                        all_dels[c][di] -= q
                        q = 0.0
                        break
            if float(all_dels[c][di] or 0) <= 1e-9:
                all_dels[c][di] = 0.0
                _moved_off += 1
    if _moved_off:
        print(f"  Эфтек: перенесено отгрузок с пн/вс: {_moved_off}", flush=True)

    # 2. Разгрузить дни тяжелее 18 т
    for di in range(n - 1, -1, -1):
        guard = 0
        while _day_w(di) > EFTEC_MAX_TRUCK_WEIGHT_G + 1e-6 and guard < 20000:
            guard += 1
            moved = False
            for c in sorted(ef, key=lambda x: -float(all_dels[x][di] or 0) * w[x]):
                pkg = pkgs[c]
                if float(all_dels[c][di] or 0) < pkg:
                    continue
                step_w = pkg * w[c]
                for k in range(di - 1, -1, -1):
                    if not demday[c][k] or not ok_day[k]:
                        continue
                    if _day_w(k) + step_w <= EFTEC_MAX_TRUCK_WEIGHT_G + 1e-6:
                        all_dels[c][di] -= pkg
                        all_dels[c][k] += pkg
                        moved = True
                        break
                if moved:
                    break
            if not moved:
                break

def _enforce_vm_truck_rows(all_dels):
    """ВМ Авто: отгрузка целыми рядами кузова, суммарная длина рядов за день
    не больше длины кузова (13.4 м). Излишек переносится на более ранние дни
    с потребностью — как у Эфтек по весу."""
    vm = [c for c in mrp_codes if c in VM_ROW_QTY and c in all_dels]
    if not vm:
        return
    n = len(all_dates)
    demday = {c: [d > 0 for d in _delivery_series_demand(c)] for c in vm}

    def _rows(c, di):
        q = float(all_dels[c][di] or 0)
        return math.ceil(q / VM_ROW_QTY[c]) if q > 0 else 0

    def _day_len(di):
        return sum(_rows(c, di) * VM_ROW_LEN_M[c] for c in vm)

    # 1. Округлить каждую отгрузку до целого ряда
    for c in vm:
        step = VM_ROW_QTY[c]
        for di in range(n):
            q = float(all_dels[c][di] or 0)
            if q > 0:
                all_dels[c][di] = float(math.ceil(q / step) * step)

    # 2. Разгрузить дни, где груз длиннее кузова
    # Длины дней держим в кеше и правим инкрементально: пересчёт _day_len на
    # каждый перенесённый ряд делал прогон неприемлемо долгим (зависание).
    dlen = [_day_len(di) for di in range(n)]
    _moved = 0
    for di in range(n - 1, -1, -1):
        guard = 0
        while dlen[di] > VM_TRUCK_LEN_M + 1e-9 and guard < 2000:
            guard += 1
            moved = False
            for c in sorted(vm, key=lambda x: -_rows(x, di) * VM_ROW_LEN_M[x]):
                if _rows(c, di) <= 0:
                    continue
                step = VM_ROW_QTY[c]
                L = VM_ROW_LEN_M[c]
                for k in range(di - 1, -1, -1):
                    if not demday[c][k]:
                        continue
                    if dlen[k] + L <= VM_TRUCK_LEN_M + 1e-9:
                        all_dels[c][di] -= step
                        all_dels[c][k] = float(all_dels[c][k] or 0) + step
                        dlen[di] -= L
                        dlen[k] += L
                        moved = True
                        _moved += 1
                        break
                if moved:
                    break
            if not moved:
                break
    if _moved:
        print(f"  ВМ Авто: перенесено рядов из-за длины кузова: {_moved}", flush=True)


def _enforce_ural_truck_composition(all_dels):
    """Уралэластотехника: целые фуры, все коды группы едут вместе.
      A01 + A08  → по 360 шт каждого A01 и по 120 шт каждого A08
                   (8402106AKJ20A — 240, применяемость 2);
      только A01 → по 480 шт каждого кода;
      только A08 → не менее 4 комплектов (4 × норма смешанной фуры).
    Количество фур k подбирается по самому «голодному» коду дня.
    """
    a01 = [c for c in URAL_A01_CODES if c in all_dels]
    a08 = [c for c in URAL_A08_CODES if c in all_dels]
    if not a01 and not a08:
        return
    for di in range(len(all_dates)):
        q01 = {c: float(all_dels[c][di] or 0) for c in a01}
        q08 = {c: float(all_dels[c][di] or 0) for c in a08}
        need01 = any(v > 0 for v in q01.values())
        need08 = any(v > 0 for v in q08.values())
        if not need01 and not need08:
            continue
        if need01 and need08:
            k = 1
            for c, v in list(q01.items()) + list(q08.items()):
                if v > 0:
                    k = max(k, math.ceil(v / URAL_MIXED_QTY[c]))
            for c in a01 + a08:
                all_dels[c][di] = float(k * URAL_MIXED_QTY[c])
        elif need01:
            k = 1
            for v in q01.values():
                if v > 0:
                    k = max(k, math.ceil(v / URAL_A01_ONLY_QTY))
            for c in a01:
                all_dels[c][di] = float(k * URAL_A01_ONLY_QTY)
        else:
            kits = URAL_A08_MIN_KITS
            for c, v in q08.items():
                if v > 0:
                    kits = max(kits, math.ceil(v / URAL_MIXED_QTY[c]))
            for c in a08:
                all_dels[c][di] = float(kits * URAL_MIXED_QTY[c])


def _lear_round_headrest_day_sum(all_dels):
    """Подголовники Lear: СУММА упаковок дня кратна 3 (таблица микса машины
    с пенами: 6-32, 9-30, 12-28 … 24-20; согласовано по эталону). Кратность
    по отдельным позициям не применяется. Добивка — коду с наибольшей
    суммарной потребностью."""
    hr = [c for c in mrp_codes
          if normalize_supplier(get_info(c)[1]) == LEAR_SUPPLIER
          and _lear_product_kind(c) == 'headrest' and c in all_dels]
    if not hr:
        return
    n = len(all_dates)
    pkgs = {c: _pkg_size(c) for c in hr}
    rem = {c: _code_3m_demand(c) for c in hr}
    for di in range(n):
        tot = sum(int(round(float(all_dels[c][di] or 0) / pkgs[c])) for c in hr
                  if all_dels[c][di] > 0)
        if tot <= 0 or tot % 3 == 0:
            continue
        need_add = 3 - (tot % 3)
        cand = sorted([c for c in hr if all_dels[c][di] > 0],
                      key=lambda x: -rem[x]) or hr
        for _ in range(need_add):
            c = cand[0]
            all_dels[c][di] = float(all_dels[c][di] or 0) + pkgs[c]

def _cap_smc_day_deliveries(raw_qty):
    """Лимит 7 паллет на фуру (сумма по всем позициям SMC в день)."""
    if not raw_qty:
        return raw_qty
    pkg_map = {c: max(1, bom.get(c, {}).get('package', 1)) for c in raw_qty}
    pallets = {c: _smc_pallet_count(c, raw_qty[c], pkg_map[c]) for c in raw_qty}
    total = sum(pallets.values())
    if total <= SMC_MAX_PALLETS_PER_TRUCK:
        return raw_qty
    scale = SMC_MAX_PALLETS_PER_TRUCK / total
    out = {}
    for c, qty in raw_qty.items():
        pkg = pkg_map[c]
        upp = _smc_units_per_pallet(c, pkg)
        allowed_pal = max(0, int(pallets[c] * scale))
        new_qty = allowed_pal * upp
        new_qty = (new_qty // pkg) * pkg if pkg else new_qty
        out[c] = new_qty
    return out

def _smooth_smc_trucks(all_dels, smc_codes):
    """SMC ≤ 7 паллет/день (строго). Избыток переносим на более ранние дни с потребностью."""
    if not smc_codes:
        return
    n = len(all_dates)
    SOFT = _soft_truck_cap(SMC_MAX_PALLETS_PER_TRUCK)
    pkgs = {c: max(1, bom.get(c, {}).get('package', 1)) for c in smc_codes}
    demday = {c: [d > 0 for d in _delivery_series_demand(c)] for c in smc_codes}

    def _day_pallets(di):
        return sum(
            _smc_pallet_count(c, all_dels[c][di], pkgs[c])
            for c in smc_codes if all_dels[c][di] > 0)

    load = [_day_pallets(di) for di in range(n)]
    for di in range(n - 1, -1, -1):
        guard = 0
        while load[di] > SOFT + 1e-9 and guard < 100000:
            guard += 1
            moved = False
            for c in smc_codes:
                pkg = pkgs[c]
                if all_dels[c][di] < pkg:
                    continue
                for k in range(di - 1, -1, -1):
                    if not demday[c][k]:
                        continue
                    all_dels[c][di] -= pkg
                    all_dels[c][k] += pkg
                    load[di] = _day_pallets(di)
                    load[k] = _day_pallets(k)
                    if load[k] <= SOFT + 1e-9:
                        moved = True
                        break
                    all_dels[c][di] += pkg
                    all_dels[c][k] -= pkg
                    load[di] = _day_pallets(di)
                    load[k] = _day_pallets(k)
                if moved:
                    break
            if not moved:
                raw = {c: all_dels[c][di] for c in smc_codes if all_dels[c][di] > 0}
                capped = _cap_smc_day_deliveries(raw)
                for c in smc_codes:
                    all_dels[c][di] = capped.get(c, 0.0)
                load[di] = _day_pallets(di)
                break

def _demand_day_delivery_qty(ss_prev, dem_d, safety, pkg):
    """Поставка только в день спроса: если после спроса остаток < страхового."""
    if dem_d <= 0:
        return 0.0
    proj = ss_prev - dem_d
    if proj >= safety:
        return 0.0
    need = safety - proj
    del_d = _ceiling_pkg_qty(max(0, need), pkg)
    return del_d if del_d > 0 else float(pkg)

def _next_working_days(dems):
    """Для каждого индекса — индекс СЛЕДУЮЩЕГО дня с потребностью (или None)."""
    n = len(dems)
    nwd = [None] * n
    _nd = None
    for di in range(n - 1, -1, -1):
        nwd[di] = _nd
        if dems[di] > 0:
            _nd = di
    return nwd

def _prev_working_days(dems):
    """Для каждого индекса — индекс ПРЕДЫДУЩЕГО дня с потребностью (или None)."""
    n = len(dems)
    pwd = [None] * n
    _pd = None
    for di in range(n):
        pwd[di] = _pd
        if dems[di] > 0:
            _pd = di
    return pwd

def _delivery_arrival_index(dems, ship_di):
    """День прихода отгрузки, отправленной в ship_di: следующий день потребности
    строго после ship_di (отгрузка дня D приходит в начале следующего дня)."""
    n = len(dems)
    for x in range(ship_di + 1, n):
        if dems[x] > 0:
            return x
    return None

def _next_demand_after(dems, di):
    for j in range(di + 1, len(dems)):
        if dems[j] > 0:
            return j
    return None

def _ndkt_transition_target_max(target):
    """Верхняя граница переходящего остатка (+20%)."""
    return float(target) * (1.0 + NDKT_SS_TOLERANCE_UP)

def _ndkt_anchor_snap_di(code):
    """Последний день с ручным снимком остатка (якорь графика) или None."""
    snap_map = manual_stock_by_date.get(code, {})
    if not snap_map:
        return None
    best = None
    for di in range(len(all_dates)):
        if snap_map.get(all_dates[di][0]) is not None:
            best = di
    return best

_ndkt_tab_batches_cache = {}

def _ndkt_tab_batches(tab):
    """Партии вкладки за весь горизонт, слитые по нормализованному id:
    [{'bvn', 'days': {gdi: (mn, day_key, bv_key)}, 'first': первый день}]."""
    if tab in _ndkt_tab_batches_cache:
        return _ndkt_tab_batches_cache[tab]
    merged = {}
    for mn, _, _ in MONTHS:
        for bv, day_qty in plan_batches.get(tab, {}).get(mn, {}).items():
            bvn = _normalize_batch_id(bv)
            for d, cars in (day_qty or {}).items():
                try:
                    dd = int(d)
                except (TypeError, ValueError):
                    continue
                if not cars or (mn, dd) not in _GLOBAL_DI:
                    continue
                merged.setdefault(bvn, {})[_GLOBAL_DI[(mn, dd)]] = (mn, d, bv)
    out = [{'bvn': bvn, 'days': days, 'first': min(days)}
           for bvn, days in merged.items() if days]
    _ndkt_tab_batches_cache[tab] = out
    return out

def _ndkt_batch_part_demand_on_day(code, tab, mn, bv, day_key):
    """Потребность code в конкретной партии на конкретный день (машины × прим.)."""
    bi = _get_batch_info(tab, bv)
    cfg_key = f"{bi.get('model','')}_{bi.get('drive','')}_{bi.get('config','')}"
    applicable = bom.get(code, {}).get('configs', {})
    if not applicable:
        return 0.0
    matched, q = cfg_match(cfg_key, applicable)
    if not matched:
        return 0.0
    cars = plan_batches.get(tab, {}).get(mn, {}).get(bv, {}).get(day_key, 0) or 0
    return float(cars) * float(q)

_ndkt_targets_cache = {}

def _ndkt_morning_targets(code):
    """Цель остатка на НАЧАЛО дня di (длина n+1; [n] — за горизонтом):
    • 100% потребности code в ПЕРВОЙ производственной партии дня di
      (партия, начатая раньше всех среди идущих в день di на вкладке детали;
      по обеим вкладкам AS_in_F_A / AS_in_H_B суммарно), но не менее 80 шт;
    • для остальных колёс (нет доли в первых партиях дня) — 80 шт (+20% допуск).
    Потребность поздних партий дня закрывается поставкой текущего дня."""
    if code in _ndkt_targets_cache:
        return _ndkt_targets_cache[code]
    n = len(all_dates)
    base = _ndkt_safety_base(code)
    fb_share = [0.0] * (n + 1)
    for tab in _ndkt_part_tabs(code):
        batches = _ndkt_tab_batches(tab)
        for di in range(n):
            today = [b for b in batches if di in b['days']]
            if not today:
                continue
            fb = min(today, key=lambda b: (b['first'], b['bvn']))
            mn, day_key, bv = fb['days'][di]
            fb_share[di] += _ndkt_batch_part_demand_on_day(code, tab, mn, bv, day_key)
    # Цель утра обязана покрыть ВЕСЬ спрос дня, а не только первую партию.
    # Прежняя версия закрывала лишь первую партию, считая, что поздние партии
    # закроет поставка того же дня. Но Del[d] приходит в Ss[d+1], то есть
    # ВНУТРИ дня d её ещё нет — отсюда провалы и минусы (128 дн×код на эталоне).
    _dems_full = _delivery_series_demand(code)
    targets = [max(fb_share[di], base,
                   float(_dems_full[di]) if di < n else 0.0)
               for di in range(n + 1)]
    _ndkt_targets_cache[code] = targets
    return targets

def _ndkt_apply_delivery_schedule(code, supplier, dems, dels, frozen_del=None):
    """NDKT: 📦 ставится в день потребности (или канун первой партии) и равна
    потребности ТЕКУЩЕГО дня + добор до цели утра СЛЕДУЮЩЕГО дня (кратно упаковке):
        Del[d] = potr[d] + цель_утра[d+1] − Ss[d],   Del[d] → Ss[d+1].
    Цель утра: 100% первой производственной партии дня (своя вкладка),
    остальные колёса — 80 шт (+20% допуск). Излишек от кратности упаковки
    гасится пропуском/уменьшением следующих поставок. История до последнего
    ручного снимка не корректируется."""
    frozen_del = frozen_del or set()
    n = len(dems)
    pkg = _pkg_size(code)
    snap_map = manual_stock_by_date.get(code, {})
    anchor = _ndkt_anchor_snap_di(code)
    start = anchor if anchor is not None else 0     # Del[якорь] → Ss[якорь+1] действует
    targets = _ndkt_morning_targets(code)
    has_demand = _code_3m_demand(code) > 0
    prev_ss = float(_opening_stock_for_ss(code))
    for di in range(n):
        dt = all_dates[di][0]
        snap = snap_map.get(dt)
        morning = float(snap) if snap is not None else prev_ss
        if di >= start and di not in frozen_del:
            dels[di] = 0.0
            next_snapped = (di + 1 < n and
                            snap_map.get(all_dates[di + 1][0]) is not None)
            allowed = dems[di] > 0 or (di + 1 < n and dems[di + 1] > 0)
            if has_demand and allowed and not next_snapped:
                need = float(dems[di]) + targets[di + 1] - morning
                if need > 1e-6:
                    dels[di] = float(math.ceil(need / pkg) * pkg)
        prev_ss = morning - float(dems[di]) + float(dels[di] or 0)

def _build_ndkt_deliveries(code, supplier):
    dems = _delivery_series_demand(code)
    dels = [0.0] * len(dems)
    _ndkt_apply_delivery_schedule(code, supplier, dems, dels)
    return dels

def _next_delivery_after(dels, di):
    """Индекс следующей отгрузки строго после di (или None)."""
    for j in range(di + 1, len(dels)):
        if float(dels[j] or 0) > 0:
            return j
    return None


def _target_until_next_delivery(dems, dels, di, safety):
    """Остаток на конец дня di должен покрыть спрос до ПРИХОДА следующей поставки.

    Отгрузка дня j приходит в начале дня j+1, поэтому покрыть надо дни
    di+1 … j. Раньше цель считалась только до следующего дня потребности
    (_target_at_next_demand) — для поставщиков со слотовым графиком (неделя,
    полнедели, месяц) этого не хватает, отсюда провалы и минусы в остатках.

    Возвращает None, если следующей поставки нет (хвост горизонта закрывает
    _ensure_non_negative_balances).
    """
    nxt = _next_delivery_after(dels, di)
    if nxt is None:
        return None
    cover = sum(dems[di + 1:nxt + 1])
    # Допуск применяем ТОЛЬКО к страховому запасу: в него просесть можно,
    # недовезти фактический спрос — нельзя.
    return safety * (1.0 - DELIVERY_SS_TARGET_TOLERANCE) + cover


def _target_at_next_demand(dems, di, safety, code=None, supplier=None):
    """Целевой остаток к началу следующего дня потребности после di."""
    nd = _next_demand_after(dems, di)
    if nd is None:
        return safety
    nd_dem = dems[nd]
    return safety + max(nd_dem, math.ceil(DELIVERY_START_COVER * nd_dem))

def _delivery_qty_to_target(code, supplier, qty, safety_target):
    """Объём поставки до цели (для не-NDKT)."""
    if qty <= 0:
        return 0
    return _round_delivery_qty(code, qty)

def _nearest_demand_day(dems, from_di, prefer_before=True):
    before = after = None
    for j in range(from_di, -1, -1):
        if dems[j] > 0:
            before = j
            break
    for j in range(from_di + 1, len(dems)):
        if dems[j] > 0:
            after = j
            break
    if prefer_before and before is not None:
        return before
    if after is not None:
        return after
    return before

def _enforce_no_delivery_without_demand(dems, dels):
    """📦 только в РЕЙСОВЫЕ дни (не вс/праздники) с potr>0. Поставку, попавшую на
    нерейсовый день или день без спроса, переносим на ближайший предыдущий рейсовый
    день со спросом (иначе — на ближайший предыдущий рейсовый день)."""
    n = len(dels)
    for di in range(n):
        q = float(dels[di] or 0)
        if q <= 0:
            continue
        if dems[di] > 0 and _is_ru_ship_day(all_dates[di][0]):
            continue
        tgt = None
        for k in range(di - 1, -1, -1):
            if dems[k] > 0 and _is_ru_ship_day(all_dates[k][0]):
                tgt = k
                break
        if tgt is None:
            for k in range(di - 1, -1, -1):
                if _is_ru_ship_day(all_dates[k][0]):
                    tgt = k
                    break
        if tgt is not None:
            dels[tgt] = float(dels[tgt] or 0) + q
            dels[di] = 0.0

def _week_key_for_di(di):
    dt = all_dates[di][0]
    iso = dt.isocalendar()
    return (iso[0], iso[1])

def _consolidate_deliveries_by_period(code, dems, dels, pkg, slot_map, period_key_fn):
    """Суммирует поставки периода в слот (день потребности до/после слота)."""
    totals = {}
    for di, q in enumerate(dels):
        if q > 0 and dems[di] > 0:
            pk = period_key_fn(di)
            totals[pk] = totals.get(pk, 0.0) + q
    for di in range(len(dels)):
        dels[di] = 0.0
    for pk, total in totals.items():
        if total <= 0:
            continue
        slots = slot_map.get(pk, [])
        target = None
        for di in slots:
            if dems[di] > 0:
                target = di
                break
        if target is None and slots:
            target = _nearest_demand_day(dems, slots[0], prefer_before=True)
        if target is None:
            for di in range(len(dems)):
                if period_key_fn(di) == pk and dems[di] > 0:
                    target = di
                    break
        if target is not None:
            dels[target] += _round_delivery_qty(code, total, pkg)

def _consolidate_to_supplier_slots(code, supplier, dems, dels):
    """Календарь поставщика: объём в день слота или ближайший день потребности до/после."""
    mode = _delivery_mode(supplier)
    if mode == 'lookahead':
        return
    pkg = max(1, bom.get(code, {}).get('package', 1))
    if mode == 'monthly_early':
        slot_map = {}
        for di in _monthly_early_delivery_indices():
            mn = all_dates[di][1]
            slot_map.setdefault(mn, []).append(di)
        _consolidate_deliveries_by_period(code, dems, dels, pkg, slot_map, lambda di: all_dates[di][1])
    elif mode == 'monthly_lag':
        slot_map = {}
        for di in _ecotexis_delivery_indices():
            mn = all_dates[di][1]
            slot_map.setdefault(mn, []).append(di)
        _consolidate_deliveries_by_period(code, dems, dels, pkg, slot_map, lambda di: all_dates[di][1])
    elif mode in ('smc_weekly', 'weekly', 'avtokom_weekly'):
        wd = (SMC_WEEKDAY if mode == 'smc_weekly'
              else 1 if mode == 'avtokom_weekly'   # вторник (эталон: 09/16/23/30.06)
              else PUREM_WEEKDAY)
        slot_map = {}
        for di in _weekday_delivery_indices(wd):
            slot_map.setdefault(_week_key_for_di(di), []).append(di)
        _consolidate_deliveries_by_period(code, dems, dels, pkg, slot_map, _week_key_for_di)
    elif mode == 'twice_weekly':
        # МТС-авто: 2 слота в неделю (MTS_WEEKDAYS). Период = половина недели.
        def _tw_key(di):
            wk = _week_key_for_di(di)
            return (wk, 0 if all_dates[di][0].weekday() < MTS_WEEKDAYS[1] else 1)
        slot_map = {}
        for di, (dt, _, _) in enumerate(all_dates):
            if dt.weekday() in MTS_WEEKDAYS:
                slot_map.setdefault(_tw_key(di), []).append(di)
        _consolidate_deliveries_by_period(code, dems, dels, pkg, slot_map, _tw_key)
    elif mode == 'semimonthly':
        # Ителма: 2 отгрузки в месяц — первый день потребности каждой половины
        def _half_key(di):
            return (all_dates[di][1], 0 if all_dates[di][2] < 16 else 1)
        slot_map = {}
        for di in range(len(all_dates)):
            hk = _half_key(di)
            if hk not in slot_map:
                slot_map[hk] = [di]
        _consolidate_deliveries_by_period(code, dems, dels, pkg, slot_map, _half_key)

def _reconcile_demand_day_deliveries(code, supplier, dems, dels, frozen_del=None):
    """Донабор 📦 только в дни потребности — целевой остаток к следующему дню спроса.
    NDKT не трогаем: его график строит ТОЛЬКО _ndkt_rebuild_truck_schedule
    (целые фуры с миксом колёс) — иначе донабор ломает фурование."""
    if _is_ndkt_supplier(supplier) or _is_ndkt_part(code):
        return
    if code == MSA_FULL_TRUCK_CODE:
        # 5304100XKN02A: пересборка целыми фурами 528/576, 1–2 раза в неделю
        _rebuild_msa_full_truck_deliveries(code, dems, dels, frozen_del)
        return
    safety = _delivery_safety_qty(code, supplier)
    n = len(dems)
    ss_path = _compute_ss_path(code, dels)
    pkg = max(1, bom.get(code, {}).get('package', 1))
    for di in range(n):
        ss = ss_path[di]
        if dems[di] <= 0:
            continue
        # Del[di] приходит в Ss[di+1]; снимок дня di+1 делает её бесполезной
        if (di + 1 < n and
                manual_stock_by_date.get(code, {}).get(all_dates[di + 1][0]) is not None):
            continue
        target = _soft_ss_target(
            _target_at_next_demand(dems, di, safety, code=code, supplier=supplier),
            code=code, supplier=supplier)
        # Покрытие до прихода СЛЕДУЮЩЕЙ поставки (важно для слотовых графиков)
        _pt = _target_until_next_delivery(dems, dels, di, safety)
        if _pt is not None and _pt > target:
            target = _pt
        projected = ss - dems[di] + float(dels[di] or 0)
        if projected < target:
            dels[di] += _delivery_qty_to_target(
                code, supplier, max(0, target - projected), safety)
    _enforce_no_delivery_without_demand(dems, dels)

def _build_initial_deliveries(code, supplier):
    """Del[di] только если potr[di]>0. Объём закрывает спрос до СЛЕДУЮЩЕГО дня потребности
    (в т.ч. через дни без потребности — поставка до/после, но не В эти дни)."""
    dems = _delivery_series_demand(code)
    safety = _delivery_safety_qty(code, supplier)
    n = len(dems)
    dels = [0.0] * n
    ss_prev = float(_opening_stock_for_ss(code))
    for di in range(n):
        dt = all_dates[di][0]
        if di > 0:
            ss_start = ss_prev - dems[di - 1] + float(dels[di - 1] or 0)
        else:
            ss_start = ss_prev
        snap = manual_stock_by_date.get(code, {}).get(dt)
        if snap is not None:
            ss = float(snap)
        else:
            ss = ss_start
        # Del[di] приходит в Ss[di+1]; если di+1 закрыт ручным снимком — поставка бесполезна
        _next_snapped = (di + 1 < n and
                         manual_stock_by_date.get(code, {}).get(all_dates[di + 1][0]) is not None)
        if dems[di] > 0 and not _next_snapped:
            pkg = max(1, bom.get(code, {}).get('package', 1))
            target = _soft_ss_target(
                _target_at_next_demand(dems, di, safety, code=code, supplier=supplier),
                code=code, supplier=supplier)
            projected = ss - dems[di] + float(dels[di] or 0)
            if projected < target:
                dels[di] += _delivery_qty_to_target(
                    code, supplier, max(0, target - projected), safety)
        ss_prev = ss
    _enforce_no_delivery_without_demand(dems, dels)
    return dels

def _ss_after_day(ss_prev, dem_d, del_d, code, dt):
    """Остаток после дня: ручной снимок всегда в приоритете, иначе prev − спрос + поставка."""
    snap = manual_stock_by_date.get(code, {}).get(dt)
    if snap is not None:
        return float(snap)
    return ss_prev - dem_d + del_d

def _delivery_arrival_qty(dems, dels, di, pwd=None):
    """Приход в день di = поставка предыдущего календарного дня (Del[di-1]), в т.ч. 0."""
    if di <= 0:
        return 0.0
    return float(dels[di - 1] or 0)

def _compute_ss_path_raw(code, dels):
    """Цепочка Ss без снимков (для решений по Del — снимки не блокируют эффект поставки)."""
    dems = _delivery_series_demand(code)
    ss_prev = float(_opening_stock_for_ss(code))
    path = []
    for di in range(len(dems)):
        dem_prev = dems[di - 1] if di > 0 else 0
        ss_d = ss_prev - dem_prev + _delivery_arrival_qty(dems, dels, di)
        path.append(ss_d)
        ss_prev = ss_d
    return path

def _compute_ss_path(code, dels):
    """Остаток на начало каждого дня (как Ss в График_Поставок: Del[d-1] → Ss[d]).
    Ручной снимок Ввод_Остатков ВСЕГДА подменяет Ss своего дня (даже при дефиците)."""
    dems = _delivery_series_demand(code)
    n = len(dems)
    ss_prev = float(_opening_stock_for_ss(code))
    path = []
    for di in range(n):
        dt = all_dates[di][0]
        dem_prev = dems[di - 1] if di > 0 else 0
        computed = ss_prev - dem_prev + _delivery_arrival_qty(dems, dels, di)
        snap = manual_stock_by_date.get(code, {}).get(dt)
        if snap is not None:
            ss_d = float(snap)
        else:
            ss_d = computed
        path.append(ss_d)
        ss_prev = ss_d
    return path

def _excel_ss_path(code, dels):
    """Тот же Ss, что пересчитывает Excel (G + цепочка Del/potr + снимки)."""
    return _compute_ss_path(code, dels)

def _count_excel_sim_negatives(all_dels):
    """Дефициты по логике формул График_Поставок (не внутренний boost).
    NDKT: потребность дня закрывается поставкой ТЕКУЩЕГО дня (приход в течение
    дня), поэтому остаток КОНЦА дня сравниваем со страховым минимумом, а
    покрытие дня = утро + Del ≥ potr."""
    n_neg = n_below = 0
    for code in mrp_codes:
        dems = _delivery_series_demand(code)
        dels = all_dels.get(code, [])
        is_ndkt = _is_ndkt_part(code)
        prev = float(_opening_stock_for_ss(code))
        for di, ss in enumerate(_excel_ss_path(code, dels)):
            if ss + 1e-6 < 0:
                n_neg += 1
            elif dems[di] > 0:
                if is_ndkt:
                    morning = prev
                    del_d = float(dels[di] or 0) if di < len(dels) else 0.0
                    if morning + del_d + 1e-6 < dems[di]:
                        n_below += 1
                elif ss + 1e-6 < dems[di]:
                    n_below += 1
            prev = ss
    return n_neg, n_below

def _count_negative_ss(code, dels):
    dems = _delivery_series_demand(code)
    is_ndkt = _is_ndkt_part(code)
    n = 0
    prev = float(_opening_stock_for_ss(code))
    for di, ss in enumerate(_compute_ss_path(code, dels)):
        if ss + 1e-6 < 0:
            n += 1
        elif dems[di] > 0:
            if is_ndkt:
                del_d = float(dels[di] or 0) if di < len(dels) else 0.0
                if prev + del_d + 1e-6 < dems[di]:
                    n += 1
            elif ss + 1e-6 < dems[di]:
                n += 1
        prev = ss
    return n

def _ensure_ndkt_balances(code, supplier, dems, dels, frozen_del=None):
    """NDKT: пересчёт Del (raw Ss) + trim лишних 📦."""
    _ndkt_apply_delivery_schedule(code, supplier, dems, dels, frozen_del)

def _ndkt_first_batch_morning_targets(code):
    """Цель остатка на УТРО дня (длина n+1): 100% потребности code в ПЕРВЫХ
    партиях плана — 1-я партия AS_in_F_A и 1-я партия AS_in_H_B (по их дням).
    Эти партии закрываются из переходящих остатков В ЛЮБОМ СЛУЧАЕ
    (поставка текущего дня не в счёт)."""
    n = len(all_dates)
    targets = [0.0] * (n + 1)
    for tab in NDKT_AS_TABS:
        batches = _ndkt_tab_batches(tab)
        if not batches:
            continue
        fb = min(batches, key=lambda b: (b['first'], b['bvn']))
        for di, (mn, day_key, bv) in fb['days'].items():
            dem = _ndkt_batch_part_demand_on_day(code, tab, mn, bv, day_key)
            if dem > 0 and di < n:
                targets[di] += dem
    return targets

def _ndkt_enforce_first_batches(all_dels, frozen_del_by_code=None):
    """Первые 2 партии плана (1-я AS_in_F_A и 1-я AS_in_H_B) закрываются на 100%
    из переходящих остатков: на утро дня партии Ss ≥ потребности кода в партии.
    Добор — предзавозом на ближайший «выживающий» день накануне."""
    frozen_del_by_code = frozen_del_by_code or {}
    fixed = 0
    for code in mrp_codes:
        if not _is_ndkt_part(code) or _code_3m_demand(code) <= 0:
            continue
        targets = _ndkt_first_batch_morning_targets(code)
        if not any(t > 0 for t in targets):
            continue
        dels = all_dels.get(code)
        if dels is None:
            continue
        dems = _delivery_series_demand(code)
        frozen = frozen_del_by_code.get(code, set())
        pkg = _pkg_size(code)
        snap_map = manual_stock_by_date.get(code, {})
        n = len(all_dates)
        for di in range(1, n):
            t = targets[di] if di < len(targets) else 0.0
            if t <= 0:
                continue
            ss = _compute_ss_path(code, dels)
            morning = ss[di - 1]
            if morning + 1e-6 >= t:
                continue
            # день добора: канун партии (предзавоз) или ближайший более ранний
            # день, который переживает очистку и не упирается в ручной снимок
            j = di - 1
            while j >= 0:
                survives = (dems[j] > 0 or (j + 1 < n and dems[j + 1] > 0))
                next_snapped = (j + 1 < n
                                and snap_map.get(all_dates[j + 1][0]) is not None)
                if j not in frozen and survives and not next_snapped:
                    break
                j -= 1
            if j < 0:
                continue
            need = t - morning
            dels[j] += float(_round_delivery_qty(code, need, pkg))
            fixed += 1
    if fixed:
        print(f"    NDKT: добор под 100% первых партий плана (AS_in_F_A/AS_in_H_B): "
              f"{fixed} поставок", flush=True)

def _ndkt_rebuild_truck_schedule(all_dels, frozen_del_by_code=None, quiet=False):
    """NDKT: пересборка дневных поставок (минимальные запасы, микс колёс).

    Шаг 1 — объём дня (по каждому колесу): ТОЛЬКО в день потребности —
      спрос текущего дня + добор до страхового запаса к концу дня
      (240/480 по коду), кратно упаковке (стопка 24 / бокс 30); накануне дня
      ПЕРВОЙ партии плана (AS_in_F_A / AS_in_H_B) — минимум 100% её
      потребности из остатков. Будущие дни НЕ предвозим и машины впрок
      НЕ добиваем — утро следующего дня ≈ страховому запасу.
    Шаг 2 — фурование (подсчёт машин): загрузка дня = Σ qty/вместимость
      (XST33A 540, прочие 384, микс по формуле x/540 + y/384 ≤ 1 на машину);
      машин в день = ceil(загрузки), последняя машина может ехать неполной.
    Ручные 📦 (frozen) не меняются и учитываются в загрузке их дня."""
    frozen_del_by_code = frozen_del_by_code or {}
    codes = [c for c in mrp_codes if _is_ndkt_part(c) and _code_3m_demand(c) > 0]
    if not codes:
        return
    n = len(all_dates)
    dems = {c: _delivery_series_demand(c) for c in codes}
    pkgs = {c: _pkg_size(c) for c in codes}
    saf = {c: _safety_qty(c, NDKT_SUPPLIER) for c in codes}
    caps = {c: _ndkt_truck_cap(c) for c in codes}
    fb_t = {c: _ndkt_first_batch_morning_targets(c) for c in codes}
    snaps = {c: manual_stock_by_date.get(c, {}) for c in codes}
    ss = {c: float(_opening_stock_for_ss(c)) for c in codes}
    nxt = {c: [None] * n for c in codes}
    for c in codes:
        _nd = None
        for di in range(n - 1, -1, -1):
            nxt[c][di] = _nd
            if dems[c][di] > 0:
                _nd = di
    rem_future = {c: [0.0] * (n + 1) for c in codes}
    for c in codes:
        for di in range(n - 1, -1, -1):
            rem_future[c][di] = rem_future[c][di + 1] + dems[c][di]
    # Рейсовые дни: без воскресений и праздников РФ (согласовано по эталону НДКТ)
    workday = [_is_ru_ship_day(all_dates[di][0]) for di in range(n)]

    def _span_extra(c, di):
        """Потребность нерейсовых дней сразу после di (предзавоз выходных)."""
        extra = 0.0
        j = di + 1
        while j < n and not workday[j]:
            extra += dems[c][j]
            j += 1
        return extra

    for di in range(n):
        dt = all_dates[di][0]
        # Ручной снимок Ввод_Остатков = остаток на УТРО дня (синхронно с _compute_ss_path)
        for c in codes:
            _snap_v = snaps[c].get(dt)
            if _snap_v is not None:
                ss[c] = float(_snap_v)
        frozen_today = {c for c in codes if di in frozen_del_by_code.get(c, set())}
        add = {}
        for c in codes:
            if c in frozen_today:
                continue
            all_dels[c][di] = 0.0
            if not workday[di]:
                continue   # вс/праздник: рейсов нет — потребление из остатков
            morning = ss[c]
            proj = morning - dems[c][di]
            need = 0.0
            extra = _span_extra(c, di)
            if dems[c][di] > 0 or extra > 0:
                # Поставка в день потребности (приход в течение дня): спрос дня
                # + спрос ближайших нерейсовых дней (предзавоз вс/праздников)
                # + выход на страховой запас. Будущие рейсовые дни не предвозим.
                need = max(need, dems[c][di] + extra + saf[c] - morning)
            # канун первой партии плана: утро завтра ≥ 100% её потребности
            if (di + 1 < n and di + 1 < len(fb_t[c]) and fb_t[c][di + 1] > 0
                    and snaps[c].get(all_dates[di + 1][0]) is None):
                need = max(need, fb_t[c][di + 1] - proj)
            # день без потребности, без выходных впереди и без кануна партии — не возим
            if (dems[c][di] <= 0 and extra <= 0 and not (
                    di + 1 < n and di + 1 < len(fb_t[c]) and fb_t[c][di + 1] > 0)):
                need = 0.0
            # поставка бесполезна, если Ss завтра зафиксирован ручным снимком
            if (di + 1 < n and snaps[c].get(all_dates[di + 1][0]) is not None
                    and dems[c][di] <= 0):
                need = 0.0
            if need > 1e-6:
                add[c] = float(_ceiling_pkg_qty(need, pkgs[c]))
        for c, q in add.items():
            all_dels[c][di] = q
        for c in codes:
            # конец дня = утро − спрос + поставка дня (снимок утра уже применён)
            ss[c] = ss[c] - dems[c][di] + float(all_dels[c][di] or 0)

    # ── Лимит машин в день (согласовано: max 8) — пики предвозятся назад ──
    def _day_load(di):
        return sum(float(all_dels[c][di] or 0) / caps[c] for c in codes)
    for di in range(n - 1, -1, -1):
        guard = 0
        while _day_load(di) > NDKT_MAX_TRUCKS_PER_DAY + 1e-9 and guard < 5000:
            guard += 1
            moved = False
            for c in sorted(codes, key=lambda x: -float(all_dels[x][di] or 0)):
                if di in frozen_del_by_code.get(c, set()):
                    continue
                if float(all_dels[c][di] or 0) < pkgs[c]:
                    continue
                share = pkgs[c] / caps[c]
                for k in range(di - 1, -1, -1):
                    if not workday[k] or k in frozen_del_by_code.get(c, set()):
                        continue
                    if _day_load(k) + share <= NDKT_MAX_TRUCKS_PER_DAY + 1e-9:
                        all_dels[c][di] -= pkgs[c]
                        all_dels[c][k] += pkgs[c]
                        moved = True
                        break
                if moved:
                    break
            if not moved:
                break   # некуда переносить — день остаётся пиковым

    n_days_ship = n_trucks = n_partial = 0
    for di in range(n):
        L = _day_load(di)
        if L > 1e-9:
            trucks = math.ceil(L - 1e-9)
            n_days_ship += 1
            n_trucks += trucks
            if trucks - L > 1e-6:
                n_partial += 1
    if not quiet and n_days_ship:
        print(f"    NDKT фурование: {n_days_ship} дн. поставок, {n_trucks} машин "
              f"(лимит {NDKT_MAX_TRUCKS_PER_DAY}/день, без вс/праздников), "
              f"неполных: {n_partial}", flush=True)

def _ensure_non_negative_balances(code, supplier, dems, dels, frozen_del=None, msa_load=None):
    """Донабор 📦: Ss[d] ≥ potr[d] и Ss[d] ≥ 0 (кратно упаковке); сверху, где
    остаётся вместимость дня, добор до страхового запаса (min_ss = потр + страх).

    Поставку добавляем ТОЛЬКО на день, который переживёт очистку
    _enforce_no_delivery_without_demand (день с потребностью или канун дня
    потребности — предзавоз), иначе донабор тут же стирается и дефицит
    возвращается. Дни с ручным «снимком» остатка (manual_stock_by_date)
    пропускаем: их Ss зафиксирован вручную и поставками не меняется — иначе
    донабор зацикливается и раздувает 📦 предыдущего дня до абсурда.

    Базовая потребность (min_ss = потр) добирается ВСЕГДА, безусловно — это
    реальный спрос, а не запас прочности. Страховая надбавка сверх потребности
    (см. коммит 3ef938e) — необязательная, и для бамперов MSA (msa_load, места
    в фуре 21/день, общие на всех кодов-бамперов сразу) ограничена свободной
    вместимостью дня отгрузки: раньше надбавка добавлялась вслепую, а
    _enforce_truck_limits потом резал её вместе с урезанием превышения — за
    15 итераций это раздувало перегруз фуры на порядок (макс +1190% → +4743%
    на эталоне), потому что каждый цикл добавлял надбавку заново на тот же
    день. Теперь донабор сам не просит больше, чем есть места, — резать
    вообще нечего."""
    frozen_del = frozen_del or set()
    if _is_ndkt_supplier(supplier) or _is_ndkt_part(code):
        return   # NDKT: график строит только фурование (_ndkt_rebuild_truck_schedule)
    pkg = _pkg_size(code)
    n = len(dems)
    snap_map = manual_stock_by_date.get(code, {})
    _safety_floor = _delivery_safety_qty(code, supplier) * (
        1.0 - DELIVERY_SS_TARGET_TOLERANCE)

    _is_eftec = (normalize_supplier(supplier) == EFTEC_SUPPLIER)

    def _survivable_ship_day(j):
        # Поставка дня j сохраняется, если j — день потребности либо канун
        # дня потребности (предзавоз Del[j] → Ss[j+1]).
        # Del[j] бесполезна, если Ss[j+1] зафиксирован ручным снимком.
        if _is_eftec and all_dates[j][0].weekday() in EFTEC_FORBIDDEN_WEEKDAYS:
            return False   # иначе донабор и перенос с пн/вс гоняют друг друга
        if j + 1 < n and snap_map.get(all_dates[j + 1][0]) is not None:
            return False
        return dems[j] > 0 or (j + 1 < n and dems[j + 1] > 0)

    # Ближайший ПРЕДЫДУЩИЙ день, на котором поставка выживет — считаем один раз.
    prev_ship = [None] * n
    _last = None
    for j in range(n):
        prev_ship[j] = _last
        if j not in frozen_del and _survivable_ship_day(j):
            _last = j

    # Раньше здесь было n*64 итераций, и КАЖДАЯ пересчитывала весь путь Ss
    # ради одного исправления (_compute_ss_path → 5888 полных проходов на код).
    # После поднятия порога до страхового запаса это стало основным тормозом.
    # Теперь: один проход вперёд с накопителем добавок (добавка на ранний день
    # только поднимает Ss всех последующих, поэтому назад возвращаться не нужно).
    for _ in range(8):
        ss = _compute_ss_path(code, dels)
        fixed = 0
        bump = 0.0
        for di in range(n):
            if snap_map.get(all_dates[di][0]) is not None:
                bump = 0.0     # ручной снимок подменяет Ss — накопитель сбрасываем
                continue
            cur = ss[di] + bump
            base_min = max(0.0, float(dems[di]))
            min_ss = base_min + _safety_floor
            if cur + 1e-6 >= min_ss:
                continue
            ship_di = prev_ship[di]
            if ship_di is None:
                continue
            need_base = max(0.0, base_min - cur)
            need_extra = max(0.0, min_ss - cur) - need_base   # страховая надбавка
            if msa_load is not None and need_extra > 1e-9:
                free_slots = MSA_MAX_CONTAINERS_PER_TRUCK - msa_load[ship_di]
                need_extra = 0.0 if free_slots <= 0 else min(need_extra, free_slots * pkg)
            add_raw = need_base + need_extra
            if add_raw <= 1e-9:
                continue
            add = _round_delivery_qty(code, add_raw, pkg)
            if add <= 0:
                add = float(pkg)
            dels[ship_di] = float(dels[ship_di] or 0) + add
            bump += add
            if msa_load is not None:
                msa_load[ship_di] += int(round(add / pkg))
            fixed += 1
        if not fixed:
            break
    _enforce_no_delivery_without_demand(dems, dels)

def _recompute_ss_from_deliveries(code, dels):
    """Остаток на начало дня D = остаток D-1 − спрос D-1 + Del D-1 (как в График_Поставок)."""
    dems = _delivery_series_demand(code)
    ss_path = _compute_ss_path(code, dels)
    return [(dels[di], ss_path[di]) for di in range(len(dems))]

def _msa_trim_day_slots(front_qty_by_code, rear_qty_by_code):
    """Места стальной тары под накладки. Комплект пары: 480+480 передних
    (960 шт) = 3 места → 320 шт/место; 180+180 задних (360 шт) = 3 места
    → 120 шт/место. Пропорционально объёму пары (кратность задних — 60)."""
    front_pcs = sum(q for q in front_qty_by_code.values() if q > 0)
    rear_pcs = sum(q for q in rear_qty_by_code.values() if q > 0)
    front_slots = math.ceil(front_pcs / 320) if front_pcs > 0 else 0
    rear_slots = math.ceil(rear_pcs / 120) if rear_pcs > 0 else 0
    return front_slots, rear_slots

def _msa_container_slots(raw_qty):
    """Тарные места (стальная тара) дневной отгрузки MSA:
    бамперы — 1 упаковка = 1 стальная тара; накладки — парные комплекты
    (480+480 или 180+180) по 3 стальных места за комплект.
    5304100XKN02A не учитывается — едет отдельной фурой (528/576)."""
    bumper_slots = 0
    front_q = {}
    rear_q = {}
    for c, q in raw_qty.items():
        if q <= 0 or c == MSA_FULL_TRUCK_CODE:
            continue
        if c in MSA_FRONT_TRIM_CODES:
            front_q[c] = q
        elif c in MSA_REAR_TRIM_CODES:
            rear_q[c] = q
        else:
            pkg = max(1, bom.get(c, {}).get('package', 1))
            bumper_slots += math.ceil(q / pkg)
    front_slots, rear_slots = _msa_trim_day_slots(front_q, rear_q)
    return bumper_slots, front_slots, rear_slots

def _cap_msa_day_deliveries(raw_qty):
    """Ограничение фуры MSA: ≤ 21 контейнер/день. Накладки сохраняются (резервируют
    места), бамперы при нехватке мест уменьшаются (избыток уходит в переходящий запас)."""
    if not raw_qty:
        return raw_qty
    bs, fs, rs = _msa_container_slots(raw_qty)
    if bs + fs + rs <= MSA_MAX_CONTAINERS_PER_TRUCK:
        return raw_qty
    avail = max(0, MSA_MAX_CONTAINERS_PER_TRUCK - (fs + rs))
    if bs <= 0:
        return raw_qty  # только накладки — бамперы не уменьшить
    scale = avail / bs
    out = {}
    for c, q in raw_qty.items():
        if c in MSA_FRONT_TRIM_CODES or c in MSA_REAR_TRIM_CODES or c == MSA_FULL_TRUCK_CODE:
            out[c] = q
        else:
            pkg = max(1, bom.get(c, {}).get('package', 1))
            cont = int(math.floor((q / pkg) * scale))
            out[c] = cont * pkg
    return out

# ── Переходящий запас комплектно: пары передний↔задний (модель/конфиг/цвет) ──
# Поставки пары планируются ВМЕСТЕ, чтобы остаток переднего ≈ остатку заднего.
_MSA_PAIR_PREFIX = [
    ('2803120XST33', '2804104AST33'),   # A01 comfort/elite/premium
    ('2803130XST33', '2804105AST33'),   # A01 Tech Plus
    ('2803104XKN61', '2804KN260004'),   # B02 elite
    ('2803105XKN61', '2804KN260005'),   # B02/B04 premium/TechPlus
]
def _msa_pairs():
    out = []
    for _fp, _rp in _MSA_PAIR_PREFIX:
        for _suf in _BUMPER_COLOR_SUFFIX:
            _f, _r = _fp + _suf, _rp + _suf
            if _f in bom and _r in bom:
                out.append((_f, _r))
    return out

def _build_paired_deliveries(fcode, rcode, fsupp, rsupp):
    """Совместный график поставок пары: доставляем обе стороны в один день,
    каждую до её страхового запаса, чтобы переходящий остаток был комплектным."""
    dems_f = _delivery_series_demand(fcode)
    dems_r = _delivery_series_demand(rcode)
    pkg_f = max(1, bom.get(fcode, {}).get('package', 1))
    pkg_r = max(1, bom.get(rcode, {}).get('package', 1))
    saf_f = _delivery_safety_qty(fcode, fsupp)
    saf_r = _delivery_safety_qty(rcode, rsupp)
    n = len(dems_f)
    dels_f = [0.0] * n
    dels_r = [0.0] * n
    ss_f = float(_opening_stock(fcode))
    ss_r = float(_opening_stock(rcode))
    for di in range(n):
        dt = all_dates[di][0]
        demf, demr = dems_f[di], dems_r[di]
        projf, projr = ss_f - demf, ss_r - demr
        trigger = ((demf > 0 and projf < saf_f) or (demr > 0 and projr < saf_r))
        if trigger:
            df = _round_delivery_qty(fcode, max(0, saf_f - projf), pkg_f)
            dr = _round_delivery_qty(rcode, max(0, saf_r - projr), pkg_r)
            if df <= 0:
                df = float(pkg_f)
            if dr <= 0:
                dr = float(pkg_r)
            dels_f[di] = df
            dels_r[di] = dr
        ss_f = _ss_after_day(ss_f, demf, dels_f[di], fcode, dt)
        ss_r = _ss_after_day(ss_r, demr, dels_r[di], rcode, dt)
    return dels_f, dels_r

def _msa_bumper_load_by_day(all_dels, msa_codes):
    """Текущая загрузка тарных мест MSA по дням (слоты накладок + упаковки бамперов).
    Общая точка правды для _smooth_msa_trucks и капасити-донабора (_ensure_all_ss),
    чтобы обе стороны видели одну и ту же загрузку дня."""
    n = len(all_dates)
    bumpers = [c for c in msa_codes
               if c not in MSA_FRONT_TRIM_CODES and c not in MSA_REAR_TRIM_CODES
               and c != MSA_FULL_TRUCK_CODE]
    pkgs = {c: max(1, bom.get(c, {}).get('package', 1)) for c in bumpers}
    load = [0] * n
    for di in range(n):
        fq = {c: all_dels[c][di] for c in MSA_FRONT_TRIM_CODES if c in all_dels}
        rq = {c: all_dels[c][di] for c in MSA_REAR_TRIM_CODES if c in all_dels}
        fs, rs = _msa_trim_day_slots(fq, rq)
        load[di] = fs + rs + sum(int(round(all_dels[c][di] / pkgs[c])) for c in bumpers)
    return load, bumpers, pkgs

def _smooth_msa_trucks(all_dels, msa_codes):
    """Фура ≤ 21 контейнер/день. Избыток переносим ТОЛЬКО на более ранний рабочий день
    С ПОТРЕБНОСТЬЮ у этой же детали (не на пустые дни). Если такого дня нет — объём
    остаётся на дне потребности (фура может превысить 21, но поставок в пустые дни нет).
    Перенос только назад по времени → отрицательных остатков не возникает."""
    CAP = MSA_MAX_CONTAINERS_PER_TRUCK
    SOFT = _soft_truck_cap(CAP)
    load, bumpers, pkgs = _msa_bumper_load_by_day(all_dels, msa_codes)
    demday = {c: [d > 0 for d in _delivery_series_demand(c)] for c in bumpers}
    for di in range(len(all_dates) - 1, -1, -1):
        guard = 0
        while load[di] > SOFT and guard < 100000:
            guard += 1
            moved = False
            for c in bumpers:
                if all_dels[c][di] < pkgs[c]:
                    continue
                dj = -1
                for k in range(di - 1, -1, -1):
                    if load[k] < SOFT and demday[c][k]:
                        dj = k
                        break
                if dj >= 0:
                    all_dels[c][di] -= pkgs[c]
                    all_dels[c][dj] += pkgs[c]
                    load[di] -= 1
                    load[dj] += 1
                    moved = True
                    break
            if not moved:
                while load[di] > CAP + 1e-9:
                    cut = False
                    for c in sorted(bumpers, key=lambda x: all_dels[x][di], reverse=True):
                        if all_dels[c][di] >= pkgs[c]:
                            all_dels[c][di] -= pkgs[c]
                            load[di] -= 1
                            cut = True
                            break
                    if not cut:
                        break
                break

def _group_day_pcs(all_dels, group_codes, di):
    return sum(all_dels[c][di] for c in group_codes if c in all_dels)

def _cap_group_day_pcs(raw_qty, max_pcs):
    """Пропорционально уменьшает отгрузки группы, если сумма > max_pcs."""
    if not raw_qty:
        return raw_qty
    total = sum(raw_qty.values())
    if total <= max_pcs:
        return raw_qty
    scale = max_pcs / total
    out = {}
    for c, q in raw_qty.items():
        new_q = _floor_pkg_qty(q * scale, _pkg_size(c))
        out[c] = new_q
    return out

def _smooth_msa_bumper_group_cap(all_dels, group_codes):
    """Сумма отгрузок группы бамперов XST33/AST33 в день ≤ 480 шт.
    Избыток переносим на более ранние дни с потребностью по той же детали."""
    codes = [c for c in group_codes if c in all_dels]
    if not codes:
        return
    n = len(all_dates)
    CAP = MSA_BUMPER_GROUP_MAX_PCS_PER_DAY
    SOFT = _soft_truck_cap(CAP)
    pkgs = {c: max(1, bom.get(c, {}).get('package', 1)) for c in codes}
    demday = {c: [d > 0 for d in _delivery_series_demand(c)] for c in codes}
    load = [_group_day_pcs(all_dels, codes, di) for di in range(n)]
    for di in range(n - 1, -1, -1):
        guard = 0
        while load[di] > SOFT + 1e-9 and guard < 100000:
            guard += 1
            moved = False
            for c in codes:
                pkg = pkgs[c]
                if all_dels[c][di] < pkg:
                    continue
                for k in range(di - 1, -1, -1):
                    if not demday[c][k]:
                        continue
                    if load[k] + pkg <= SOFT + 1e-9:
                        all_dels[c][di] -= pkg
                        all_dels[c][k] += pkg
                        load[di] -= pkg
                        load[k] += pkg
                        moved = True
                        break
                if moved:
                    break
            if not moved:
                capped = _cap_group_day_pcs(
                    {c: all_dels[c][di] for c in codes if all_dels[c][di] > 0}, CAP)
                for c in codes:
                    all_dels[c][di] = capped.get(c, 0.0)
                load[di] = _group_day_pcs(all_dels, codes, di)
                break

def _apply_msa_trim_delivery_rules(all_dels):
    """Накладки: передние кратно 480 (мин. 480), задние кратно 180 + ПАРНОСТЬ."""
    for code in MSA_FRONT_TRIM_CODES | MSA_REAR_TRIM_CODES:
        if code not in all_dels:
            continue
        pkg = max(1, bom.get(code, {}).get('package', 1))
        for di in range(len(all_dels[code])):
            q = all_dels[code][di]
            if q > 0:
                all_dels[code][di] = _round_delivery_qty(code, q, pkg)
    _sync_msa_trim_pairs(all_dels)
    _balance_msa_trim_pairs(all_dels)
    _sync_msa_trim_pairs(all_dels)

def _sync_msa_trim_pairs(all_dels):
    """Парность накладок: в день отгрузки обе детали пары едут одинаковым
    объёмом (передние кратно 480; задние кратно 60, минимум 180)."""
    for codes, mult, min_q in ((MSA_FRONT_TRIM_CODES, MSA_FRONT_TRIM_MULTIPLE,
                                MSA_FRONT_TRIM_MIN_QTY),
                               (MSA_REAR_TRIM_CODES, MSA_REAR_TRIM_MULTIPLE,
                                MSA_REAR_TRIM_MIN_QTY)):
        pair = [c for c in sorted(codes) if c in all_dels]
        if len(pair) != 2:
            continue
        a, b = pair
        n = min(len(all_dels[a]), len(all_dels[b]))
        for di in range(n):
            qa = float(all_dels[a][di] or 0)
            qb = float(all_dels[b][di] or 0)
            if qa <= 0 and qb <= 0:
                continue
            q = max(math.ceil(qa / mult), math.ceil(qb / mult)) * mult
            q = max(q, min_q)
            all_dels[a][di] = float(q)
            all_dels[b][di] = float(q)

def _balance_msa_trim_pairs(all_dels):
    """Балансировка остатков накладок ПЕРЁД/ЗАД: задние возим в ТЕ ЖЕ ДНИ, что и
    передние, и в объёме, выводящем остаток задних на уровень передних — чтобы
    уровень запасов перед/зад на каждый день был одинаковым (применяемость одна).
    Кратность задних 60 (мин 180); 60 делит 480, поэтому уровни сходятся.
    Передний график (согласованный) не меняем."""
    front = sorted(c for c in MSA_FRONT_TRIM_CODES if c in all_dels)
    rear = sorted(c for c in MSA_REAR_TRIM_CODES if c in all_dels)
    if len(front) != 2 or len(rear) != 2:
        return
    fdel = all_dels[front[0]]
    n = len(fdel)
    dems_f = _delivery_series_demand(front[0])
    dems_r = _delivery_series_demand(rear[0])
    ss_f = float(_opening_stock_for_ss(front[0]))
    ss_r = float(_opening_stock_for_ss(rear[0]))
    new_r = [0.0] * n
    for di in range(n):
        fd = float(fdel[di] or 0)
        df = dems_f[di] if di < len(dems_f) else 0.0
        dr = dems_r[di] if di < len(dems_r) else 0.0
        if fd > 0:
            front_end = ss_f + fd - df
            rear_proj = ss_r - dr
            need = front_end - rear_proj
            if need > 0:
                q = math.ceil(need / MSA_REAR_TRIM_MULTIPLE) * MSA_REAR_TRIM_MULTIPLE
                new_r[di] = float(max(q, MSA_REAR_TRIM_MIN_QTY))
        ss_f = ss_f + fd - df
        ss_r = ss_r + new_r[di] - dr
    for c in rear:
        all_dels[c] = list(new_r)


def _msa_full_truck_qty(need):
    """Минимальный объём ≥ need из комбинаций фур 528/576 (5304100XKN02A)."""
    if need <= 0:
        return 0
    s_min, s_max = min(MSA_FULL_TRUCK_SIZES), max(MSA_FULL_TRUCK_SIZES)
    best = None
    max_b = int(math.ceil(need / s_max)) + 1
    for b in range(max_b + 1):
        rest = need - b * s_max
        a = int(math.ceil(rest / s_min)) if rest > 0 else 0
        total = a * s_min + b * s_max
        if total >= need and (best is None or total < best):
            best = total
    return int(best or s_min)

def _rebuild_msa_full_truck_deliveries(code, dems, dels, frozen_del=None):
    """5304100XKN02A: отгрузка целыми фурами (528/576), 1–2 раза в неделю.
    Слоты недели: первый день потребности и первый день потребности с четверга.
    Объём фуры покрывает потребность до следующего слота. Ручные 📦 (frozen)
    не трогаем."""
    frozen_del = frozen_del or set()
    n = len(dems)
    slots = []
    cur_week = None
    week_dis = []
    def _flush_week(w_dis):
        if not w_dis:
            return
        slots.append(w_dis[0])
        second = next((di for di in w_dis if all_dates[di][0].weekday() >= 3
                       and di != w_dis[0]), None)
        if second is not None:
            slots.append(second)
    for di in range(n):
        if dems[di] <= 0:
            continue
        wk = all_dates[di][0].isocalendar()[:2]
        if wk != cur_week:
            _flush_week(week_dis)
            cur_week = wk
            week_dis = []
        week_dis.append(di)
    _flush_week(week_dis)
    slot_set = set(slots)
    ss = float(_opening_stock_for_ss(code))
    for di in range(n):
        dt = all_dates[di][0]
        if di not in frozen_del:
            dels[di] = 0.0
            if di in slot_set:
                nxt = next((s for s in slots if s > di), n)
                period = sum(dems[di:nxt])
                need = period - ss
                if need > 1e-6:
                    dels[di] = float(_msa_full_truck_qty(need))
        ss = _ss_after_day(ss, dems[di], float(dels[di] or 0), code, dt)

def _consolidate_single_fixed_delivery(all_dels, code, fixed_qty):
    """Одна поставка фиксированного объёма (Калуга) в первый день потребности."""
    if code not in all_dels or fixed_qty <= 0:
        return
    dels = all_dels[code]
    if sum(dels) <= 0:
        return
    dems = _delivery_series_demand(code)
    target = next((di for di in range(len(dems)) if dems[di] > 0), None)
    if target is None:
        return
    for di in range(len(dels)):
        dels[di] = 0.0
    dels[target] = float(_round_delivery_qty(code, fixed_qty))

def _build_ahead_deliveries(code, supplier):
    """ПРЕДЗАВОЗ: поставка приходит в конце предыдущего рабочего дня (дня потребности),
    поэтому на НАЧАЛО дня потребности остаток уже покрывает спрос (≥100%, минимум 50%) —
    отрицательных остатков внутри дня нет. del[d] на рабочем дне d закрывает СЛЕДУЮЩИЙ
    рабочий день. Цель конца дня = страх.запас + спрос следующего дня потребности.
    (Пары передний↔задний согласуются сами: одинаковый спрос → одни дни поставок.)"""
    dems = _delivery_series_demand(code)
    pkg = max(1, bom.get(code, {}).get('package', 1))
    safety = _delivery_safety_qty(code, supplier)
    n = len(dems)
    dels = [0.0] * n
    arrival = [0.0] * n          # отгрузка, ПРИХОДЯЩАЯ в этот день (с пред. рабочего дня)
    nxt_wd = [None] * n
    _nd = None
    for di in range(n - 1, -1, -1):
        nxt_wd[di] = _nd
        if dems[di] > 0:
            _nd = di
    _first_wd = next((di for di in range(n) if dems[di] > 0), None)
    ss = float(_opening_stock(code))
    # Деталь производится в 1-й день горизонта: предыдущего дня нет → считаем, что предзавоз
    # был до горизонта (начальный остаток покрывает 1-й день), иначе день 0 уходит в минус.
    if _first_wd == 0:
        ss = max(ss, safety + dems[0])
    for di in range(n):
        dt = all_dates[di][0]
        _snap = manual_stock_by_date.get(code, {}).get(dt)
        if _snap is not None:
            ss = float(_snap)
        else:
            ss = ss - dems[di] + arrival[di]
        # СТАРТОВАЯ поставка: на день перед первым днём потребности (приходит в первый день),
        # т.к. предыдущего дня потребности в горизонте нет — иначе первый день уходит в минус.
        if _first_wd is not None and _first_wd > 0 and di == _first_wd - 1:
            _d0 = _round_delivery_qty(code, max(0, safety + dems[_first_wd] - ss), pkg)
            if _d0 > 0:
                dels[di] += _d0
                arrival[_first_wd] += _d0
        if dems[di] > 0:
            # отгрузка этого рабочего дня ПРИХОДИТ на следующий рабочий день и закрывает его
            _nw = nxt_wd[di]
            _target = (safety + dems[_nw]) if _nw is not None else safety
            _d = _round_delivery_qty(code, max(0, _target - ss), pkg)
            if _d > 0:
                dels[di] += _d
                if _nw is not None:
                    arrival[_nw] += _d
    return dels

def _schedules_from_all_dels(all_dels):
    schedules = {}
    for code in mrp_codes:
        pairs = _recompute_ss_from_deliveries(code, all_dels[code])
        schedules[code] = {
            'dels': [p[0] for p in pairs],
            'ss': [p[1] for p in pairs],
            'pairs': {di: pairs[di] for di in range(len(pairs))},
        }
    return schedules

def _finalize_all_dels(all_dels, frozen_del_by_code=None, quiet=False):
    """Reconcile + кратность упаковки + дonабor Ss (с учётом замороженных ручных 📦)."""
    frozen_del_by_code = frozen_del_by_code or {}
    neg_before, below_before = _count_excel_sim_negatives(all_dels)
    for _fix_pass in range(5):
        for code in mrp_codes:
            dems = _delivery_series_demand(code)
            _, supp, _ = get_info(code)
            _reconcile_demand_day_deliveries(
                code, supp, dems, all_dels[code],
                frozen_del=frozen_del_by_code.get(code))
        _enforce_delivery_pkg_multiples(all_dels, mrp_codes)
        for code in mrp_codes:
            dems = _delivery_series_demand(code)
            _, supp, _ = get_info(code)
            _ensure_non_negative_balances(
                code, supp, dems, all_dels[code],
                frozen_del=frozen_del_by_code.get(code, set()))
        neg, below = _count_excel_sim_negatives(all_dels)
        if neg + below == 0:
            break
    _enforce_delivery_pkg_multiples(all_dels, mrp_codes)
    neg_after, below_after = _count_excel_sim_negatives(all_dels)
    total_before = neg_before + below_before
    total_after = neg_after + below_after
    if not quiet:
        if total_before:
            print(f"  Ss дефicit до дonабora: {total_before} дн×код", flush=True)
        if total_after:
            print(f"  ⚠️  Ss дефicit после дonабora: {total_after} "
                  f"(Ss<0: {neg_after}, Ss<potr: {below_after})", flush=True)
        elif total_before:
            print(f"  ✅ Дефицит Ss устранён дonабorом поставок", flush=True)
    return total_after

def _frozen_manual_deliveries():
    """Дни с ручными 📦 (оранж. ячейки) — не менять при дonаборе."""
    di_by_date = {all_dates[i][0]: i for i in range(len(all_dates))}
    out = {}
    for code, daymap in manual_delivery_by_date.items():
        frozen = set()
        for dt in daymap:
            di = di_by_date.get(dt)
            if di is not None:
                frozen.add(di)
        if frozen:
            out[code] = frozen
            nc = normalize_code(code)
            if nc != code:
                out.setdefault(nc, set()).update(frozen)
    return out

def _ensure_all_ss(all_dels, frozen_del_by_code=None):
    frozen_del_by_code = frozen_del_by_code or {}
    msa_codes = [c for c in mrp_codes
                 if normalize_supplier(get_info(c)[1]) == MSA_SUPPLIER]
    msa_load, msa_bumpers, _ = (
        _msa_bumper_load_by_day(all_dels, msa_codes) if msa_codes else ([], [], {}))
    msa_bumpers = set(msa_bumpers)
    for code in mrp_codes:
        dems = _delivery_series_demand(code)
        _, supp, _ = get_info(code)
        _ensure_non_negative_balances(
            code, supp, dems, all_dels[code],
            frozen_del=frozen_del_by_code.get(code, set()),
            msa_load=msa_load if code in msa_bumpers else None)
    _enforce_delivery_pkg_multiples(all_dels, mrp_codes)

def _settle_deliveries(all_dels, frozen_del_by_code=None):
    """Финализация → фурование NDKT (целые машины, 100% первых партий) →
    лимиты фур → дonабor Ss → финальное фурование NDKT (последнее слово)."""
    frozen_del_by_code = frozen_del_by_code or {}
    _finalize_all_dels(all_dels, frozen_del_by_code, quiet=False)
    _ndkt_rebuild_truck_schedule(all_dels, frozen_del_by_code)
    last_total = -1
    for _ in range(15):   # было 10; 30 давало слишком долгий прогон
        _enforce_truck_limits(all_dels)
        _ensure_all_ss(all_dels, frozen_del_by_code)
        neg, below = _count_excel_sim_negatives(all_dels)
        total = neg + below
        if total == 0:
            _ndkt_rebuild_truck_schedule(all_dels, frozen_del_by_code, quiet=True)
            _lear_round_headrest_day_sum(all_dels)
            _smooth_eftec_weight(all_dels)
            return 0
        if total == last_total:
            break
        last_total = total
    _ndkt_rebuild_truck_schedule(all_dels, frozen_del_by_code, quiet=True)
    _lear_round_headrest_day_sum(all_dels)
    _smooth_eftec_weight(all_dels)
    neg, below = _count_excel_sim_negatives(all_dels)
    total = neg + below
    if total:
        print(f"  ⚠️  Ss дефicit после стабилизации: {total} дн×код "
              f"(Ss<0: {neg}, Ss<potr: {below})", flush=True)
    return total

def _enforce_truck_limits(all_dels):
    """Лимиты фур: перенос назад, при невозможности — урезание до номинала (строго)."""
    smc_codes = [c for c in mrp_codes if _delivery_mode(get_info(c)[1]) == 'smc_weekly']
    if smc_codes:
        _smooth_smc_trucks(all_dels, smc_codes)
    msa_codes = [c for c in mrp_codes
                 if normalize_supplier(get_info(c)[1]) == MSA_SUPPLIER]
    if msa_codes:
        _smooth_msa_trucks(all_dels, msa_codes)
        _apply_msa_trim_delivery_rules(all_dels)
        _smooth_msa_bumper_group_cap(all_dels, MSA_XST33_FRONT_BUMPER_CODES)
        _smooth_msa_bumper_group_cap(all_dels, MSA_AST33_REAR_BUMPER_CODES)
    lear_codes = [c for c in mrp_codes
                  if normalize_supplier(get_info(c)[1]) == LEAR_SUPPLIER]
    if lear_codes:
        _smooth_lear_trucks(all_dels, lear_codes)
        _lear_round_headrest_day_sum(all_dels)
    _smooth_eftec_weight(all_dels)
    _enforce_delivery_pkg_multiples(all_dels, mrp_codes)
    _enforce_vm_truck_rows(all_dels)
    # Правило фуры Уралэластотехники строже кратности упаковки — применяем последним
    _enforce_ural_truck_composition(all_dels)

def build_all_delivery_schedules():
    all_dels = {}
    for code in mrp_codes:
        _, supp, _ = get_info(code)
        dems = _delivery_series_demand(code)
        dels = _build_initial_deliveries(code, supp)
        _consolidate_to_supplier_slots(code, supp, dems, dels)
        _reconcile_demand_day_deliveries(code, supp, dems, dels)
        all_dels[code] = dels

    # Пары передний↔задний согласуются автоматически (одинаковый спрос → одни дни
    # поставок и близкие объёмы) — комплектный переходящий запас сохраняется.
    lear_codes = [c for c in mrp_codes
                  if normalize_supplier(get_info(c)[1]) == LEAR_SUPPLIER]
    _untyped = [c for c in lear_codes if _lear_product_kind(c) is None]
    if _untyped:
        print(f"  ⚠️  Lear без примечания Пена/Подголовник: {len(_untyped)} кодов "
              f"(укажите в листе Упаковка col F)", flush=True)
    for code, qty in KALUGA_SINGLE_DELIVERY_QTY.items():
        if code in all_dels:
            _consolidate_single_fixed_delivery(all_dels, code, qty)

    for code in mrp_codes:
        dems = _delivery_series_demand(code)
        _, supp, _ = get_info(code)
        _reconcile_demand_day_deliveries(code, supp, dems, all_dels[code])

    _settle_deliveries(all_dels)
    _report_truck_tolerance_usage(all_dels)
    return all_dels, _schedules_from_all_dels(all_dels)

def _report_truck_tolerance_usage(all_dels):
    """Дни с загрузкой фуры выше номинала / выше допуска (для согласования лимитов)."""
    n = len(all_dates)
    smc_codes = [c for c in mrp_codes if _delivery_mode(get_info(c)[1]) == 'smc_weekly']
    msa_codes = [c for c in mrp_codes
                 if normalize_supplier(get_info(c)[1]) == MSA_SUPPLIER]
    lear_codes = [c for c in mrp_codes
                  if normalize_supplier(get_info(c)[1]) == LEAR_SUPPLIER]
    bump_f = MSA_XST33_FRONT_BUMPER_CODES
    bump_r = MSA_AST33_REAR_BUMPER_CODES
    pkgs_smc = {c: max(1, bom.get(c, {}).get('package', 1)) for c in smc_codes}
    pkgs_msa = {c: max(1, bom.get(c, {}).get('package', 1)) for c in msa_codes
                if c not in MSA_FRONT_TRIM_CODES and c not in MSA_REAR_TRIM_CODES
                and c != MSA_FULL_TRUCK_CODE}

    def _count_over(load_fn, nominal, soft):
        over_nom = over_soft = 0
        max_pct = 0.0
        for di in range(n):
            load = load_fn(di)
            if load <= nominal + 1e-9:
                continue
            over_nom += 1
            pct = (load / nominal - 1.0) * 100 if nominal else 0
            max_pct = max(max_pct, pct)
            if load > soft + 1e-9:
                over_soft += 1
        return over_nom, over_soft, max_pct

    smc_fn = lambda di: sum(
        _smc_pallet_count(c, all_dels[c][di], pkgs_smc[c])
        for c in smc_codes if all_dels.get(c, [0]*n)[di] > 0)
    msa_fn = lambda di: (
        sum(int(round(all_dels[c][di] / pkgs_msa[c])) for c in pkgs_msa if all_dels[c][di] > 0)
        + sum(_msa_trim_day_slots(
            {c: all_dels[c][di] for c in MSA_FRONT_TRIM_CODES if c in all_dels},
            {c: all_dels[c][di] for c in MSA_REAR_TRIM_CODES if c in all_dels})))
    lear_fn = lambda di: _lear_day_pallet_slots(
        {c: all_dels[c][di] for c in lear_codes if all_dels[c][di] > 0})
    bf_fn = lambda di: _group_day_pcs(all_dels, bump_f, di)
    br_fn = lambda di: _group_day_pcs(all_dels, bump_r, di)

    lines = []
    for label, fn, cap in [
        ('SMC', smc_fn, SMC_MAX_PALLETS_PER_TRUCK),
        ('MSA', msa_fn, MSA_MAX_CONTAINERS_PER_TRUCK),
        ('Lear', lear_fn, LEAR_MAX_PALLET_SLOTS),
        ('MSA бамперы XST33', bf_fn, MSA_BUMPER_GROUP_MAX_PCS_PER_DAY),
        ('MSA бамперы AST33', br_fn, MSA_BUMPER_GROUP_MAX_PCS_PER_DAY),
    ]:
        soft = _soft_truck_cap(cap)
        on, os, mp = _count_over(fn, cap, soft)
        if on:
            lines.append(f"{label}: {on} дн. >{cap} (макс +{mp:.0f}%), {os} дн. >допуска {soft:.0f}")
    # Эфтек: 18 т — лимит ОДНОЙ машины; день может требовать несколько ТС.
    # Смузер уже перенёс всё, что можно, на ранние дни; остаток — реальные 2+ ТС.
    _ef_codes = [c for c in mrp_codes
                 if normalize_supplier(get_info(c)[1]) == EFTEC_SUPPLIER]
    if _ef_codes:
        _ef_multi = 0
        _ef_trucks = 0
        for di in range(n):
            _w = sum(float(all_dels.get(c, [0] * n)[di] or 0) * _eftec_unit_weight_g(c)
                     for c in _ef_codes)
            if _w > 1e-6:
                _t = -(-int(_w) // EFTEC_MAX_TRUCK_WEIGHT_G)
                _ef_trucks += max(1, _t)
                if _w > EFTEC_MAX_TRUCK_WEIGHT_G + 1e-6:
                    _ef_multi += 1
        if _ef_multi:
            lines.append(f"Эфтек: {_ef_multi} дн. требуют 2+ ТС (по 18 т), всего ТС: {_ef_trucks}")
    # YAPP: 240 шт = 1 машина, допустимо до 4 машин в день (согласовано)
    yapp_codes = [c for c in mrp_codes if 'YAPP' in str(get_info(c)[1] or '').upper()]
    if yapp_codes:
        _yp_over = sum(1 for di in range(n)
                       if sum(float(all_dels.get(c, [0] * n)[di] or 0)
                              for c in yapp_codes) > 960 + 1e-6)
        if _yp_over:
            lines.append(f"YAPP: {_yp_over} дн. >960 шт (лимит 4 машины × 240)")
    tol_pct = int(TRUCK_CAPACITY_TOLERANCE * 100)
    ss_tol = int(DELIVERY_SS_TARGET_TOLERANCE * 100)
    truck_msg = "строго по номиналу" if tol_pct == 0 else f"+{tol_pct}%"
    print(f"  Допуски: фура {truck_msg}, целевой Ss −{ss_tol}% (пол Ss≥potr)", flush=True)
    if lines:
        print(f"  Загрузка фур выше номинала: {' | '.join(lines)}", flush=True)
    else:
        print(f"  Загрузка фур: все дни ≤ номинала", flush=True)
print("  Расчёт графиков поставок (Ecoal'yance / Ecotexis / SMC / Purem / Lear)...")
all_dels, DELIVERY_SCHEDULES = build_all_delivery_schedules()
_ndkt_codes = [c for c in mrp_codes if _is_ndkt_part(c) and _code_3m_demand(c) > 0]
if _ndkt_codes:
    _ndkt_max_trucks = 0.0
    for _di in range(len(all_dates)):
        _load_n = sum(float(all_dels.get(c, [0]*len(all_dates))[_di] or 0) / _ndkt_truck_cap(c)
                      for c in _ndkt_codes)
        _ndkt_max_trucks = max(_ndkt_max_trucks, _load_n)
    print(f"    NDKT: {len(_ndkt_codes)} поз. | упрощённая логика: цель = страх.запас по коду "
          f"(240/480) + спрос след. дня | ПЕРВЫЕ партии плана (AS_in_F_A + AS_in_H_B) — 100% из остатков | "
          f"упаковка: XST33A бокс 30, прочие колёса стопка 24 | микс колёс (фура 540/384), "
          f"макс. {_ndkt_max_trucks:.1f} фуры/день", flush=True)
_n_smc = sum(1 for c in mrp_codes if _delivery_mode(get_info(c)[1]) == 'smc_weekly')
_n_lear = sum(1 for c in mrp_codes if normalize_supplier(get_info(c)[1]) == LEAR_SUPPLIER)
print(f"    Расписаний: {len(DELIVERY_SCHEDULES)} | SMC: {_n_smc} (≤{SMC_MAX_PALLETS_PER_TRUCK} палл.) | "
      f"Lear: {_n_lear} (≤{LEAR_MAX_PALLET_SLOTS} палл., {LEAR_FOAM_PKGS_PER_TRUCK} пен/фура)", flush=True)

# ═══ CKD: прогноз +3 месяца (месячные потребности и заказы, без расчёта по дням) ═══
# Источник: самая свежая папка «New CKD …» — файл-сводка «план по моделям × месяцы».
# Распределение плана моделей по конфигурациям — СРЕДНИЕ доли конфигураций
# (и средний цветовой микс для бамперов/красок) из 3 расчётных месяцев Plan-Fact.
print("\nПрогноз CKD (+3 месяца, месячные потребности/заказы)...")

FORECAST_MONTHS = []   # [(mnum, year, 'Sep 2026')]
_fc_last_mn = MONTHS[-1][0]
_fc_last_yr = MONTH_YEAR[_fc_last_mn]
for _i in range(1, 5):   # горизонт прогноза расширен до декабря (+4 мес.)
    _m = (_fc_last_mn - 1 + _i) % 12 + 1
    _y = _fc_last_yr + (_fc_last_mn - 1 + _i) // 12
    FORECAST_MONTHS.append((_m, _y, f"{_MONTH_EN_SHORT[_m]} {_y}"))
FORECAST_LABEL = {(m, y): lbl for m, y, lbl in FORECAST_MONTHS}

_CKD_RU_MON = {'янв': 1, 'фев': 2, 'мар': 3, 'апр': 4, 'май': 5, 'мая': 5, 'июн': 6,
               'июл': 7, 'авг': 8, 'сен': 9, 'окт': 10, 'ноя': 11, 'дек': 12}
_CKD_EN_MON = {'jan': 1, 'feb': 2, 'mar': 3, 'apr': 4, 'may': 5, 'jun': 6, 'jul': 7,
               'aug': 8, 'sep': 9, 'oct': 10, 'nov': 11, 'dec': 12}

def _parse_ckd_month_label(v):
    """«Сен'26» / «Oct'26» / дата → (месяц, год) или None."""
    if v is None:
        return None
    if isinstance(v, (datetime.datetime, datetime.date)):
        return (v.month, v.year)
    s = str(v).strip().lower().replace('’', "'")
    m = re.match(r"^([а-яёa-z]{3,8})\.?\s*'?\s*(\d{2,4})$", s)
    if not m:
        return None
    mn = _CKD_RU_MON.get(m.group(1)[:3]) or _CKD_EN_MON.get(m.group(1)[:3])
    if not mn:
        return None
    yr = int(m.group(2))
    if yr < 100:
        yr += 2000
    return (mn, yr)

def _find_ckd_dir():
    """Самая свежая папка с CKD-планами (по mtime xlsx внутри)."""
    cands = [d for d in (glob.glob(os.path.join(ROOT, '*CKD*'))
                         + glob.glob(os.path.join(ROOT, 'Входные данные', '*CKD*')))
             if os.path.isdir(d)]
    if not cands:
        return None
    def _key(d):
        xl = [p for p in glob.glob(os.path.join(d, '*.xlsx')) if not _is_excel_lock_file(p)]
        return max([os.path.getmtime(p) for p in xl] or [os.path.getmtime(d)])
    return max(cands, key=_key)

_CKD_MODEL_RE = re.compile(r'^(A01|A08|B02|B04|B06|B16)\b', re.I)

# Алиасы моделей в файлах CKD (порядок важен: F7x раньше F7, H3/H7 точные)
_CKD_FILE_MODEL_ALIASES = [
    ('f7x', 'B04'), ('f7', 'B02'), ('jolion', 'A01'),
    ('dargo', 'B06'), ('h7', 'B16'), ('h3', 'A08'),
]

def _ckd_alias_model(text):
    t = re.sub(r'\s+', '', str(text or '').lower())   # «J O L I O N» → «jolion»
    if not t or 'total' in t or 'итог' in t:
        return None
    for al, mdl in _CKD_FILE_MODEL_ALIASES:
        if al in t:
            return mdl
    return None

def _ckd_rus_year(rus_m):
    """Год месяца RUS: скользящее окно от первого расчётного месяца."""
    base_m = MONTHS[0][0]
    base_y = MONTH_YEAR[base_m]
    return base_y if rus_m >= base_m else base_y + 1

def _ckd_name_month(s):
    mn = _CKD_RU_MON.get(str(s or '').strip().lower()[:3]) or _CKD_EN_MON.get(str(s or '').strip().lower()[:3])
    return mn

def _load_ckd_plan_model_files(d):
    """{(mnum, year): {model: cars}} из файлов моделей CKD (Production request).
    Месяцы в файлах — китайские «06月…»; сдвиг CHN→RUS берётся из имени файла:
    «… Sep - Feb RUS (June - Oct CHN)» → June CHN = Sep RUS. Складываются
    колонки Total конфигурационных строк (строки Total/Итого пропускаются)."""
    plan = {}
    src_files = []
    for path in sorted(glob.glob(os.path.join(d, '*.xlsx'))):
        if _is_excel_lock_file(path):
            continue
        base = os.path.basename(path)
        m = re.search(r'([A-Za-zА-Яа-я]+)\s*[-–]\s*[A-Za-zА-Яа-я]+\s+RUS\s*\(\s*([A-Za-zА-Яа-я]+)',
                      base, re.I)
        if not m:
            continue
        rus_start = _ckd_name_month(m.group(1))
        chn_start = _ckd_name_month(m.group(2))
        if not rus_start or not chn_start:
            continue
        offset = (rus_start - chn_start) % 12
        try:
            wb_c = load_workbook(path, read_only=True, data_only=True)
        except Exception:
            continue
        file_rows = 0
        for ws_c in wb_c.worksheets:
            rows = list(ws_c.iter_rows(min_row=1, max_row=120, max_col=60, values_only=True))
            hdr_i = month_cols = None
            for i, row in enumerate(rows):
                cols = {}
                for ci, v in enumerate(row):
                    mm = re.match(r'^\s*(\d{1,2})\s*月', str(v or ''))
                    if mm:
                        cols[ci] = int(mm.group(1))
                if len(cols) >= 2:
                    hdr_i, month_cols = i, cols
                    break
            if hdr_i is None or hdr_i + 1 >= len(rows):
                continue
            sub = rows[hdr_i + 1]
            # колонка Total каждого месяца — до начала следующего месяца
            mc_sorted = sorted(month_cols.items())
            tot_cols = {}
            for k, (ci, chn_m) in enumerate(mc_sorted):
                nxt = mc_sorted[k + 1][0] if k + 1 < len(mc_sorted) else len(sub)
                for j in range(ci, min(nxt, len(sub))):
                    sv = str(sub[j] or '').lower()
                    if 'total' in sv or '合计' in sv:
                        tot_cols[chn_m] = j
                        break
            if not tot_cols:
                continue
            # колонка Sequence (признак конфигурационной строки)
            seq_col = next((ci for ci, v in enumerate(rows[hdr_i])
                            if 'sequence' in str(v or '').lower()), 6)
            cur_model = None
            for row in rows[hdr_i + 2:]:
                a = str(row[0] or '').strip() if row else ''
                if a and 'comment' in a.lower():
                    break
                if a:
                    if 'total' in a.lower() or 'итог' in a.lower():
                        continue
                    mdl = _ckd_alias_model(a)
                    if mdl:
                        cur_model = mdl
                if cur_model is None:
                    continue
                b = str(row[1] or '').strip() if len(row) > 1 else ''
                seq_v = row[seq_col] if seq_col < len(row) else None
                try:
                    has_seq = float(seq_v) >= 0
                except (TypeError, ValueError):
                    has_seq = False
                if not (len(b) >= 6 or has_seq):
                    continue
                for chn_m, tc in tot_cols.items():
                    v = row[tc] if tc < len(row) else None
                    try:
                        q = float(v)
                    except (TypeError, ValueError):
                        continue
                    if q <= 0:
                        continue
                    rus_m = (chn_m - 1 + offset) % 12 + 1
                    key = (rus_m, _ckd_rus_year(rus_m))
                    plan.setdefault(key, {})
                    plan[key][cur_model] = plan[key].get(cur_model, 0) + q
                    file_rows += 1
            if file_rows:
                break   # первая таблица файла (Production request) обработана
        try:
            wb_c.close()
        except Exception:
            pass
        if file_rows:
            src_files.append(base)
    return plan, src_files

def _load_ckd_plan():
    """{(mnum, year): {model: cars}} из самой свежей папки CKD.
    Приоритет: файлы моделей (Production request, актуальные объёмы);
    сводка «Сен'26-Фев'27» — только как резерв."""
    d = _find_ckd_dir()
    if not d:
        print("  ⚠️  Папка CKD не найдена — прогноз +3 мес. пропущен")
        return {}, None
    try:
        plan_mf, src_files = _load_ckd_plan_model_files(d)
    except Exception as _mf_e:
        print(f"  ⚠️  Ошибка чтения файлов моделей CKD: {_mf_e}")
        plan_mf, src_files = {}, []
    cov_mf = sum(1 for (m, y, _) in FORECAST_MONTHS if (m, y) in plan_mf)
    if cov_mf == len(FORECAST_MONTHS):
        for (fm, fy, flbl) in FORECAST_MONTHS:
            tot_m = sum(plan_mf.get((fm, fy), {}).values())
            print(f"  План CKD {flbl}: {tot_m:.0f} авто "
                  f"({', '.join(f'{k}={v:.0f}' for k, v in sorted(plan_mf[(fm, fy)].items()))})")
        print(f"  Источник: файлы моделей CKD (папка «{os.path.basename(d)}»): "
              f"{len(src_files)} файлов")
        return plan_mf, os.path.join(d, src_files[0]) if src_files else d
    best = None   # (coverage, n_months, plan, path)
    for path in sorted(glob.glob(os.path.join(d, '*.xlsx')), key=os.path.getsize):
        if _is_excel_lock_file(path):
            continue
        try:
            wb_c = load_workbook(path, read_only=True, data_only=True)
        except Exception:
            continue
        for ws_c in wb_c.worksheets:
            months_cols = None
            plan = {}
            got = set()
            stop = False
            for row in ws_c.iter_rows(min_row=1, max_row=200, max_col=40, values_only=True):
                if stop:
                    break
                if months_cols is None:
                    cols = {}
                    for ci, v in enumerate(row):
                        lb = _parse_ckd_month_label(v)
                        if lb:
                            cols[ci] = lb
                    if len(cols) >= 2:
                        months_cols = cols
                    continue
                label = None
                for v in row[:6]:
                    if v is None:
                        continue
                    mm = _CKD_MODEL_RE.match(str(v).strip())
                    if mm:
                        label = mm.group(1).upper()
                        break
                if label is None:
                    if got and any('итог' in str(v or '').lower() for v in row[:6]):
                        stop = True   # конец первой таблицы (план производства)
                    continue
                if label in got:
                    continue
                got.add(label)
                for ci, (mn, yy) in months_cols.items():
                    v = row[ci] if ci < len(row) else None
                    try:
                        q = float(v)
                    except (TypeError, ValueError):
                        continue
                    if q > 0:
                        plan.setdefault((mn, yy), {})[label] = q
            if plan:
                cov = sum(1 for (m, y, _) in FORECAST_MONTHS if (m, y) in plan)
                if best is None or (cov, len(plan)) > (best[0], best[1]):
                    best = (cov, len(plan), plan, path)
        try:
            wb_c.close()
        except Exception:
            pass
    if not best or best[0] == 0:
        if plan_mf:
            print(f"  План CKD: только файлы моделей (частичное покрытие {cov_mf}/3)")
            return plan_mf, (os.path.join(d, src_files[0]) if src_files else d)
        print(f"  ⚠️  В папке «{os.path.basename(d)}» не найден план моделей на "
              f"{', '.join(l for _, _, l in FORECAST_MONTHS)} — прогноз пропущен")
        return {}, None
    merged = dict(best[2])
    merged.update(plan_mf)   # месяцы из файлов моделей точнее сводки
    print(f"  План CKD: сводка «{os.path.basename(best[3])}» (папка «{os.path.basename(d)}»), "
          f"месяцев: {best[1]}; из файлов моделей: {len(plan_mf)} мес. (приоритет)")
    return merged, best[3]

CKD_PLAN, CKD_SRC = {}, None
try:
    CKD_PLAN, CKD_SRC = _load_ckd_plan()
except Exception as _ckd_e:
    print(f"  ⚠️  Ошибка чтения плана CKD: {_ckd_e}")

def _ckd_day_profile(tab, mdl):
    """Средний профиль распределения по дням месяца (доля на день) для модели
    на вкладке за 3 расчётных месяца Plan-Fact (та же логика, что в прошлых
    периодах плана). mdl=None — профиль всей вкладки (fallback)."""
    prof, tot = {}, 0.0
    for mn, _, _ in MONTHS:
        for bv, day_qty in plan_batches.get(tab, {}).get(mn, {}).items():
            bi = _get_batch_info(tab, bv)
            if mdl is not None and bi.get('model', '') != mdl:
                continue
            for d, cars in (day_qty or {}).items():
                try:
                    dd = int(d)
                except (TypeError, ValueError):
                    continue
                v = float(cars or 0)
                if v <= 0:
                    continue
                prof[dd] = prof.get(dd, 0.0) + v
                tot += v
    if tot <= 0:
        return {}
    return {d: v / tot for d, v in prof.items()}

# Государственные праздники РФ (нерабочие дни, ст. 112 ТК РФ): (месяц, день)
_RU_HOLIDAYS_MD = {
    (1, 1), (1, 2), (1, 3), (1, 4), (1, 5), (1, 6), (1, 7), (1, 8),  # новогодние + Рождество
    (2, 23),   # День защитника Отечества
    (3, 8),    # Международный женский день
    (5, 1),    # Праздник Весны и Труда
    (5, 9),    # День Победы
    (6, 12),   # День России
    (11, 4),   # День народного единства
}

def _ckd_is_working_day(fy, fm, d):
    """Рабочий день прогноза: не воскресенье и не гос. праздник РФ."""
    if (fm, d) in _RU_HOLIDAYS_MD:
        return False
    return datetime.date(fy, fm, d).weekday() != 6   # 6 = воскресенье

def _ckd_distribute_days(qty, profile, fm, fy):
    """Целочисленное распределение qty по дням прогнозного месяца по профилю
    (метод наибольших остатков). Воскресенья и гос. праздники РФ — выходные:
    их доля переносится на ближайший следующий рабочий день (в конце месяца —
    на предыдущий)."""
    total = int(round(qty))
    if total <= 0:
        return {}
    ndays = _cal.monthrange(fy, fm)[1]
    work_days = [d for d in range(1, ndays + 1) if _ckd_is_working_day(fy, fm, d)]
    if not work_days:
        work_days = list(range(1, ndays + 1))

    def _to_work(d):
        if d in work_days:
            return d
        nxt = next((w for w in work_days if w > d), None)
        return nxt if nxt is not None else work_days[-1]

    shares = {}
    for d, s in (profile or {}).items():
        dd = _to_work(min(max(1, int(d)), ndays))
        shares[dd] = shares.get(dd, 0.0) + s
    ssum = sum(shares.values())
    if ssum <= 0:
        shares = {d: 1.0 for d in work_days}
        ssum = float(len(work_days))
    ideal = {d: total * s / ssum for d, s in shares.items()}
    base = {d: int(v) for d, v in ideal.items()}
    rest = total - sum(base.values())
    for d, _ in sorted(ideal.items(), key=lambda kv: (-(kv[1] - int(kv[1])), kv[0]))[:max(0, rest)]:
        base[d] += 1
    return {d: v for d, v in sorted(base.items()) if v > 0}

def _inject_ckd_synthetic_batches():
    """Синтетические партии прогнозных месяцев: план модели × средние доли
    (привод, конфигурация) модели на каждой вкладке Plan-Fact за 3 расчётных
    месяца; цвета — средний цветовой микс модели (для бамперов/красок).
    Объёмы — целые, распределены по дням месяца по среднему дневному профилю
    модели из предыдущих периодов плана."""
    injected = 0
    for tab in list(plan_batches.keys()):
        mix, colors, tot = {}, {}, {}
        for mn, _, _ in MONTHS:
            for bv, day_qty in plan_batches.get(tab, {}).get(mn, {}).items():
                bi = _get_batch_info(tab, bv)
                mdl = bi.get('model', '')
                if not mdl:
                    continue
                cars = sum(day_qty.values())
                if cars <= 0:
                    continue
                key = (bi.get('drive', ''), bi.get('config', ''))
                mix.setdefault(mdl, {})
                mix[mdl][key] = mix[mdl].get(key, 0) + cars
                tot[mdl] = tot.get(mdl, 0) + cars
                bc = _batch_color_lookup(bv)
                if bc:
                    bsum = sum(bc.values())
                    if bsum > 0:
                        cd = colors.setdefault(mdl, {})
                        for ck, cq in bc.items():
                            cd[ck] = cd.get(ck, 0) + cars * (cq / bsum)
        tab_profile = _ckd_day_profile(tab, None)
        model_profiles = {mdl: (_ckd_day_profile(tab, mdl) or tab_profile) for mdl in tot}
        for (fm, fy, flbl) in FORECAST_MONTHS:
            plan_m = CKD_PLAN.get((fm, fy), {})
            if not plan_m:
                continue
            ndays_f = _cal.monthrange(fy, fm)[1]
            month_data = plan_batches.setdefault(tab, {}).setdefault(fm, {})
            for mdl, cars_m in plan_m.items():
                if tot.get(mdl, 0) <= 0:
                    continue
                # Целочисленное распределение плана модели по конфигурациям
                # (метод наибольших остатков): сумма по конфигурациям = план модели,
                # без потерь на округлении.
                ideal_cfg = {k: cars_m * c / tot[mdl] for k, c in mix[mdl].items()}
                tgt_m = int(round(cars_m))
                base_cfg = {k: int(v) for k, v in ideal_cfg.items()}
                rest_m = tgt_m - sum(base_cfg.values())
                for k, _ in sorted(ideal_cfg.items(),
                                   key=lambda kv: (-(kv[1] - int(kv[1])), str(kv[0])))[:max(0, rest_m)]:
                    base_cfg[k] += 1
                for (drv, cfg), q_cfg in base_cfg.items():
                    if q_cfg <= 0:
                        continue
                    day_map = _ckd_distribute_days(q_cfg, model_profiles.get(mdl), fm, fy)
                    if not day_map:
                        continue
                    q_int = sum(day_map.values())
                    bv = re.sub(r'\W+', '', f"RFC{fm:02d}{mdl}{drv or 'X'}{cfg or 'X'}").upper()
                    month_data[bv] = day_map
                    bi_n = {'model': mdl, 'drive': drv, 'config': cfg, 'total': q_int}
                    batch_info_by_tab.setdefault(tab, {}).setdefault(bv, bi_n)
                    batch_info.setdefault(bv, dict(bi_n))
                    cd = colors.get(mdl)
                    if cd:
                        csum = sum(cd.values())
                        if csum > 0:
                            batch_color[bv] = {ck: q_int * cv / csum for ck, cv in cd.items()}
                    injected += 1
    return injected

forecast_demand = {}
forecast_orders = {}
if CKD_PLAN and not CALC_ORDERS:
    print("  Прогноз CKD: пропущен (включается вместе с галочкой «Месячные заказы»)")
if CKD_PLAN and CALC_ORDERS:
    _n_inj = _inject_ckd_synthetic_batches()
    print(f"  Синтетических партий прогноза: {_n_inj} "
          f"(доли конфигураций — средние за {PERIOD_LABEL})")

    def _forecast_fallback_ratio(code):
        """Коды с особой логикой (EcoAlliance и т.п.): 3-мес. потребность,
        масштабированная пропорционально плану CKD моделей на вкладках детали."""
        tabs_f = resolve_plan_tabs(part_tab_map.get(code, ''))
        if not tabs_f:
            return {}
        models_f, cars_hist = set(), 0.0
        for tab in tabs_f:
            for mn, _, _ in MONTHS:
                for bv, day_qty in plan_batches.get(tab, {}).get(mn, {}).items():
                    bi = _get_batch_info(tab, bv)
                    if bi.get('model'):
                        models_f.add(bi['model'])
                        cars_hist += sum(day_qty.values())
        if cars_hist <= 0:
            return {}
        hist = _code_3m_demand(code)
        out = {}
        for (fm, fy, _) in FORECAST_MONTHS:
            cars_fm = sum(q for mdl, q in CKD_PLAN.get((fm, fy), {}).items() if mdl in models_f)
            out[(fm, fy)] = round(hist * cars_fm / cars_hist, 2)
        return out

    for code in mrp_codes:
        forecast_demand[code] = {}
        for (fm, fy, flbl) in FORECAST_MONTHS:
            try:
                d_f = get_daily_demand(code, fm)
            except Exception:
                d_f = {}
            forecast_demand[code][(fm, fy)] = round(sum((d_f or {}).values()), 2)
        if sum(forecast_demand[code].values()) <= 0 and _code_3m_demand(code) > 0:
            _fb = _forecast_fallback_ratio(code)
            if _fb and sum(_fb.values()) > 0:
                forecast_demand[code] = _fb

    # Заказы прогноза: старт — расчётный остаток на конец горизонта из графика
    # поставок; далее остаток переходит: stock += заказ − потребность.
    for code in mrp_codes:
        _ss_list = DELIVERY_SCHEDULES.get(code, {}).get('ss') or []
        stk_f = max(0.0, float(_ss_list[-1])) if _ss_list else float(stock.get(code, 0) or 0)
        saf_f = _safety_qty(code, bom.get(code, {}).get('supplier', ''))
        pkg_f = _pkg_size(code)
        forecast_orders[code] = {}
        for (fm, fy, flbl) in FORECAST_MONTHS:
            dem_f = forecast_demand.get(code, {}).get((fm, fy), 0) or 0
            net_f = max(0.0, dem_f + saf_f - stk_f)
            oq_f = math.ceil(net_f / pkg_f) * pkg_f if net_f > 0 else 0
            forecast_orders[code][(fm, fy)] = {
                'demand': dem_f, 'stock': round(stk_f, 1), 'order': oq_f}
            stk_f = max(0.0, stk_f + oq_f - dem_f)
    _n_fc = sum(1 for c in mrp_codes if sum(forecast_demand.get(c, {}).values()) > 0)
    print(f"  Прогнозная потребность построена: {_n_fc} кодов")
else:
    print("  Прогноз CKD не построен (нет данных плана)")

_GP_DATA_START = 4
_GP_MANUAL_DEL_FILL = fill("FDEBD0")

def _gp_del_col_idx(di):
    return 8 + di * 3 + 2

def _calc_del_baseline(delivery_schedules, code, di):
    _ds = delivery_schedules.get(code, {}).get('dels', [])
    _v = float(_ds[di]) if di < len(_ds) else 0.0
    if _v and round(_v) > 0:
        return int(_round_delivery_qty(code, _v))
    return 0

def _is_manual_del_fill(cell):
    """Оранжевая заливка = явный ручной ввод 📦 (не авто-число прошлого прогона)."""
    try:
        fg = cell.fill.fgColor
        if fg and getattr(fg, 'rgb', None):
            return str(fg.rgb).upper().endswith('FDEBD0')
    except Exception:
        pass
    return False

def _load_manual_delivery_overrides(out_path, delivery_schedules):
    """Сохраняет ручные 📦 из предыдущего График_Поставок (только числа).
    Формулы VLOOKUP замораживаются в значение; авто-ячейки (как в расчёте) не переносятся."""
    manual = {}
    if not out_path or not os.path.exists(out_path):
        return manual

    ndays = len(all_dates)
    if ndays == 0:
        return manual
    del_cols = [_gp_del_col_idx(di) for di in range(ndays)]
    max_col = del_cols[-1]

    baselines = {}
    for code, sched in delivery_schedules.items():
        ds = sched.get('dels', [])
        bl = []
        for di in range(ndays):
            v = float(ds[di]) if di < len(ds) else 0.0
            bl.append(int(round(v)) if v and round(v) > 0 else 0)
        baselines[code] = bl
        nc = normalize_code(code)
        if nc != code:
            baselines.setdefault(nc, bl)

    print("  Загрузка ручных поставок (📦)...", flush=True)
    try:
        wb_f = load_workbook(out_path, data_only=False, read_only=True)
        wb_v = load_workbook(out_path, data_only=True, read_only=True)
    except PermissionError:
        print(f"  ⚠️  Закройте {os.path.basename(out_path)} в Excel — файл занят, "
              f"ручные поставки пропущены.", flush=True)
        return manual
    except Exception as e:
        print(f"  ⚠️  Ручные поставки: не удалось прочитать {os.path.basename(out_path)}: {e}",
              flush=True)
        return manual
    if 'График_Поставок' not in wb_f.sheetnames:
        wb_f.close()
        wb_v.close()
        return manual
    ws_f = wb_f['График_Поставок']
    ws_v = wb_v['График_Поставок']
    n_cells = 0
    n_converted = 0
    n_skipped = 0
    empty_streak = 0
    rows_f = ws_f.iter_rows(min_row=_GP_DATA_START, min_col=1, max_col=max_col, values_only=False)
    rows_v = ws_v.iter_rows(min_row=_GP_DATA_START, min_col=1, max_col=max_col, values_only=True)
    for row_f, row_v in zip(rows_f, rows_v):
        raw_code = row_f[0].value
        if raw_code in (None, ''):
            empty_streak += 1
            if empty_streak >= 30:
                break
            continue
        empty_streak = 0
        code = normalize_code(str(raw_code).strip())
        if not code or code.lower() in ('nan', 'код', 'none'):
            continue
        bl = baselines.get(code)
        if bl is None:
            continue
        for di, col_1 in enumerate(del_cols):
            col_idx = col_1 - 1
            if col_idx >= len(row_f):
                continue
            cell_f = row_f[col_idx]
            raw_f = cell_f.value
            if raw_f is None or raw_f == '':
                continue
            dt = all_dates[di][0]
            is_formula = (cell_f.data_type == 'f'
                          or (isinstance(raw_f, str) and str(raw_f).lstrip().startswith('=')))
            if is_formula:
                raw_v = row_v[col_idx] if col_idx < len(row_v) else None
                if raw_v is None or raw_v == '':
                    n_skipped += 1
                    continue
                try:
                    live_v = float(raw_v)
                except (TypeError, ValueError):
                    n_skipped += 1
                    continue
                n_converted += 1
            else:
                try:
                    live_v = float(raw_f)
                except (TypeError, ValueError):
                    n_skipped += 1
                    continue
            if abs(live_v - bl[di]) > 1e-6 and not is_formula:
                snapped = _round_delivery_qty(code, live_v) if live_v > 0 else 0
                manual.setdefault(code, {})[dt] = int(snapped) if snapped == int(snapped) else snapped
                n_cells += 1
    wb_f.close()
    wb_v.close()
    if n_skipped:
        print(f"  ⚠️  Пропущено битых/пустых формул 📦: {n_skipped}", flush=True)
    if n_cells:
        print(f"  Ручные поставки (📦): {n_cells} ячеек из {os.path.basename(out_path)}"
              f" ({n_converted} VLOOKUP → значение)", flush=True)
    return manual

_MAN_DEL_FIRST_COL = 4
_MAN_DEL_DATE_ROW = 3
_MAN_DEL_SUB_ROW = 4
_MAN_DEL_DATA_START = 5

def _parse_input_delivery_date(val, by_md):
    import datetime as _dt, re as _re
    if isinstance(val, (_dt.datetime, _dt.date)):
        return by_md.get((val.month, val.day))
    m = _re.match(r'^(\d{1,2})[.\-/](\d{1,2})', str(val).strip())
    if m:
        return by_md.get((int(m.group(2)), int(m.group(1))))
    return None

def _load_manual_deliveries_from_input_sheet(out_path):
    """Лист Ввод_Поставок (сетка: 2 колонки на день — серая авто-справка + жёлтый
    ввод). Читаем ТОЛЬКО жёлтые колонки. Колонка->дата по шапке (дд.мм), что
    надёжно при сдвиге горизонта. Значения замораживаются, приоритет над авто."""
    res = {}
    if not out_path or not os.path.exists(out_path):
        return res
    try:
        wb = load_workbook(out_path, data_only=True)
    except Exception:
        return res
    if 'Ввод_Поставок' not in wb.sheetnames:
        wb.close()
        return res
    ws = wb['Ввод_Поставок']
    by_md = {}
    for _d, _, _ in all_dates:
        by_md.setdefault((_d.month, _d.day), _d)
    col_date = {}
    for di in range(len(all_dates)):
        auto_col = _MAN_DEL_FIRST_COL + di * 2
        hv = ws.cell(_MAN_DEL_DATE_ROW, auto_col).value
        d = _parse_input_delivery_date(hv, by_md) if hv not in (None, '') else None
        if d is None:
            d = all_dates[di][0]
        col_date[auto_col + 1] = d
    n = 0
    for r in range(_MAN_DEL_DATA_START, ws.max_row + 1):
        code_raw = ws.cell(r, 1).value
        if code_raw in (None, ''):
            continue
        code = normalize_code(str(code_raw).strip())
        if not code:
            continue
        for input_col, d in col_date.items():
            if d is None:
                continue
            v = ws.cell(r, input_col).value
            if v in (None, ''):
                continue
            try:
                q = float(v)
            except (TypeError, ValueError):
                continue
            if q < 0:
                continue
            res.setdefault(code, {})[d] = int(round(q)) if float(q) == int(q) else q
            n += 1
    wb.close()
    if n:
        print(f"  Ручные поставки из Ввод_Поставок: {n} (приоритет над авто)", flush=True)
    return res

manual_delivery_by_date = _load_manual_delivery_overrides(OUT, DELIVERY_SCHEDULES)
_man_del_input_sheet = _load_manual_deliveries_from_input_sheet(OUT)
for _mc, _mdm in _man_del_input_sheet.items():
    manual_delivery_by_date.setdefault(_mc, {}).update(_mdm)
_auto_dels_backup = {c: list(all_dels[c]) for c in mrp_codes}

def _apply_manual_deliveries_to_schedules(manual_by_date, all_dels, auto_backup=None):
    """Учитываем ручные 📦 в остатке Ss и убираем дефицит, НЕ трогая ручные дни.

    Ручные поставки пишутся в колонку 📦 и сохраняются при пересчёте, но раньше
    остаток (Ss) считался по АВТО-графику и ручные значения игнорировал — из-за
    этого Ss мог показывать ложный «минус» (или прятать реальный). Теперь Ss
    пересчитывается по фактическим 📦 (авто + ручные), а дефицит закрывается
    донабором на соседних авто-днях (кратно упаковке), при этом ручные дни
    «заморожены» — их значения остаются ровно такими, как ввёл пользователь."""
    if not manual_by_date:
        return
    di_by_date = {all_dates[i][0]: i for i in range(len(all_dates))}
    frozen_del_by_code = {}
    n_codes = 0
    for code in mrp_codes:
        daymap = (manual_by_date.get(code)
                  or manual_by_date.get(normalize_code(code)) or {})
        if not daymap:
            continue
        dels = [float(x or 0) for x in all_dels.get(code, [0.0] * len(all_dates))]
        if len(dels) < len(all_dates):
            dels += [0.0] * (len(all_dates) - len(dels))
        frozen = set()
        for dt, qty in daymap.items():
            di = di_by_date.get(dt)
            if di is None:
                continue
            try:
                dels[di] = float(qty or 0)
            except (TypeError, ValueError):
                continue
            frozen.add(di)
        if not frozen:
            continue
        all_dels[code] = dels
        frozen_del_by_code[code] = frozen
        n_codes += 1
    if not n_codes:
        return
    total = _settle_deliveries(all_dels, frozen_del_by_code)
    if total > 0:
        # Ручной ввод поставок ВСЕГДА сохраняется (его дни заморожены). Если после
        # донабора на соседних авто-днях остаётся дефицит — это реальная нехватка,
        # её показываем в дашборде, но НЕ затираем то, что ввёл пользователь.
        # (Раньше здесь был откат к авто-графику с manual_by_date.clear() — из-за
        #  него ручной ввод «не сохранялся».)
        print(f"  ⚠️  Ручные 📦 заморожены; остаточный дефицит Ss: {total} "
              f"дн×код ({n_codes} кодов) — показываем в Риск_Дефицита", flush=True)
    for code in mrp_codes:
        if code not in all_dels:
            continue
        pairs = _recompute_ss_from_deliveries(code, all_dels[code])
        DELIVERY_SCHEDULES[code] = {
            'dels': [p[0] for p in pairs],
            'ss':   [p[1] for p in pairs],
            'pairs': {di: pairs[di] for di in range(len(pairs))},
        }
        for dt, qty in (manual_by_date.get(code) or manual_by_date.get(normalize_code(code)) or {}).items():
            di = di_by_date.get(dt)
            if di is not None:
                try:
                    DELIVERY_SCHEDULES[code]['dels'][di] = float(qty or 0)
                except (TypeError, ValueError):
                    pass
    if n_codes:
        print(f"  Ss пересчитан с учётом ручных 📦: {n_codes} кодов "
              f"(ручные дни заморожены, дonабor + лимиты фур)", flush=True)

_apply_manual_deliveries_to_schedules(manual_delivery_by_date, all_dels, _auto_dels_backup)

print("  Финальная сверка Ss (как в Excel)...", flush=True)
_settle_deliveries(all_dels, _frozen_manual_deliveries() if manual_delivery_by_date else {})
DELIVERY_SCHEDULES = _schedules_from_all_dels(all_dels)
_ex_neg, _ex_below = _count_excel_sim_negatives(all_dels)
if _ex_neg + _ex_below == 0:
    print(f"  ✅ Ss по цепочке Excel: 0 дефицитов", flush=True)
else:
    print(f"  ⚠️  Ss по цепочке Excel: Ss<0={_ex_neg}, Ss<potr={_ex_below}", flush=True)

# ── Переходящие остатки месяцев (согласовано с Excel-формулами листов):
#    нач. остаток М2 = расчётный остаток конца М1 из графика поставок;
#    нач. остаток М3 = М2 − потребность М2 + заказ М2. ──
_m1_days = MONTHS[0][2]
for code in mrp_codes:
    _ss_l = DELIVERY_SCHEDULES.get(code, {}).get('ss') or []
    if len(_ss_l) >= _m1_days:
        _m2_open = max(0.0, float(_ss_l[_m1_days - 1]))
        projected_stock.setdefault(code, {})[MONTHS[1][0]] = _m2_open
        _dem2 = sum(demand[code].get(MONTHS[1][0], {}).values())
        _ord2 = calc_order(code, MONTHS[1][0], stock_override=_m2_open)['order_qty']
        projected_stock[code][MONTHS[2][0]] = max(0.0, _m2_open - _dem2 + _ord2)
print(f"  Переходящие остатки: нач. {MONTH_SHORT[MONTHS[1][0]]} = конец {MONTH_SHORT[MONTHS[0][0]]} "
      f"из графика поставок; нач. {MONTH_SHORT[MONTHS[2][0]]} = остаток − потребность + заказ "
      f"{MONTH_SHORT[MONTHS[1][0]]}", flush=True)


# ── Упаковка (редактируемый) ────────────────────────────────
print("  Упаковка (редактируемый)...", flush=True)
ws_pkg = wb_out.create_sheet('Упаковка')
ws_pkg.freeze_panes = 'C3'

# Header
ws_pkg.row_dimensions[1].height = 45
t_pkg = ws_pkg.cell(1, 1)
t_pkg.value = ('УПАКОВКА — КРАТНОСТЬ ЗАКАЗА | '
               f'Источник: {os.path.basename(PKG_FILE) if PKG_FILE else "Упаковка локала.xlsx"} (Haval_Stock) | '
               '1 упаковка = 1 код (не смешивать) | отгрузка только кратно упаковке | '
               'Col C — РУЧНАЯ ПРАВКА (сохраняется между запусками) | Col F = Примечание (Lear) | '
               'Col G — авто-значение (НЕ менять)')
t_pkg.font = Font(bold=True, size=11, color="FFFFFF", name="Arial")
t_pkg.fill = fill("1F3864")
t_pkg.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
ws_pkg.merge_cells('A1:G1')

pkg_col_w = [20, 42, 14, 8, 20, 30, 10]
for ci, w in enumerate(pkg_col_w, 1):
    ws_pkg.column_dimensions[get_column_letter(ci)].width = w
ws_pkg.row_dimensions[2].height = 30

for ci, h in enumerate(['Код', 'Наименование / Применяемость / Цвет', 'Уп. (шт) ← РЕДАКТИРОВАТЬ', 'Ед.', 'Поставщик', 'Примечание', 'Уп. авто\n(не менять)'], 1):
    hcell(ws_pkg, 2, ci, h, H_FILL)

# DataValidation: только целые числа >= 1
from openpyxl.worksheet.datavalidation import DataValidation
dv_pkg = DataValidation(type="whole", operator="greaterThanOrEqual", formula1="1",
                        allow_blank=False, showDropDown=False)
dv_pkg.error = 'Введите целое число ≥ 1'
dv_pkg.errorTitle = 'Неверное значение'
ws_pkg.add_data_validation(dv_pkg)

# Write all parts with packaging
pkg_row = 3
pkg_row_map = {}  # code -> row number (for LIVE reading)
for code in all_codes:
    info = bom.get(code, {})
    pkg = info.get('package', 1)
    name = info.get('name', '')
    unit = info.get('unit', 'шт')
    supp = info.get('supplier', '')
    fb = GRY_F if pkg_row % 2 == 0 else NO_F

    ws_pkg.row_dimensions[pkg_row].height = 15

    # Col A: code (read-only style)
    c = ws_pkg.cell(pkg_row, 1); c.value = code
    c.font = Font(size=9, name="Arial", bold=True); c.alignment = Alignment(vertical='center')
    if fb: c.fill = fb

    # Col B: name
    c = ws_pkg.cell(pkg_row, 2); c.value = _name_with_appl(code, name)
    c.font = Font(size=9, name="Arial"); c.alignment = Alignment(horizontal='left', vertical='center')
    if fb: c.fill = fb

    # Col C: package qty (EDITABLE — yellow highlight)
    c = ws_pkg.cell(pkg_row, 3); c.value = int(pkg)
    c.font = Font(size=10, name="Arial", bold=True, color="1F3864")
    c.fill = fill("FFFACC")  # жёлтый = редактируемая ячейка
    c.alignment = Alignment(horizontal='center', vertical='center')
    c.number_format = '0'
    # sqref задаётся одним диапазоном после цикла

    # Col D: unit
    c = ws_pkg.cell(pkg_row, 4); c.value = unit
    c.font = Font(size=9, name="Arial"); c.alignment = Alignment(horizontal='center', vertical='center')
    if fb: c.fill = fb

    # Col E: supplier
    c = ws_pkg.cell(pkg_row, 5); c.value = supp[:25]
    c.font = Font(size=9, name="Arial"); c.alignment = Alignment(horizontal='left', vertical='center')
    if fb: c.fill = fb

    # Col F: note (EDITABLE for Lear: Пена / Подголовник)
    c = ws_pkg.cell(pkg_row, 6)
    c.value = _pkg_note_for_code(code)
    c.font = Font(size=9, name="Arial")
    c.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
    if normalize_supplier(supp) == LEAR_SUPPLIER:
        c.fill = fill("FFFACC")
    elif fb:
        c.fill = fb

    # Col G: авто-значение упаковки (база для детекта ручной правки col C)
    c = ws_pkg.cell(pkg_row, 7)
    c.value = int(pkg_base_by_code.get(code, int(pkg)))
    c.font = Font(size=9, name="Arial", color="808080")
    c.alignment = Alignment(horizontal='center', vertical='center')
    c.number_format = '0'
    if fb: c.fill = fb

    pkg_row_map[code] = pkg_row
    pkg_row += 1

# Conditional formatting: pkg=1 = grey (default), pkg>1 = green (explicitly set)
from openpyxl.formatting.rule import CellIsRule
pkg_cf_range = f'C3:C{pkg_row-1}'
dv_pkg.sqref = pkg_cf_range   # один диапазон вместо per-row конкатенации
ws_pkg.conditional_formatting.add(pkg_cf_range,
    CellIsRule(operator='greaterThan', formula=['1'],
               fill=fill("C6EFCE"), font=Font(bold=True, color="375623", name="Arial")))

print(f"    Упаковка: {len(all_codes)} деталей")

# ── Ввод_Остатков ─────────────────────────────────────────────
print("  Ввод_Остатков...")
ws_man = wb_out.create_sheet('Ввод_Остатков')
_n_date_cols_man = len(all_dates)
_total_cols_man = _MAN_STOCK_OPEN_COL + _n_date_cols_man
stock_input_date_col = {}
stock_input_row = {}
for di, (dt, _, _) in enumerate(all_dates):
    stock_input_date_col[dt] = _MAN_STOCK_FIRST_DATE_COL + di

ws_man.freeze_panes = 'F4'
ws_man.row_dimensions[1].height = 40
ws_man.row_dimensions[2].height = 16
ws_man.row_dimensions[_MAN_STOCK_HDR_ROW].height = 20

_MAN_FIX_H = ['Код детали', 'Наименование / Применяемость / Цвет', 'Поставщик', 'Ед.', 'Остаток\n(нач.мес.)\n✏️']
_MAN_FIX_W = [22, 60, 16, 6, 18]
for ci, (h, w) in enumerate(zip(_MAN_FIX_H, _MAN_FIX_W), 1):
    bg = H_FILL if ci != _MAN_STOCK_OPEN_COL else fill("B8860B")
    ws_man.merge_cells(f'{get_column_letter(ci)}2:{get_column_letter(ci)}3')
    hcell(ws_man, 2, ci, h, bg)
    _hc = ws_man.cell(2, ci)
    _hc.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    if ci == _MAN_STOCK_OPEN_COL:
        _hc.font = Font(bold=True, color="FFFFFF", size=9, name="Arial")
    ws_man.column_dimensions[get_column_letter(ci)].width = w

if stock_sources:
    src_names_display = sorted(set(n for s in stock_src.values() for n in s))
    msg = (f"✅ ОСТАТКИ ЗАГРУЖЕНЫ: {len(stock)} деталей из {len(src_names_display)} файла(ов): {', '.join(n[:20] for n in src_names_display)}"
           f"  |  ✏️ E = остаток на 1-е число первого месяца (не меняется от снимков по дням)  |  "
           f"жёлтые колонки = снимок на дату → только Ss этого дня в График_Поставок")
    fc = "1F6B00"; bg = fill("E2EFDA")
else:
    msg = ("ОСТАТКИ НЕ ЗАГРУЖЕНЫ. Положите xlsx в папку MRP  |  "
           "✏️ E — нач.мес.; жёлтые ячейки — снимок по дате (файл приоритетнее ручного ввода)")
    fc = "C00000"; bg = fill("FFD7D7")
c1 = ws_man.cell(1, 1); c1.value = msg
c1.font = Font(bold=True, size=9, color=fc, name="Arial")
c1.fill = bg; c1.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
ws_man.merge_cells(f'A1:{get_column_letter(_total_cols_man)}1')

for mnum_m, mlabel_m, ndays_m in MONTHS:
    first_di = sum(nd for mn, _, nd in MONTHS if mn < mnum_m)
    last_di = first_di + ndays_m - 1
    col_first = _MAN_STOCK_FIRST_DATE_COL + first_di
    col_last = _MAN_STOCK_FIRST_DATE_COL + last_di
    mc = ws_man.cell(2, col_first)
    mc.value = mlabel_m
    mc.font = Font(bold=True, size=10, color="FFFFFF", name="Arial")
    mc.fill = M_FILL[mnum_m]
    mc.alignment = Alignment(horizontal='center', vertical='center')
    ws_man.merge_cells(f'{get_column_letter(col_first)}2:{get_column_letter(col_last)}2')

for di, (dt, mnum_h, _) in enumerate(all_dates):
    ci = _MAN_STOCK_FIRST_DATE_COL + di
    c_sh = ws_man.cell(_MAN_STOCK_HDR_ROW, ci)
    c_sh.value = dt.strftime('%d.%m')
    c_sh.font = Font(bold=True, color="FFFFFF", size=8, name="Arial")
    c_sh.fill = M_FILL[mnum_h]
    c_sh.alignment = Alignment(horizontal='center', vertical='center')
    ws_man.column_dimensions[get_column_letter(ci)].width = 5.5

for ri, code in enumerate(all_codes, _MAN_STOCK_DATA_START):
    stock_input_row[code] = ri
    ws_man.row_dimensions[ri].height = 14
    fb = GRY_F if ri % 2 == 0 else NO_F
    name, supp, unit = get_info(code)
    stk = _opening_stock(code)
    snaps = manual_stock_by_date.get(code, {})
    for ci, v in enumerate([code, _name_with_appl(code,name), supp, unit, stk], 1):
        c = ws_man.cell(ri, ci); c.value = v
        c.font = Font(size=9, name="Arial")
        c.alignment = Alignment(horizontal='left' if ci <= 2 else 'center', vertical='center')
        if ci == _MAN_STOCK_OPEN_COL:
            c.fill = YEL_F; c.number_format = '#,##0.##'
        elif fb:
            c.fill = fb
    for di, (dt, _, _) in enumerate(all_dates):
        ci = _MAN_STOCK_FIRST_DATE_COL + di
        snap_v = snaps.get(dt)
        c = ws_man.cell(ri, ci)
        if snap_v is not None:
            c.value = snap_v
        c.fill = YEL_F
        c.number_format = '#,##0.##'
        c.font = Font(size=9, name="Arial")
        c.alignment = Alignment(horizontal='center', vertical='center')

# ── Ввод_Поставок (сетка: серая авто-справка + жёлтый ввод; фильтр по поставщику) ─
print("  Ввод_Поставок...")
ws_md = wb_out.create_sheet('Ввод_Поставок')
_n_del_days = len(all_dates)
_md_total_cols = (_MAN_DEL_FIRST_COL - 1) + _n_del_days * 2
ws_md.row_dimensions[1].height = 42
ws_md.row_dimensions[_MAN_DEL_DATE_ROW].height = 15
ws_md.row_dimensions[_MAN_DEL_SUB_ROW].height = 14
_md_msg = ("РУЧНЫЕ ПОСТАВКИ: серые колонки «авто» — справка из График_Поставок (не редактировать). "
           "Впишите количество в ЖЁЛТУЮ колонку нужного дня — значение ЗАМОРАЖИВАЕТСЯ и имеет приоритет "
           "над авто-графиком. Пусто = берётся авто. Фильтр по поставщику — стрелка в шапке столбца «Поставщик». "
           "Убрать ручную поставку — очистите жёлтую ячейку и пересчитайте.")
_mc1 = ws_md.cell(1, 1)
_mc1.value = _md_msg
_mc1.font = Font(bold=True, size=9, color="1F6B00", name="Arial")
_mc1.fill = fill("E2EFDA")
_mc1.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
ws_md.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max(_md_total_cols, 4))
ws_md.merge_cells(start_row=2, start_column=1, end_row=_MAN_DEL_DATE_ROW, end_column=3)
_cdh = ws_md.cell(2, 1, 'Деталь / Поставщик')
_cdh.font = Font(bold=True, color="FFFFFF", size=9, name="Arial")
_cdh.fill = H_FILL
_cdh.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
for _ci, (_h, _w) in enumerate([('Код детали', 22), ('Наименование / Применяемость', 46), ('Поставщик', 16)], 1):
    _c = ws_md.cell(_MAN_DEL_SUB_ROW, _ci, _h)
    _c.font = Font(bold=True, color="FFFFFF", size=9, name="Arial")
    _c.fill = H_FILL
    _c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws_md.column_dimensions[get_column_letter(_ci)].width = _w
for _mn, _ml, _nd in MONTHS:
    _fdi = sum(nd for mn, _, nd in MONTHS if mn < _mn)
    _ldi = _fdi + _nd - 1
    _cf = _MAN_DEL_FIRST_COL + _fdi * 2
    _cl = _MAN_DEL_FIRST_COL + _ldi * 2 + 1
    _mc = ws_md.cell(2, _cf, _ml)
    _mc.font = Font(bold=True, size=10, color="FFFFFF", name="Arial")
    _mc.fill = M_FILL[_mn]
    _mc.alignment = Alignment(horizontal='center', vertical='center')
    ws_md.merge_cells(start_row=2, start_column=_cf, end_row=2, end_column=_cl)
for di, (dt, mnum_h, _) in enumerate(all_dates):
    auto_col = _MAN_DEL_FIRST_COL + di * 2
    in_col = auto_col + 1
    _cd = ws_md.cell(_MAN_DEL_DATE_ROW, auto_col, dt.strftime('%d.%m'))
    _cd.font = Font(bold=True, color="FFFFFF", size=8, name="Arial")
    _cd.fill = M_FILL[mnum_h]
    _cd.alignment = Alignment(horizontal='center', vertical='center')
    ws_md.merge_cells(start_row=_MAN_DEL_DATE_ROW, start_column=auto_col, end_row=_MAN_DEL_DATE_ROW, end_column=in_col)
    _ca = ws_md.cell(_MAN_DEL_SUB_ROW, auto_col, 'авто')
    _ca.font = Font(size=7, name="Arial", color="888888")
    _ca.fill = GRY_F
    _ca.alignment = Alignment(horizontal='center', vertical='center')
    _cx = ws_md.cell(_MAN_DEL_SUB_ROW, in_col, '✏️')
    _cx.font = Font(size=7, name="Arial", color="9C5700")
    _cx.fill = _GP_MANUAL_DEL_FILL
    _cx.alignment = Alignment(horizontal='center', vertical='center')
    ws_md.column_dimensions[get_column_letter(auto_col)].width = 5.0
    ws_md.column_dimensions[get_column_letter(in_col)].width = 5.5
ws_md.freeze_panes = f"{get_column_letter(_MAN_DEL_FIRST_COL)}{_MAN_DEL_DATA_START}"
for _ri, code in enumerate(mrp_codes, _MAN_DEL_DATA_START):
    ws_md.row_dimensions[_ri].height = 13
    try:
        _nm, _supp, _u = get_info(code)
    except Exception:
        _nm, _supp = '', ''
    _fb = GRY_F if _ri % 2 == 0 else NO_F
    _c0 = ws_md.cell(_ri, 1, code); _c0.font = Font(size=8, name="Arial")
    _c1 = ws_md.cell(_ri, 2, _name_with_appl(code, _nm)); _c1.font = Font(size=8, name="Arial")
    _c1.alignment = Alignment(horizontal='left', vertical='center')
    _c2 = ws_md.cell(_ri, 3, _supp); _c2.font = Font(size=8, name="Arial")
    _c2.alignment = Alignment(horizontal='center', vertical='center')
    if _fb:
        _c0.fill = _fb; _c1.fill = _fb; _c2.fill = _fb
    _dels = DELIVERY_SCHEDULES.get(code, {}).get('dels', [])
    _man = manual_delivery_by_date.get(code) or manual_delivery_by_date.get(normalize_code(code)) or {}
    for di, (dt, _, _) in enumerate(all_dates):
        auto_col = _MAN_DEL_FIRST_COL + di * 2
        in_col = auto_col + 1
        _av = _dels[di] if di < len(_dels) else 0
        _ca = ws_md.cell(_ri, auto_col)
        if _av and round(_av) > 0:
            _ca.value = int(round(_av))
        _ca.font = Font(size=8, name="Arial", color="555555")
        _ca.fill = GRY_F
        _ca.number_format = '#,##0'
        _ca.alignment = Alignment(horizontal='center', vertical='center')
        _cx = ws_md.cell(_ri, in_col)
        _mv = _man.get(dt)
        if _mv is not None:
            try:
                _mvf = float(_mv)
                _cx.value = int(round(_mvf)) if _mvf == int(_mvf) else _mvf
            except (TypeError, ValueError):
                pass
        _cx.fill = _GP_MANUAL_DEL_FILL
        _cx.number_format = '#,##0'
        _cx.font = Font(size=8, bold=True, name="Arial", color="9C5700")
        _cx.alignment = Alignment(horizontal='center', vertical='center')
_md_last_row = _MAN_DEL_DATA_START + len(mrp_codes) - 1
ws_md.auto_filter.ref = f"A{_MAN_DEL_SUB_ROW}:C{max(_md_last_row, _MAN_DEL_DATA_START)}"

# ── Precompute: Ss-колонка последнего дня месяца в График_Поставок ─────────────
# График_Поставок: col A=Код, для дня di:
#   Ss-col  = 8 + di*2       (дата дд.мм, остаток)
#   Del-col = 8 + di*2 + 1   (📦 поставка — только если есть спрос в этот день)
_gp_last_mon_col = {}  # mnum -> (col_letter, col_num) Ss-колонки последнего дня месяца
for _di, (_dt, _mn, _d) in enumerate(all_dates):
    _cn = 8 + _di * 3
    _gp_last_mon_col[_mn] = (get_column_letter(_cn), _cn)
# После цикла: _gp_last_mon_col[5] = Ss последнего дня мая (May 31)
# Для июля: используем Потребность_Jun col I (Дефицит = остаток после июньского спроса)

# ── Потребность_May/Jun/Jul ───────────────────────────────────
FH=['Код детали','Наименование / Применяемость / Цвет','Поставщик','Ед.','Вкладка плана',
    'Тип / норма','Остаток','Потребность','Дефицит']
FW=[22,60,16,6,22,30,10,14,12]
for mi,(mnum,mlabel,n_days) in enumerate(MONTHS):
    sname=f"Потребность_{mlabel[:3]}"
    print(f"  {sname}...")
    ws=wb_out.create_sheet(sname)
    ws.freeze_panes='E3'   # A–D + строки 1–2; дни месяца прокручиваются вправо
    ws.row_dimensions[1].height=45; ws.row_dimensions[2].height=14
    for ci,(h,w) in enumerate(zip(FH,FW),1):
        hcell(ws,1,ci,h,H_FILL); ws.column_dimensions[get_column_letter(ci)].width=w
    for d in range(1,n_days+1):
        ci=len(FH)+d; hcell(ws,1,ci,str(d),M_FILL[mnum])
        ws.column_dimensions[get_column_letter(ci)].width=6.5
    _note2 = "Только детали с ✓ в BOM_Детальный (+ химия/Additional). Без применяемости — не в расчёте."
    if mi == 0:
        # 1-й месяц: потребность = ОСТАТОК потребности с даты предоставления остатков
        _hH = ws.cell(1, 8); _hH.value = 'Потребность\n(с даты остатков)'
        _rc_col = len(FH) + n_days + 1
        hcell(ws, 1, _rc_col, 'Расчёт\nс даты', fill("B8860B"))
        ws.column_dimensions[get_column_letter(_rc_col)].width = 10
        _note2 = ("Потребность 1-го месяца = остаток потребности С ДАТЫ ПРЕДОСТАВЛЕНИЯ ОСТАТКОВ "
                  "по каждому коду (последняя колонка «Расчёт с даты»). " + _note2)
    ws.cell(2,1).value=_note2
    ws.cell(2,1).font=Font(italic=True,size=8,color="555555",name="Arial")
    ws.merge_cells(f'A2:{get_column_letter(len(FH)+n_days)}2')
    for ri,code in enumerate(mrp_codes,3):
        ws.row_dimensions[ri].height=14
        fb=GRY_F if ri%2==0 else NO_F
        name,supp,unit=get_info(code)
        daily=demand[code].get(mnum,{})
        # 1-й месяц: остаток потребности с даты предоставления остатков кода
        mtotal = _remaining_demand_for_month(code, mnum) if mi == 0 else sum(daily.values())
        stk=_opening_stock(code); tab=part_tab_map.get(code,'—'); norm_str=get_norm_str(code)
        vals=[code,_name_with_appl(code,name),supp,unit,tab,norm_str,stk,
              round(mtotal, 2), None]
        for ci,v in enumerate(vals,1):
            c=ws.cell(ri,ci); c.value=v
            c.font=Font(size=9,name="Arial")
            c.alignment=Alignment(horizontal='left' if ci<=2 else 'center',vertical='center')
            if ci==6: c.fill=NRM_F
            elif ci==7: c.fill=CYN_F; c.number_format='#,##0.##'
            elif ci==8:
                if mtotal:
                    c.fill=BLU_F; c.font=Font(bold=True,size=9,name="Arial")
                c.number_format='#,##0' if unit in ('шт','pcs','шт.') else '#,##0.##'
            elif fb and ci not in (6,7,8): c.fill=fb
        # col G (Остаток): для текущего месяца — из Ввод_Остатков; для Jun/Jul — из График_Поставок
        if mnum == STOCK_AS_OF_MONTH:
            _cg = ws.cell(ri, 7)
            _cg.value = f"=IFERROR(VLOOKUP(A{ri},{_xlsheet_ref('Ввод_Остатков', '$A:$E')},5,0),{int(stk)})"
            _cg.fill = CYN_F; _cg.number_format = '#,##0.##'
            _cg.font = Font(size=9, name="Arial")
            _cg.alignment = Alignment(horizontal='center', vertical='center')
        elif mi == 1 and MONTHS[0][0] in _gp_last_mon_col:
            # Месяц 2: переходящий остаток = последний день месяца 1 из График_Поставок
            _prev_col_l, _prev_col_n = _gp_last_mon_col[MONTHS[0][0]]
            _cg = ws.cell(ri, 7)
            _pkg_r = max(1, bom.get(code, {}).get('package', 1))
            _cg.value = _graph_stock_formula(ri, _prev_col_l, _prev_col_n, _pkg_r, int(stk))
            _cg.fill = CYN_F; _cg.number_format = '#,##0.##'
            _cg.font = Font(size=9, name="Arial")
            _cg.alignment = Alignment(horizontal='center', vertical='center')
        elif mi == 2:
            # Месяц 3: нач. остаток = Остаток М2 − Потребность М2 + Заказ М2 (из Заказы_М2)
            _zk2 = _xlsheet_ref('Заказы_' + MONTH_SHORT[MONTHS[1][0]], '$A:$K')
            _cg = ws.cell(ri, 7)
            _cg.value = (f"=MAX(0,IFERROR(VLOOKUP(A{ri},{_zk2},6,0),{int(stk)})"
                         f"-IFERROR(VLOOKUP(A{ri},{_zk2},7,0),0)"
                         f"+IFERROR(VLOOKUP(A{ri},{_zk2},11,0),0))")
            _cg.fill = CYN_F; _cg.number_format = '#,##0.##'
            _cg.font = Font(size=9, name="Arial")
            _cg.alignment = Alignment(horizontal='center', vertical='center')
        dc=ws.cell(ri,9); dc.value=f"=G{ri}-IF(ISBLANK(H{ri}),0,H{ri})"
        dc.number_format='#,##0.#'; dc.font=Font(size=9,name="Arial")
        dc.alignment=Alignment(horizontal='center',vertical='center')
        for d in range(1,n_days+1):
            ci=len(FH)+d; c=ws.cell(ri,ci); q=daily.get(d)
            if q and q>0:
                c.value=int(round(q)) if unit in ('шт','pcs','шт.') else round(q,2)
                c.number_format='#,##0' if unit in ('шт','pcs','шт.') else '#,##0.##'
                c.font=Font(size=9,name="Arial")
            if fb: c.fill=fb
            c.alignment=Alignment(horizontal='center',vertical='center')
        if mi == 0:
            # отметка: с какой даты считается оставшаяся потребность кода
            c = ws.cell(ri, len(FH) + n_days + 1)
            c.value = _calc_from_date(code).strftime('%d.%m.%Y')
            c.font = Font(size=8, name="Arial", color="B8860B", bold=True)
            c.alignment = Alignment(horizontal='center', vertical='center')

# ── Нормы_Расхода (г/кг материалы) ──────────────────────────
print("  Нормы_Расхода...")

def _chem_cars(code, month_num, fut_only=False):
    """Кол-во применяемых авт для chem-материала в данном месяце."""
    norms_c = chem_norms[code]
    applicable_models = set(m for m, n in norms_c.items() if n > 0)
    color_filter = paint_colors.get(code, False)
    if isinstance(color_filter, list) and color_filter:
        # Краска с цветовым фильтром: авт берём обратным расчётом из потребности
        total_d = sum(demand[code].get(month_num, {}).values())
        first_norm = list(norms_c.values())[0] if norms_c else 1
        if first_norm == 0: return 0
        if fut_only:
            fut_d = _remaining_demand_for_month(code, month_num)
            return round(fut_d / first_norm, 1)
        return round(total_d / first_norm, 1)
    plan_tab = part_tab_map.get(code, '')
    tabs = resolve_plan_tabs(plan_tab)
    total = 0
    for tab in tabs:
        if tab not in plan_batches or month_num not in plan_batches[tab]: continue
        for bv, day_qty in plan_batches[tab][month_num].items():
            bi = _get_batch_info(tab, bv)
            if bi.get('model', '') not in applicable_models: continue
            for day, cars in day_qty.items():
                if fut_only and month_num == _first_calc_month_num():
                    _sd_c = _calc_from_date(code)
                    if (_sd_c.month == month_num and day < _sd_c.day):
                        continue
                total += cars
    return round(total, 1)

CHEM_CODES = [c for c in all_codes if c in chem_norms]
norms_rows = {}
for _code in CHEM_CODES:
    _row = {}
    for _mn, _, _ in MONTHS:
        _total_d = sum(demand[_code].get(_mn, {}).values())
        _cars = _chem_cars(_code, _mn)
        _blended = round(_total_d / _cars, 6) if _cars > 0 else (list(chem_norms[_code].values())[0] if chem_norms[_code] else 0)
        _row[f'norm_{_mn}'] = _blended
        _row[f'cars_{_mn}'] = _cars
        if _mn == _first_calc_month_num():
            _row['cars_may_fut'] = _chem_cars(_code, _mn, fut_only=True)
    norms_rows[_code] = _row

_m0, _m0lbl, _ = MONTHS[0]; _m1, _m1lbl, _ = MONTHS[1]; _m2, _m2lbl, _ = MONTHS[2]
_ms0 = MONTH_SHORT[_m0]; _ms1 = MONTH_SHORT[_m1]; _ms2 = MONTH_SHORT[_m2]
NRM_HDR = ['Код','Наименование / Применяемость / Цвет','Ед.','Поставщик',
           f'Норма\n{_ms0}\n(ред.)',f'Норма\n{_ms1}\n(ред.)',f'Норма\n{_ms2}\n(ред.)',
           f'Авт_{_ms0}\n(полн)',f'Авт_{_ms0}\n(ост.)',
           f'Авт_{_ms1}',f'Авт_{_ms2}']
NRM_W   = [22, 60, 6, 18, 10, 10, 10, 10, 10, 10, 10]
ws_n = wb_out.create_sheet('Нормы_Расхода')
ws_n.freeze_panes = 'E3'; ws_n.row_dimensions[1].height = 40; ws_n.row_dimensions[2].height = 40
_t = ws_n.cell(1, 1)
_t.value = "НОРМЫ РАСХОДА | г/кг материалы | Колонки «Норма» (жёлтые) — редактируемые; пересчёт потребности — в листах Заказы"
_t.font = Font(bold=True, size=11, color="FFFFFF", name="Arial")
_t.fill = H_FILL; _t.alignment = Alignment(horizontal='center', vertical='center')
ws_n.merge_cells(f'A1:{get_column_letter(len(NRM_HDR))}1')
EDIT_F = fill("FFF2CC")
STAT_F = fill("EAF4FB")
for ci, (h, w) in enumerate(zip(NRM_HDR, NRM_W), 1):
    hcell(ws_n, 2, ci, h, H_FILL)
    ws_n.column_dimensions[get_column_letter(ci)].width = w
for ri, _code in enumerate(CHEM_CODES, 3):
    ws_n.row_dimensions[ri].height = 14
    _name, _supp, _unit = get_info(_code)
    _fb = GRY_F if ri % 2 == 0 else NO_F
    _nr = norms_rows[_code]
    _m0n, _m1n, _m2n = MONTHS[0][0], MONTHS[1][0], MONTHS[2][0]
    _vals = [_code, _name_with_appl(_code,_name), _unit, _supp,
             _nr[f'norm_{_m0n}'], _nr[f'norm_{_m1n}'], _nr[f'norm_{_m2n}'],
             _nr[f'cars_{_m0n}'], _nr.get('cars_may_fut', _nr[f'cars_{_m0n}']),
             _nr[f'cars_{_m1n}'], _nr[f'cars_{_m2n}']]
    for ci, v in enumerate(_vals, 1):
        c = ws_n.cell(ri, ci)
        c.value = round(v, 4) if isinstance(v, float) else v
        c.font = Font(size=9, name="Arial")
        c.alignment = Alignment(horizontal='left' if ci <= 2 else 'center', vertical='center')
        if ci in (5, 6, 7):
            c.fill = EDIT_F; c.number_format = '#,##0.0000'
        elif ci in (8, 9, 10, 11):
            c.fill = STAT_F; c.number_format = '#,##0.#'
        elif _fb:
            c.fill = _fb
_tr = len(CHEM_CODES) + 3
hcell(ws_n, _tr, 1, 'ИТОГО авт', H_FILL); ws_n.merge_cells(f'A{_tr}:D{_tr}')
for _ci in range(8, 12):
    _c = ws_n.cell(_tr, _ci)
    _c.value = f"=SUM({get_column_letter(_ci)}3:{get_column_letter(_ci)}{_tr-1})"
    _c.font = Font(bold=True, size=9, color="FFFFFF", name="Arial"); _c.fill = H_FILL
    _c.number_format = '#,##0.#'; _c.alignment = Alignment(horizontal='center', vertical='center')

# ── Заказы_May/Jun/Jul (опционально: MRP_CALC_ORDERS) ────────
if not CALC_ORDERS:
    print("  Заказы_*: расчёт отключён (включается галочкой «Месячные заказы»)")
    for _mn_o, _ml_o, _ in MONTHS:
        _ws_stub = wb_out.create_sheet(f"Заказы_{_ml_o[:3]}")
        _c_stub = _ws_stub.cell(1, 1)
        _c_stub.value = ("Расчёт месячных заказов отключён. Включите галочку "
                         "«Месячные заказы» в run_mrp_gui (или MRP_CALC_ORDERS=1) и перезапустите.")
        _c_stub.font = Font(bold=True, size=11, color="C00000", name="Arial")
        _ws_stub.column_dimensions['A'].width = 110
for mi_ord,(mnum,mlabel,n_days) in enumerate(MONTHS if CALC_ORDERS else []):
    oname=f"Заказы_{mlabel[:3]}"; print(f"  {oname}...")
    ws_o=wb_out.create_sheet(oname)
    ws_o.freeze_panes='J3'; ws_o.row_dimensions[1].height=50
    oh=['Код','Наименование / Применяемость / Цвет','Поставщик','Ед.','Уп.\n(шт)','Остаток',
        f'Потребн.\n{mlabel[:3]}','Страх.\nзапас','Чистая\nпотребн.',
        'Кол-во\nупаковок','ЗАКАЗ\n(итого)','Статус']
    ow=[22,60,16,6,7,10,14,12,14,10,14,18]
    if mi_ord == 0:
        oh = oh + ['Расчёт\nс даты']
        ow = ow + [10]
    for ci,(h,w) in enumerate(zip(oh,ow),1):
        hcell(ws_o,1,ci,h,H_FILL); ws_o.column_dimensions[get_column_letter(ci)].width=w
    if mi_ord == 0:
        _o_note = (f"Заказ=CEILING((Потребность+Страх_запас-Остаток)/Упаковка)×Упаковка | "
                   f"Потребность = ОСТАТОК потребности с даты предоставления остатков кода (col M)")
    elif mi_ord == 1:
        _o_note = (f"Заказ=CEILING((Потребность+Страх_запас-Остаток)/Упаковка)×Упаковка | "
                   f"Остаток = расчётный остаток конца {MONTH_SHORT[MONTHS[0][0]]} из График_Поставок")
    else:
        _o_note = (f"Заказ=CEILING((Потребность+Страх_запас-Остаток)/Упаковка)×Упаковка | "
                   f"Остаток = Остаток {MONTH_SHORT[MONTHS[1][0]]} − Потребность {MONTH_SHORT[MONTHS[1][0]]} + Заказ {MONTH_SHORT[MONTHS[1][0]]}")
    ws_o.cell(2,1).value=_o_note
    ws_o.cell(2,1).font=Font(italic=True,size=9,color="555555",name="Arial")
    ws_o.merge_cells('A2:L2')
    for ri,code in enumerate(mrp_codes,3):
        ws_o.row_dimensions[ri].height=14
        fb=GRY_F if ri%2==0 else NO_F
        name,supp,unit=get_info(code)
        # Для июня/июля используем прогнозный остаток на 1-е число месяца
        proj_stk = projected_stock.get(code, {}).get(mnum) if mnum != STOCK_AS_OF_MONTH else None
        oi=calc_order(code, mnum, stock_override=proj_stk)
        dem=oi['demand']; stk=oi['stock']; saf=oi['safety']
        net=oi['net']; pkg=oi['pkg_size']; pkgs=oi['packages']; oq=oi['order_qty']
        if dem==0: st='—'
        elif oq==0: st='✅ Достаточно'
        elif stk==0: st='🔴 НЕТ ОСТАТКА'
        else: st='📦 Заказ'
        vals=[code,_name_with_appl(code,name),supp,unit,pkg,stk,dem,saf,net,pkgs,oq,st]
        fmts=[None,None,None,None,'#,##0','#,##0.#','#,##0.#','#,##0.#','#,##0.#','#,##0','#,##0.#',None]
        for ci,(v,fmt) in enumerate(zip(vals,fmts),1):
            c=ws_o.cell(ri,ci); c.value=v
            c.font=Font(size=9,name="Arial",bold=(ci==11))
            c.alignment=Alignment(horizontal='left' if ci<=2 else 'center',vertical='center')
            if fmt: c.number_format=fmt
            if ci==5: c.fill=ORG_F
            elif ci==6: c.fill=CYN_F
            elif ci==7 and dem: c.fill=BLU_F; c.font=Font(bold=True,size=9,name="Arial")
            elif ci==11: c.fill=fill("E8F5E9") if oq>0 else (fb or NO_F)
            elif ci==12:
                if 'НЕТ' in str(v): c.fill=RED_F; c.font=Font(size=9,bold=True,name="Arial",color="C00000")
                elif '✅' in str(v): c.fill=fill("E2EFDA")
                elif '📦' in str(v): c.fill=BLU_F
                elif fb: c.fill=fb
            elif fb and ci not in (5,6,7,11,12): c.fill=fb
        if mi_ord == 0:
            # отметка: дата предоставления остатков кода — от неё остаток потребности
            _cM = ws_o.cell(ri, 13)
            _cM.value = _calc_from_date(code).strftime('%d.%m.%Y')
            _cM.font = Font(size=8, name="Arial", color="B8860B", bold=True)
            _cM.alignment = Alignment(horizontal='center', vertical='center')
        # ── Формулы для авто-пересчёта при ручном вводе остатка (Ввод_Остатков col E) ──
        # Для текущего месяца (May): col F = VLOOKUP, col I/J/K/L = Excel-формулы
        if mnum == _first_calc_month_num() and code not in chem_norms:
            # Оставшаяся потребность с даты запуска расчёта (CALC_RUN_DATE)
            d_future = _remaining_demand_for_month(code, mnum)
            # F: остаток из Ввод_Остатков (при ручном изменении — пересчёт мгновенный)
            _cF = ws_o.cell(ri, 6)
            _cF.value = f"=IFERROR(VLOOKUP(A{ri},{_xlsheet_ref('Ввод_Остатков', '$A:$E')},5,0),0)"
            _cF.font = Font(size=9, name="Arial"); _cF.fill = CYN_F
            _cF.number_format = '#,##0.#'; _cF.alignment = Alignment(horizontal='center', vertical='center')
            # I: чистая потребность = оставшаяся_потребность + страх.запас - остаток
            _cI = ws_o.cell(ri, 9)
            _cI.value = f"=MAX(0,{round(d_future,2)}+H{ri}-F{ri})"
            _cI.font = Font(size=9, name="Arial"); _cI.number_format = '#,##0.#'
            _cI.alignment = Alignment(horizontal='center', vertical='center')
            if fb: _cI.fill = fb
            # J: кол-во упаковок
            _cJ = ws_o.cell(ri, 10)
            _cJ.value = f"=IF(I{ri}>0,CEILING(I{ri}/E{ri},1),0)"
            _cJ.font = Font(size=9, name="Arial"); _cJ.number_format = '#,##0'
            _cJ.alignment = Alignment(horizontal='center', vertical='center')
            if fb: _cJ.fill = fb
            # K: ЗАКАЗ итого
            _cK = ws_o.cell(ri, 11)
            _cK.value = f"=J{ri}*E{ri}"
            _cK.font = Font(bold=True, size=9, name="Arial"); _cK.number_format = '#,##0.#'
            _cK.alignment = Alignment(horizontal='center', vertical='center')
            _cK.fill = fill("E8F5E9") if oq > 0 else (fb or NO_F)
            # L: статус
            _cL = ws_o.cell(ri, 12)
            _cL.value = f'=IF(G{ri}=0,"—",IF(K{ri}=0,"✅ Достаточно",IF(F{ri}=0,"🔴 НЕТ ОСТАТКА","📦 Заказ")))'
            _cL.font = Font(size=9, name="Arial")
            _cL.alignment = Alignment(horizontal='center', vertical='center')
            if fb: _cL.fill = fb

        # ── Месяц 2 (non-chem): переходящий остаток = конец месяца 1 из График_Поставок ──
        elif mi_ord == 1 and code not in chem_norms and MONTHS[0][0] in _gp_last_mon_col:
            _prev_col_l, _prev_col_n = _gp_last_mon_col[MONTHS[0][0]]
            _cF = ws_o.cell(ri, 6)
            _pkg_o = max(1, bom.get(code, {}).get('package', 1))
            _cF.value = _graph_stock_formula(ri, _prev_col_l, _prev_col_n, _pkg_o, max(0, int(stk)))
            _cF.font = Font(size=9, name="Arial"); _cF.fill = CYN_F
            _cF.number_format = '#,##0.#'; _cF.alignment = Alignment(horizontal='center', vertical='center')
            # I: чистая потребность = спрос + страх.запас - остаток
            _cI = ws_o.cell(ri, 9)
            _cI.value = f"=MAX(0,G{ri}+H{ri}-F{ri})"
            _cI.font = Font(size=9, name="Arial"); _cI.number_format = '#,##0.#'
            _cI.alignment = Alignment(horizontal='center', vertical='center')
            if fb: _cI.fill = fb
            # J: упаковки
            _cJ = ws_o.cell(ri, 10)
            _cJ.value = f"=IF(I{ri}>0,CEILING(I{ri}/E{ri},1),0)"
            _cJ.font = Font(size=9, name="Arial"); _cJ.number_format = '#,##0'
            _cJ.alignment = Alignment(horizontal='center', vertical='center')
            if fb: _cJ.fill = fb
            # K: заказ итого
            _cK = ws_o.cell(ri, 11)
            _cK.value = f"=J{ri}*E{ri}"
            _cK.font = Font(bold=True, size=9, name="Arial"); _cK.number_format = '#,##0.#'
            _cK.alignment = Alignment(horizontal='center', vertical='center')
            _cK.fill = fill("E8F5E9") if oq > 0 else (fb or NO_F)
            # L: статус
            _cL = ws_o.cell(ri, 12)
            _cL.value = f'=IF(G{ri}=0,"—",IF(K{ri}=0,"✅ Достаточно",IF(F{ri}=0,"🔴 НЕТ ОСТАТКА","📦 Заказ")))'
            _cL.font = Font(size=9, name="Arial")
            _cL.alignment = Alignment(horizontal='center', vertical='center')
            if fb: _cL.fill = fb

        # ── Месяц 3 (non-chem): нач. остаток = Остаток М2 − Потребность М2 + Заказ М2 ──
        elif mi_ord == 2 and code not in chem_norms:
            _zk2 = _xlsheet_ref('Заказы_' + MONTH_SHORT[MONTHS[1][0]], '$A:$K')
            _cF = ws_o.cell(ri, 6)
            _cF.value = (f"=MAX(0,IFERROR(VLOOKUP(A{ri},{_zk2},6,0),{max(0,int(stk))})"
                         f"-IFERROR(VLOOKUP(A{ri},{_zk2},7,0),0)"
                         f"+IFERROR(VLOOKUP(A{ri},{_zk2},11,0),0))")
            _cF.font = Font(size=9, name="Arial"); _cF.fill = CYN_F
            _cF.number_format = '#,##0.#'; _cF.alignment = Alignment(horizontal='center', vertical='center')
            _cI = ws_o.cell(ri, 9)
            _cI.value = f"=MAX(0,G{ri}+H{ri}-F{ri})"
            _cI.font = Font(size=9, name="Arial"); _cI.number_format = '#,##0.#'
            _cI.alignment = Alignment(horizontal='center', vertical='center')
            if fb: _cI.fill = fb
            _cJ = ws_o.cell(ri, 10)
            _cJ.value = f"=IF(I{ri}>0,CEILING(I{ri}/E{ri},1),0)"
            _cJ.font = Font(size=9, name="Arial"); _cJ.number_format = '#,##0'
            _cJ.alignment = Alignment(horizontal='center', vertical='center')
            if fb: _cJ.fill = fb
            _cK = ws_o.cell(ri, 11)
            _cK.value = f"=J{ri}*E{ri}"
            _cK.font = Font(bold=True, size=9, name="Arial"); _cK.number_format = '#,##0.#'
            _cK.alignment = Alignment(horizontal='center', vertical='center')
            _cK.fill = fill("E8F5E9") if oq > 0 else (fb or NO_F)
            _cL = ws_o.cell(ri, 12)
            _cL.value = f'=IF(G{ri}=0,"—",IF(K{ri}=0,"✅ Достаточно",IF(F{ri}=0,"🔴 НЕТ ОСТАТКА","📦 Заказ")))'
            _cL.font = Font(size=9, name="Arial")
            _cL.alignment = Alignment(horizontal='center', vertical='center')
            if fb: _cL.fill = fb

        # ── Формулы для г/кг материалов (пересчёт при смене нормы) ──
        if code in chem_norms:
            NR = 'Нормы_Расхода!$A:$K'
            _nc = 5 + mi_ord           # норма-колонка (col5=M0, col6=M1, col7=M2)
            _cc = 8 if mi_ord == 0 else (10 if mi_ord == 1 else 11)  # авт_полн
            _fc = 9 if mi_ord == 0 else (10 if mi_ord == 1 else 11)  # авт_для_net
            def _vl(col, r=ri): return f"VLOOKUP($A{r},{NR},{col},0)"
            def _vls(col, r=ri): return f"IFERROR(VLOOKUP($A{r},{NR},{col},0),0)"
            # F6 = остаток: май → Ввод_Остатков; июн → конец мая из График_Поставок; июл → Потребность_Jun col I
            if mnum == STOCK_AS_OF_MONTH:
                _cF = ws_o.cell(ri, 6)
                _cF.value = f"=IFERROR(VLOOKUP(A{ri},{_xlsheet_ref('Ввод_Остатков', '$A:$E')},5,0),0)"
                _cF.font = Font(size=9, name="Arial"); _cF.fill = CYN_F
                _cF.number_format = '#,##0.#'; _cF.alignment = Alignment(horizontal='center', vertical='center')
            elif mi_ord == 1 and MONTHS[0][0] in _gp_last_mon_col:
                _prev_col_l, _prev_col_n = _gp_last_mon_col[MONTHS[0][0]]
                _cF = ws_o.cell(ri, 6)
                _pkg_ch = max(1, bom.get(code, {}).get('package', 1))
                _cF.value = _graph_stock_formula(ri, _prev_col_l, _prev_col_n, _pkg_ch, max(0, int(stk)))
                _cF.font = Font(size=9, name="Arial"); _cF.fill = CYN_F
                _cF.number_format = '#,##0.#'; _cF.alignment = Alignment(horizontal='center', vertical='center')
            elif mi_ord == 2:
                _zk2 = _xlsheet_ref('Заказы_' + MONTH_SHORT[MONTHS[1][0]], '$A:$K')
                _cF = ws_o.cell(ri, 6)
                _cF.value = (f"=MAX(0,IFERROR(VLOOKUP(A{ri},{_zk2},6,0),{max(0,int(stk))})"
                             f"-IFERROR(VLOOKUP(A{ri},{_zk2},7,0),0)"
                             f"+IFERROR(VLOOKUP(A{ri},{_zk2},11,0),0))")
                _cF.font = Font(size=9, name="Arial"); _cF.fill = CYN_F
                _cF.number_format = '#,##0.#'; _cF.alignment = Alignment(horizontal='center', vertical='center')
            # G7 = потребность за полный месяц
            _c7 = ws_o.cell(ri, 7)
            _c7.value = f"=IFERROR({_vl(_nc)}*{_vl(_cc)},0)"
            _c7.font = Font(bold=True, size=9, name="Arial"); _c7.fill = BLU_F
            _c7.number_format = '#,##0.#'; _c7.alignment = Alignment(horizontal='center', vertical='center')
            # H8 = страховой запас (3-мес. среднее)
            _f_saf = (f"=IFERROR(({_vls(5)}*{_vls(8)}+"
                      f"{_vls(6)}*{_vls(10)}+"
                      f"{_vls(7)}*{_vls(11)})/92*{SAFETY_DAYS},0)")
            _c8 = ws_o.cell(ri, 8)
            _c8.value = _f_saf
            _c8.font = Font(size=9, name="Arial"); _c8.number_format = '#,##0.#'
            _c8.alignment = Alignment(horizontal='center', vertical='center')
            if fb: _c8.fill = fb
            # I9 = чистая потребность = норм×авт_нетто + страх - остаток
            _c9 = ws_o.cell(ri, 9)
            _c9.value = f"=MAX(0,IFERROR({_vl(_nc)}*{_vl(_fc)},0)+H{ri}-F{ri})"
            _c9.font = Font(size=9, name="Arial"); _c9.number_format = '#,##0.#'
            _c9.alignment = Alignment(horizontal='center', vertical='center')
            if fb: _c9.fill = fb
            # J10 = кол-во упаковок
            _c10 = ws_o.cell(ri, 10)
            _c10.value = f"=IF(I{ri}>0,CEILING(I{ri}/E{ri},1),0)"
            _c10.font = Font(size=9, name="Arial"); _c10.number_format = '#,##0'
            _c10.alignment = Alignment(horizontal='center', vertical='center')
            if fb: _c10.fill = fb
            # K11 = заказ итого
            _c11 = ws_o.cell(ri, 11)
            _c11.value = f"=J{ri}*E{ri}"
            _c11.font = Font(bold=True, size=9, name="Arial")
            _c11.number_format = '#,##0.#'; _c11.alignment = Alignment(horizontal='center', vertical='center')
            _c11.fill = fill("E8F5E9") if oq > 0 else (fb or NO_F)
            # L12 = статус
            _c12 = ws_o.cell(ri, 12)
            _c12.value = f'=IF(G{ri}=0,"—",IF(K{ri}=0,"✅ Достаточно",IF(F{ri}=0,"🔴 НЕТ ОСТАТКА","📦 Заказ")))'
            _c12.font = Font(size=9, name="Arial")
            _c12.alignment = Alignment(horizontal='center', vertical='center')
            if fb: _c12.fill = fb

# ── График_Поставок ── Ss | Del на каждый день ──
# Для каждого дня d — 2 колонки:
#   ci_ss  = 8 + di*2      → Остаток (дата дд.мм; = prev_Ss - Спрос_d + Поставка_d)
#   ci_del = 8 + di*2 + 1  → Поставка (📦; только если спрос_d > 0 и остаток < страх.)
print("  График_Поставок (Del|Ss по дням)...", flush=True)
from openpyxl.formatting.rule import FormulaRule as _FR_gp
ws_g = wb_out.create_sheet('График_Поставок')
ws_g.freeze_panes = 'H4'   # A–G + строки 1–3; дни (Ss|потр|📦) прокручиваются вправо
ws_g.row_dimensions[1].height = 40
ws_g.row_dimensions[2].height = 16
ws_g.row_dimensions[3].height = 20

_n_day_cols_gp = len(all_dates)           # 92 дня
_total_cols_gp = 7 + _n_day_cols_gp * 3  # 7 fix + 3*92 (остаток|потребность|поставка)
_POTREB_DAY_COL_START = len(FH)         # 9 фикс. колонок → день 1 = col 10

def _gp_demand_ref(gp_ri, mnum, day_d):
    """Дневной спрос из Потребность_* (строка на 1 выше, чем в График_Поставок)."""
    sname = f"Потребность_{MONTH_SHORT[mnum]}"
    col = get_column_letter(_POTREB_DAY_COL_START + day_d)
    prow = gp_ri - 1
    cell = _xlsheet_ref(sname, f"{col}{prow}")
    return f"IF(ISBLANK({cell}),0,{cell})"

def _gp_local_demand_ref(gp_ri, di):
    """Дневной спрос из колонки 'потр' на текущем листе График_Поставок."""
    cell = f"{get_column_letter(_gp_ci_dem(di))}{gp_ri}"
    return f"IF(ISBLANK({cell}),0,{cell})"

def _gp_prev_day_demand_ref(gp_ri, di):
    """Для остатка на день D вычитается потребность предыдущего дня."""
    if di <= 0:
        return "0"
    return _gp_local_demand_ref(gp_ri, di - 1)

def _gp_ci_ss(di):
    return 8 + di * 3

def _gp_ci_dem(di):
    return 8 + di * 3 + 1

def _gp_ci_del(di):
    return 8 + di * 3 + 2

def _gp_prev_ss_ref(gp_ri, di):
    if di == 0:
        return f"G{gp_ri}"
    return f"{get_column_letter(_gp_ci_ss(di - 1))}{gp_ri}"

def _gp_arrival_del_ref(gp_ri, di):
    """Del предыдущего КАЛЕНДАРНОГО дня — ровно одна ячейка, без pwd/выходных."""
    if di <= 0:
        return ""
    return f"+{get_column_letter(_gp_ci_del(di - 1))}{gp_ri}"

def _stock_input_cell_ref(code, dt):
    """Excel-ссылка на жёлтую ячейку снимка остатка (лист Ввод_Остатков)."""
    ri = stock_input_row.get(code)
    ci = stock_input_date_col.get(dt)
    if not ri or not ci:
        return None
    return _xlsheet_ref('Ввод_Остатков', f"{get_column_letter(ci)}{ri}")

def _gp_ss_formula(gp_ri, di, code=None):
    """Ss[d] = Ss[d-1] − potr[d-1] + Del[d-1]; ручной снимок Ввод_Остатков ВСЕГДА подменяет Ss дня."""
    prev_ss = _gp_prev_ss_ref(gp_ri, di)
    dem_expr = _gp_prev_day_demand_ref(gp_ri, di)
    if di <= 0:
        calc = f"({prev_ss}-{dem_expr})"
    else:
        del_ref = f"{get_column_letter(_gp_ci_del(di - 1))}{gp_ri}"
        calc = f"({prev_ss}-{dem_expr}+{del_ref})"
    if code is not None:
        man_ref = _stock_input_cell_ref(code, all_dates[di][0])
        if man_ref:
            return f"=IF(ISNUMBER({man_ref}),{man_ref},{calc})"
    return f"={calc}"

def _gp_del_formula(gp_ri, di):
    """📦 предзавоз накануне (Ss[d+1]≥potr) + дonабор в день спроса до мягкой цели."""
    prev_ss = _gp_prev_ss_ref(gp_ri, di)
    dem_today = _gp_local_demand_ref(gp_ri, di)
    if di + 1 < len(all_dates):
        dem_next = _gp_local_demand_ref(gp_ri, di + 1)
    else:
        dem_next = "0"
    nd_ref = dem_next
    for look in range(2, 8):
        if di + look >= len(all_dates):
            break
        nd_ref = (
            f"IF({get_column_letter(_gp_ci_dem(di + look))}{gp_ri}>0,"
            f"{get_column_letter(_gp_ci_dem(di + look))}{gp_ri},{nd_ref})"
        )
    proj = f"({prev_ss}-{dem_today})"
    target = f"MAX(F{gp_ri},CEILING({DELIVERY_START_COVER}*({nd_ref}),1))"
    soft = f"({target})*(1-{DELIVERY_SS_TARGET_TOLERANCE})"
    qty_soft = f"CEILING(MAX(0,{soft}-{proj})/E{gp_ri},1)*E{gp_ri}"
    eve_qty = f"CEILING(MAX(0,{dem_next}-{proj})/E{gp_ri},1)*E{gp_ri}"
    eve_part = f"IF(AND({dem_next}>0,{proj}<{dem_next}),{eve_qty},0)"
    dem_part = f"IF(AND({dem_today}>0,{proj}<{soft}),{qty_soft},0)"
    return f"=IF({dem_today}>0,{dem_part},{eve_part})"

# Строка 1: заголовок
t = ws_g.cell(1, 1)
t.value = (f"ГРАФИК ПОСТАВОК | {PERIOD_LABEL}  |  "
           "🟢 Поставка  🔴 Дефицит (Ss<0 или Ss<potr)  🟡 Ниже страх.запаса  "
           "Фуры: строго по номиналу (без допуска)  |  "
           f"Допуск целевого Ss −{int(DELIVERY_SS_TARGET_TOLERANCE*100)}% (пол Ss≥potr)  |  "
           "G = остаток на 1-е число (Ввод_Остатков col E)  |  📦 Del[d] → Ss[d+1]  |  "
           "📦 кратно упаковке (1 уп = 1 код)  |  предзавоз накануне  |  "
           "жёлтые снимки → Ss дня (всегда, в т.ч. при дефиците)  |  "
           "✏️ 📦 оранж. = ручной ввод  |  "
           "Ecoal'yance: 1×/мес (1–5) | Ecotexis: +35д от заказа | "
           f"Purem/SMC: 1×/нед | SMC ≤{SMC_MAX_PALLETS_PER_TRUCK} палл. | "
           f"Lear: {LEAR_MAX_PALLET_SLOTS} палл. | "
           f"NDKT: цель переходящего остатка = страх.запас по коду (240/480 шт); "
           f"ПЕРВЫЕ партии плана (AS_in_F_A/AS_in_H_B) — 100% из остатков; "
           f"упаковка: 3101100XST33A — бокс 30 шт, прочие колёса — стопка 24 шт; "
           f"колёса миксуются в машине (фура: XST33A 540 шт, прочие 384 шт)")
t.font = Font(bold=True, size=9, color="FFFFFF", name="Arial")
t.fill = H_FILL
t.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
ws_g.merge_cells(f'A1:{get_column_letter(_total_cols_gp)}1')

# Строки 2–3: шапка вне объединений «строка2+3» для A-G (ломает автофильтр и даты).
# Строка 2: A-G — пустой блок; H+ — названия месяцев.
# Строка 3: A-G — заголовки + H+ — даты (единственная строка автофильтра).
_GP_FIX_H = ['Код детали','Наименование / Применяемость / Цвет','Поставщик','Ед.','Уп.','Страх.\nшт','Остаток\n(нач.мес.)']
_GP_FIX_W = [22, 60, 16, 5, 5, 9, 11]
ws_g.merge_cells('A2:G2')
_h2 = ws_g.cell(2, 1)
_h2.fill = H_FILL
_h2.alignment = Alignment(horizontal='center', vertical='center')
for ci, (h, w) in enumerate(zip(_GP_FIX_H, _GP_FIX_W), 1):
    hcell(ws_g, 3, ci, h, H_FILL)
    ws_g.cell(3, ci).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws_g.column_dimensions[get_column_letter(ci)].width = w

# Месячные метки в строке 2 над днями
for mnum_m, mlabel_m, ndays_m in MONTHS:
    first_di = sum(nd for mn,_,nd in MONTHS if mn < mnum_m)
    last_di  = first_di + ndays_m - 1
    col_first = 8 + first_di * 3
    col_last  = 8 + last_di  * 3 + 2
    mc = ws_g.cell(2, col_first)
    mc.value = mlabel_m
    mc.font  = Font(bold=True, size=10, color="FFFFFF", name="Arial")
    mc.fill  = M_FILL[mnum_m]
    mc.alignment = Alignment(horizontal='center', vertical='center')
    ws_g.merge_cells(f'{get_column_letter(col_first)}2:{get_column_letter(col_last)}2')

# Строка 3: чередующиеся заголовки «дд.мм» | «📦»
_DEL_FILL = fill("C6EFCE")   # зелёный — колонка поставки
for di, (dt, mnum_h, d_h) in enumerate(all_dates):
    ci_ss  = _gp_ci_ss(di)
    ci_dem = _gp_ci_dem(di)
    ci_del = _gp_ci_del(di)
    c_sh = ws_g.cell(3, ci_ss)
    c_sh.value = dt.strftime('%d.%m')
    c_sh.font  = Font(bold=True, color="FFFFFF", size=8, name="Arial")
    c_sh.fill  = M_FILL[mnum_h]
    c_sh.alignment = Alignment(horizontal='center', vertical='center')
    ws_g.column_dimensions[get_column_letter(ci_ss)].width = 5.5
    c_qh = ws_g.cell(3, ci_dem)
    c_qh.value = 'потр'
    c_qh.font  = Font(bold=True, size=8, name="Arial", color="9C5700")
    c_qh.fill  = fill("FFF2CC")
    c_qh.alignment = Alignment(horizontal='center', vertical='center')
    ws_g.column_dimensions[get_column_letter(ci_dem)].width = 5.5
    c_dh = ws_g.cell(3, ci_del)
    c_dh.value = '📦'
    c_dh.font  = Font(bold=True, size=8, name="Arial")
    c_dh.fill  = _DEL_FILL
    c_dh.alignment = Alignment(horizontal='center', vertical='center')
    ws_g.column_dimensions[get_column_letter(ci_del)].width = 5.5

# Данные: строки 4+
_gp_data_start = 4
_gp_n = len(mrp_codes)
for _gp_i, code in enumerate(mrp_codes, _gp_data_start):
    if (_gp_i - _gp_data_start) > 0 and (_gp_i - _gp_data_start) % 100 == 0:
        print(f"    ... {_gp_i - _gp_data_start}/{_gp_n}", flush=True)
    ri = _gp_i
    ws_g.row_dimensions[ri].height = 14
    fb = GRY_F if ri % 2 == 0 else NO_F
    name, supp, unit = get_info(code)
    pkg  = max(1, bom.get(code, {}).get('package', 1))
    sd   = get_safety_days(code, supp)
    total_3m  = sum(sum(demand[code].get(mn, {}).values()) for mn, _, _ in MONTHS)
    n_days_3m = sum(nd for _, _, nd in MONTHS)
    avg_daily  = total_3m / n_days_3m if n_days_3m else 0
    safety_qty = _safety_qty(code, supp)
    _op_stk = int(round(_opening_stock_for_ss(code)))

    # A-E: статика
    for ci, v in enumerate([code, _name_with_appl(code,name), supp, unit, pkg], 1):
        c = ws_g.cell(ri, ci); c.value = v
        c.font = Font(size=9, name="Arial")
        c.alignment = Alignment(horizontal='left' if ci <= 2 else 'center', vertical='center')
        if fb: c.fill = fb

    # F: страховой запас (шт)
    cf = ws_g.cell(ri, 6); cf.value = safety_qty
    cf.font = Font(size=9, name="Arial"); cf.number_format = '#,##0.#'
    cf.alignment = Alignment(horizontal='center', vertical='center')
    cf.fill = fill("FFF2CC")

    # G: остаток на 1-е число — формула из Ввод_Остатков (0 в col E → fallback из расчёта)
    _op_stk = int(round(_opening_stock_for_ss(code)))
    _vlk_g = f"VLOOKUP(A{ri},{_xlsheet_ref('Ввод_Остатков', '$A:$E')},5,0)"
    cg = ws_g.cell(ri, 7)
    cg.value = f"=IFERROR(IF({_vlk_g}=0,{_op_stk},{_vlk_g}),{_op_stk})"
    cg.font = Font(size=9, name="Arial", bold=True); cg.number_format = '#,##0'
    cg.alignment = Alignment(horizontal='center', vertical='center')
    cg.fill = CYN_F

    # H+: Ss | потр | Del — Ss по формуле; жёлтый снимок Ввод_Остатков подменяет Ss этого дня.
    for di, (dt, mnum_d, day_d) in enumerate(all_dates):
        ci_ss  = _gp_ci_ss(di)
        ci_dem = _gp_ci_dem(di)
        ci_del = _gp_ci_del(di)
        _ss_list = DELIVERY_SCHEDULES.get(code, {}).get('ss', [])

        c_ss = ws_g.cell(ri, ci_ss)
        c_ss.value = _gp_ss_formula(ri, di, code)
        c_ss.font = Font(size=8, name="Arial")
        c_ss.number_format = '#,##0'
        c_ss.alignment = Alignment(horizontal='center', vertical='center')
        if fb: c_ss.fill = fb

        c_dem = ws_g.cell(ri, ci_dem)
        c_dem.value = "=" + _gp_demand_ref(ri, mnum_d, day_d)
        c_dem.font = Font(size=8, name="Arial", color="9C5700")
        c_dem.number_format = '#,##0'
        c_dem.alignment = Alignment(horizontal='center', vertical='center')
        c_dem.fill = fill("FFF8E1")

        # Поставка (📦): авто — только в пустые ячейки; ручные — число (VLOOKUP→значение)
        _ds_dels = DELIVERY_SCHEDULES.get(code, {}).get('dels', [])
        _del_v = _ds_dels[di] if di < len(_ds_dels) else 0
        _manual_del = manual_delivery_by_date.get(code, {}).get(dt)
        if _manual_del is None:
            _manual_del = manual_delivery_by_date.get(normalize_code(code), {}).get(dt)
        c_del = ws_g.cell(ri, ci_del)
        if _manual_del is not None:
            try:
                _mv = float(_manual_del)
                c_del.value = int(round(_mv)) if _mv == int(_mv) else _mv
            except (TypeError, ValueError):
                c_del.value = None
            c_del.fill = _GP_MANUAL_DEL_FILL
        elif _del_v and round(_del_v) > 0:
            c_del.value = f"={int(round(_del_v))}"   # авто 📦 как формула → вписанное ЧИСЛО = ручной ввод
        else:
            c_del.value = "=0"
        c_del.font = Font(size=8, name="Arial", bold=True)
        c_del.number_format = '#,##0'
        c_del.alignment = Alignment(horizontal='center', vertical='center')

# Условное форматирование
_gp_last_row = _gp_data_start + len(mrp_codes) - 1
_gp_filter_last_row = max(_gp_last_row, _gp_data_start)
# Автофильтр: строка 3 = заголовок; кнопки фильтра только A-G (H+ без dropdown).
ws_g.auto_filter.ref = f"A3:{get_column_letter(_total_cols_gp)}{_gp_filter_last_row}"
try:
    from openpyxl.worksheet.filters import FilterColumn as _FilterColumn
    if ws_g.auto_filter.filterColumn is None:
        ws_g.auto_filter.filterColumn = []
    else:
        ws_g.auto_filter.filterColumn.clear()
    for _gp_fcol in range(7, min(_total_cols_gp, 263)):
        ws_g.auto_filter.filterColumn.append(_FilterColumn(colId=_gp_fcol, hiddenButton=True))
except Exception as _gp_af_err:
    print(f"  ⚠️  Автофильтр График_Поставок (скрытие H+): {_gp_af_err}", flush=True)
# Del колонки (I, K, M, ...): зелёный если > 0
_del_range = f"H{_gp_data_start}:{get_column_letter(_total_cols_gp)}{_gp_last_row}"
ws_g.conditional_formatting.add(_del_range, _FR_gp(
    formula=[f"AND(MOD(COLUMN(H{_gp_data_start})-8,3)=2,H{_gp_data_start}>0)"],
    fill=PatternFill("solid", fgColor="C6EFCE"),
    font=Font(color="375623", bold=True, size=8, name="Arial")))
# Ss колонки: красный если < 0 или дефицит на день спроса.
# NDKT: потребность дня закрывается поставкой ТЕКУЩЕГО дня → красный только
# если Ss + Del дня < potr; остальные поставщики: Ss < potr.
_ss_range = f"H{_gp_data_start}:{get_column_letter(_total_cols_gp)}{_gp_last_row}"
ws_g.conditional_formatting.add(_ss_range, _FR_gp(
    formula=[f"AND(MOD(COLUMN(H{_gp_data_start})-8,3)=0,"
             f"OR(H{_gp_data_start}<0,"
             f"AND(OFFSET(H{_gp_data_start},0,1)>0,"
             f"H{_gp_data_start}+IF(ISNUMBER(SEARCH(\"NDKT\",$C{_gp_data_start})),"
             f"OFFSET(H{_gp_data_start},0,2),0)<OFFSET(H{_gp_data_start},0,1))))"],
    fill=PatternFill("solid", fgColor="FFC7CE"),
    font=Font(color="9C0006", bold=True, size=8, name="Arial")))
# Ss колонки: жёлтый если ниже страхового
ws_g.conditional_formatting.add(_ss_range, _FR_gp(
    formula=[f"AND(MOD(COLUMN(H{_gp_data_start})-8,3)=0,H{_gp_data_start}>=0,H{_gp_data_start}<$F{_gp_data_start})"],
    fill=PatternFill("solid", fgColor="FFEB9C"),
    font=Font(color="9C5700", size=8, name="Arial")))

# ── Аналитика_Оборачиваемости ─────────────────────────────────
print("  Аналитика_Оборачиваемости...")
ws_t=wb_out.create_sheet('Аналитика_Оборачиваемости')
ws_t.freeze_panes='E3'; ws_t.row_dimensions[1].height=50; ws_t.row_dimensions[2].height=35
t=ws_t.cell(1,1); t.value=f"АНАЛИТИКА ОБОРАЧИВАЕМОСТИ | Еженедельно | {PERIOD_LABEL}"
t.font=Font(bold=True,size=13,color="FFFFFF",name="Arial"); t.fill=H3
t.alignment=Alignment(horizontal='center',vertical='center')
ws_t.merge_cells(f'A1:{get_column_letter(4+len(WEEKS)*2+3)}1')
for ci,(h,w) in enumerate(zip(['Код','Наименование / Применяемость / Цвет','Поставщик','Ед.'],[22,60,16,6]),1):
    hcell(ws_t,2,ci,h,H_FILL); ws_t.column_dimensions[get_column_letter(ci)].width=w
for wi,w_info in enumerate(WEEKS):
    bc=5+wi*2; wf=fill(WEEK_CLR[wi])
    lbl=f"{w_info['label']}\n{w_info['start'].strftime('%d.%m')}–{w_info['end'].strftime('%d.%m')}"
    for mi,mn in enumerate(['Потребн.','Оборач.\n(дн.)']):
        c=ws_t.cell(2,bc+mi); c.value=lbl+f"\n{mn}" if mi==0 else mn
        c.font=Font(bold=True,color="FFFFFF",size=8,name="Arial"); c.fill=wf
        c.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)
        ws_t.column_dimensions[get_column_letter(bc+mi)].width=9
sb=5+len(WEEKS)*2
for si,(h,w) in enumerate(zip(['Ср.оборач.\n(дн.)','Тренд','Статус'],[10,14,18])):
    c=ws_t.cell(2,sb+si); c.value=h
    c.font=Font(bold=True,color="FFFFFF",size=9,name="Arial"); c.fill=H_FILL
    c.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)
    ws_t.column_dimensions[get_column_letter(sb+si)].width=w
for ri,code in enumerate(mrp_codes,3):
    ws_t.row_dimensions[ri].height=14; fb=GRY_F if ri%2==0 else NO_F
    name,supp,unit=get_info(code); stk=stock.get(code,0)
    wd=weekly_demand_map(code); pkg=max(1,bom.get(code,{}).get('package',1))
    for ci,v in enumerate([code,_name_with_appl(code,name),supp,unit],1):
        c=ws_t.cell(ri,ci); c.value=v; c.font=Font(size=9,name="Arial")
        c.alignment=Alignment(horizontal='left' if ci<=2 else 'center',vertical='center')
        if fb: c.fill=fb
    wtos=[]; run_stk=stk
    for wi,w_info in enumerate(WEEKS):
        bc=5+wi*2; wq=wd.get(wi,0)
        pkgs_w=math.ceil(wq/pkg) if wq>0 else 0; oq_w=pkgs_w*pkg
        d_in_w=(w_info['end']-w_info['start']).days+1
        dc=wq/d_in_w if d_in_w else 0
        to=round(run_stk/dc,1) if dc>0 else (999.0 if run_stk>0 else 0.0)
        wtos.append(to); run_stk=max(0,run_stk-wq+oq_w)
        c1=ws_t.cell(ri,bc); c1.value=round(wq,1); c1.number_format='#,##0.#'
        c1.font=Font(size=9,name="Arial"); c1.alignment=Alignment(horizontal='center',vertical='center')
        if fb: c1.fill=fb
        c2=ws_t.cell(ri,bc+1); c2.value=to if to<999 else None
        c2.number_format='#,##0.#'; c2.font=Font(size=9,name="Arial")
        c2.alignment=Alignment(horizontal='center',vertical='center')
        c2.fill=(RED_F if 0<to<3 else ORG_F if 0<to<7 else fill("E2EFDA") if 0<to<999 else (fb or NO_F))
    vt=[t for t in wtos if 0<t<999]
    avg_to=round(sum(vt)/len(vt),1) if vt else 0
    h2=len(vt)//2
    trend=(("↑ Растёт" if sum(vt[-h2:])/h2>sum(vt[:h2])/h2*1.1 else
            "↓ Падает" if sum(vt[-h2:])/h2<sum(vt[:h2])/h2*0.9 else
            "→ Стабильно") if h2>=1 else "—")
    st_t=("⚪ Нет данных" if avg_to==0 and stk==0 else
          "🔴 Критично" if 0<avg_to<3 else "🟡 Внимание" if 0<avg_to<7 else
          "🟡 Сверхзапас" if avg_to>60 else "✅ Норма")
    c_a=ws_t.cell(ri,sb); c_a.value=avg_to; c_a.number_format='#,##0.#'
    c_a.font=Font(size=9,bold=True,name="Arial"); c_a.alignment=Alignment(horizontal='center',vertical='center')
    c_a.fill=(RED_F if 0<avg_to<3 else ORG_F if 0<avg_to<7 else fill("E2EFDA") if avg_to>0 else (fb or NO_F))
    c_tr=ws_t.cell(ri,sb+1); c_tr.value=trend
    c_tr.font=Font(size=9,name="Arial"); c_tr.alignment=Alignment(horizontal='center',vertical='center')
    if fb: c_tr.fill=fb
    c_st=ws_t.cell(ri,sb+2); c_st.value=st_t
    c_st.font=Font(size=9,bold=('Критично' in st_t or 'Норма' in st_t),name="Arial")
    c_st.alignment=Alignment(horizontal='center',vertical='center')
    c_st.fill=(RED_F if 'Критично' in st_t else ORG_F if 'Внимание' in st_t else
               fill("E2EFDA") if 'Норма' in st_t else YEL_F if 'Сверх' in st_t else (fb or NO_F))

# ── BOM_Linkage ───────────────────────────────────────────────
print("  BOM_Linkage...")
ws_bl=wb_out.create_sheet('BOM_Linkage')
ws_bl.freeze_panes='F4'; ws_bl.row_dimensions[1].height=40
ws_bl.row_dimensions[2].height=20; ws_bl.row_dimensions[3].height=35
t=ws_bl.cell(1,1); t.value="BOM LINKAGE — Применяемость → Вкладка плана → Потребность"
t.font=Font(bold=True,size=12,color="FFFFFF",name="Arial"); t.fill=H3
t.alignment=Alignment(horizontal='center',vertical='center'); ws_bl.merge_cells('A1:N1')
ws_bl.cell(2,1).value="FIX#1=B16 elite сиденья | FIX#2=ALAA005669 B06 | Жёлтый=исправлена применяемость"
ws_bl.cell(2,1).font=Font(italic=True,size=9,color="1F3864",name="Arial")
ws_bl.cell(2,1).fill=fill("E8F5E9"); ws_bl.merge_cells('A2:N2')
lh=['Код','Наименование / Применяемость / Цвет','Поставщик','Ед.','Тип/Норма','Вкладка плана',
    f'Уп.\n(шт)',f'Потребн.\n{MONTH_SHORT[MONTHS[0][0]]}',f'Потребн.\n{MONTH_SHORT[MONTHS[1][0]]}',f'Потребн.\n{MONTH_SHORT[MONTHS[2][0]]}',
    'Итого\n3 мес.','Остаток','Статус','Раздел BOM']
lw=[22,60,16,6,28,22,7,12,12,12,12,10,16,25]
for ci,(h,w) in enumerate(zip(lh,lw),1):
    hcell(ws_bl,3,ci,h,H_FILL); ws_bl.column_dimensions[get_column_letter(ci)].width=w
# Порядок строк — по разделам BOM (перегруппировано), затем коды вне разделов
_bl_codes = []
_bl_seen = set()
_bl_mrp_set = set(mrp_codes)
for _sec, _codes_in_sec in bom_sections.items():
    for _c in _codes_in_sec:
        if _c in _bl_seen or _c not in _bl_mrp_set:
            continue
        _bl_codes.append(_c); _bl_seen.add(_c)
for _c in mrp_codes:
    if _c not in _bl_seen:
        _bl_codes.append(_c); _bl_seen.add(_c)
for ri,code in enumerate(_bl_codes,4):
    ws_bl.row_dimensions[ri].height=14; fb=GRY_F if ri%2==0 else NO_F
    name,supp,unit=get_info(code)
    tab=part_tab_map.get(code,'—'); norm_str=get_norm_str(code)
    pkg=bom.get(code,{}).get('package',1); stk=stock.get(code,0)
    _dm = [round(sum(demand[code].get(mn,{}).values()),1) for mn,_,__ in MONTHS]
    d_may,d_jun,d_jul = _dm[0],_dm[1],_dm[2]
    d_tot=round(sum(_dm),1)
    sec=bom.get(code,{}).get('section','')[:25]
    if code in B16_ELITE: st="FIX#1: B16 elite"
    elif code=='ALAA005669': st="FIX#2: B06 paint"
    elif d_tot==0: st="⚠️ Нет спроса"
    else: st="✅ OK"
    vals=[code,_name_with_appl(code,name),supp,unit,norm_str,tab,pkg,
          d_may if d_may>0 else None,d_jun if d_jun>0 else None,d_jul if d_jul>0 else None,
          d_tot if d_tot>0 else None,stk if stk>0 else None,st,sec]
    for ci,v in enumerate(vals,1):
        c=ws_bl.cell(ri,ci); c.value=v; c.font=Font(size=9,name="Arial")
        c.alignment=Alignment(horizontal='left' if ci<=2 else 'center',vertical='center')
        if ci==5: c.fill=NRM_F
        elif ci==6: c.fill=CYN_F
        elif ci in (8,9,10) and v: c.fill=BLU_F; c.number_format='#,##0.#'
        elif ci==11 and v: c.fill=fill("E8F5E9"); c.number_format='#,##0.#'; c.font=Font(bold=True,size=9,name="Arial")
        elif ci==13:
            if 'FIX' in str(v): c.fill=YEL_F; c.font=Font(bold=True,size=9,name="Arial",color="856404")
            elif '⚠️' in str(v): c.fill=ORG_F
            elif '✅' in str(v): c.fill=fill("E2EFDA")
            elif fb: c.fill=fb
        elif fb and ci not in (5,6,8,9,10,11,13): c.fill=fb

# ── Сводка_3мес ──────────────────────────────────────────────
print("  Сводка_3мес...")
ws_s=wb_out.create_sheet('Сводка_3мес')
ws_s.freeze_panes='H2'; ws_s.row_dimensions[1].height=50
SH=['Код','Наименование / Применяемость / Цвет','Поставщик','Ед.','Уп.(шт)','Остаток']
SW=[22,60,16,6,7,10]
for ci,(h,w) in enumerate(zip(SH,SW),1):
    hcell(ws_s,1,ci,h,H_FILL); ws_s.column_dimensions[get_column_letter(ci)].width=w
for mi,(mnum,mlabel,_) in enumerate(MONTHS):
    base=len(SH)+mi*4+1
    for off,(txt,w) in enumerate(zip([f'Потребн.\n{mlabel[:3]}',f'Заказ\n{mlabel[:3]}',
                                       f'Дефицит\n{mlabel[:3]}',f'Покрытие\n{mlabel[:3]}'],[14,12,12,10])):
        c=ws_s.cell(1,base+off); c.value=txt
        c.font=Font(bold=True,color="FFFFFF",size=9,name="Arial"); c.fill=M_FILL[mnum]
        c.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)
        ws_s.column_dimensions[get_column_letter(base+off)].width=w
_sv_rc_col = len(SH) + len(MONTHS) * 4 + 1
_c_rc = ws_s.cell(1, _sv_rc_col)
_c_rc.value = f'Расчёт {MONTH_SHORT[MONTHS[0][0]]}\nс даты'
_c_rc.font = Font(bold=True, color="FFFFFF", size=9, name="Arial")
_c_rc.fill = fill("B8860B")
_c_rc.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
ws_s.column_dimensions[get_column_letter(_sv_rc_col)].width = 10
if not CALC_ORDERS:
    _c_off = ws_s.cell(2, 1)
    _c_off.value = "Сводка отключена вместе с месячными заказами (галочка «Месячные заказы» в run_mrp_gui)"
    _c_off.font = Font(bold=True, size=10, color="C00000", name="Arial")
for ri,code in enumerate(mrp_codes if CALC_ORDERS else [],2):
    ws_s.row_dimensions[ri].height=14; fb=GRY_F if ri%2==0 else NO_F
    name,supp,unit=get_info(code); stk=_opening_stock(code); pkg=bom.get(code,{}).get('package',1)
    for ci,v in enumerate([code,_name_with_appl(code,name),supp,unit,pkg,stk],1):
        c=ws_s.cell(ri,ci); c.value=v; c.font=Font(size=9,name="Arial")
        c.alignment=Alignment(horizontal='left' if ci<=2 else 'center',vertical='center')
        if ci==5: c.fill=ORG_F
        elif ci==6: c.fill=CYN_F; c.number_format='#,##0.##'
        elif fb: c.fill=fb
    for mi,(mnum,mlabel,_) in enumerate(MONTHS):
        base=len(SH)+mi*4+1
        proj_stk_rd = projected_stock.get(code, {}).get(mnum) if mnum != STOCK_AS_OF_MONTH else None
        oi=calc_order(code, mnum, stock_override=proj_stk_rd); tot=oi['demand']; oq=oi['order_qty']
        stk_rd = oi['stock']
        defc=round(stk_rd-tot,1); cov=round(stk_rd/tot*100,1) if tot>0 else None
        for off,(val,fmt) in enumerate(zip([round(tot,1) if tot else None,round(oq,1) if oq else None,
                                            round(defc,1) if tot else None,cov],
                                           ['#,##0.#','#,##0.#','#,##0.#','0.0"%"'])):
            c=ws_s.cell(ri,base+off); c.value=val; c.number_format=fmt
            c.font=Font(size=9,name="Arial"); c.alignment=Alignment(horizontal='center',vertical='center')
            if off==0 and tot: c.fill=BLU_F; c.font=Font(bold=True,size=9,name="Arial")
            elif off==1 and oq: c.fill=fill("E8F5E9")
            elif off==2 and tot: c.fill=RED_F if defc<0 else fill("E2EFDA")
            elif off==3 and cov: c.fill=fill("E2EFDA") if cov>=100 else RED_F
            elif fb: c.fill=fb
    c = ws_s.cell(ri, _sv_rc_col)
    c.value = _calc_from_date(code).strftime('%d.%m.%Y')
    c.font = Font(size=8, name="Arial", color="B8860B", bold=True)
    c.alignment = Alignment(horizontal='center', vertical='center')

# ── Прогноз_CKD: +3 месяца — месячные потребности и заказы (без дней) ──
print("  Прогноз_CKD...")
ws_f = wb_out.create_sheet('Прогноз_CKD')
ws_f.freeze_panes = 'G3'
ws_f.row_dimensions[1].height = 55
_fc_src_name = os.path.basename(CKD_SRC) if CKD_SRC else '—'
_t_f = ws_f.cell(1, 1)
_t_f.value = (f"ПРОГНОЗ +3 МЕС. ПО ПЛАНУ CKD | Источник: {_fc_src_name} | "
              f"План моделей распределён по конфигурациям СРЕДНИМИ долями за {PERIOD_LABEL} | "
              f"Остаток на старт = расчётный остаток конца {MONTH_SHORT[MONTHS[-1][0]]} (График_Поставок), "
              f"далее переходит | Заказ = CEILING((Потребн.+Страх−Остаток)/Уп.)×Уп.")
_t_f.font = Font(bold=True, size=10, color="FFFFFF", name="Arial")
_t_f.fill = fill("1F3864")
_t_f.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
_FH_F = ['Код', 'Наименование / Применяемость / Цвет', 'Поставщик', 'Ед.', 'Уп.\n(шт)',
         'Остаток\nна старт', 'Страх.\nзапас']
_FW_F = [22, 60, 16, 6, 7, 11, 10]
ws_f.merge_cells(f'A1:{get_column_letter(len(_FH_F) + 2 * len(FORECAST_MONTHS))}1')
for ci, (h, w) in enumerate(zip(_FH_F, _FW_F), 1):
    hcell(ws_f, 2, ci, h, H_FILL)
    ws_f.column_dimensions[get_column_letter(ci)].width = w
_FC_FILLS = ['2E75B6', '1F7A4C', '7B3F91']
for mi_f, (fm, fy, flbl) in enumerate(FORECAST_MONTHS):
    base = len(_FH_F) + mi_f * 2 + 1
    for off, txt in enumerate([f'Потребн.\n{flbl[:3]}', f'Заказ\n{flbl[:3]}']):
        c = ws_f.cell(2, base + off)
        c.value = txt
        c.font = Font(bold=True, color="FFFFFF", size=9, name="Arial")
        c.fill = fill(_FC_FILLS[mi_f % 3])
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        ws_f.column_dimensions[get_column_letter(base + off)].width = 13
if not CALC_ORDERS:
    _c_off = ws_f.cell(3, 1)
    _c_off.value = "Прогноз CKD отключён вместе с месячными заказами (галочка в run_mrp_gui)"
    _c_off.font = Font(bold=True, size=10, color="C00000", name="Arial")
_fc_row = 3
for code in (mrp_codes if CALC_ORDERS else []):
    _fd = forecast_demand.get(code, {})
    _fo = forecast_orders.get(code, {})
    name, supp, unit = get_info(code)
    fb = GRY_F if _fc_row % 2 == 0 else NO_F
    ws_f.row_dimensions[_fc_row].height = 14
    _ss_list_f = DELIVERY_SCHEDULES.get(code, {}).get('ss') or []
    _stk0 = max(0.0, float(_ss_list_f[-1])) if _ss_list_f else float(stock.get(code, 0) or 0)
    _saf0 = _safety_qty(code, bom.get(code, {}).get('supplier', ''))
    vals_f = [code, _name_with_appl(code, name), supp, unit, _pkg_size(code),
              round(_stk0, 1), round(_saf0, 1)]
    for ci, v in enumerate(vals_f, 1):
        c = ws_f.cell(_fc_row, ci)
        c.value = v
        c.font = Font(size=9, name="Arial")
        c.alignment = Alignment(horizontal='left' if ci <= 2 else 'center', vertical='center')
        if ci == 5:
            c.fill = ORG_F
        elif ci == 6:
            c.fill = CYN_F; c.number_format = '#,##0.#'
        elif fb:
            c.fill = fb
    for mi_f, (fm, fy, flbl) in enumerate(FORECAST_MONTHS):
        base = len(_FH_F) + mi_f * 2 + 1
        dem_v = _fd.get((fm, fy), 0) or 0
        oq_v = (_fo.get((fm, fy), {}) or {}).get('order', 0) or 0
        c = ws_f.cell(_fc_row, base)
        c.value = round(dem_v, 1) if dem_v else None
        c.number_format = '#,##0.#'
        c.font = Font(bold=bool(dem_v), size=9, name="Arial")
        c.alignment = Alignment(horizontal='center', vertical='center')
        if dem_v:
            c.fill = BLU_F
        elif fb:
            c.fill = fb
        c = ws_f.cell(_fc_row, base + 1)
        c.value = round(oq_v, 1) if oq_v else None
        c.number_format = '#,##0.#'
        c.font = Font(size=9, name="Arial")
        c.alignment = Alignment(horizontal='center', vertical='center')
        if oq_v:
            c.fill = fill("E8F5E9")
        elif fb:
            c.fill = fb
    _fc_row += 1
print(f"    Прогноз_CKD: {_fc_row - 3} кодов × {len(FORECAST_MONTHS)} мес.")

# ── Планы_Производств: исходные планы, по которым считается потребность ──
print("  Планы_Производств...")
ws_pp = wb_out.create_sheet('Планы_Производств')
ws_pp.freeze_panes = 'A4'
ws_pp.row_dimensions[1].height = 40
_t_pp = ws_pp.cell(1, 1)
_t_pp.value = (f"ПЛАНЫ ПРОИЗВОДСТВ — данные, по которым рассчитана потребность | "
               f"Plan-Fact: {os.path.basename(PF_FILE) if PF_FILE else '—'} | "
               f"Прогнозные месяцы: план CKD ({_fc_src_name}), распределённый средними долями конфигураций")
_t_pp.font = Font(bold=True, size=11, color="FFFFFF", name="Arial")
_t_pp.fill = fill("1F3864")
_t_pp.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
ws_pp.merge_cells(f'A1:{get_column_letter(8 + 31)}1')
_PP_H = ['Вкладка', 'Месяц', 'Партия', 'Модель', 'Привод', 'Конфигурация', 'Источник', 'Итого']
_PP_W = [12, 10, 14, 8, 8, 12, 16, 9]
ws_pp.row_dimensions[3].height = 16
for ci, (h, w) in enumerate(zip(_PP_H, _PP_W), 1):
    hcell(ws_pp, 3, ci, h, H_FILL)
    ws_pp.column_dimensions[get_column_letter(ci)].width = w
for d in range(1, 32):
    ci = len(_PP_H) + d
    hcell(ws_pp, 3, ci, str(d), fill("44546A"))
    ws_pp.column_dimensions[get_column_letter(ci)].width = 5.5
_pp_months = [(mn, mlabel) for mn, mlabel, _ in MONTHS]
_pp_forecast = [(fm, flbl) for fm, fy, flbl in FORECAST_MONTHS]
# Очерёдность: сначала ВСЕ планы из источника Plan-Fact, затем прогноз CKD
_pp_order = ([(tab, mn, mlabel) for tab in TAB_COLS if tab in plan_batches
              for mn, mlabel in _pp_months]
             + [(tab, fm, flbl) for tab in TAB_COLS if tab in plan_batches
                for fm, flbl in _pp_forecast])
_pp_row = 4
for tab, mn, mlabel in _pp_order:
    month_data = plan_batches.get(tab, {}).get(mn) or {}
    if not month_data:
        continue
    for bv in sorted(month_data.keys(), key=lambda b: (min(month_data[b].keys() or [99]), str(b))):
        day_qty = month_data[bv] or {}
        bi = _get_batch_info(tab, bv)
        is_fc_batch = str(bv).startswith('RFC')
        src_lbl = 'CKD прогноз (ср. доли)' if is_fc_batch else 'Plan-Fact'
        _tot_f = sum(day_qty.values())
        tot_b = int(round(_tot_f)) if abs(_tot_f - round(_tot_f)) < 1e-9 else round(_tot_f, 1)
        fb = GRY_F if _pp_row % 2 == 0 else NO_F
        ws_pp.row_dimensions[_pp_row].height = 13
        vals_pp = [tab, mlabel, str(bv), bi.get('model', ''), bi.get('drive', ''),
                   bi.get('config', ''), src_lbl, tot_b]
        for ci, v in enumerate(vals_pp, 1):
            c = ws_pp.cell(_pp_row, ci)
            c.value = v
            c.font = Font(size=8, name="Arial", bold=(ci == 3))
            c.alignment = Alignment(horizontal='left' if ci in (1, 7) else 'center',
                                    vertical='center')
            if ci == 7 and is_fc_batch:
                c.fill = fill("FFF2CC")
            elif ci == 8:
                c.fill = BLU_F; c.number_format = '#,##0.#'
            elif fb:
                c.fill = fb
        for d, q in day_qty.items():
            try:
                dd = int(d)
            except (TypeError, ValueError):
                continue
            if 1 <= dd <= 31 and q:
                c = ws_pp.cell(_pp_row, len(_PP_H) + dd)
                _qf = float(q)
                c.value = int(round(_qf)) if abs(_qf - round(_qf)) < 1e-9 else round(_qf, 1)
                c.number_format = '#,##0.#'
                c.font = Font(size=8, name="Arial")
                c.alignment = Alignment(horizontal='center', vertical='center')
                if is_fc_batch:
                    c.fill = fill("FFF2CC")
        _pp_row += 1
print(f"    Планы_Производств: {_pp_row - 4} строк (сначала Plan-Fact, затем CKD)")

# ── Риск_Дефицита ─────────────────────────────────────────────
print("  Риск_Дефицита...")


def _risk_active_from_month(code):
    """Месяц старта потребности (Litum / частичный грунт)."""
    if code in REPLACEMENT_NEW_CODES:
        _, fm = REPLACEMENT_NEW_CODES[code]
        if fm is not None:
            return fm
    if code in LITUM_PRIMER_CODES:
        return LITUM_PRIMER_CODES[code]
    return None


def _risk_avg_daily(code):
    """Ср/день = потребность периода / дни активности (для Litum — только с from_month)."""
    total = sum(sum(demand[code].get(mn, {}).values()) for mn, _, _ in MONTHS)
    if total <= 0:
        return 0.0
    from_m = _risk_active_from_month(code)
    if from_m is not None:
        days = sum(nd for mn, _, nd in MONTHS if mn >= from_m)
    else:
        days = sum(nd for _, _, nd in MONTHS)
    return total / days if days > 0 else 0.0


def _risk_balance_series(code):
    """Дневной баланс по графику поставок (все дни с потребностью в горизонте)."""
    pairs = DELIVERY_SCHEDULES.get(code, {}).get('pairs', {})
    stk0 = float(stock.get(code, 0) or 0)
    series = []
    for di, (dt, mnum, d) in enumerate(all_dates):
        del_d, bal = pairs.get(di, (0.0, stk0))
        series.append({'dt': dt, 'mn': mnum, 'd': d, 'bal': bal, 'del': del_d})
    return series


def _risk_simulation(code):
    """Дефицит по симуляции графика; справочно — ближайшие поставки и дата закрытия."""
    series = _risk_balance_series(code)
    upcoming = [(s['dt'], s['del']) for s in series if s['del'] > 0]

    first_def = None
    min_bal = float(stock.get(code, 0) or 0)
    closure_date = None
    in_def = False

    for s in series:
        if s['bal'] < min_bal:
            min_bal = s['bal']
        if s['bal'] < 0:
            if first_def is None:
                first_def = s['dt']
            in_def = True
        elif in_def and s['bal'] >= 0:
            closure_date = s['dt']
            break

    nearest = upcoming[0] if upcoming else (None, None)
    closure_after_nearest = None
    if nearest[0] and first_def and nearest[0] >= first_def:
        for s in series:
            if s['dt'] >= nearest[0] and s['bal'] >= 0:
                closure_after_nearest = s['dt']
                break

    return {
        'first_deficit': first_def,
        'min_balance': min_bal,
        'nearest_del_date': nearest[0],
        'nearest_del_qty': nearest[1] if nearest[0] else None,
        'closure_date': closure_after_nearest or closure_date,
        'upcoming': upcoming[:3],
    }


ws_r = wb_out.create_sheet('Риск_Дефицита')
ws_r.freeze_panes = 'A4'
ws_r.row_dimensions[1].height = 40
ws_r.row_dimensions[2].height = 28
ws_r.row_dimensions[3].height = 35
_RISK_LAST_COL = 17
tt = ws_r.cell(1, 1)
tt.value = (f"РИСК ДЕФИЦИТА | {PERIOD_LABEL} | {len(mrp_codes)} деталей | {with_demand} с потребностью")
tt.font = Font(bold=True, size=13, color="FFFFFF", name="Arial")
tt.fill = fill("C00000")
tt.alignment = Alignment(horizontal='center', vertical='center')
ws_r.merge_cells(f'A1:{get_column_letter(_RISK_LAST_COL)}1')
lg = ws_r.cell(2, 1)
lg.value = ("🔴 Дефицит = отрицательный баланс по симуляции графика поставок (остаток из Ввод_Остатков/файлов)  |  "
            "🟡 Сверхзапас > 3× норматива  |  Кол. 15–17: справочно ближайшие поставки и закрытие дефицита")
lg.font = Font(italic=True, size=9, color="555555", name="Arial")
lg.alignment = Alignment(wrap_text=True, vertical='center')
ws_r.merge_cells(f'A2:{get_column_letter(_RISK_LAST_COL)}2')
RH = ['Код', 'Наименование / Применяемость / Цвет', 'Поставщик', 'Ед.', 'Уп.', 'Остаток', 'Ср/день',
      'Норм.\nдней', 'Норм.\nтреб.', 'Дефицит', 'Дата\nдефицита', 'Дн.\nзапаса',
      'Статус', 'Комментарий', 'Ближ.\nпоставка', 'Кол-во\nпоставки', 'Закрытие\nдефицита']
RW = [22, 60, 16, 6, 7, 11, 10, 7, 12, 12, 12, 9, 16, 30, 12, 11, 12]
for ci, (h, w) in enumerate(zip(RH, RW), 1):
    hcell(ws_r, 3, ci, h, H_FILL)
    ws_r.column_dimensions[get_column_letter(ci)].width = w

risk_rows = []
for code in mrp_codes:
    name, supp, unit = get_info(code)
    stk = float(_opening_stock(code) or 0)
    total_d = sum(sum(demand[code].get(mn, {}).values()) for mn, _, _ in MONTHS)
    if total_d <= 0:
        continue
    avg = _risk_avg_daily(code)
    if avg <= 0:
        continue
    norm_days = get_safety_days(code, supp)
    norm_qty = round(avg * norm_days, 2)
    if norm_qty < 0.01:
        continue

    sim = _risk_simulation(code)
    first_def = sim['first_deficit']
    is_def = first_def is not None
    sd_val = round(stk / avg, 1) if avg > 0 else 9999
    is_ov = (sd_val > norm_days * 3) and sd_val < 9000 and not is_def
    if not (is_def or is_ov):
        continue

    pkg = bom.get(code, {}).get('package', 1)
    if is_def:
        def_qty = round(max(0.0, -sim['min_balance']), 1)
        if def_qty <= 0:
            def_qty = round(max(0.0, norm_qty - stk), 1)
        st = '🔴 ДЕФИЦИТ'
        parts = []
        if code in live_opening_stock:
            parts.append('остаток ручной')
        parts.append(f"мин. баланс {sim['min_balance']:.1f} {unit}")
        if sim['nearest_del_date']:
            parts.append(f"поставка {sim['nearest_del_qty']:.0f} {unit} "
                         f"{sim['nearest_del_date'].strftime('%d.%m')}")
        if sim['closure_date']:
            parts.append(f"закрытие {sim['closure_date'].strftime('%d.%m')}")
        comment = '; '.join(parts)
    else:
        def_qty = None
        st = '🟡 СВЕРХЗАПАС'
        comment = f"Запас {sd_val:.0f} дн., норм {norm_days} дн."

    if is_def and first_def and avg > 0:
        _sd = get_stock_date(code)
        sd_val = max(0.0, (first_def - _sd).days)
        sd_val = round(sd_val, 1)

    risk_rows.append({
        'code': code, 'name': _name_with_appl(code,name), 'supplier': supp, 'unit': unit, 'pkg': pkg,
        'stk': stk, 'avg': round(avg, 2), 'norm_days': norm_days, 'norm_qty': norm_qty,
        'def_qty': def_qty, 'def_date': first_def, 'sd_val': sd_val if sd_val < 9000 else None,
        'status': st, 'comment': comment, 'is_def': is_def,
        'nearest_del_date': sim['nearest_del_date'],
        'nearest_del_qty': sim['nearest_del_qty'],
        'closure_date': sim['closure_date'],
    })

risk_rows.sort(key=lambda x: (0 if x['is_def'] else 1, x['def_date'] or datetime.date(2099, 1, 1)))
for ri, r in enumerate(risk_rows, 4):
    ws_r.row_dimensions[ri].height = 15
    bg = RED_F if r['is_def'] else ORG_F
    vals = [
        r['code'], r['name'], r['supplier'], r['unit'], r['pkg'], r['stk'],
        r['avg'], r['norm_days'], r['norm_qty'], r['def_qty'],
        r['def_date'].strftime('%d.%m.%Y') if r['def_date'] else '',
        r['sd_val'], r['status'], r['comment'],
        r['nearest_del_date'].strftime('%d.%m.%Y') if r.get('nearest_del_date') else '',
        r.get('nearest_del_qty'),
        r['closure_date'].strftime('%d.%m.%Y') if r.get('closure_date') else '',
    ]
    for ci, v in enumerate(vals, 1):
        c = ws_r.cell(ri, ci)
        c.value = v
        c.fill = bg
        c.font = Font(size=9, name="Arial", bold=(ci in (1, 10, 13)),
                      color="C00000" if r['is_def'] and ci in (10, 13) else "000000")
        c.alignment = Alignment(horizontal='left' if ci in (2, 14) else 'center', vertical='center')
        if isinstance(v, float) and ci in (6, 7, 9, 10, 12, 16):
            c.number_format = '#,##0.0'
        if ci == 8 and isinstance(v, int):
            c.number_format = '0'

n_def = sum(1 for r in risk_rows if r['is_def'])
n_ov = len(risk_rows) - n_def
rs = 4 + len(risk_rows) + 1
cs = ws_r.cell(rs, 1)
cs.value = (f"ИТОГО: 🔴 Дефицит (симуляция): {n_def}  |  🟡 Сверхзапас: {n_ov}  |  "
            f"В расчёте: {len(mrp_codes)} (всего в BOM: {len(all_codes)})")
cs.font = Font(bold=True, size=10, name="Arial")
cs.fill = fill("F0F0F0")
ws_r.merge_cells(f'A{rs}:{get_column_letter(_RISK_LAST_COL)}{rs}')

# ── BOM_Детальный (editable live BOM) ────────────────────────
print("  BOM_Детальный (редактируемый)...")
from openpyxl.styles import Border, Side, numbers
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule, FormulaRule, ColorScaleRule

ws_bom = wb_out.create_sheet('BOM_Детальный')
ws_bom.freeze_panes = 'F5'
ws_bom.row_dimensions[1].height = 50
ws_bom.row_dimensions[2].height = 25
ws_bom.row_dimensions[3].height = 25
ws_bom.row_dimensions[4].height = 35

CFG_LABELS = [
    'A01\n2WD\ncomfort','A01\n2WD\nelite','A01\n2WD\npremium',
    'A01\n4WD\nelite','A01\n4WD\npremium','A01\n4WD\nTech+',
    'A08\n2WD\nelite','A08\n2WD\npremium','A08\n4WD\nelite',
    'A08\n4WD\npremium','A08\n4WD\nTech+',
    'B02\n2WD\nprem','B02\n2WD\nelite','B02\n4WD\nelite','B02\n4WD\nTech+',
    'B04\n4WD\nTech+','B04\n4WD\nprem',
    'B06\n4WD\nelite','B06\n4WD\nprem','B06\n4WD\nTech+',
    'B16\n4WD\nprem','B16\n4WD\nTech+',
]
MODEL_CLR = {
    'A01':fill("1F78B4"),'A08':fill("33A02C"),
    'B02':fill("E31A1C"),'B04':fill("FF7F00"),
    'B06':fill("6A3D9A"),'B16':fill("B15928"),
}
MODEL_COLS = {  # which columns belong to each model
    'A01': list(range(5,11)), 'A08': list(range(11,16)),
    'B02': list(range(16,20)),'B04': list(range(20,22)),
    'B06': list(range(22,25)),'B16': list(range(25,27)),
}

THIN_S = Side(style='thin', color='CCCCCC')
MED_S  = Side(style='medium', color='888888')
BORD_T = Border(left=THIN_S,right=THIN_S,top=THIN_S,bottom=THIN_S)
BORD_M = Border(left=MED_S,right=THIN_S,top=MED_S,bottom=MED_S)

# Row 1: Title
t_bom = ws_bom.cell(1, 1)
t_bom.value = (f'BOM_ДЕТАЛЬНЫЙ — ПРИМЕНЯЕМОСТЬ ПО {len(CFG_KEYS)} КОНФИГУРАЦИЯМ | '
               'Колонки F–{col}: ✓=в расчёте, пусто=нет | '
               'Или лист «Исключить_применяемость» (код в col A) → сохранить → перезапустить'
               .format(col=get_column_letter(BOM_LAST_CFG_COL)))
t_bom.font = Font(bold=True, size=12, color="FFFFFF", name="Arial")
t_bom.fill = fill("1F3864")
t_bom.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
ws_bom.merge_cells(f'A1:{get_column_letter(BOM_LAST_CFG_COL)}1')

# Row 2: Model group headers
for model, cols in MODEL_COLS.items():
    c = ws_bom.cell(2, cols[0]+1)
    c.value = model
    c.font = Font(bold=True, size=11, color="FFFFFF", name="Arial")
    c.fill = MODEL_CLR.get(model, H_FILL)
    c.alignment = Alignment(horizontal='center', vertical='center')
    if len(cols) > 1:
        ws_bom.merge_cells(f'{get_column_letter(cols[0]+1)}2:{get_column_letter(cols[-1]+1)}2')

# Row 3: Drive headers
DRIVE_GROUPS = {
    (5,7):'2WD', (8,10):'4WD',
    (11,12):'2WD', (13,15):'4WD',
    (16,17):'2WD', (18,19):'4WD',
    (20,21):'4WD',
    (22,24):'4WD',
    (25,26):'4WD',
}
for (c_start, c_end), drive in DRIVE_GROUPS.items():
    c = ws_bom.cell(3, c_start+1)
    c.value = drive
    c.font = Font(bold=True, size=9, color="FFFFFF", name="Arial")
    c.fill = fill("2E75B6") if drive=='2WD' else fill("375623")
    c.alignment = Alignment(horizontal='center', vertical='center')
    if c_start != c_end:
        ws_bom.merge_cells(f'{get_column_letter(c_start+1)}3:{get_column_letter(c_end+1)}3')

# Row 4: Column headers
for ci, h in enumerate(['№', 'Код', 'Наименование', 'Ед.', 'Поставщик'], 1):
    c = ws_bom.cell(4, ci)
    c.value = h; c.font = Font(bold=True, color="FFFFFF", size=9, name="Arial")
    c.fill = H_FILL; c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
# Упаковка (col28) и Примечания (col29) — редактируемые столбцы
c_upak_hdr = ws_bom.cell(4, 28)
c_upak_hdr.value = 'Упак.\n(шт)'
c_upak_hdr.font = Font(bold=True, color="FFFFFF", size=9, name="Arial")
c_upak_hdr.fill = fill("E67E22")  # оранжевый — редактируемое поле
c_upak_hdr.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
for ci, lbl in enumerate(CFG_LABELS, 6):
    c = ws_bom.cell(4, ci)
    c.value = lbl; c.font = Font(bold=True, color="FFFFFF", size=8, name="Arial")
    # Color by model
    model = CFG_KEYS[ci-6].split('_')[0]
    c.fill = MODEL_CLR.get(model, H_FILL)
    c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    if CFG_KEYS[ci-6] == 'B02_2WD_elite':
        _gold = Side(style='medium', color='FFD700')
        c.border = Border(left=_gold, right=_gold, top=_gold, bottom=_gold)
c_notes = ws_bom.cell(4, 29)
c_notes.value = 'Примечания'; c_notes.font = Font(bold=True, color="FFFFFF", size=9, name="Arial")
c_notes.fill = H_FILL; c_notes.alignment = Alignment(horizontal='center', vertical='center')
ws_bom.column_dimensions[get_column_letter(28)].width = 7  # Упак. col

# Column widths
ws_bom.column_dimensions['A'].width = 5
ws_bom.column_dimensions['B'].width = 20
ws_bom.column_dimensions['C'].width = 40
ws_bom.column_dimensions['D'].width = 5
ws_bom.column_dimensions['E'].width = 14
for i in range(6, BOM_LAST_CFG_COL + 2):
    ws_bom.column_dimensions[get_column_letter(i)].width = 6.5
ws_bom.column_dimensions[get_column_letter(BOM_LAST_CFG_COL + 1)].width = 20

# Data validation for config cells: only ✓ or empty
dv = DataValidation(type="custom", formula1='OR(F5="✓",ISNUMBER(F5))',
                    allow_blank=True, showErrorMessage=True, errorStyle='warning')
dv.error = 'Обычно ✓ (=1 на конфигурацию). Можно ввести ЧИСЛО (шт. на конфигурацию). Текст игнорируется при расчёте.'
dv.errorTitle = 'Применяемость: ✓ или число'
ws_bom.add_data_validation(dv)
_bom_cfg_end_row = 5 + len(all_codes) + len(bom_sections) * 2
dv.sqref = f'F5:{get_column_letter(BOM_LAST_CFG_COL)}{_bom_cfg_end_row}'

# Conditional formatting: ✓ = green background
from openpyxl.formatting.rule import CellIsRule
cf_range = f'F5:{get_column_letter(BOM_LAST_CFG_COL)}{_bom_cfg_end_row}'
ws_bom.conditional_formatting.add(cf_range,
    CellIsRule(operator='equal', formula=['"✓"'], fill=fill("C6EFCE"),
               font=Font(bold=True, color="375623", name="Arial")))

# Collect ordered parts with sections
ordered_parts = []
seen = set()
# First: use bom_sections order (from BOM source)
for sec, codes in bom_sections.items():
    ordered_parts.append(('section', sec))
    for code in codes:
        if code not in seen:
            ordered_parts.append(('part', code))
            seen.add(code)
# Then: any remaining parts not in sections
for code in all_codes:
    if code not in seen:
        ordered_parts.append(('part', code))
        seen.add(code)

# Write BOM rows
row_idx = 5
part_num = 0
bom_row_map = {}  # code -> excel row number

for item in ordered_parts:
    if item[0] == 'section':
        # Section separator row
        sec_name = item[1]
        ws_bom.row_dimensions[row_idx].height = 18
        c = ws_bom.cell(row_idx, 1)
        c.value = sec_name[:100]
        c.font = Font(bold=True, size=9, color="FFFFFF", name="Arial")
        c.fill = fill("2E4057")
        c.alignment = Alignment(horizontal='left', vertical='center')
        ws_bom.merge_cells(f'A{row_idx}:{get_column_letter(29)}{row_idx}')
        row_idx += 1
        continue

    code = item[1]
    part_num += 1
    ws_bom.row_dimensions[row_idx].height = 15
    fb = GRY_F if row_idx % 2 == 0 else NO_F

    info = bom.get(code, {})
    name = info.get('name', '')
    unit = info.get('unit', 'шт')
    supp = info.get('supplier', '')
    notes = info.get('notes', '')
    cfgs = info.get('configs', set())

    # Part info columns
    for ci, v in enumerate([part_num, code, name[:60], unit, supp[:20]], 1):
        c = ws_bom.cell(row_idx, ci); c.value = v
        c.font = Font(size=9, name="Arial")
        c.alignment = Alignment(horizontal='left' if ci in (2,3,5) else 'center', vertical='center')
        if fb: c.fill = fb

    # Config checkmark columns — write ✓ or numeric qty_per_car
    bom_cfgs_dict = info.get('configs', {}) if isinstance(info.get('configs'), dict) else {}
    _check_font = Font(bold=True, size=10, color="375623", name="Arial")
    _check_fill = fill("C6EFCE")
    _center = Alignment(horizontal='center', vertical='center')
    for cfg_idx, cfg_key in enumerate(CFG_KEYS):
        ci = 6 + cfg_idx
        if cfg_key in cfgs:
            c = ws_bom.cell(row_idx, ci)
            qty_pc = bom_cfgs_dict.get(cfg_key, 1)
            c.value = qty_pc if qty_pc != 1 else '✓'
            c.font = _check_font
            c.fill = _check_fill
            c.alignment = _center
        elif fb:
            c = ws_bom.cell(row_idx, ci)
            c.fill = fb
    # data validation: один диапазон на всю секцию (быстрее)

    # Упаковка (col28) — редактируемое поле
    c_pkg = ws_bom.cell(row_idx, 28)
    pkg_val = info.get('package', 1)
    c_pkg.value = pkg_val if pkg_val and pkg_val != 1 else None
    c_pkg.font = Font(size=9, name="Arial", bold=bool(pkg_val and pkg_val != 1))
    c_pkg.alignment = Alignment(horizontal='center', vertical='center')
    pkg_fill = fill("FDEBD0") if pkg_val and pkg_val != 1 else (fb if fb else NO_F)
    c_pkg.fill = pkg_fill

    # Notes (col29 — сдвинуто на 1)
    c_n = ws_bom.cell(row_idx, 29)
    c_n.value = notes[:30] if notes else None
    c_n.font = Font(size=8, color="555555", name="Arial", italic=True)
    c_n.alignment = Alignment(horizontal='left', vertical='center')
    if fb: c_n.fill = fb

    bom_row_map[code] = row_idx
    row_idx += 1

print(f"    BOM_Детальный: {part_num} деталей, {len(bom_sections)} разделов")

# ── Исключить_применяемость (ручные исключения из расчёта потребности) ──
print("  Исключить_применяемость...")
ws_excl = wb_out.create_sheet(BOM_NO_DEMAND_EXCLUDE_SHEET)
ws_excl.row_dimensions[1].height = 36
ws_excl.row_dimensions[2].height = 22
t_excl = ws_excl.cell(1, 1)
t_excl.value = ('ИСКЛЮЧИТЬ ИЗ РАСЧЁТА ПОТРЕБНОСТИ | Код в col A → пусто col B = без потребности | '
                'col B = «да» = вернуть в расчёт | Сохранить → перезапустить mrp_v10-1.py')
t_excl.font = Font(bold=True, size=10, color="FFFFFF", name="Arial")
t_excl.fill = fill("C00000")
t_excl.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
ws_excl.merge_cells('A1:C1')
for ci, (h, w) in enumerate([('Код детали', 22), ('В расчёте? (да = включить)', 24),
                              ('Наименование', 50)], 1):
    hcell(ws_excl, 2, ci, h, H_FILL)
    ws_excl.column_dimensions[get_column_letter(ci)].width = w
_excl_row = 3
for _excl_code in sorted(BOM_NO_DEMAND_CODES):
    _excl_info = bom.get(_excl_code, {})
    for ci, val in enumerate([_excl_code, '', _excl_info.get('name', '')], 1):
        c = ws_excl.cell(_excl_row, ci)
        c.value = val
        c.font = Font(size=9, name="Arial")
        c.alignment = Alignment(horizontal='left' if ci == 3 else 'center', vertical='center')
        if ci == 1:
            c.fill = fill("FCE4D6")
    _excl_row += 1
print(f"    {BOM_NO_DEMAND_EXCLUDE_SHEET}: {_excl_row - 3} кодов")

# ── Конфигурации (reference sheet) ────────────────────────────
ws_cfg = wb_out.create_sheet('Конфигурации')
ws_cfg.row_dimensions[1].height = 40
t_cfg = ws_cfg.cell(1, 1)
t_cfg.value = f"СПРАВОЧНИК {len(CFG_KEYS)} КОНФИГУРАЦИИ | Используется в BOM_Детальный и расчёте потребности"
t_cfg.font = Font(bold=True, size=11, color="FFFFFF", name="Arial")
t_cfg.fill = H3; t_cfg.alignment = Alignment(horizontal='center', vertical='center')
ws_cfg.merge_cells('A1:H1')

cfg_hdrs = ['Ключ конфигурации','Модель','Привод','Комплектация','Код кузова','Двигатель','Описание','Вкладка плана']
cfg_widths = [22, 8, 8, 14, 16, 12, 24, 24]
for ci, (h, w) in enumerate(zip(cfg_hdrs, cfg_widths), 1):
    hcell(ws_cfg, 2, ci, h, H_FILL)
    ws_cfg.column_dimensions[get_column_letter(ci)].width = w
ws_cfg.row_dimensions[2].height = 30

CFG_DATA = [
    ('A01_2WD_comfort','A01','2WD','comfort','CC7150BA00B','GW4G15K','A01 2WD comfort','WS_in (5)'),
    ('A01_2WD_elite','A01','2WD','elite','CC7150BA01B','GW4G15K','A01 2WD elite','WS_in (5)'),
    ('A01_2WD_premium','A01','2WD','premium','CC7150BA01B','GW4G15K','A01 2WD premium','WS_in (5)'),
    ('A01_4WD_elite','A01','4WD','elite','CC6450BA26B','GW4B15L','A01 4WD elite','WS_in (5)'),
    ('A01_4WD_premium','A01','4WD','premium','CC6450BA26B','GW4B15L','A01 4WD premium','WS_in (5)'),
    ('A01_4WD_TechPlus','A01','4WD','Tech Plus','CC6450BA26B','GW4B15L','A01 4WD Tech Plus','WS_in (5)'),
    ('A08_2WD_elite','A08','2WD','elite','CC6450BZ01C','GW4G15M','A08 2WD elite','WS_in (5)'),
    ('A08_2WD_premium','A08','2WD','premium','CC6450BZ01C','GW4G15M','A08 2WD premium','WS_in (5)'),
    ('A08_4WD_elite','A08','4WD','elite','CC6450BZ20E','GW4B15L','A08 4WD elite','WS_in (5)'),
    ('A08_4WD_premium','A08','4WD','premium','CC6450BZ20E','GW4B15L','A08 4WD premium','WS_in (5)'),
    ('A08_4WD_TechPlus','A08','4WD','Tech Plus','CC6450BZ20E','GW4B15L','A08 4WD Tech Plus','WS_in (5)'),
    ('B02_2WD_premium','B02','2WD','premium','CC6480AL00C','GW4B15L','B02 2WD premium','WS_in WS_in_H'),
    ('B02_2WD_elite','B02','2WD','elite','CC6480AL05A','GW4B15L','B02 2WD elite','WS_in WS_in_H'),
    ('B02_4WD_elite','B02','4WD','elite','CC6480AL23C','GW4N20A','B02 4WD elite','WS_in WS_in_H'),
    ('B02_4WD_TechPlus','B02','4WD','Tech Plus','CC6480AL23C','GW4N20A','B02 4WD Tech Plus','WS_in WS_in_H'),
    ('B04_4WD_TechPlus','B04','4WD','Tech Plus','CC6481AL23C','GW4N20A','B04 4WD Tech Plus','WS_in WS_in_H'),
    ('B04_4WD_Premium','B04','4WD','Premium','CC6481AL23C','GW4N20A','B04 4WD Premium','WS_in WS_in_H'),
    ('B06_4WD_elite','B06','4WD','elite','CC6460AZ26A','GW4N20A','B06 4WD elite','AS_in_H_B'),
    ('B06_4WD_premium','B06','4WD','premium','CC6460AZ26A','GW4N20A','B06 4WD premium','AS_in_H_B'),
    ('B06_4WD_TechPlus','B06','4WD','Tech Plus','CC6460AZ26A','GW4N20A','B06 4WD Tech Plus','AS_in_H_B'),
    ('B16_4WD_premium','B16','4WD','premium','CC6470CF27A','GW4N20A','B16 4WD premium','AS_in_H_B'),
    ('B16_4WD_TechPlus','B16','4WD','Tech Plus','CC6470CF27A','GW4N20A','B16 4WD Tech Plus','AS_in_H_B'),
]
MODEL_BG = {'A01':fill("DBEAFE"),'A08':fill("DCFCE7"),'B02':fill("FEE2E2"),
            'B04':fill("FEF3C7"),'B06':fill("F3E8FF"),'B16':fill("FFF7ED")}
for ri, row_data in enumerate(CFG_DATA, 3):
    ws_cfg.row_dimensions[ri].height = 16
    model = row_data[1]
    bg = MODEL_BG.get(model, NO_F)
    for ci, v in enumerate(row_data, 1):
        c = ws_cfg.cell(ri, ci); c.value = v
        c.font = Font(size=10, name="Arial", bold=(ci==1))
        c.fill = bg
        c.alignment = Alignment(horizontal='center' if ci != 3 else 'left', vertical='center')

print(f"    Конфигурации: {len(CFG_DATA)} конфигураций")

# ═══ Фильтры по колонкам + удобный VLOOKUP на КАЖДОЙ вкладке ═══
# Включаем автофильтр на строке заголовка (там, где «Код»...), и разъединяем
# объединённые ячейки В ТЕЛЕ таблицы (ниже заголовка) — иначе фильтр и VLOOKUP
# ломаются. Заливки/шрифты (форматирование) сохраняются; объединения шапки/титула — тоже.
def _enable_column_filter(ws):
    hdr = None
    for r in range(1, min(ws.max_row, 8) + 1):
        for c in (1, 2, 3):
            v = ws.cell(r, c).value
            if v and 'Код' in str(v):
                hdr = r
                break
        if hdr:
            break
    if not hdr:
        return
    for rng in list(ws.merged_cells.ranges):
        if rng.min_row > hdr:               # объединения в теле — разъединяем (формат остаётся)
            ws.unmerge_cells(str(rng))
    if ws.max_row > hdr and ws.max_column >= 1:
        ws.auto_filter.ref = f"A{hdr}:{get_column_letter(ws.max_column)}{ws.max_row}"

for _ws in wb_out.worksheets:
    if _ws.title == 'Ввод_Поставок':
        continue
    try:
        _enable_column_filter(_ws)
    except Exception as _ef:
        print(f"  ⚠️  Фильтр не добавлен для «{_ws.title}»: {_ef}")
print("  Автофильтр по колонкам включён на всех вкладках; тело без объединений (VLOOKUP-friendly)")

def _apply_workbook_sheet_order(wb):
    """Порядок вкладок: Риск → Ввод → График → Потребности → Заказы → Сводка → прочие."""
    demand_titles = [f"Потребность_{mlabel[:3]}" for _, mlabel, _ in MONTHS]
    order_titles = [f"Заказы_{mlabel[:3]}" for _, mlabel, _ in MONTHS]
    priority = (
        ['Риск_Дефицита', 'Ввод_Остатков', 'Планы_Производств', 'График_Поставок', 'Ввод_Поставок']
        + demand_titles
        + order_titles
        + ['Сводка_3мес', 'Прогноз_CKD']
    )
    by_title = {ws.title: ws for ws in wb.worksheets}
    ordered = []
    seen = set()
    for title in priority:
        ws = by_title.get(title)
        if ws and title not in seen:
            ordered.append(ws)
            seen.add(title)
    for ws in wb.worksheets:
        if ws.title not in seen:
            ordered.append(ws)
            seen.add(ws.title)
    wb._sheets = ordered

_apply_workbook_sheet_order(wb_out)
print(f"  Порядок листов: {[s.title for s in wb_out.worksheets[:12]]}...")

# ═══ Сохраняем ═══════════════════════════════════════════════
_out_sheet_titles = [s.title for s in wb_out.worksheets]
OUT_SAVED = _save_workbook_safe(wb_out, OUT)
try:
    wb_out.close()
except Exception:
    pass
print(f"\n✅ Готово! Файл сохранён: {OUT_SAVED}")
print(f"   Расчётный период: {PERIOD_LABEL}")
print(f"   Листы: {_out_sheet_titles}")

if _MASTER_BOM_SYNC_PATH and os.path.exists(_MASTER_BOM_SYNC_PATH):
    print("\nMaster BOM sync...")
    _sync_master_bom_detailed(bom, _MASTER_BOM_SYNC_PATH)

# ═══ Выгрузка графиков поставок по поставщикам ════════════════
# Пересчитываем Del/Ss в Python (не читаем формулы из Excel — openpyxl их не вычисляет).
# Логика идентична График_Поставок:
#   del_d = CEILING(MAX(0, safety-(ss_prev-dem_d-dem_next))/pkg, 1)*pkg  если ss_prev-dem_d-dem_next < safety
#   ss_d  = ss_prev - dem_d + del_d
# Выгружаем неделю Пн–Вс: текущую (корректировка) или следующую (новый график).
print("\nВыгрузка графиков поставок по поставщикам...")

def _prompt_supplier_export_week_mode():
    """current = корректировка текущей недели; next = новая (следующая) неделя."""
    env = os.environ.get('MRP_EXPORT_WEEK', '').strip().lower()
    if env in ('current', 'cur', '1', 'текущая', 'т', 'this'):
        return 'current'
    if env in ('next', 'new', '2', 'новая', 'н', 'следующая'):
        return 'next'
    if not sys.stdin.isatty():
        return 'next'
    print("  Папка «Графики_поставщиков» — какую неделю выгрузить?")
    print("    [1] Корректировать ТЕКУЩУЮ неделю  (Пн–Вс этой недели, пересчёт)")
    print("    [2] Считать НОВУЮ неделю           (Пн–Вс следующей недели)")
    while True:
        try:
            ans = input("  Выбор [1/2, Enter=2]: ").strip().lower()
        except EOFError:
            return 'next'
        if ans in ('', '2', 'n', 'next', 'новая', 'н', 'след'):
            return 'next'
        if ans in ('1', 'c', 'current', 'текущая', 'т', 'cur'):
            return 'current'
        print("  ⚠️  Введите 1 или 2")

try:
    import math as _math

    _export_week_mode = _prompt_supplier_export_week_mode()

    # ── Границы недели для выгрузки ───────────────────────────────
    _today_run = datetime.date.today()
    _wd = _today_run.weekday()                              # 0=Mon
    _cur_mon = _today_run - datetime.timedelta(days=_wd)   # Пн текущей недели
    _cur_sun = _cur_mon + datetime.timedelta(days=6)       # Вс текущей недели
    _next_mon = _cur_mon + datetime.timedelta(days=7)      # Пн следующей недели
    _next_sun = _next_mon + datetime.timedelta(days=6)     # Вс следующей недели

    if _export_week_mode == 'current':
        _week_mon, _week_sun = _cur_mon, _cur_sun
        _week_mode_ru = 'текущая неделя (корректировка)'
    else:
        _week_mon, _week_sun = _next_mon, _next_sun
        _week_mode_ru = 'новая неделя'

    print(f"  Режим: {_week_mode_ru}")
    print(f"  Период выгрузки: {_week_mon.strftime('%d.%m.%Y')} – {_week_sun.strftime('%d.%m.%Y')}")

    # Индексы дней из all_dates, попадающих в выбранную неделю
    _nw_indices = [(di, dt) for di, (dt, mn, d) in enumerate(all_dates)
                   if _week_mon <= dt <= _week_sun]

    if not EXPORT_SUPPLIERS:
        print("  Выгрузка поставщикам отключена (галочка «Выгружать графики поставщиков» "
              "в run_mrp_gui или MRP_EXPORT_SUPPLIERS=1)")
    elif not _nw_indices:
        print("  ⚠️  Выбранная неделя выходит за пределы расчётного горизонта — выгрузка пропущена.")
    else:
        _export_dir = os.path.join(ROOT, "Графики_поставщиков")
        os.makedirs(_export_dir, exist_ok=True)

        _date_tag = _today_run.strftime('%Y-%m-%d')
        _week_label = f"{_week_mon.strftime('%d.%m')}–{_week_sun.strftime('%d.%m.%Y')}"
        _week_file_tag = (
            f"{'cur' if _export_week_mode == 'current' else 'new'}"
            f"_{_week_mon.strftime('%Y%m%d')}-{_week_sun.strftime('%Y%m%d')}"
        )

        # ── Для каждой детали пересчитываем Del/Ss по всему горизонту ──
        # Нам нужен ss на начало следующей недели, поэтому считаем с di=0.
        # daily_dem_all[code] = список len(all_dates) значений спроса
        def _sim_schedule(code):
            """Возвращает {di: (del_d, ss_d)} — из DELIVERY_SCHEDULES."""
            return DELIVERY_SCHEDULES.get(code, {}).get('pairs', {})

        # ── Группируем коды по поставщику ────────────────────────
        _by_supp = {}
        for _code in mrp_codes:
            _supp = bom.get(_code, {}).get('supplier', '') or '—'
            _supp = _supp.strip() or '—'
            _by_supp.setdefault(_supp, []).append(_code)

        _n_exported = 0

        for _supp_name in sorted(_by_supp):
            _codes_s = _by_supp[_supp_name]

            # Пересчёт расписания и фильтрация по наличию поставок на неделю
            _rows_out = []   # (code, name, unit, pkg, safety_qty, {di: del_d})
            for _code in _codes_s:
                _sched = _sim_schedule(_code)
                _week_dels = {di: _sched[di][0] for di, _ in _nw_indices
                              if _sched.get(di, (0,))[0] > 0}
                _total3 = sum(sum(demand[_code].get(mn, {}).values()) for mn, _, _ in MONTHS)
                # Показываем ВСЕ позиции поставщика с потребностью (а не только те,
                # что отгружаются на следующей неделе) — кол-во позиций = кол-ву в потребности.
                if _total3 <= 0 and not _week_dels:
                    continue
                _info   = bom.get(_code, {})
                _pkg    = max(1, _info.get('package', 1))
                _safety = _safety_qty(_code, _supp_name)
                _rows_out.append((_code,
                                  _info.get('name', ''),
                                  _info.get('unit', 'шт'),
                                  _pkg,
                                  _safety,
                                  _week_dels))

            if not _rows_out:
                continue

            # ── Строим книгу Excel для поставщика ────────────────
            _wb_s = Workbook()
            _ws_s = _wb_s.active
            _ws_s.title = "График поставок"

            _n_day_cols = len(_nw_indices)
            _total_c    = 6 + _n_day_cols   # A=Код B=Наим C=Ед D=Уп E=Страх.зап F=ИТОГО + дни

            # Строка 1: заголовок
            _ws_s.merge_cells(f'A1:{get_column_letter(_total_c)}1')
            _hc = _ws_s.cell(1, 1)
            _hc.value = (f"ГРАФИК ПОСТАВОК  |  {_supp_name}  |  "
                         f"{_week_mode_ru}: {_week_label}  |  "
                         f"Сформировано {_today_run.strftime('%d.%m.%Y')}")
            _hc.font      = Font(bold=True, size=11, color="FFFFFF", name="Arial")
            _hc.fill      = H_FILL
            _hc.alignment = Alignment(horizontal='center', vertical='center')
            _ws_s.row_dimensions[1].height = 30

            # Строка 2: заголовки колонок
            _ws_s.row_dimensions[2].height = 28
            _ch = ['Код детали', 'Наименование / Применяемость / Цвет', 'Ед.', 'Уп.(шт)', 'Страх.\nзапас', 'ИТОГО\nнеделя']
            _cw = [22, 60, 6, 7, 10, 10]
            for _di, _dt in _nw_indices:
                _ch.append(_dt.strftime('%d.%m\n%a').replace('Mon','Пн').replace('Tue','Вт')
                           .replace('Wed','Ср').replace('Thu','Чт').replace('Fri','Пт')
                           .replace('Sat','Сб').replace('Sun','Вс'))
                _cw.append(10)
            for _ci, (_h, _w) in enumerate(zip(_ch, _cw), 1):
                _c = _ws_s.cell(2, _ci)
                _c.value = _h
                _c.font  = Font(bold=True, size=9, color="FFFFFF", name="Arial")
                _c.fill  = H2
                _c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                _ws_s.column_dimensions[get_column_letter(_ci)].width = _w

            # Строки данных
            for _ri, (_code, _name, _unit, _pkg, _safety, _wdels) in enumerate(_rows_out, 3):
                _ws_s.row_dimensions[_ri].height = 15
                _fb = GRY_F if _ri % 2 == 0 else NO_F
                _week_total = sum(_wdels.values())

                _static = [_code, _name_with_appl(_code,_name), _unit, _pkg, _safety, _week_total]
                for _ci, _v in enumerate(_static, 1):
                    _c = _ws_s.cell(_ri, _ci)
                    _c.value = _v
                    _c.font  = Font(size=9, name="Arial",
                                    bold=(_ci == 6),
                                    color=("375623" if _ci == 6 else "000000"))
                    _c.alignment = Alignment(
                        horizontal='left' if _ci <= 2 else 'center',
                        vertical='center')
                    if _ci == 6:
                        _c.fill = fill("C6EFCE")
                        _c.number_format = '#,##0'
                    elif _ci in (4, 5):
                        _c.number_format = '#,##0.#'
                        if _fb: _c.fill = _fb
                    elif _fb:
                        _c.fill = _fb

                # Колонки дней
                for _col_i, (_di, _dt) in enumerate(_nw_indices, 7):
                    _dv = _wdels.get(_di, 0)
                    _c  = _ws_s.cell(_ri, _col_i)
                    _c.value = int(_dv) if _dv > 0 else None
                    _c.alignment = Alignment(horizontal='center', vertical='center')
                    if _dv > 0:
                        _c.font  = Font(bold=True, size=9, color="375623", name="Arial")
                        _c.fill  = fill("C6EFCE")
                        _c.number_format = '#,##0'
                    else:
                        _c.font = Font(size=9, name="Arial")

                        if _fb: _c.fill = _fb

            # Итоговая строка
            _tot_r = 3 + len(_rows_out)
            _ws_s.row_dimensions[_tot_r].height = 16
            for _ci in range(1, _total_c + 1):
                _c = _ws_s.cell(_tot_r, _ci)
                if _ci >= 6:
                    _c.value = f"=SUM({get_column_letter(_ci)}3:{get_column_letter(_ci)}{_tot_r-1})"
                    _c.number_format = '#,##0'
                    _c.font = Font(bold=True, size=9, color="FFFFFF", name="Arial")
                    _c.fill = H_FILL
                    _c.alignment = Alignment(horizontal='center', vertical='center')
                else:
                    _c.value = 'ИТОГО' if _ci == 1 else None
                    _c.font = Font(bold=True, size=9, color="FFFFFF", name="Arial")
                    _c.fill = H_FILL
                    _c.alignment = Alignment(horizontal='left' if _ci == 1 else 'center',
                                             vertical='center')

            # Сохраняем — в подпапку поставщика
            _safe_name = re.sub(r'[\\/:*?"<>|]', '_', _supp_name)
            _supp_subdir = os.path.join(_export_dir, _safe_name)
            os.makedirs(_supp_subdir, exist_ok=True)
            _out_path  = os.path.join(_supp_subdir, f"{_safe_name}_{_week_file_tag}.xlsx")
            _wb_s.save(_out_path)
            _n_exported += 1
            print(f"    ✅ {_supp_name}: {len(_rows_out)} позиций → {os.path.basename(_out_path)}")

        print(f"  Экспорт завершён ({_week_mode_ru}): {_n_exported} файлов → "
              f"папка '{os.path.basename(_export_dir)}/'")

except Exception as _exp_err:
    print(f"  ⚠️  Ошибка выгрузки графиков: {_exp_err}")

# v10-1: устранение отрицательных остатков в График_Поставок (уч

# ═══ Выгрузка МЕСЯЧНЫХ ЗАКАЗОВ по поставщикам (галочка «Месячные заказы») ═══
# Для каждого поставщика — файл в «Monthly orders/<Поставщик>/» с листами
# Horisontal template for Docs (для согласования/подписания) и calc/calc_проверка
# (для проверки). Оформление и масштаб печати — как в образцах.
#
# Нумерация (DP-PR-092): LOC + ГГГГ + ММ(месяц выпуска) + порядковый номер.
# Горизонт: N+5 → 6 мес. (AVTOKOM, MSA, AAT, Yapp, Fuyao, Itelma, Lear, MTS, VM, SGK, GSK, Technoform);
#   N+3 → Autotechnika (4 мес.); N+2 → все остальные (3 мес.). Смещение N+2 для части поставщиков — в monthly_orders_export.
if CALC_ORDERS:
    print("\nВыгрузка месячных заказов по поставщикам...")
    try:
        import monthly_orders_export as _mo
        _n_ord = _mo.export_monthly_orders(
            root=ROOT,
            bom=bom,
            mrp_codes=mrp_codes,
            projected_stock=projected_stock,
            forecast_demand=forecast_demand,
            forecast_orders=forecast_orders,
            months=MONTHS,
            forecast_months=FORECAST_MONTHS,
            month_year=MONTH_YEAR,
            stock_as_of_month=STOCK_AS_OF_MONTH,
            calc_order=calc_order,
            get_info=get_info,
            pkg_size_fn=_pkg_size,
        )
        print(f"  Месячные заказы выгружены: {_n_ord} файлов → {os.path.join(ROOT, 'Monthly orders')}")
    except Exception as _ord_err:
        import traceback
        print(f"  ⚠️  Ошибка выгрузки заказов: {_ord_err}")
        traceback.print_exc()
