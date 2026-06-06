"""
MRP v9 — устранение замечаний
=============================
[1] остатки бамперов = улица + линия + буфер (а не один столбец)
[2] цветная декомпозиция плана для бамперов и красок через
    Order_calculation_statistics_UPDATED.xlsx (батч → цвет → шт)
[3] задние бампера XST33* — только на дорестайл-партиях
[4] B16 elite (подголовники) — ТОЛЬКО B16_4WD_elite
[5] потребность в краске = норма × кол-во кузовов нужных цветов
[6] BOM_Детальный — корректный offset столбцов (A01_2WD_comfort не пуст)
[7] case B04_4WD_Premium → premium (нормализация)
[8] Сводка_3мес: порядок колонок данных = порядок шапки
[9] Риск_Дефицита: «Ср/день» = потребность месяца / число дней месяца
[10] авто-сканирование всех файлов остатков
"""
import pandas as pd
import re
import math, os, json, datetime, warnings, glob
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.utils.datetime import from_excel
warnings.filterwarnings('ignore')

# ══════════════════════════════════════════════════════════════
# АВТО-ОПРЕДЕЛЕНИЕ ФАЙЛОВ
# Все входящие файлы ищутся по паттерну — просто положи новый файл
# в папку, старый можно оставить или удалить. Скрипт возьмёт
# самый свежий файл по каждому паттерну (по дате изменения).
# ══════════════════════════════════════════════════════════════

# ROOT = папка где лежит скрипт. Работает на любом компьютере без изменений.
ROOT = os.path.dirname(os.path.abspath(__file__)) if '__file__' in globals() else os.getcwd()

def _latest(folder, *patterns):
    """Возвращает самый свежий файл среди всех паттернов, или None."""
    candidates = []
    for pat in patterns:
        candidates.extend(glob.glob(os.path.join(folder, pat)))
        candidates.extend(glob.glob(os.path.join(folder, '**', pat), recursive=True))
    candidates = [p for p in candidates if os.path.isfile(p)]
    return max(candidates, key=os.path.getmtime) if candidates else None

def _require(path, label):
    if path and os.path.exists(path): return path
    print(f"  ⚠️  Файл не найден: {label}")
    return path or ''

# ── Основные файлы ────────────────────────────────────────────
PF_FILE  = _require(_latest(ROOT, "*Plan-Fact*.xlsx", "*Plan_Fact*.xlsx",
                             "*план-факт*.xlsx", "*планфакт*.xlsx"),
                    "Plan-Fact (план производства)")

BOM_FILE = _require(_latest(ROOT, "Master_BOM_актуальный.xlsx",
                             "Master_BOM*.xlsx"),
                    "Master_BOM")
BOM_NEW  = _latest(ROOT, "Master_BOM_актуальный.xlsx") or BOM_FILE

PKG_FILE = _require(_latest(ROOT, "Упаковка локала*.xlsx"),
                    "Упаковка локала")  # только точный паттерн

ADD_FILE = _require(_latest(ROOT, "Additional.xlsx", "Additional*.xlsx"),
                    "Additional")

# ── Теоретические остатки (высший приоритет, перекрывает все другие) ──
THEOR_STOCK_FILE = _latest(ROOT,
                            "Теоретические остатки*.xlsx",
                            "теоретические*остатки*.xlsx",
                            "Теор*остатки*.xlsx")

OUT      = os.path.join(ROOT, "MRP_System_v9.xlsx")

# ── Paint statistics: авто-генерация из GWM-файла ─────────────
# 1. Ищем входной GWM-файл (各车型成套批次统计表) в папке MRP
# 2. Если найден — запускаем build_order_calc_v2 для генерации статистики
# 3. Используем свежий Order_calculation_statistics_UPDATED.xlsx
_PAINT_OUT = os.path.join(ROOT, "Order_calculation_statistics_UPDATED.xlsx")
_GWM_INPUT = _latest(ROOT, "*批次*統計*.xlsx", "*各车型*.xlsx",
                     "*批次*统计*.xlsx", "*GWM*批次*.xlsx")
# Исключаем Plan-Fact и выходной файл
if _GWM_INPUT and any(x in os.path.basename(_GWM_INPUT) for x in
                      ["Order_calculation", "Plan-Fact", "Plan_Fact"]):
    _GWM_INPUT = None

if _GWM_INPUT:
    print(f"[INIT] GWM-файл найден: {os.path.basename(_GWM_INPUT)}")
    print(f"[INIT] Генерация paint statistics...")
    try:
        # Ищем build_order_calc_v2.py рядом со скриптом MRP или в соседних папках
        import sys as _sys
        _boc_candidates = [
            os.path.join(ROOT, "build_order_calc_v2.py"),
            os.path.join(os.path.dirname(ROOT), "files FINAL", "build_order_calc_v2.py"),
            os.path.join(os.path.dirname(ROOT), "files", "build_order_calc_v2.py"),
            os.path.join(os.path.expanduser("~"), "Downloads", "files FINAL", "build_order_calc_v2.py"),
            os.path.join(os.path.expanduser("~"), "Downloads", "files", "build_order_calc_v2.py"),
        ]
        _boc_path = next((p for p in _boc_candidates if os.path.exists(p)), None)
        if _boc_path:
            _boc_dir = os.path.dirname(_boc_path)
            if _boc_dir not in _sys.path:
                _sys.path.insert(0, _boc_dir)
            # Импортируем функцию и запускаем
            import importlib.util as _ilu
            _spec = _ilu.spec_from_file_location("build_order_calc_v2", _boc_path)
            _boc_mod = _ilu.module_from_spec(_spec)
            _spec.loader.exec_module(_boc_mod)
            _boc_mod.build_output(_GWM_INPUT, _PAINT_OUT)
            print(f"[INIT] ✅ Paint statistics обновлена → {os.path.basename(_PAINT_OUT)}")
        else:
            print(f"[INIT] ⚠️  build_order_calc_v2.py не найден — используем существующий файл статистики")
    except Exception as _boc_e:
        print(f"[INIT] ⚠️  Ошибка генерации paint statistics: {_boc_e}")
else:
    print(f"[INIT] ℹ️  GWM-файл не найден — используем существующий файл статистики")

# Итоговый путь к файлу статистики
PAINT_STATS = (_PAINT_OUT if os.path.exists(_PAINT_OUT) else
               _latest(ROOT, "*Order*calculation*.xlsx", "*涂装*.xlsx") or
               os.path.join(ROOT, "Замечания", "Paiint calcualtion",
                            "Order_calculation_statistics_UPDATED.xlsx"))

# ── Бамперы ───────────────────────────────────────────────────
BUMPER_FILE = _latest(ROOT, "*бампер*.xlsx", "*Bumper*.xlsx",
                      "*Пересчет бамперов*.xlsx") or ''

# ── Файлы остатков — авто-сканирование ───────────────────────
# Всё что НЕ является системным файлом и содержит нужные ключевые слова
_SYSTEM_FILES = {
    os.path.normcase(p) for p in [
        OUT, PF_FILE, BOM_FILE, BOM_NEW, PKG_FILE, ADD_FILE,
        THEOR_STOCK_FILE,                          # теоретические остатки — отдельная загрузка
        os.path.join(ROOT, "MRP_System_v8.xlsx"),
        os.path.join(ROOT, "MRP_System_v9_new.xlsx"),
        os.path.join(ROOT, "Stock in days.xlsx"),  # справочник норм — не файл остатков
    ] if p
}
# Группы типов файлов остатков: (название_группы, [ключевые_слова])
# Из каждой группы берётся ТОЛЬКО самый свежий файл по дате изменения.
_STOCK_GROUPS = [
    ('AAT',         ['aat']),
    ('Warehouse',   ['warehouse']),
    ('GSK',         ['gsk']),
    ('Purem',       ['пюрэм', 'пюрем', 'purem']),
    ('Bumpers',     ['пересчет', 'бампер', 'bumper']),
    ('Accounting',  ['accounting']),
    ('Остатки',     ['остатк', 'остаток']),
    ('Lear',        ['lear']),
    ('Fuyao',       ['fuyao']),
    ('Yapp',        ['yapp']),
    ('STP',         ['stp']),
    ('GSK2',        ['gsk']),
]
_SKIP_IN_STOCK = ['mrp_system','plan-fact','plan_fact','master_bom',
                   'упаковка','additional','order_calc','order calculation','涂装']

# Сортируем все xlsx по дате изменения (сначала новые)
_all_xlsx = sorted(glob.glob(os.path.join(ROOT, "*.xlsx")),
                   key=os.path.getmtime, reverse=True)

STOCK_FILES = []
_seen_groups = set()   # группы, для которых файл уже найден
_skipped_old = []      # старые файлы — пропускаем

for p in _all_xlsx:
    if os.path.normcase(p) in _SYSTEM_FILES: continue
    base = os.path.basename(p).lower()
    if any(x in base for x in _SKIP_IN_STOCK): continue
    # Проверяем к какой группе относится файл
    matched_group = None
    for grp_name, kws in _STOCK_GROUPS:
        if any(kw in base for kw in kws):
            matched_group = grp_name
            break
    if matched_group is None: continue
    if matched_group in _seen_groups:
        _skipped_old.append(os.path.basename(p))  # уже есть более свежий
    else:
        _seen_groups.add(matched_group)
        STOCK_FILES.append(p)

print(f"[INIT] ROOT = {ROOT}")
print(f"[INIT] Plan-Fact: {os.path.basename(PF_FILE) if PF_FILE else '❌ НЕ НАЙДЕН'}")
print(f"[INIT] BOM:       {os.path.basename(BOM_FILE) if BOM_FILE else '❌ НЕ НАЙДЕН'}")
print(f"[INIT] Теор.ост.: {os.path.basename(THEOR_STOCK_FILE) if THEOR_STOCK_FILE else 'не найден (необязательно)'}")
print(f"[INIT] Бамперы:   {os.path.basename(BUMPER_FILE) if BUMPER_FILE else '❌ НЕ НАЙДЕН'}")
print(f"[INIT] PaintStats:{os.path.basename(PAINT_STATS) if PAINT_STATS and os.path.exists(PAINT_STATS) else '❌ НЕ НАЙДЕН'}")
print(f"[INIT] Найдено {len(STOCK_FILES)} файлов остатков (берётся самый свежий из каждого типа):")
for p in STOCK_FILES:
    print(f"       ✅ {os.path.basename(p)}")
if _skipped_old:
    print(f"[INIT] Пропущено устаревших:")
    for nm in _skipped_old:
        print(f"       ⏩ {nm}")

# ── Устаревшие коды (не поставляются, исключаем из расчёта) ──
OBSOLETE_CODES = {
    '7005110XKJ22A',     # больше не поставляется (заменён на другой)
    '3101100xst33A',     # дубль 3101100XST33A (lowercase) — удалить
    # ── Дорестайл-задние бампера B02 XKN61 (выходят из оборота, потребности нет) ──
    '2804111XKN61A8T','2804111XKN61A9C','2804111XKN61AC3','2804111XKN61AH4',
    '2804112XKN61A8T','2804112XKN61A9C','2804112XKN61AC3','2804112XKN61AH4',
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
}

# ── Правила графиков поставок по поставщикам (канонические имена из BOM) ──
ECOALYANCE_SUPPLIER = "Ecoal'yance"
ECOTEXIS_SUPPLIER   = 'Ecotexis'
SMC_SUPPLIER        = 'SMC'
PUREM_SUPPLIER      = 'Purem'
SMC_MAX_PALLETS_PER_TRUCK = 7
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
    # LK015530 (SGK): в BOM ошибочно помечен на B02+B04; должен только B04
    'LK015530': {'B04_4WD_TechPlus', 'B04_4WD_premium'},
    # 1101100XKJ23A (Yapp): в BOM нет отметок о применяемости → без override
    # код использует все конфиги (некорректно). XKJ23A = A08, применяется на всех A08.
    '1101100XKJ23A': {'A08_2WD_elite', 'A08_2WD_premium',
                      'A08_4WD_elite', 'A08_4WD_premium', 'A08_4WD_TechPlus'},
    # TODO (Корнеева): Экоальянс 1205112XST11A / 1205113XST11A
    #   Должно: A01 4*2 new + A08 4*4 old + B02 4*2 old
    #   Сложно без учёта партий (RAS2117, RBU1084, RBA1022) — требует batch-логики.
    # TODO (Корнеева): 1205526XGW01B / 1205527XGW01A — неверные остатки в файлах.
    #   Требует корректировки данных в источниках остатков.
    # TODO (Ильина): ALAB001234 — расчёт не соответствует действительности.
    #   Требует уточнения норм расхода (chem_norms) у ответственного.
}

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
    if not cfgs or 'B02_2WD_elite' in cfgs:
        return False
    if not any(k.startswith('B02_') for k in cfgs):
        return False
    has_prem = 'B02_2WD_premium' in cfgs
    has_elite = 'B02_4WD_elite' in cfgs
    has_tech = 'B02_4WD_TechPlus' in cfgs
    # все B02 (3 существующие конфигурации до elite 2WD)
    if B02_ALL_THREE <= set(cfgs):
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
           '4x2':'2WD','4x4':'4WD'}
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
    _mm = _re.search(r'[._](\d{2})[._](\d{2})[._](\d{2})\.xlsx$', _bn, _re.IGNORECASE)
    if _mm:
        try:
            _d, _m, _y = int(_mm.group(1)), int(_mm.group(2)), 2000+int(_mm.group(3))
            if 1<=_d<=31 and 1<=_m<=12:
                _stock_date = datetime.date(_y, _m, _d)
                break
        except: pass
if _stock_date is None:
    _stock_date = _today  # fallback — текущая дата

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

print(f"[INIT] Расчётный период: {PERIOD_LABEL}  (STOCK_AS_OF: {STOCK_AS_OF_DAY:02d}.{STOCK_AS_OF_MONTH:02d}.{_stock_date.year})")

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
SAFETY_DAYS_FILE = os.path.join(ROOT, "Stock in days.xlsx")  # Per-part safety stock file

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
import datetime as _dt

def _wh16_qty_norm(raw_qty, raw_weight, unit_str):
    """WH16: qty_result = M × F, нормализованное в базовые единицы."""
    try:
        q = float(raw_qty) if raw_qty is not None else 0
        w = float(str(raw_weight).replace(',','.')) if raw_weight is not None else 1
        unit = str(unit_str).lower().strip().rstrip('.')
        total = q * w
        if unit in ('g','г','г.'): return total/1000   # г → кг
        if unit in ('ml','мл'):    return total/1000   # мл → л
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
    if len(ws_rows) < 3: return result
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
    if not triplets: return result
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
                if total > 0:
                    result[code] = result.get(code, 0) + total
                break  # нашли последнюю дату с данными для этой строки
    return result

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


def _parse_file_date(fpath):
    """Извлекает дату из имени файла (паттерн .ДД.ММ.ГГ.xlsx).
    Fallback — дата изменения файла."""
    bn = os.path.basename(fpath)
    mm = _re.search(r'[._](\d{2})[._](\d{2})[._](\d{2})\.xlsx$', bn, _re.IGNORECASE)
    if mm:
        try:
            d, m, y = int(mm.group(1)), int(mm.group(2)), 2000 + int(mm.group(3))
            if 1 <= d <= 31 and 1 <= m <= 12:
                return datetime.date(y, m, d)
        except: pass
    # Fallback: дата изменения файла
    try:
        return datetime.date.fromtimestamp(os.path.getmtime(fpath))
    except:
        return datetime.date.today()


def load_all_stock(file_list):
    stock = {}; sources = {}; dates = {}   # dates: code -> datetime.date
    for fpath in file_list:
        if not os.path.exists(fpath) or os.path.getsize(fpath) == 0:
            print(f"  ⚠️  Файл не найден: {os.path.basename(fpath)}"); continue
        fname = os.path.basename(fpath)
        fdate = _parse_file_date(fpath)   # дата этого файла
        try: wb = load_workbook(fpath, read_only=True, data_only=True)
        except Exception as e:
            print(f"  ❌ Ошибка чтения {fname}: {e}"); continue
        file_stock = {}
        for sname in wb.sheetnames:
            try:
                ws = wb[sname]
                rows = list(ws.iter_rows(values_only=True))
            except: continue
            if len(rows) < 2: continue
            h0 = rows[0]; h1 = rows[1] if len(rows) > 1 else []
            code_col = qty_col = data_start = None
            # SKIP aggregated summary sheets from WH16 (no weight info)
            if sname in ('Low Stock  Warnng','3 months not used','expired','approaching expiration','information','warehouse used table'):
                continue
            # SPECIAL CASE: Пересчет бамперов — улица+линия+буфер
            fname_lower = fname.lower()
            if 'бампер' in fname_lower or 'пересчет' in fname_lower:
                partial = _parse_bumper_recount(rows)
                if partial:
                    for code, qty in partial.items():
                        file_stock[code] = file_stock.get(code, 0) + qty
                    continue  # обработано
            # SPECIAL CASE: accounting_for_local_materials → последняя дата, Склад+Линия
            if sname == 'пересчет локал':
                partial = _parse_local_accounting(rows)
                for code, qty in partial.items():
                    file_stock[code] = file_stock.get(code, 0) + qty
                continue
            # SPECIAL CASE: WH16 'warehouse 16 list' → SUM(M * F) per code
            if sname == 'warehouse 16 list':
                partial = _parse_wh16_sheet(rows)
                for code, qty in partial.items():
                    file_stock[code] = file_stock.get(code, 0) + qty
                continue
            # SPECIAL CASE: AAT — простая таблица: Код / Количество / Адрес / Примечание
            if rows and rows[0] and len(rows[0]) >= 2:
                h_str = ' '.join(str(x or '').lower() for x in rows[0][:4])
                if ('код' in h_str and 'количеств' in h_str and 'инвентар' in h_str):
                    for r in rows[1:]:
                        if not r or not r[0]: continue
                        code = normalize_code(str(r[0]).strip())
                        if not re.match(r'^[A-Za-z0-9]{8,20}$', code): continue
                        try: q = float(r[1]) if r[1] is not None else 0
                        except: q = 0
                        if q > 0:
                            file_stock[code] = file_stock.get(code, 0) + q
                    continue
            # FORMAT A: ГСК — date headers + "склад"/"китайские" subheader
            has_sklad = any(i < len(h1) and h1[i] and 'склад' in str(h1[i]).lower()
                            for i in range(1, min(len(h1), 10)))
            if has_sklad:
                date_sklad = [i for i in range(1, len(h0))
                              if i < len(h1) and h1[i] and 'склад' in str(h1[i]).lower()]
                target_col = next((c for c in reversed(date_sklad)
                    if any(rows[r][c] is not None for r in range(2, min(8, len(rows)))
                           if c < len(rows[r]))), date_sklad[-1] if date_sklad else None)
                if target_col is None: continue
                code_col, qty_col, data_start = 0, target_col, 2
            # FORMAT B: ААТ — date headers (datetime or text), last col with data = current stock
            elif any(isinstance(h0[i], (_dt.datetime, _dt.date)) or
                     (isinstance(h0[i], str) and re.search(r'\d{1,2}[.,/]\d{2}[.,/]\d{4}|\d{4}-\d{2}-\d{2}|\d+[а-яА-Яa-zA-Z]+\s*\d{4}', str(h0[i])))
                     for i in range(1, min(len(h0), 5)) if h0[i] is not None):
                date_cols = [i for i, v in enumerate(h0)
                             if i > 0 and v is not None and (
                                 isinstance(v, (_dt.datetime, _dt.date)) or
                                 (isinstance(v, str) and re.search(r'\d', v)))]
                target_col = next((c for c in reversed(date_cols)
                    if any(rows[r][c] is not None for r in range(1, min(6, len(rows)))
                           if c < len(rows[r]))), date_cols[-1] if date_cols else None)
                if target_col is None: continue
                code_col, qty_col, data_start = 0, target_col, 1
            else:
                # FORMAT C: Named columns — search headers for code+qty columns
                for hi, hrow in enumerate(rows[:5]):
                    for j, v in enumerate(hrow):
                        vs = str(v).lower().replace('\n',' ') if v else ''
                        if ('коддетали' in vs.replace(' ','') or 'партномер' in vs
                                or 'компонент' in vs) and code_col is None:
                            code_col = j
                        if ('остаток' in vs and 'выдач' in vs) and qty_col is None:
                            qty_col = j
                        if 'всего' in vs or '合计' in str(v or ''):
                            if code_col == 0: qty_col = j
                    if code_col is not None and qty_col is not None:
                        data_start = hi + 1; break
                if code_col is None or qty_col is None: continue
            # Read data rows
            for row in rows[data_start:]:
                if code_col >= len(row): continue
                code = normalize_code(str(row[code_col]).strip()) if row[code_col] else ''
                if not code or code in ('nan','None',''): continue
                if not re.match(r'^[A-Za-z0-9]{8,20}$', code): continue
                if qty_col >= len(row): continue
                try:
                    q = float(row[qty_col]) if row[qty_col] is not None else 0
                    if q >= 0: file_stock[code] = file_stock.get(code, 0) + q
                except: pass
        wb.close()
        print(f"  ✅ {fname} [{fdate}]: {len(file_stock)} деталей с остатком")
        for code, qty in file_stock.items():
            stock[code] = stock.get(code, 0) + qty
            sources[code] = sources.get(code, set()); sources[code].add(fname[:25])
            # Запоминаем дату: берём более свежую если код встречается в нескольких файлах
            if code not in dates or fdate > dates[code]:
                dates[code] = fdate
    return stock, sources, dates

stock, stock_src, stock_date_map = load_all_stock(STOCK_FILES)
stock_sources = [(list(s)[0] if s else '?', qty) for code, (qty, s)
                 in {c: (stock.get(c,0), stock_src.get(c,set())) for c in stock}.items()]
print(f"  Итого: {len(stock)} деталей с ненулевым остатком из {len(STOCK_FILES)} файла(ов)")
if not stock:
    print("  ℹ️  Добавьте файлы остатков в список STOCK_FILES в начале скрипта")

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
                            _theor_filled += 1
                    except: pass

                # Остатки цеха — прибавляются всегда; дата обновляется если теор.файл свежее
                if _col_workshop is not None and _col_workshop < len(_tr):
                    try:
                        _wval = float(_tr[_col_workshop]) if _tr[_col_workshop] is not None else 0.0
                        if _wval > 0:
                            stock[_code] = stock.get(_code, 0) + _wval
                            stock_src.setdefault(_code, set()).add(os.path.basename(THEOR_STOCK_FILE)[:25] + ' [цех]')
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
    if value is None or value == '':
        return None
    if isinstance(value, datetime.datetime):
        return value.date()
    if isinstance(value, datetime.date):
        return value
    if isinstance(value, (int, float)) and value > 0:
        try:
            return from_excel(value).date()
        except Exception:
            return None
    s = str(value).strip()
    if not s:
        return None
    for fmt in ('%d.%m.%Y', '%d.%m.%y', '%Y-%m-%d', '%d/%m/%Y', '%d/%m/%y', '%d-%m-%Y', '%d-%m-%y'):
        try:
            return datetime.datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    return None

def get_stock_date(code):
    """Дата актуальности остатков для данного кода.
    Если код есть в stock_date_map — берём её, иначе глобальная _stock_date."""
    return stock_date_map.get(code, _stock_date)

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

def load_bom_from_live_output(path):
    """
    Читает лист BOM_Детальный из MRP_System_v9.xlsx после ручного редактирования.
    Возвращает (bom, bom_sections, bom_ordered_codes, live_packages) или None.
    """
    wb = load_workbook(path, read_only=True, data_only=True)
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

BOM_SOURCE = None
BOM_LIVE_ACTIVE = False
live_packages = {}
bom = {}
bom_sections = {}
bom_ordered_codes = []
BOM_COL_OFFSET = 0

if os.path.exists(OUT):
    try:
        _live = load_bom_from_live_output(OUT)
        if _live:
            bom, bom_sections, bom_ordered_codes, live_packages = _live
            BOM_SOURCE = OUT
            BOM_LIVE_ACTIVE = True
            _master_for_mq = BOM_NEW if os.path.exists(BOM_NEW) else BOM_FILE
            _merge_bom_summary_model_qty(bom, _master_for_mq)
            print(f"\nLoading BOM from MRP output (LIVE MODE — BOM_Детальный из {os.path.basename(OUT)})")
            print(f"  LIVE: {len(bom)} деталей | упаковок в BOM: {len(live_packages)} | разделов: {len(bom_sections)}")
    except Exception as _live_e:
        print(f"  ⚠️  Ошибка чтения LIVE BOM_Детальный: {_live_e}")

if not BOM_LIVE_ACTIVE:
    BOM_SOURCE = BOM_NEW if os.path.exists(BOM_NEW) else BOM_FILE
    print(f"\nLoading Master BOM from: {os.path.basename(BOM_SOURCE)}")

    wb_bom = load_workbook(BOM_SOURCE, read_only=True, data_only=True)
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
    wb_sec = load_workbook(BOM_NEW if os.path.exists(BOM_NEW) else BOM_FILE, read_only=True)
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

bom_mode = 'LIVE (BOM_Детальный из MRP)' if BOM_LIVE_ACTIVE else f'Источник: {os.path.basename(BOM_SOURCE)}'
print(f"  {len(bom)} деталей в BOM | {len(bom_sections)} разделов | Режим: {bom_mode}")

# ═══ STEP 3: Packaging ════════════════════════════════════════
print("Loading packaging...")
# Приоритет источников упаковки:
# 1) LIVE: Заказы_May col E из существующего MRP_System_v8.xlsx
# 2) Fallback: Haval_Stock в Упаковка_локала.xlsx
# Ручные остатки: Ввод_Остатков col E (Остаток) и col F (Дата остатка)

manual_stock_override = {}  # code -> qty (ручной ввод из Ввод_Остатков)
manual_stock_date_override = {}  # code -> date (ручная дата актуальности остатка)
manual_deliveries     = {}  # не используется, сохраняется для совместимости simulate_deliveries

if os.path.exists(OUT):
    try:
        wb_live = load_workbook(OUT, read_only=True, data_only=True)
        # A. Упаковки из листа 'Упаковка' (LIVE, приоритет #1)
        pkg_loaded = 0
        if 'Упаковка' in wb_live.sheetnames:
            for row in wb_live['Упаковка'].iter_rows(min_row=3, values_only=True):
                code = str(row[0]).strip() if row[0] else ''
                pkg_v = row[2] if len(row) > 2 else None
                if code and pkg_v is not None:
                    try:
                        p = int(float(pkg_v))
                        if p >= 1 and code in bom:
                            bom[code]['package'] = p; pkg_loaded += 1
                    except: pass
            print(f"  Упаковок из листа 'Упаковка' (LIVE): {pkg_loaded}")

        # Упаковки из Заказы_May col E (index 4)
        pkg_loaded = 0
        if f"Заказы_{MONTH_SHORT[STOCK_AS_OF_MONTH]}" in wb_live.sheetnames:
            for row in wb_live[f"Заказы_{MONTH_SHORT[STOCK_AS_OF_MONTH]}"].iter_rows(min_row=3, values_only=True):
                code = str(row[0]).strip() if row[0] else ''
                if not code or code in ('nan','Код'): continue
                try:
                    p = float(row[4]) if row[4] else 1
                    if p >= 1 and code in bom:
                        bom[code]['package'] = p
                        pkg_loaded += 1
                except: pass
        print(f"  Упаковок из MRP (LIVE): {pkg_loaded}")
        # B. Ручные остатки из Ввод_Остатков — col E (индекс 4), дата — col F (индекс 5)
        manual_count = 0
        manual_date_count = 0
        if 'Ввод_Остатков' in wb_live.sheetnames:
            for row in wb_live['Ввод_Остатков'].iter_rows(min_row=3, values_only=True):
                code = str(row[0]).strip() if row[0] else ''
                if not code or code == 'nan': continue
                val = row[4] if len(row) > 4 else None  # col E = Остаток
                if val is not None:
                    try:
                        q = float(val)
                        if q > 0:
                            manual_stock_override[code] = q
                            manual_count += 1
                    except: pass
                date_val = row[5] if len(row) > 5 else None  # col F = Дата остатка
                manual_date = _parse_manual_stock_date(date_val)
                if manual_date is not None:
                    manual_stock_date_override[code] = manual_date
                    manual_date_count += 1
        print(f"  Ручных остатков из Ввод_Остатков: {manual_count}")
        print(f"  Ручных дат остатков из Ввод_Остатков: {manual_date_count}")
        wb_live.close()
    except Exception as e:
        print(f"  ⚠️  Ошибка чтения LIVE файла: {e}")

# Применяем ручные остатки ПОВЕРХ файловых данных
for code, qty in manual_stock_override.items():
    stock[code] = qty  # ручной ввод заменяет данные из файлов остатков
for code, stock_dt in manual_stock_date_override.items():
    stock_date_map[code] = stock_dt  # ручная дата меняет горизонт спроса/поставок

# Fallback упаковки из PKG_FILE (Haval_Stock)
wb_pkg=load_workbook(PKG_FILE,read_only=True)
for row in wb_pkg['Haval_Stock'].iter_rows(min_row=2,values_only=True):
    code=str(row[0]).strip() if row[0] else ''
    if not code or code.lower()=='nan': continue
    try: pkg=max(1,int(row[3])) if row[3] and row[3]!=0 else 1
    except: pkg=1
    if code not in bom:
        bom[code]={'name':str(row[1]).strip() if row[1] else '','unit':str(row[4]).strip() if row[4] else 'шт',
                   'supplier':str(row[2]).strip() if row[2] else '','model_qty':{},'configs':set(),'section':'Упаковка','package':pkg}
    elif bom[code].get('package',1) == 1:  # только если LIVE не дал значение
        bom[code]['package']=pkg
wb_pkg.close()
for code in bom:
    if 'package' not in bom[code]: bom[code]['package']=1

# ── Применяем ручные переопределения упаковки (PACKAGE_OVERRIDES) ──
# В LIVE-режиме не перезаписываем упаковку, заданную в BOM_Детальный col «Упак.»
pkg_override_count = 0
for code, pkg_val in PACKAGE_OVERRIDES.items():
    if BOM_LIVE_ACTIVE and code in live_packages:
        continue
    if code in bom:
        bom[code]['package'] = pkg_val
        pkg_override_count += 1
    else:
        bom[code] = {'name': code, 'unit': 'pcs', 'supplier': '',
                     'model_qty': {}, 'configs': {}, 'section': 'Упаковка',
                     'package': pkg_val, 'notes': ''}
        pkg_override_count += 1
print(f"  Переопределено упаковок (PACKAGE_OVERRIDES): {pkg_override_count}"
      + (" | LIVE: сохранены упаковки из BOM_Детальный" if BOM_LIVE_ACTIVE else ""))

# ── Применяем переопределения применяемости BOM (BOM_APPLICABILITY_OVERRIDES) ──
# В LIVE-режиме применяемость берётся только из BOM_Детальный (ручные ✓)
bom_appl_count = 0
if not BOM_LIVE_ACTIVE:
    for code, allowed_cfgs in BOM_APPLICABILITY_OVERRIDES.items():
        if code in bom:
            bom[code]['configs'] = {k: 1 for k in allowed_cfgs}
            bom_appl_count += 1
            print(f"  BOM применяемость {code}: {sorted(allowed_cfgs)}")
else:
    print("  BOM применяемость: LIVE — используются правки из BOM_Детальный")
print(f"  Переопределено применяемостей BOM: {bom_appl_count}")

# B02_2WD_elite: дополняем применяемость по правилам «все B02 / 2WD / Elite»
_b02_elite_added = 0
for _code, _info in bom.items():
    _before = 'B02_2WD_elite' in _info.get('configs', {})
    _info['configs'] = apply_b02_2wd_elite_marks(_info.get('configs', {}), _code)
    if not _before and 'B02_2WD_elite' in _info.get('configs', {}):
        _b02_elite_added += 1
if _b02_elite_added:
    print(f"  B02_2WD_elite: добавлена применяемость для {_b02_elite_added} деталей")


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

# ── Default упаковка для бамперов = 8 шт (если в BOM пусто или 1) ──
DEFAULT_BUMPER_PKG = 8
for code in list(bom.keys()):
    info = bom[code]
    is_bumper = code in bumper_meta or (code.startswith('2803') or code.startswith('2804'))
    if is_bumper and (not info.get('package') or info.get('package') == 1):
        info['package'] = DEFAULT_BUMPER_PKG

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

    # drive: колонка между model и config (применяем смещение)
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

plan_batches={}; batch_info={}
for tab,cols in TAB_COLS.items():
    try: df=pd.read_excel(PF_FILE,sheet_name=tab,header=None)
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
            bv=str(row.iloc[_bc]).strip()
            if not bv.startswith('R'): continue
            try:
                mr=MODEL_MAP.get(str(row.iloc[_mc]).strip() if _mc<len(row) and pd.notna(row.iloc[_mc]) else '','')
                dr=DRIVE_MAP.get(str(row.iloc[_dc]).strip() if _dc<len(row) and pd.notna(row.iloc[_dc]) else '','')
                cr=CONFIG_MAP.get(str(row.iloc[_cc]).strip() if _cc<len(row) and pd.notna(row.iloc[_cc]) else '','')
                if bv not in batch_info:
                    batch_info[bv]={'model':mr,'drive':dr,'config':cr,'total':0}
            except:
                if bv not in batch_info:
                    batch_info[bv]={'model':'','drive':'','config':'','total':0}
            for d in range(1,n_days+1):
                col_idx=_d1+d-1
                if col_idx<len(row) and pd.notna(row.iloc[col_idx]):
                    try:
                        q=float(row.iloc[col_idx])
                        if q>0:
                            month_data.setdefault(bv,{})[d]=month_data.get(bv,{}).get(d,0)+q
                            batch_info[bv]['total']+=q
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
if os.path.exists(PAINT_STATS):
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
            batch = str(batch).strip()
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
else:
    print(f"  ⚠️  Файл не найден: {PAINT_STATS}")

# ═══ STEP 6: Demand (V7 logic) ════════════════════════════════
def cars_by_color(month_num, color_key, models=None, configs=None, applicable=None):
    """Возвращает {day: cars} — кузова данного цвета в данном месяце.
    Дедуплицирует батчи по R-номеру (один батч может быть в нескольких tabs).
    applicable — dict BOM-применяемости (фильтр по конфигурации партии).
    """
    out = {}
    seen = set()
    for tab, months_data in plan_batches.items():
        if month_num not in months_data: continue
        for bv, day_qty in months_data[month_num].items():
            if bv in seen: continue
            seen.add(bv)
            bi = batch_info.get(bv, {})
            b_model = bi.get('model','')
            b_config = bi.get('config','')
            if models and b_model not in models: continue
            if configs and b_config not in configs: continue
            if applicable:
                cfg_key = f"{b_model}_{bi.get('drive','')}_{b_config}"
                matched, _ = cfg_match(cfg_key, applicable)
                if not matched:
                    continue
            bc_split = batch_color.get(bv, {})
            if not bc_split: continue
            total_in_batch = sum(bc_split.values())
            if total_in_batch <= 0: continue
            color_cars = bc_split.get(color_key, 0)
            if color_cars <= 0: continue
            share = color_cars / total_in_batch
            for day, cars in day_qty.items():
                out[day] = out.get(day, 0) + cars * share
    return out


def _primer_demand_for_paints(primer_code, month_num, paint_codes):
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
                    month_num, color_key, models={model_key}, applicable=applicable)
                for day, cars in cars_d.items():
                    daily[day] = daily.get(day, 0) + cars * norm
    return daily


def get_primer_demand(primer_code, month_num):
    """YIERTE-грунт: все связанные краски минус перешедшие на Litum. Litum-грунт: одна краска."""
    sp_litum = _primer_split_for_litum(primer_code)
    if sp_litum:
        if month_num < sp_litum['from_month']:
            return {}
        return _primer_demand_for_paints(
            primer_code, month_num, [sp_litum['paint_new']])
    if primer_code not in PRIMER_LINKED_PAINTS:
        return {}
    linked = [p for p in PRIMER_LINKED_PAINTS[primer_code]
              if p not in _primer_excluded_paints(primer_code, month_num)]
    return _primer_demand_for_paints(primer_code, month_num, linked)

def cars_filter(month_num, models=None, configs=None, predicate=None):
    """Возвращает {day: cars} с произвольным фильтром по партиям."""
    out = {}
    for tab, months_data in plan_batches.items():
        if month_num not in months_data: continue
        for bv, day_qty in months_data[month_num].items():
            bi = batch_info.get(bv, {})
            if models and bi.get('model','') not in models: continue
            if configs and bi.get('config','') not in configs: continue
            if predicate and not predicate(bv, bi): continue
            for day, cars in day_qty.items():
                out[day] = out.get(day, 0) + cars
    return out

# ── Спец-правило для задних бамперов B02 XST33 (только дорестайл) ──
# Дорестайл-партии — это B02 2WD premium (CC6480AL00C); рестайл — B02 4WD elite/Tech+
PRERESTYLE_B02_CONFIGS = {'premium'}   # пока: только premium 2WD = pre-restyle
PRERESTYLE_B02_DRIVES  = {'2WD'}

# (HEADREST_X2_CODES, OBSOLETE_CODES, CFG_FUZZY_ALT, cfg_match определены в начале файла)


def get_daily_demand_raw_bumper(code, month_num, b_models, tabs):
    """Helper: compute bumper demand without V7 scaling (for ratio calculation)"""
    d={}
    for tab in tabs:
        if tab not in plan_batches or month_num not in plan_batches[tab]: continue
        for bv,day_qty in plan_batches[tab][month_num].items():
            if b_models and batch_info.get(bv,{}).get('model','') not in b_models: continue
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
            return get_primer_demand(code, month_num)
        norms=chem_norms[code]
        color_filter=paint_colors.get(code,False)
        # ─ Цветная краска: norm × кузова нужного(их) цвета(ов) ─
        if isinstance(color_filter, list) and color_filter:
            for color_key in color_filter:
                # для каждого цвета — соберём кузова по моделям с подходящей нормой
                for model_key, norm in norms.items():
                    if norm <= 0: continue
                    cars_d = cars_by_color(month_num, color_key, models={model_key})
                    for day, cars in cars_d.items():
                        daily[day] = daily.get(day, 0) + cars * norm
            return daily
        # ─ Краска без цветового фильтра: norm × все кузова ─
        for tab in tabs:
            if tab not in plan_batches or month_num not in plan_batches[tab]: continue
            for bv,day_qty in plan_batches[tab][month_num].items():
                bi=batch_info.get(bv,{})
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
            cars_d = cars_by_color(month_num, color_key)
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
        applicable_b = bom.get(code, {}).get('configs', {})
        if not isinstance(applicable_b, dict):
            applicable_b = {k: 1 for k in applicable_b}
        if code in bom and not applicable_b:
            return {}
        # Эвристика только для бамперов без строки в BOM
        if not applicable_b and code not in bom:
            if any(p in code for p in ('XKN61', 'KN260004', 'KN260005')):
                applicable_b = {'B02_4WD_elite': 1, 'B02_4WD_TechPlus': 1}
            elif any(p in code for p in ('XST33', 'AST33', 'AKN02')):
                applicable_b = {'B02_2WD_premium': 1}
        if not applicable_b:
            return {}
        # ── Дедуплицируем партии: одна и та же партия R-номер встречается в ≥1 tab ──
        seen_batches = set()
        unique_batches = []   # list of (bv, day_qty)
        for tab, months_data in plan_batches.items():
            if month_num not in months_data: continue
            for bv, day_qty in months_data[month_num].items():
                if bv in seen_batches: continue
                seen_batches.add(bv)
                unique_batches.append((bv, day_qty))
        # Если есть цветовая раскраска — split по цветам
        if color_key and batch_color:
            for bv, day_qty in unique_batches:
                bi = batch_info.get(bv, {})
                cfg_key = f"{bi.get('model','')}_{bi.get('drive','')}_{bi.get('config','')}"
                matched, _ = cfg_match(cfg_key, applicable_b)
                if not matched: continue
                bc_split = batch_color.get(bv, {})
                if not bc_split: continue
                total_in_batch = sum(bc_split.values())
                if total_in_batch <= 0: continue
                color_cars = bc_split.get(color_key, 0)
                if color_cars <= 0: continue
                share = color_cars / total_in_batch
                for day, cars in day_qty.items():
                    daily[day] = daily.get(day, 0) + cars * share
            # Бампера — целые единицы (округляем итог по дням)
            return {d: round(v) for d, v in daily.items()}
        # fallback: все кузова из applicable_b (без цветовой раскраски)
        for bv, day_qty in unique_batches:
            bi = batch_info.get(bv, {})
            cfg_key = f"{bi.get('model','')}_{bi.get('drive','')}_{bi.get('config','')}"
            matched, qty = cfg_match(cfg_key, applicable_b)
            if not matched: continue
            for day, cars in day_qty.items():
                daily[day] = daily.get(day, 0) + cars * qty
        return {d: round(v) for d, v in daily.items()}

    # 3. BOM general
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
            bi=batch_info.get(bv,{})
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
            for day,cars in day_qty.items():
                daily[day]=daily.get(day,0)+cars*qty_per_car
    return daily

# ═══ STEP 7: Pre-compute demand ═══════════════════════════════
print("\nPre-computing demand...")
demand={}
for code in mrp_codes:
    demand[code]={}
    for mnum,_,_ in MONTHS:
        demand[code][mnum]=get_daily_demand(code,mnum)

with_demand=sum(1 for c in mrp_codes if any(demand[c].get(m) for m,_,_ in MONTHS))
print(f"  Parts with demand > 0: {with_demand} / {len(mrp_codes)}")
for code in ['6803112XKN08A','ALAA005669','1101100AGW01A','2803104XKN61A8T']:
    t=sum(demand.get(code,{}).get(MONTHS[0][0],{}).values())
    print(f"  {code}: May={t:.1f}  tab={part_tab_map.get(code,'?')}")

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

def calc_order(code, month_num, stock_override=None):
    """Расчёт заказа на месяц.
    Использует дату актуальности остатков из stock_date_map (per-code).
    Для месяца в котором лежит дата остатков — учитываем только будущие дни.
    Для предыдущих месяцев — потребность = 0 (уже прошло).
    Для будущих месяцев — полная потребность.
    """
    _sdate = get_stock_date(code)   # дата остатков для этого кода
    _as_month = _sdate.month
    _as_day   = _sdate.day

    d_full = sum(demand[code].get(month_num, {}).values())
    n_days = next(nd for mn, _, nd in MONTHS if mn == month_num)
    # Определяем потребность с учётом даты остатков
    if month_num == _as_month:
        # Месяц остатков: берём только дни ПОСЛЕ даты остатков
        d_future = sum(q for d, q in demand[code].get(month_num, {}).items()
                       if d > _as_day)
    elif month_num < _as_month:
        # Этот месяц уже прошёл относительно даты остатков — ничего не заказываем
        d_future = 0.0
    else:
        d_future = d_full
    # Страховой запас = SAFETY_DAYS × среднесуточная потребность за 3 месяца
    d_total_3m = sum(sum(demand[code].get(mn, {}).values()) for mn, _, _ in MONTHS)
    n_days_total = sum(nd for _, _, nd in MONTHS)
    avg_daily = d_total_3m / n_days_total if n_days_total else 0
    safety = avg_daily * SAFETY_DAYS
    stk = stock_override if stock_override is not None else stock.get(code, 0)
    net = max(0, d_future + safety - stk)
    pkg = max(1, bom.get(code, {}).get('package', 1))
    pkgs = math.ceil(net / pkg) if net > 0 else 0
    return {'demand': round(d_full, 2), 'stock': stk, 'safety': round(safety, 2),
            'net': round(net, 2), 'pkg_size': pkg, 'packages': pkgs,
            'order_qty': round(pkgs * pkg, 2)}

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
    s = stock.get(code, 0)
    _sdate = get_stock_date(code)
    _as_month = _sdate.month; _as_day = _sdate.day
    ps = {}
    for i, (mn, _, _) in enumerate(MONTHS):
        ps[mn] = max(0.0, s)
        # Считаем прогнозный остаток на начало следующего месяца
        if i == 0:
            # Первый месяц: учитываем только остаток потребности с даты остатков кода
            remaining = sum(q for d, q in demand[code].get(mn, {}).items() if d > _as_day)
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

def _graph_stock_formula(ri, prev_col_l, prev_col_n, pkg, fallback):
    """Переходящий остаток из График_Поставок, округление вверх до упаковки."""
    pkg = max(1, int(pkg) if pkg else 1)
    fb = int(fallback) if fallback else 0
    return (
        f"=IFERROR(CEILING(VLOOKUP(A{ri},График_Поставок!$A:${prev_col_l},"
        f"{prev_col_n},0)/{pkg},1)*{pkg},{fb})"
    )

def _delivery_series_demand(code):
    _code_sdate = get_stock_date(code)
    return [
        round(demand[code].get(mn, {}).get(d, 0), 4) if dt > _code_sdate else 0.0
        for (dt, mn, d) in all_dates
    ]

def _delivery_safety_qty(code, supplier):
    sd = get_safety_days(code, supplier)
    total_3m = sum(sum(demand[code].get(mn, {}).values()) for mn, _, _ in MONTHS)
    n_days_3m = sum(nd for _, _, nd in MONTHS)
    avg_daily = total_3m / n_days_3m if n_days_3m else 0
    return round(avg_daily * sd, 4)

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
        if mn == STOCK_AS_OF_MONTH:
            order_sent = max(order_sent, datetime.date(y, mn, STOCK_AS_OF_DAY))
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

def _build_initial_deliveries(code, supplier):
    dems = _delivery_series_demand(code)
    pkg = max(1, bom.get(code, {}).get('package', 1))
    safety = _delivery_safety_qty(code, supplier)
    mode = _delivery_mode(supplier)
    n = len(dems)
    dels = [0.0] * n

    if mode == 'monthly_early':
        del_days = _monthly_early_delivery_indices()
    elif mode == 'monthly_lag':
        del_days = _ecotexis_delivery_indices()
    elif mode == 'smc_weekly':
        del_days = _weekday_delivery_indices(SMC_WEEKDAY)
    elif mode == 'weekly':
        del_days = _weekday_delivery_indices(PUREM_WEEKDAY)
    else:
        del_days = None

    if del_days is not None:
        del_set = set(del_days)
        next_map = {
            del_days[i]: (del_days[i + 1] if i + 1 < len(del_days) else n)
            for i in range(len(del_days))
        }
        ss_prev = float(stock.get(code, 0))
        for di in range(n):
            dem_d = dems[di]
            del_d = 0.0
            if di in del_set:
                del_d = _period_delivery_qty(
                    ss_prev, dems, di, next_map[di], safety, pkg)
            dels[di] = del_d
            ss_prev = ss_prev - dem_d + del_d
        return dels

    ss_prev = float(stock.get(code, 0))
    for di, (dt, _, _) in enumerate(all_dates):
        dem_d = dems[di]
        del_d = 0.0
        if dt.weekday() != 6:
            la = _lookahead_demand(dems, di)
            if la > 0 and (ss_prev - dem_d - la) < safety:
                need = safety - (ss_prev - dem_d - la)
                del_d = _ceiling_pkg_qty(need, pkg) if need > 0 else pkg
        dels[di] = del_d
        ss_prev = ss_prev - dem_d + del_d
    return dels

def _recompute_ss_from_deliveries(code, dels):
    dems = _delivery_series_demand(code)
    ss_prev = float(stock.get(code, 0))
    pairs = []
    for di, dem_d in enumerate(dems):
        del_d = dels[di]
        ss_d = ss_prev - dem_d + del_d
        pairs.append((del_d, ss_d))
        ss_prev = ss_d
    return pairs

def build_all_delivery_schedules():
    all_dels = {}
    for code in mrp_codes:
        _, supp, _ = get_info(code)
        all_dels[code] = _build_initial_deliveries(code, supp)

    smc_codes = [c for c in mrp_codes if _delivery_mode(get_info(c)[1]) == 'smc_weekly']
    for di in _weekday_delivery_indices(SMC_WEEKDAY):
        raw = {c: all_dels[c][di] for c in smc_codes if all_dels[c][di] > 0}
        if not raw:
            continue
        capped = _cap_smc_day_deliveries(raw)
        for c, q in capped.items():
            all_dels[c][di] = q

    schedules = {}
    for code in mrp_codes:
        pairs = _recompute_ss_from_deliveries(code, all_dels[code])
        schedules[code] = {
            'dels': [p[0] for p in pairs],
            'ss': [p[1] for p in pairs],
            'pairs': {di: pairs[di] for di in range(len(pairs))},
        }
    return schedules

print("  Расчёт графиков поставок (Ecoal'yance / Ecotexis / SMC / Purem)...")
DELIVERY_SCHEDULES = build_all_delivery_schedules()
_n_smc = sum(1 for c in mrp_codes if _delivery_mode(get_info(c)[1]) == 'smc_weekly')
print(f"    Расписаний: {len(DELIVERY_SCHEDULES)} | SMC позиций: {_n_smc} | лимит фуры: {SMC_MAX_PALLETS_PER_TRUCK} палл.")

# ── Упаковка (редактируемый) ────────────────────────────────
print("  Упаковка (редактируемый)...")
ws_pkg = wb_out.create_sheet('Упаковка')
ws_pkg.freeze_panes = 'C3'

# Header
ws_pkg.row_dimensions[1].height = 45
t_pkg = ws_pkg.cell(1, 1)
t_pkg.value = ('УПАКОВКА — КРАТНОСТЬ ЗАКАЗА | '
               'Измените значение в колонке C → перезапустите mrp_v10.py → заказ пересчитается')
t_pkg.font = Font(bold=True, size=11, color="FFFFFF", name="Arial")
t_pkg.fill = fill("1F3864")
t_pkg.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
ws_pkg.merge_cells('A1:F1')

pkg_col_w = [20, 42, 14, 8, 20, 30]
for ci, w in enumerate(pkg_col_w, 1):
    ws_pkg.column_dimensions[get_column_letter(ci)].width = w
ws_pkg.row_dimensions[2].height = 30

for ci, h in enumerate(['Код', 'Наименование', 'Уп. (шт) ← РЕДАКТИРОВАТЬ', 'Ед.', 'Поставщик', 'Примечание'], 1):
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
    c = ws_pkg.cell(pkg_row, 2); c.value = name[:60]
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

    # Col F: note
    c = ws_pkg.cell(pkg_row, 6); c.value = ''
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
ws_man=wb_out.create_sheet('Ввод_Остатков')
ws_man.freeze_panes='A4'; ws_man.row_dimensions[1].height=40; ws_man.row_dimensions[2].height=14
for ci,(h,w) in enumerate(zip(['Код детали','Наименование','Поставщик','Ед.','Остаток\n✏️ ручной ввод','Дата остатка\n✏️ дд.мм.гггг'],[22,44,16,6,18,14]),1):
    hcell(ws_man,1,ci,h,H_FILL if ci not in (5,6) else fill("B8860B")); ws_man.column_dimensions[get_column_letter(ci)].width=w
for _ci in (5, 6):
    _he = ws_man.cell(1,_ci)
    _he.font=Font(bold=True,color="FFFFFF",size=9,name="Arial")
    _he.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)
if stock_sources:
    src_names_display = sorted(set(n for s in stock_src.values() for n in s))
    msg=(f"✅ ОСТАТКИ ЗАГРУЖЕНЫ: {len(stock)} деталей из {len(src_names_display)} файла(ов): {', '.join(n[:20] for n in src_names_display)}"
         f"  |  ✏️ Измените Остаток (col E) и Дату остатка (col F) → при следующем запуске пересчитаются: График_Поставок, Потребность и Заказы всех месяцев.")
    fc="1F6B00"; bg=fill("E2EFDA")
else:
    msg="ОСТАТКИ НЕ ЗАГРУЖЕНЫ. Загрузите xlsx/csv с двумя колонками: [Код детали, Кол-во]  |  ✏️ Колонка E — ручной ввод остатков, F — дата остатка"
    fc="C00000"; bg=fill("FFD7D7")
c2=ws_man.cell(2,1); c2.value=msg; c2.font=Font(bold=True,size=9,color=fc,name="Arial")
c2.fill=bg; c2.alignment=Alignment(horizontal='left',vertical='center')
ws_man.merge_cells('A2:F2')
for ri,code in enumerate(all_codes,3):
    ws_man.row_dimensions[ri].height=14
    fb=GRY_F if ri%2==0 else NO_F
    name,supp,unit=get_info(code); stk=stock.get(code,0)
    stock_dt = get_stock_date(code)
    for ci,v in enumerate([code,name[:50],supp,unit,stk,stock_dt],1):
        c=ws_man.cell(ri,ci); c.value=v
        c.font=Font(size=9,name="Arial")
        c.alignment=Alignment(horizontal='left' if ci<=2 else 'center',vertical='center')
        if ci==5: c.fill=YEL_F; c.number_format='#,##0.##'
        elif ci==6: c.fill=YEL_F; c.number_format='DD.MM.YYYY'
        elif fb: c.fill=fb

# ── Precompute: Ss-колонка последнего дня мая в График_Поставок ─────────────
# График_Поставок: col A=Код, для дня di:
#   Del-col = 8 + di*2       (чётный offset)
#   Ss-col  = 8 + di*2 + 1   (нечётный offset) ← именно её используем для VLOOKUP
# Данные начинаются с row 4 (_gp_data_start=4).
_gp_last_mon_col = {}  # mnum -> (col_letter, col_num) Ss-колонки последнего дня месяца
for _di, (_dt, _mn, _d) in enumerate(all_dates):
    _cn = 8 + _di * 2 + 1   # Ss-колонка (нечётный offset)
    _gp_last_mon_col[_mn] = (get_column_letter(_cn), _cn)
# После цикла: _gp_last_mon_col[5] = Ss последнего дня мая (May 31)
# Для июля: используем Потребность_Jun col I (Дефицит = остаток после июньского спроса)

# ── Потребность_May/Jun/Jul ───────────────────────────────────
FH=['Код детали','Наименование','Поставщик','Ед.','Вкладка плана',
    'Тип / норма','Остаток','Потребность','Дефицит']
FW=[22,42,16,6,22,30,10,14,12]
for mi,(mnum,mlabel,n_days) in enumerate(MONTHS):
    sname=f"Потребность_{mlabel[:3]}"
    print(f"  {sname}...")
    ws=wb_out.create_sheet(sname)
    ws.freeze_panes='J3'; ws.row_dimensions[1].height=45; ws.row_dimensions[2].height=14
    for ci,(h,w) in enumerate(zip(FH,FW),1):
        hcell(ws,1,ci,h,H_FILL); ws.column_dimensions[get_column_letter(ci)].width=w
    for d in range(1,n_days+1):
        ci=len(FH)+d; hcell(ws,1,ci,str(d),M_FILL[mnum])
        ws.column_dimensions[get_column_letter(ci)].width=6.5
    ws.cell(2,1).value="Только детали с ✓ в BOM_Детальный (+ химия/Additional). Без применяемости — не в расчёте."
    ws.cell(2,1).font=Font(italic=True,size=8,color="555555",name="Arial")
    ws.merge_cells(f'A2:{get_column_letter(len(FH)+n_days)}2')
    for ri,code in enumerate(mrp_codes,3):
        ws.row_dimensions[ri].height=14
        fb=GRY_F if ri%2==0 else NO_F
        name,supp,unit=get_info(code)
        daily=demand[code].get(mnum,{}); mtotal=sum(daily.values())
        stk=stock.get(code,0); tab=part_tab_map.get(code,'—'); norm_str=get_norm_str(code)
        vals=[code,name[:50],supp,unit,tab,norm_str,stk,
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
            _cg.value = f"=IFERROR(VLOOKUP(A{ri},Ввод_Остатков!$A:$E,5,0),{int(stk)})"
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
            # Месяц 3: остаток берётся из Потребность_M1 col I (Дефицит = остаток после спроса месяца 2)
            _cg = ws.cell(ri, 7)
            _cg.value = f"=IFERROR(VLOOKUP(A{ri},Потребность_{MONTH_SHORT[MONTHS[1][0]]}!$A:$I,9,0),{int(stk)})"
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
            fut_d = sum(q for d, q in demand[code].get(month_num, {}).items() if d > STOCK_AS_OF_DAY)
            return round(fut_d / first_norm, 1)
        return round(total_d / first_norm, 1)
    plan_tab = part_tab_map.get(code, '')
    tabs = resolve_plan_tabs(plan_tab)
    total = 0
    for tab in tabs:
        if tab not in plan_batches or month_num not in plan_batches[tab]: continue
        for bv, day_qty in plan_batches[tab][month_num].items():
            bi = batch_info.get(bv, {})
            if bi.get('model', '') not in applicable_models: continue
            for day, cars in day_qty.items():
                if fut_only and month_num == STOCK_AS_OF_MONTH and day <= STOCK_AS_OF_DAY: continue
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
        if _mn == STOCK_AS_OF_MONTH:
            _row['cars_may_fut'] = _chem_cars(_code, _mn, fut_only=True)
    norms_rows[_code] = _row

_m0, _m0lbl, _ = MONTHS[0]; _m1, _m1lbl, _ = MONTHS[1]; _m2, _m2lbl, _ = MONTHS[2]
_ms0 = MONTH_SHORT[_m0]; _ms1 = MONTH_SHORT[_m1]; _ms2 = MONTH_SHORT[_m2]
NRM_HDR = ['Код','Наименование','Ед.','Поставщик',
           f'Норма\n{_ms0}\n(ред.)',f'Норма\n{_ms1}\n(ред.)',f'Норма\n{_ms2}\n(ред.)',
           f'Авт_{_ms0}\n(полн)',f'Авт_{_ms0}\n(ост.)',
           f'Авт_{_ms1}',f'Авт_{_ms2}']
NRM_W   = [22, 44, 6, 18, 10, 10, 10, 10, 10, 10, 10]
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
    _vals = [_code, _name[:50], _unit, _supp,
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

# ── Заказы_May/Jun/Jul ────────────────────────────────────────
for mi_ord,(mnum,mlabel,n_days) in enumerate(MONTHS):
    oname=f"Заказы_{mlabel[:3]}"; print(f"  {oname}...")
    ws_o=wb_out.create_sheet(oname)
    ws_o.freeze_panes='J3'; ws_o.row_dimensions[1].height=50
    oh=['Код','Наименование','Поставщик','Ед.','Уп.\n(шт)','Остаток',
        f'Потребн.\n{mlabel[:3]}','Страх.\nзапас','Чистая\nпотребн.',
        'Кол-во\nупаковок','ЗАКАЗ\n(итого)','Статус']
    ow=[22,42,16,6,7,10,14,12,14,10,14,18]
    for ci,(h,w) in enumerate(zip(oh,ow),1):
        hcell(ws_o,1,ci,h,H_FILL); ws_o.column_dimensions[get_column_letter(ci)].width=w
    ws_o.cell(2,1).value=f"Заказ=CEILING((Потребность+Страх_запас-Остаток)/Упаковка)×Упаковка | Страховой запас={SAFETY_DAYS} дн."
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
        vals=[code,name[:50],supp,unit,pkg,stk,dem,saf,net,pkgs,oq,st]
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
        # ── Формулы для авто-пересчёта при ручном вводе остатка (Ввод_Остатков col E) ──
        # Для текущего месяца (May): col F = VLOOKUP, col I/J/K/L = Excel-формулы
        if mnum == STOCK_AS_OF_MONTH and code not in chem_norms:
            # d_future = потребность с STOCK_AS_OF_DAY+1 до конца месяца (используется в расчёте заказа)
            d_future = sum(q for d, q in demand[code].get(mnum, {}).items() if d > STOCK_AS_OF_DAY)
            # F: остаток из Ввод_Остатков (при ручном изменении — пересчёт мгновенный)
            _cF = ws_o.cell(ri, 6)
            _cF.value = f"=IFERROR(VLOOKUP(A{ri},Ввод_Остатков!$A:$E,5,0),0)"
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

        # ── Месяц 3 (non-chem): переходящий остаток из Потребность_M2 col I ──────
        elif mi_ord == 2 and code not in chem_norms:
            _cF = ws_o.cell(ri, 6)
            _cF.value = f"=IFERROR(VLOOKUP(A{ri},Потребность_{MONTH_SHORT[MONTHS[1][0]]}!$A:$I,9,0),{max(0,int(stk))})"
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
                _cF.value = f"=IFERROR(VLOOKUP(A{ri},Ввод_Остатков!$A:$E,5,0),0)"
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
                _cF = ws_o.cell(ri, 6)
                _cF.value = f"=IFERROR(VLOOKUP(A{ri},Потребность_{MONTH_SHORT[MONTHS[1][0]]}!$A:$I,9,0),{max(0,int(stk))})"
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

# ── График_Поставок ── структура как DeliveryMay/June: Del | Ss на каждый день ──
# Для каждого дня d — 2 колонки:
#   ci_del = 8 + di*2      → Поставка (авто-расчёт, =0 если не нужна)
#   ci_ss  = 8 + di*2 + 1  → Остаток  (= prev_Ss - Спрос_d + Поставка_d)
# Логика поставки: если (Ss_d - Спрос_{d+1}) < страховой → поставить CEILING(...)*pkg
# Это зеркало логики DeliveryMay, но автоматизированное.
print("  График_Поставок (Del|Ss по дням)...")
from openpyxl.formatting.rule import FormulaRule as _FR_gp
ws_g = wb_out.create_sheet('График_Поставок')
ws_g.freeze_panes = 'H4'
ws_g.row_dimensions[1].height = 40
ws_g.row_dimensions[2].height = 16
ws_g.row_dimensions[3].height = 20

_n_day_cols_gp = len(all_dates)           # 92 дня
_total_cols_gp = 7 + _n_day_cols_gp * 2  # 7 fix + 2*92 = 191

# Строка 1: заголовок
t = ws_g.cell(1, 1)
t.value = (f"ГРАФИК ПОСТАВОК | {PERIOD_LABEL}  |  "
           "🟢 Поставка  🔴 Дефицит  🟡 Ниже страх.запаса  "
           "Остаток G = живой (из Ввод_Остатков)  |  "
           "Ecoal'yance: 1×/мес (1–5) | Ecotexis: +35д от заказа | "
           f"Purem/SMC: 1×/нед | SMC ≤{SMC_MAX_PALLETS_PER_TRUCK} палл./фура")
t.font = Font(bold=True, size=9, color="FFFFFF", name="Arial")
t.fill = H_FILL
t.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
ws_g.merge_cells(f'A1:{get_column_letter(_total_cols_gp)}1')

# Строка 2: месячные метки (объединённые)
_GP_FIX_H = ['Код детали','Наименование','Поставщик','Ед.','Уп.','Страх.\nшт','Остаток\n(Ввод.Ост.)']
_GP_FIX_W = [22, 42, 16, 5, 5, 9, 11]
for ci, (h, w) in enumerate(zip(_GP_FIX_H, _GP_FIX_W), 1):
    hcell(ws_g, 2, ci, h, H_FILL)
    hcell(ws_g, 3, ci, h, H_FILL)
    ws_g.column_dimensions[get_column_letter(ci)].width = w
# Fixed headers stay unmerged so Excel can expose autofilter dropdowns
# for Код/Наименование/Поставщик and the other static columns.

# Месячные метки в строке 2 над днями
for mnum_m, mlabel_m, ndays_m in MONTHS:
    first_di = sum(nd for mn,_,nd in MONTHS if mn < mnum_m)
    last_di  = first_di + ndays_m - 1
    col_first = 8 + first_di * 2
    col_last  = 8 + last_di  * 2 + 1
    mc = ws_g.cell(2, col_first)
    mc.value = mlabel_m
    mc.font  = Font(bold=True, size=10, color="FFFFFF", name="Arial")
    mc.fill  = M_FILL[mnum_m]
    mc.alignment = Alignment(horizontal='center', vertical='center')
    ws_g.merge_cells(f'{get_column_letter(col_first)}2:{get_column_letter(col_last)}2')

# Строка 3: чередующиеся заголовки «📦» | «дд.мм»
_DEL_FILL = fill("C6EFCE")   # зелёный — колонка поставки
for di, (dt, mnum_h, d_h) in enumerate(all_dates):
    ci_del = 8 + di * 2
    ci_ss  = 8 + di * 2 + 1
    # Заголовок Del
    c_dh = ws_g.cell(3, ci_del)
    c_dh.value = '📦'
    c_dh.font  = Font(bold=True, size=8, name="Arial")
    c_dh.fill  = _DEL_FILL
    c_dh.alignment = Alignment(horizontal='center', vertical='center')
    ws_g.column_dimensions[get_column_letter(ci_del)].width = 5.5
    # Заголовок Ss
    c_sh = ws_g.cell(3, ci_ss)
    c_sh.value = dt.strftime('%d.%m')
    c_sh.font  = Font(bold=True, color="FFFFFF", size=8, name="Arial")
    c_sh.fill  = M_FILL[mnum_h]
    c_sh.alignment = Alignment(horizontal='center', vertical='center')
    ws_g.column_dimensions[get_column_letter(ci_ss)].width = 5.5

# Данные: строки 4+
_gp_data_start = 4
for ri, code in enumerate(mrp_codes, _gp_data_start):
    ws_g.row_dimensions[ri].height = 14
    fb = GRY_F if ri % 2 == 0 else NO_F
    name, supp, unit = get_info(code)
    pkg  = max(1, bom.get(code, {}).get('package', 1))
    sd   = get_safety_days(code, supp)
    total_3m  = sum(sum(demand[code].get(mn, {}).values()) for mn, _, _ in MONTHS)
    n_days_3m = sum(nd for _, _, nd in MONTHS)
    avg_daily  = total_3m / n_days_3m if n_days_3m else 0
    safety_qty = round(avg_daily * sd, 2)
    stk_fallback = int(stock.get(code, 0))
    # Обнуляем спрос до даты актуальности остатков (включительно) — эти дни уже прошли
    _code_sdate = get_stock_date(code)
    daily_dem = [
        round(demand[code].get(mn, {}).get(d, 0), 4) if dt > _code_sdate else 0.0
        for (dt, mn, d) in all_dates
    ]

    # A-E: статика
    for ci, v in enumerate([code, name[:46], supp, unit, pkg], 1):
        c = ws_g.cell(ri, ci); c.value = v
        c.font = Font(size=9, name="Arial")
        c.alignment = Alignment(horizontal='left' if ci <= 2 else 'center', vertical='center')
        if fb: c.fill = fb

    # F: страховой запас (шт)
    cf = ws_g.cell(ri, 6); cf.value = safety_qty
    cf.font = Font(size=9, name="Arial"); cf.number_format = '#,##0.#'
    cf.alignment = Alignment(horizontal='center', vertical='center')
    cf.fill = fill("FFF2CC")

    # G: живой остаток из Ввод_Остатков (VLOOKUP)
    cg = ws_g.cell(ri, 7)
    cg.value = f"=IFERROR(VLOOKUP(A{ri},Ввод_Остатков!$A:$E,5,0),{stk_fallback})"
    cg.font = Font(size=9, name="Arial", bold=True); cg.number_format = '#,##0'
    cg.alignment = Alignment(horizontal='center', vertical='center')
    cg.fill = CYN_F

    # H+: Del — из Python (правила поставщика); Ss — формула от prev и Del
    _sched = DELIVERY_SCHEDULES.get(code, {})
    _sched_dels = _sched.get('dels', [0.0] * len(all_dates))
    for di, (dt, mnum_d, day_d) in enumerate(all_dates):
        ci_del = 8 + di * 2
        ci_ss  = 8 + di * 2 + 1
        prev_ss  = f"G{ri}" if di == 0 else f"{get_column_letter(8 + (di-1)*2 + 1)}{ri}"
        dem_d    = daily_dem[di]
        del_col  = get_column_letter(ci_del)
        del_val  = _sched_dels[di] if di < len(_sched_dels) else 0.0
        ss_formula = f"={prev_ss}-{dem_d}+{del_col}{ri}"

        c_del = ws_g.cell(ri, ci_del)
        c_del.value = del_val if del_val else None
        c_del.font = Font(size=8, name="Arial", bold=True)
        c_del.number_format = '#,##0'
        c_del.alignment = Alignment(horizontal='center', vertical='center')
        # Leave fill for conditional formatting

        # Write Ss cell
        c_ss = ws_g.cell(ri, ci_ss)
        c_ss.value = ss_formula
        c_ss.font = Font(size=8, name="Arial")
        c_ss.number_format = '#,##0'
        c_ss.alignment = Alignment(horizontal='center', vertical='center')
        if fb: c_ss.fill = fb

# Условное форматирование
_gp_last_row = _gp_data_start + len(mrp_codes) - 1
_gp_filter_last_row = max(_gp_last_row, 3)
ws_g.auto_filter.ref = f"A3:{get_column_letter(_total_cols_gp)}{_gp_filter_last_row}"
# Del колонки (H, J, L, ...): зелёный если > 0
_del_range = f"H{_gp_data_start}:{get_column_letter(_total_cols_gp)}{_gp_last_row}"
ws_g.conditional_formatting.add(_del_range, _FR_gp(
    formula=[f"AND(COLUMN(H{_gp_data_start})<>COLUMN(H{_gp_data_start})+1,H{_gp_data_start}>0,MOD(COLUMN(H{_gp_data_start})-8,2)=0)"],
    fill=PatternFill("solid", fgColor="C6EFCE"),
    font=Font(color="375623", bold=True, size=8, name="Arial")))
# Ss колонки: красный если < 0
_ss_range = f"I{_gp_data_start}:{get_column_letter(_total_cols_gp)}{_gp_last_row}"
ws_g.conditional_formatting.add(_ss_range, _FR_gp(
    formula=[f"AND(MOD(COLUMN(I{_gp_data_start})-8,2)=1,I{_gp_data_start}<0)"],
    fill=PatternFill("solid", fgColor="FFC7CE"),
    font=Font(color="9C0006", bold=True, size=8, name="Arial")))
# Ss колонки: жёлтый если ниже страхового
ws_g.conditional_formatting.add(_ss_range, _FR_gp(
    formula=[f"AND(MOD(COLUMN(I{_gp_data_start})-8,2)=1,I{_gp_data_start}>=0,I{_gp_data_start}<$F{_gp_data_start})"],
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
for ci,(h,w) in enumerate(zip(['Код','Наименование','Поставщик','Ед.'],[22,44,16,6]),1):
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
    for ci,v in enumerate([code,name[:50],supp,unit],1):
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
lh=['Код','Наименование','Поставщик','Ед.','Тип/Норма','Вкладка плана',
    f'Уп.\n(шт)',f'Потребн.\n{MONTH_SHORT[MONTHS[0][0]]}',f'Потребн.\n{MONTH_SHORT[MONTHS[1][0]]}',f'Потребн.\n{MONTH_SHORT[MONTHS[2][0]]}',
    'Итого\n3 мес.','Остаток','Статус','Раздел BOM']
lw=[22,42,16,6,28,22,7,12,12,12,12,10,16,25]
for ci,(h,w) in enumerate(zip(lh,lw),1):
    hcell(ws_bl,3,ci,h,H_FILL); ws_bl.column_dimensions[get_column_letter(ci)].width=w
for ri,code in enumerate(mrp_codes,4):
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
    vals=[code,name[:50],supp,unit,norm_str,tab,pkg,
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
SH=['Код','Наименование','Поставщик','Ед.','Уп.(шт)','Остаток']
SW=[22,44,16,6,7,10]
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
for ri,code in enumerate(mrp_codes,2):
    ws_s.row_dimensions[ri].height=14; fb=GRY_F if ri%2==0 else NO_F
    name,supp,unit=get_info(code); stk=stock.get(code,0); pkg=bom.get(code,{}).get('package',1)
    for ci,v in enumerate([code,name[:50],supp,unit,pkg,stk],1):
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
    """Дневной баланс по графику поставок; стартовый остаток — фактический (в т.ч. ручной)."""
    _code_sdate = get_stock_date(code)
    pairs = DELIVERY_SCHEDULES.get(code, {}).get('pairs', {})
    stk0 = float(stock.get(code, 0) or 0)
    series = []
    for di, (dt, mnum, d) in enumerate(all_dates):
        if dt <= _code_sdate:
            series.append({'dt': dt, 'mn': mnum, 'd': d, 'bal': stk0, 'del': 0.0})
        else:
            del_d, bal = pairs.get(di, (0.0, stk0))
            series.append({'dt': dt, 'mn': mnum, 'd': d, 'bal': bal, 'del': del_d})
    return series


def _risk_simulation(code):
    """Дефицит по симуляции графика; справочно — ближайшие поставки и дата закрытия."""
    _code_sdate = get_stock_date(code)
    series = _risk_balance_series(code)
    upcoming = [(s['dt'], s['del']) for s in series if s['dt'] > _code_sdate and s['del'] > 0]

    first_def = None
    min_bal = float(stock.get(code, 0) or 0)
    closure_date = None
    in_def = False

    for s in series:
        if s['dt'] <= _code_sdate:
            continue
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
RH = ['Код', 'Наименование', 'Поставщик', 'Ед.', 'Уп.', 'Остаток', 'Ср/день',
      'Норм.\nдней', 'Норм.\nтреб.', 'Дефицит', 'Дата\nдефицита', 'Дн.\nзапаса',
      'Статус', 'Комментарий', 'Ближ.\nпоставка', 'Кол-во\nпоставки', 'Закрытие\nдефицита']
RW = [22, 38, 16, 6, 7, 11, 10, 7, 12, 12, 12, 9, 16, 30, 12, 11, 12]
for ci, (h, w) in enumerate(zip(RH, RW), 1):
    hcell(ws_r, 3, ci, h, H_FILL)
    ws_r.column_dimensions[get_column_letter(ci)].width = w

risk_rows = []
for code in mrp_codes:
    name, supp, unit = get_info(code)
    stk = float(stock.get(code, 0) or 0)
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
        if code in manual_stock_override:
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
        'code': code, 'name': name[:45], 'supplier': supp, 'unit': unit, 'pkg': pkg,
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
               'Изменяйте ✓ / пусто / Упак. → сохраните файл → перезапустите mrp_v10.py')
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
dv = DataValidation(type="list", formula1='"✓,"', allow_blank=True, showDropDown=False)
dv.error = 'Введите ✓ или оставьте пустым'
dv.errorTitle = 'Неверное значение'
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

# ═══ Сохраняем ═══════════════════════════════════════════════
wb_out.save(OUT)
print(f"\n✅ Готово! Файл сохранён: {OUT}")
print(f"   Расчётный период: {PERIOD_LABEL}")
print(f"   Листы: {[s.title for s in wb_out.worksheets]}")

# ═══ Выгрузка графиков поставок по поставщикам ════════════════
# Пересчитываем Del/Ss в Python (не читаем формулы из Excel — openpyxl их не вычисляет).
# Логика идентична График_Поставок:
#   del_d = CEILING(MAX(0, safety-(ss_prev-dem_d-dem_next))/pkg, 1)*pkg  если ss_prev-dem_d-dem_next < safety
#   ss_d  = ss_prev - dem_d + del_d
# Выгружаем только дни следующей недели (Пн–Вс).
print("\nВыгрузка графиков поставок по поставщикам...")
try:
    import math as _math

    # ── Определяем "следующую неделю" ────────────────────────────
    _today_run = datetime.date.today()
    _wd = _today_run.weekday()                              # 0=Mon
    _cur_mon  = _today_run - datetime.timedelta(days=_wd)  # Пн текущей недели
    _next_mon = _cur_mon  + datetime.timedelta(days=7)     # Пн следующей недели
    _next_sun = _next_mon + datetime.timedelta(days=6)     # Вс следующей недели
    print(f"  Следующая неделя: {_next_mon.strftime('%d.%m.%Y')} – {_next_sun.strftime('%d.%m.%Y')}")

    # Индексы дней из all_dates, попадающих в следующую неделю
    _nw_indices = [(di, dt) for di, (dt, mn, d) in enumerate(all_dates)
                   if _next_mon <= dt <= _next_sun]

    if not _nw_indices:
        print("  ⚠️  Следующая неделя выходит за пределы расчётного горизонта — выгрузка пропущена.")
    else:
        _export_dir = os.path.join(ROOT, "Графики_поставщиков")
        os.makedirs(_export_dir, exist_ok=True)

        _date_tag   = _today_run.strftime('%Y-%m-%d')
        _week_label = f"{_next_mon.strftime('%d.%m')}–{_next_sun.strftime('%d.%m.%Y')}"

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
                if not _week_dels:
                    continue  # нет поставок на эту неделю
                _info   = bom.get(_code, {})
                _pkg    = max(1, _info.get('package', 1))
                _sd     = get_safety_days(_code, _supp_name)
                _total3 = sum(sum(demand[_code].get(mn, {}).values()) for mn, _, _ in MONTHS)
                _ndays3 = sum(nd for _, _, nd in MONTHS)
                _safety = round(_total3 / _ndays3 * _sd, 1) if _ndays3 else 0
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
                         f"Неделя {_week_label}  |  Сформировано {_today_run.strftime('%d.%m.%Y')}")
            _hc.font      = Font(bold=True, size=11, color="FFFFFF", name="Arial")
            _hc.fill      = H_FILL
            _hc.alignment = Alignment(horizontal='center', vertical='center')
            _ws_s.row_dimensions[1].height = 30

            # Строка 2: заголовки колонок
            _ws_s.row_dimensions[2].height = 28
            _ch = ['Код детали', 'Наименование', 'Ед.', 'Уп.(шт)', 'Страх.\nзапас', 'ИТОГО\nнеделя']
            _cw = [22, 44, 6, 7, 10, 10]
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

                _static = [_code, _name[:52], _unit, _pkg, _safety, _week_total]
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

            # Сохраняем
            _safe_name = re.sub(r'[\\/:*?"<>|]', '_', _supp_name)
            _out_path  = os.path.join(_export_dir, f"{_safe_name}_{_date_tag}.xlsx")
            _wb_s.save(_out_path)
            _n_exported += 1
            print(f"    ✅ {_supp_name}: {len(_rows_out)} позиций → {os.path.basename(_out_path)}")

        print(f"  Экспорт завершён: {_n_exported} файлов → папка '{os.path.basename(_export_dir)}/'")

except Exception as _exp_err:
    print(f"  ⚠️  Ошибка выгрузки графиков: {_exp_err}")
    import traceback; traceback.print_exc()
