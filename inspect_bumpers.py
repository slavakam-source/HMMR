"""
inspect_bumpers.py — диагностика для расчёта потребности бамперов.

Положите этот файл рядом с mrp_v10.py (в папке MRP) и запустите:
    python inspect_bumpers.py

Скрипт НИЧЕГО не меняет. Он печатает структуру Plan-Fact (лист AS_in_F_A)
и Order_calculation_statistics (paint stats), чтобы понять, как правильно
считать цвет кузова для бамперов B02 2WD premium на 06.06.

Скопируйте ВЕСЬ вывод и пришлите его — по нему я починю расчёт.
"""
import os
import glob
import pandas as pd
from openpyxl import load_workbook

ROOT = os.path.dirname(os.path.abspath(__file__)) if '__file__' in globals() else os.getcwd()


def _latest(*patterns):
    cands = []
    for pat in patterns:
        cands.extend(glob.glob(os.path.join(ROOT, pat)))
        cands.extend(glob.glob(os.path.join(ROOT, '**', pat), recursive=True))
    cands = [p for p in cands if os.path.isfile(p) and not os.path.basename(p).startswith('~$')]
    return max(cands, key=os.path.getmtime) if cands else None


def sep(title):
    print('\n' + '=' * 70)
    print(title)
    print('=' * 70)


# ─────────────────────────────────────────────────────────────
# 1. PLAN-FACT — лист AS_in_F_A
# ─────────────────────────────────────────────────────────────
PF = _latest("*Plan-Fact*.xlsx", "*Plan_Fact*.xlsx", "*план-факт*.xlsx")
sep(f"PLAN-FACT: {os.path.basename(PF) if PF else 'НЕ НАЙДЕН'}")

if PF:
    try:
        xl = pd.ExcelFile(PF)
        print("Листы:", xl.sheet_names)
    except Exception as e:
        print("Ошибка открытия:", e)
        xl = None

    if xl and 'AS_in_F_A' in xl.sheet_names:
        df = pd.read_excel(PF, sheet_name='AS_in_F_A', header=None)
        print(f"\nРазмер AS_in_F_A: {df.shape[0]} строк × {df.shape[1]} колонок")

        # Печатаем первые 12 строк целиком (заголовки месяцев/колонок)
        print("\n--- Первые 12 строк (для понимания заголовков и колонок) ---")
        for i in range(min(12, df.shape[0])):
            vals = [str(v) for v in df.iloc[i].tolist()]
            print(f"r{i}: " + " | ".join(vals[:18]))

        # Ищем строку-заголовок с 'Batch'
        hrow = None
        for i in range(min(40, df.shape[0])):
            if any('batch' in str(v).lower() for v in df.iloc[i]):
                hrow = i
                break
        print(f"\nСтрока заголовка (Batch): r{hrow}")
        if hrow is not None:
            hdr = df.iloc[hrow].tolist()
            print("Колонки заголовка (индекс: значение):")
            for ci, v in enumerate(hdr):
                if v is not None and str(v).strip() and str(v) != 'nan':
                    print(f"   [{ci}] = {repr(v)}")

            # Печатаем все строки секции с непустым batch, где модель похожа на B02
            print("\n--- Строки партий (batch начинается с 'R'), первые 60 ---")
            shown = 0
            for i in range(hrow + 1, df.shape[0]):
                row = df.iloc[i].tolist()
                # ищем ячейку, начинающуюся с R (batch)
                bcell = None
                for v in row[:6]:
                    if isinstance(v, str) and v.strip().upper().startswith('R'):
                        bcell = v.strip()
                        break
                if not bcell:
                    continue
                print(f"r{i}: " + " | ".join(str(v) for v in row[:18]))
                shown += 1
                if shown >= 60:
                    print("   ... (обрезано)")
                    break
    else:
        print("Лист AS_in_F_A не найден.")

# ─────────────────────────────────────────────────────────────
# 2. PAINT STATS — Order_calculation_statistics
# ─────────────────────────────────────────────────────────────
PS = _latest("*Order*calculation*.xlsx", "*涂装*.xlsx", "Order_calculation_statistics_UPDATED.xlsx")
sep(f"PAINT STATS: {os.path.basename(PS) if PS else 'НЕ НАЙДЕН'}")

if PS:
    wb = load_workbook(PS, read_only=True, data_only=True)
    print("Листы:", wb.sheetnames)
    sheet = 'DAP All batches' if 'DAP All batches' in wb.sheetnames else wb.sheetnames[0]
    ws = wb[sheet]
    print(f"\nЛист: {sheet}")
    rows = list(ws.iter_rows(min_row=1, max_row=4, values_only=True))
    print("\n--- Первые 4 строки (заголовки) ---")
    for ri, r in enumerate(rows, 1):
        print(f"r{ri}: " + " | ".join(str(v) for v in (r[:20] if r else [])))

    # Печатаем первые 15 строк данных целиком
    print("\n--- Первые 15 строк данных (batch + цвета) ---")
    cnt = 0
    for r in ws.iter_rows(min_row=3, values_only=True):
        if not r or all(v in (None, '') for v in r):
            continue
        print("   " + " | ".join(str(v) for v in r[:20]))
        cnt += 1
        if cnt >= 15:
            print("   ... (обрезано)")
            break
    wb.close()

sep("ГОТОВО — скопируйте весь вывод выше и пришлите его")
