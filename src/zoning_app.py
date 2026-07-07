import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import numpy as np
from scipy import stats
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import re

# ─────────────────────────────────────────────
#  CORE ANALYTICS
# ─────────────────────────────────────────────

def quartile_zone(series: pd.Series, direction: str) -> pd.Series:
    vals = series.dropna()
    if len(vals) == 0:
        return series.map(lambda x: np.nan)
    q25 = vals.quantile(0.25)
    q50 = vals.quantile(0.50)
    q75 = vals.quantile(0.75)

    def assign(v):
        if pd.isna(v):
            return np.nan
        if "больше" in str(direction).lower():
            if v > q75:    return 1
            elif v > q50:  return 2
            elif v > q25:  return 3
            else:          return 4
        else:
            if v <= q25:   return 1
            elif v <= q50: return 2
            elif v <= q75: return 3
            else:          return 4

    return series.apply(assign)


def normality_test(arr):
    arr = arr[~np.isnan(arr)]
    if len(arr) < 3:
        return None, None, "недостаточно данных"
    if len(arr) <= 50:
        stat, p = stats.shapiro(arr)
        return stat, p, "Shapiro-Wilk"
    else:
        stat, p = stats.normaltest(arr)
        return stat, p, "D'Agostino-Pearson"


def compare_zones(q1_vals, q4_vals):
    q1 = np.array(q1_vals, dtype=float)
    q1 = q1[~np.isnan(q1)]
    q4 = np.array(q4_vals, dtype=float)
    q4 = q4[~np.isnan(q4)]

    result = {
        "n_Q1": len(q1), "n_Q4": len(q4),
        "mean_Q1": float(np.nanmean(q1)) if len(q1) else np.nan,
        "mean_Q4": float(np.nanmean(q4)) if len(q4) else np.nan,
        "sd_Q1":   float(np.nanstd(q1, ddof=1)) if len(q1) > 1 else np.nan,
        "sd_Q4":   float(np.nanstd(q4, ddof=1)) if len(q4) > 1 else np.nan,
    }
    _, p1, test_name = normality_test(q1)
    _, p4, _         = normality_test(q4)
    result["normality_test"] = test_name
    result["p_norm_Q1"] = p1
    result["p_norm_Q4"] = p4
    normal = (p1 is not None and p1 > 0.05 and p4 is not None and p4 > 0.05)

    if len(q1) < 2 or len(q4) < 2:
        result.update({"stat_test": "—", "stat_value": np.nan,
                       "p_value": np.nan, "significant": "—"})
        return result

    if normal:
        stat, p = stats.ttest_ind(q1, q4, equal_var=False)
        result["stat_test"] = "t-тест Стьюдента (Велча)"
    else:
        stat, p = stats.mannwhitneyu(q1, q4, alternative="two-sided")
        result["stat_test"] = "U-критерий Манна-Уитни"

    result["stat_value"] = round(float(stat), 4)
    result["p_value"]    = round(float(p), 4)
    result["significant"] = "Да" if p < 0.05 else "Нет"
    return result


# ─────────────────────────────────────────────
#  EXCEL STYLING
# ─────────────────────────────────────────────

ZONE_FILLS = {
    1: PatternFill("solid", fgColor="C6EFCE"),
    2: PatternFill("solid", fgColor="FFEB9C"),
    3: PatternFill("solid", fgColor="FFCC99"),
    4: PatternFill("solid", fgColor="FFC7CE"),
}
HEADER_FILL  = PatternFill("solid", fgColor="1F497D")
HEADER_FONT  = Font(bold=True, color="FFFFFF", size=10)
SUBHDR_FILL  = PatternFill("solid", fgColor="D9E1F2")
SUBHDR_FONT  = Font(bold=True, size=10)
BOLD_FONT    = Font(bold=True, size=10)
NORMAL_FONT  = Font(size=10)
CENTER_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT_ALIGN   = Alignment(horizontal="left",   vertical="center", wrap_text=True)
THIN_BORDER  = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"),  bottom=Side(style="thin"))


def _sh(cell, sub=False):
    cell.fill      = SUBHDR_FILL if sub else HEADER_FILL
    cell.font      = SUBHDR_FONT if sub else HEADER_FONT
    cell.alignment = CENTER_ALIGN
    cell.border    = THIN_BORDER


def _sc(cell, zone=None, bold=False):
    cell.font      = BOLD_FONT if bold else NORMAL_FONT
    cell.alignment = LEFT_ALIGN
    cell.border    = THIN_BORDER
    if zone in ZONE_FILLS:
        cell.fill = ZONE_FILLS[zone]


def _aw(ws, mn=8, mx=45):
    for col in ws.columns:
        w = mn
        for c in col:
            if c.value:
                w = max(w, min(mx, len(str(c.value)) + 2))
        ws.column_dimensions[get_column_letter(col[0].column)].width = w


def _safe_name(name):
    return re.sub(r'[\\/*?:\[\]]', '', name)[:31]


# ─────────────────────────────────────────────
#  REPORT BUILDER
# ─────────────────────────────────────────────

def build_report(indicators_df, directions, method_names,
                 indices_df, out_path, progress_cb=None):

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    subjects  = indicators_df.index.tolist()
    n_methods = len(method_names)
    zone_matrix = pd.DataFrame(index=subjects, columns=method_names, dtype=float)

    # ── Per-method sheets ────────────────────────────────────────────────────
    for i, meth in enumerate(method_names):
        if progress_cb:
            progress_cb(int(28 * i / n_methods), "Зонирование: " + meth)

        series    = indicators_df[meth]
        direction = directions[meth]
        zones     = quartile_zone(series, direction)
        zone_matrix[meth] = zones

        vals_clean = series.dropna()
        q25 = vals_clean.quantile(0.25) if len(vals_clean) else 0
        q50 = vals_clean.quantile(0.50) if len(vals_clean) else 0
        q75 = vals_clean.quantile(0.75) if len(vals_clean) else 0

        ws = wb.create_sheet(title=_safe_name(meth))

        headers = ["№ п.п.", "Субъект РФ", "Значение показателя",
                   "Зона (квартиль)", "Ранг неблагополучия"]
        for c, h in enumerate(headers, 1):
            _sh(ws.cell(row=1, column=c, value=h))
        ws.row_dimensions[1].height = 30

        meta = [
            ("Методика:", meth, "Направление:", direction),
            ("Пороги:", f"Q25={q25:.4g}  Q50={q50:.4g}  Q75={q75:.4g}", "", ""),
        ]
        for ri, row_data in enumerate(meta, 2):
            for ci, val in enumerate(row_data, 1):
                c = ws.cell(row=ri, column=ci, value=val)
                c.font = Font(italic=True, size=9)

        for r, subj in enumerate(subjects, 4):
            val  = series.get(subj, np.nan)
            zone = zones.get(subj, np.nan)
            z_int = int(zone) if not pd.isna(zone) else None
            row_vals = [r - 3, subj,
                        round(float(val), 4) if not pd.isna(val) else "",
                        z_int if z_int else "",
                        z_int if z_int else ""]
            for c, v in enumerate(row_vals, 1):
                _sc(ws.cell(row=r, column=c, value=v), zone=z_int)

        leg_r = len(subjects) + 6
        ws.cell(row=leg_r, column=1, value="Легенда:").font = BOLD_FONT
        for z, label, color in [
            (1, "Зона 1 — наиболее благоприятная", "C6EFCE"),
            (2, "Зона 2",                          "FFEB9C"),
            (3, "Зона 3",                          "FFCC99"),
            (4, "Зона 4 — наиболее неблагоприятная","FFC7CE"),
        ]:
            c = ws.cell(row=leg_r + z, column=1, value=label)
            c.fill   = PatternFill("solid", fgColor=color)
            c.border = THIN_BORDER
        _aw(ws)

    # ── Aggregate sheet ──────────────────────────────────────────────────────
    if progress_cb:
        progress_cb(32, "Агрегирование рангов…")

    score_series  = zone_matrix.astype(float).sum(axis=1, skipna=True)
    hits_series   = (zone_matrix.astype(float) == 4).sum(axis=1)
    rank_by_score = score_series.rank(method="min").astype(int)
    rank_by_hits  = hits_series.rank(method="min", ascending=False).astype(int)

    ws_agg = wb.create_sheet(title="Агрегат_Ранги")
    agg_h  = (["№ п.п.", "Субъект РФ"] +
               ["Ранг: " + m for m in method_names] +
               ["Суммарный балл неблагополучия",
                "Число попаданий в Зону 4",
                "Итоговый ранг по баллу",
                "Итоговый ранг по числу попаданий"])
    for c, h in enumerate(agg_h, 1):
        _sh(ws_agg.cell(row=1, column=c, value=h))
    ws_agg.row_dimensions[1].height = 40

    for r, subj in enumerate(subjects, 2):
        ws_agg.cell(row=r, column=1, value=r - 1).border = THIN_BORDER
        ws_agg.cell(row=r, column=2, value=subj).border  = THIN_BORDER
        for ci, meth in enumerate(method_names, 3):
            z = zone_matrix.at[subj, meth]
            z_int = int(z) if not pd.isna(z) else None
            _sc(ws_agg.cell(row=r, column=ci, value=z_int if z_int else ""), zone=z_int)
        base = len(method_names) + 3
        for ci, val in enumerate([
            round(float(score_series[subj]), 2),
            int(hits_series[subj]),
            int(rank_by_score[subj]),
            int(rank_by_hits[subj]),
        ], base):
            c = ws_agg.cell(row=r, column=ci, value=val)
            c.border = THIN_BORDER
            c.font   = BOLD_FONT
    _aw(ws_agg)

    # ── Combined sheet ───────────────────────────────────────────────────────
    if progress_cb:
        progress_cb(55, "Объединение с индексами…")

    idx_cols = indices_df.columns.tolist()
    ws_c = wb.create_sheet(title="Субъекты_Зоны_Индексы")
    comb_h = (["№ п.п.", "Субъект РФ",
               "Суммарный балл", "Попаданий в Зону 4",
               "Итоговый ранг по баллу"] +
               ["Зона: " + m for m in method_names] + idx_cols)
    for c, h in enumerate(comb_h, 1):
        _sh(ws_c.cell(row=1, column=c, value=h))
    ws_c.row_dimensions[1].height = 40

    for r, subj in enumerate(subjects, 2):
        ws_c.cell(row=r, column=1, value=r-1).border = THIN_BORDER
        ws_c.cell(row=r, column=2, value=subj).border = THIN_BORDER
        ws_c.cell(row=r, column=3, value=round(float(score_series[subj]),2)).border = THIN_BORDER
        ws_c.cell(row=r, column=4, value=int(hits_series[subj])).border = THIN_BORDER
        ws_c.cell(row=r, column=5, value=int(rank_by_score[subj])).border = THIN_BORDER
        for ci, meth in enumerate(method_names, 6):
            z = zone_matrix.at[subj, meth]
            z_int = int(z) if not pd.isna(z) else None
            _sc(ws_c.cell(row=r, column=ci, value=z_int if z_int else ""), zone=z_int)
        base = 6 + len(method_names)
        for ci, ic in enumerate(idx_cols, base):
            v = indices_df.at[subj, ic] if subj in indices_df.index else ""
            ws_c.cell(row=r, column=ci, value=v).border = THIN_BORDER
    _aw(ws_c)

    # ── Statistical comparison ───────────────────────────────────────────────
    if progress_cb:
        progress_cb(70, "Статистический анализ по каждой методике…")

    ws_st = wb.create_sheet(title="Стат_Сравнение_Зон")
    stat_h = ["Методика", "Индекс / Заболеваемость", "Статистический тест",
              "Тест нормальности", "p-норм Q1", "p-норм Q4",
              "n Q1 (лучшая зона)", "Среднее Q1", "СО Q1",
              "n Q4 (худшая зона)", "Среднее Q4", "СО Q4",
              "Значение статистики", "p-значение",
              "Статистически значимо (p<0.05)"]
    for c, h in enumerate(stat_h, 1):
        _sh(ws_st.cell(row=1, column=c, value=h))
    ws_st.row_dimensions[1].height = 34

    rn = 2

    def _fmt(v, nd=4):
        try:
            return round(float(v), nd) if not pd.isna(v) else "—"
        except Exception:
            return str(v)

    for meth in method_names:
        method_zones = zone_matrix[meth].astype(float)
        q1_subjs = method_zones[method_zones == 1].index.tolist()
        q4_subjs = method_zones[method_zones == 4].index.tolist()

        hdr = ws_st.cell(
            row=rn,
            column=1,
            value=f"Методика: {meth} | Q1 n={len(q1_subjs)} | Q4 n={len(q4_subjs)}"
        )
        hdr.font = Font(bold=True, color="1F1F1F", size=10)
        hdr.fill = SUBHDR_FILL
        hdr.border = THIN_BORDER
        ws_st.merge_cells(start_row=rn, start_column=1, end_row=rn, end_column=len(stat_h))
        rn += 1

        info1 = ws_st.cell(
            row=rn,
            column=1,
            value="Субъекты Q1: " + (", ".join(q1_subjs[:10]) + ("…" if len(q1_subjs) > 10 else ""))
        )
        info1.font = Font(italic=True, size=9, color="1F497D")
        ws_st.merge_cells(start_row=rn, start_column=1, end_row=rn, end_column=len(stat_h))
        rn += 1

        info2 = ws_st.cell(
            row=rn,
            column=1,
            value="Субъекты Q4: " + (", ".join(q4_subjs[:10]) + ("…" if len(q4_subjs) > 10 else ""))
        )
        info2.font = Font(italic=True, size=9, color="C00000")
        ws_st.merge_cells(start_row=rn, start_column=1, end_row=rn, end_column=len(stat_h))
        rn += 1

        for ic in idx_cols:
            q1v = [indices_df.at[s, ic] for s in q1_subjs if s in indices_df.index]
            q4v = [indices_df.at[s, ic] for s in q4_subjs if s in indices_df.index]
            res = compare_zones(q1v, q4v)

            row_data = [
                meth, ic, res["stat_test"], res["normality_test"],
                _fmt(res["p_norm_Q1"]), _fmt(res["p_norm_Q4"]),
                res["n_Q1"], _fmt(res["mean_Q1"]), _fmt(res["sd_Q1"]),
                res["n_Q4"], _fmt(res["mean_Q4"]), _fmt(res["sd_Q4"]),
                _fmt(res["stat_value"]), _fmt(res["p_value"]),
                res["significant"],
            ]
            for c, val in enumerate(row_data, 1):
                cell = ws_st.cell(row=rn, column=c, value=val)
                cell.border = THIN_BORDER
                cell.font = NORMAL_FONT
                cell.alignment = LEFT_ALIGN if c in (1, 2) else CENTER_ALIGN
                if c == len(stat_h):
                    if val == "Да":
                        cell.font = Font(bold=True, color="C00000", size=10)
                    elif val == "Нет":
                        cell.font = Font(color="375623", size=10)
            rn += 1

        rn += 1
    _aw(ws_st)

    # ── Summary distribution ─────────────────────────────────────────────────
    if progress_cb:
        progress_cb(88, "Сводная таблица распределения…")

    ws_d = wb.create_sheet(title="Сводная_Распределение")
    dist_h = ["Методика", "Зона 1 (n)", "Зона 2 (n)", "Зона 3 (n)", "Зона 4 (n)",
              "Пропуски (n)", "% в Зоне 4"]
    for c, h in enumerate(dist_h, 1):
        _sh(ws_d.cell(row=1, column=c, value=h))
    ws_d.row_dimensions[1].height = 28

    for r, meth in enumerate(method_names, 2):
        col = zone_matrix[meth].astype(float)
        counts = {z: int((col == z).sum()) for z in [1,2,3,4]}
        na     = int(col.isna().sum())
        total  = sum(counts.values())
        pct4   = round(counts[4] / total * 100, 1) if total > 0 else ""
        row_data = [meth, counts[1], counts[2], counts[3], counts[4], na, pct4]
        for c, val in enumerate(row_data, 1):
            cell = ws_d.cell(row=r, column=c, value=val)
            cell.border    = THIN_BORDER
            cell.font      = NORMAL_FONT
            cell.alignment = LEFT_ALIGN if c == 1 else CENTER_ALIGN
            zone_map = {2: 1, 3: 2, 4: 3, 5: 4}
            if c in zone_map:
                cell.fill = ZONE_FILLS[zone_map[c]]
    _aw(ws_d)

    if progress_cb:
        progress_cb(96, "Сохранение файла…")
    wb.save(out_path)
    if progress_cb:
        progress_cb(100, "Готово!")


# ─────────────────────────────────────────────
#  GUI
# ─────────────────────────────────────────────

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Многокритериальное зонирование субъектов РФ")
        self.geometry("880x700")
        self.resizable(True, True)
        self.configure(bg="#F0F4F8")
        self.ind_path    = tk.StringVar()
        self.idx_path    = tk.StringVar()
        self.out_path    = tk.StringVar()
        self.method_vars = {}
        self.methods_df  = None
        self.directions  = {}
        self.indices_df  = None
        self._build_ui()

    def _btn(self, **kw):
        base = dict(font=("Segoe UI", 9), bg="#2E75B6", fg="white",
                    relief="flat", padx=8, pady=3, cursor="hand2",
                    activebackground="#1F497D")
        base.update(kw)
        return base

    def _build_ui(self):
        tk.Label(self,
                 text="Многокритериальное зонирование субъектов РФ",
                 font=("Segoe UI", 13, "bold"), bg="#1F497D", fg="white",
                 pady=10).pack(fill="x")
        tk.Label(self,
                 text="ФГБНУ «Национальный НИИ общественного здоровья имени Н.А. Семашко»",
                 font=("Segoe UI", 9), bg="#1F497D", fg="#C5D5E8",
                 pady=3).pack(fill="x")

        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=10, pady=8)

        # Tab 1 ─ Loading
        t1 = tk.Frame(nb, bg="#F0F4F8"); nb.add(t1, text="  1. Загрузка данных  ")
        self._file_row(t1, "Файл показателей (.xlsx):", self.ind_path,
                       self._load_indicators, 0)
        self._file_row(t1, "Файл индексов / заболеваемости (.xlsx):", self.idx_path,
                       self._load_indices, 1)
        self._file_row(t1, "Выходной файл отчёта (.xlsx):", self.out_path,
                       self._choose_out, 2, save=True)
        tk.Label(t1, text="Информация о загруженных данных:",
                 font=("Segoe UI", 10, "bold"), bg="#F0F4F8").grid(
                 row=3, column=0, columnspan=3, sticky="w", padx=10, pady=6)
        self.info_text = tk.Text(t1, height=12, state="disabled",
                                 font=("Consolas", 9), bg="#fff",
                                 relief="solid", borderwidth=1)
        self.info_text.grid(row=4, column=0, columnspan=3,
                            padx=10, pady=4, sticky="nsew")
        t1.rowconfigure(4, weight=1); t1.columnconfigure(1, weight=1)

        # Tab 2 ─ Methods
        t2 = tk.Frame(nb, bg="#F0F4F8"); nb.add(t2, text="  2. Выбор методик  ")
        bf = tk.Frame(t2, bg="#F0F4F8"); bf.pack(fill="x", padx=10, pady=4)
        tk.Button(bf, text="Выбрать все",  command=self._sel_all,  **self._btn()).pack(side="left", padx=4)
        tk.Button(bf, text="Снять все",    command=self._desel_all, **self._btn()).pack(side="left", padx=4)
        cf = tk.Frame(t2, bg="#F0F4F8"); cf.pack(fill="both", expand=True, padx=10, pady=4)
        self.cv = tk.Canvas(cf, bg="#fff", relief="solid", borderwidth=1)
        sb2 = ttk.Scrollbar(cf, orient="vertical", command=self.cv.yview)
        self.cv.configure(yscrollcommand=sb2.set)
        sb2.pack(side="right", fill="y"); self.cv.pack(side="left", fill="both", expand=True)
        self.meth_inner = tk.Frame(self.cv, bg="#fff")
        self.cv.create_window((0, 0), window=self.meth_inner, anchor="nw")
        self.meth_inner.bind("<Configure>",
            lambda e: self.cv.configure(scrollregion=self.cv.bbox("all")))
        self.cnt_lbl = tk.Label(t2, text="Методики не загружены.",
                                font=("Segoe UI", 9), bg="#F0F4F8", fg="#555")
        self.cnt_lbl.pack(pady=2)

        # Tab 3 ─ Run
        t3 = tk.Frame(nb, bg="#F0F4F8"); nb.add(t3, text="  3. Запуск анализа  ")
        tk.Label(t3, text="Параметры анализа",
                 font=("Segoe UI", 11, "bold"), bg="#F0F4F8").pack(pady=(16, 4))

        lf = tk.LabelFrame(t3, text=" Настройки ", bg="#F0F4F8",
                           font=("Segoe UI", 9))
        lf.pack(fill="x", padx=20, pady=4)
        params = [
            "• Зонирование: квартильное деление (Q1 — наиболее благоприятная, Q4 — наиболее неблагоприятная)",
            "• Веса методик: равные",
            "• Статистический тест: t-критерий Стьюдента (Welch) при нормальном распределении; U-критерий Манна-Уитни при ненормальном",
            "• Тест нормальности: Шапиро-Уилка (n<=50) / Д'Агостино-Пирсон (n>50)",
            "• Сравнение: субъекты Q1 (наименьший суммарный балл) vs Q4 (наибольший суммарный балл)",
            "• Пропуски: исключаются из всех расчётов",
        ]
        for p in params:
            tk.Label(lf, text=p, bg="#F0F4F8", font=("Segoe UI", 9),
                     justify="left").pack(anchor="w", padx=10, pady=2)

        self.run_btn = tk.Button(t3,
            text="   Запустить анализ и сформировать отчёт   ",
            font=("Segoe UI", 11, "bold"), bg="#1F497D", fg="white",
            activebackground="#2E75B6", relief="flat", pady=10,
            cursor="hand2", command=self._run)
        self.run_btn.pack(pady=14, padx=40, fill="x")

        self.pb = ttk.Progressbar(t3, length=600, mode="determinate")
        self.pb.pack(pady=2)
        self.pb_lbl = tk.Label(t3, text="", font=("Segoe UI", 9),
                               bg="#F0F4F8", fg="#555")
        self.pb_lbl.pack()

        self.log = tk.Text(t3, height=12, state="disabled",
                           font=("Consolas", 9), bg="#fff",
                           relief="solid", borderwidth=1)
        self.log.pack(fill="both", expand=True, padx=20, pady=8)

    def _file_row(self, parent, label, var, cmd, row, save=False):
        tk.Label(parent, text=label, font=("Segoe UI", 9),
                 bg="#F0F4F8").grid(row=row, column=0, sticky="w",
                                    padx=10, pady=6)
        tk.Entry(parent, textvariable=var, font=("Segoe UI", 9),
                 width=54, relief="solid", borderwidth=1).grid(
                 row=row, column=1, padx=4, pady=6, sticky="ew")
        tk.Button(parent, text="Сохранить" if save else "Обзор",
                  command=cmd, **self._btn()).grid(row=row, column=2,
                                                   padx=6, pady=6)

    def _log_add(self, msg):
        self.log.configure(state="normal")
        self.log.insert("end", msg + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def _set_info(self, msg):
        self.info_text.configure(state="normal")
        self.info_text.delete("1.0", "end")
        self.info_text.insert("end", msg)
        self.info_text.configure(state="disabled")

    def _load_indicators(self):
        path = filedialog.askopenfilename(
            title="Файл с показателями",
            filetypes=[("Excel", "*.xlsx *.xls")])
        if not path:
            return
        self.ind_path.set(path)
        try:
            raw = pd.read_excel(path, header=None)
            # Row 0 = column headers (№, Субъект РФ, meth1, meth2, ...)
            # Row 1 = label row "Значение показателя №N"
            # Row 2 = direction "Больше - лучше" / "Меньше - лучше"
            # Row 3+ = numeric data
            headers   = raw.iloc[0].tolist()
            dirs_row  = raw.iloc[2].tolist()
            meth_names = [str(h) for h in headers[2:] if pd.notna(h)]
            dirs_vals  = dirs_row[2:2 + len(meth_names)]
            self.directions = {m: d for m, d in zip(meth_names, dirs_vals)}

            data = raw.iloc[3:].reset_index(drop=True)
            data.columns = [str(h) for h in headers]
            data = data.dropna(subset=["Субъект РФ"])
            data = data.set_index("Субъект РФ")
            data = data.drop(columns=["№ п.п."], errors="ignore")
            for m in meth_names:
                if m in data.columns:
                    data[m] = pd.to_numeric(data[m], errors="coerce")
            self.methods_df = data[[m for m in meth_names if m in data.columns]]
            self._populate_methods()

            lines = [
                "Файл показателей загружен успешно",
                f"  Субъектов РФ : {len(self.methods_df)}",
                f"  Методик      : {len(self.directions)}",
                "",
                "Методики (первые 15):",
            ]
            for m, d in list(self.directions.items())[:15]:
                lines.append(f"  {m}  [{d}]")
            if len(self.directions) > 15:
                lines.append(f"  … ещё {len(self.directions)-15}")
            self._set_info("\n".join(lines))
        except Exception as e:
            messagebox.showerror("Ошибка загрузки показателей", str(e))

    def _load_indices(self):
        path = filedialog.askopenfilename(
            title="Файл с индексами / заболеваемостью",
            filetypes=[("Excel", "*.xlsx *.xls")])
        if not path:
            return
        self.idx_path.set(path)
        try:
            df = pd.read_excel(path)
            df = df.dropna(subset=["Субъект РФ"])
            df = df.set_index("Субъект РФ")
            df = df.drop(columns=["№ п.п."], errors="ignore")
            for c in df.columns:
                df[c] = pd.to_numeric(df[c], errors="coerce")
            self.indices_df = df
            cur = self.info_text.get("1.0", "end").rstrip()
            add = (f"\n\nФайл индексов загружен успешно"
                   f"\n  Субъектов РФ : {len(df)}"
                   f"\n  Показателей  : {len(df.columns)}"
                   f"\n  Переменные   : {', '.join(df.columns.tolist()[:8])}"
                   + (" …" if len(df.columns) > 8 else ""))
            self._set_info(cur + add)
        except Exception as e:
            messagebox.showerror("Ошибка загрузки индексов", str(e))

    def _choose_out(self):
        p = filedialog.asksaveasfilename(
            title="Путь выходного файла",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")])
        if p:
            self.out_path.set(p)

    def _populate_methods(self):
        for w in self.meth_inner.winfo_children():
            w.destroy()
        self.method_vars = {}
        for i, (name, direction) in enumerate(self.directions.items()):
            bg = "#FFFFFF" if i % 2 == 0 else "#F7F9FC"
            row_f = tk.Frame(self.meth_inner, bg=bg)
            row_f.pack(fill="x")
            var = tk.BooleanVar(value=True)
            self.method_vars[name] = var
            tk.Checkbutton(row_f, variable=var, bg=bg,
                           command=self._upd_cnt).pack(side="left", padx=4)
            tk.Label(row_f, text=name, bg=bg, font=("Segoe UI", 9),
                     anchor="w").pack(side="left", padx=2)
            col = "#375623" if "больше" in str(direction).lower() else "#C00000"
            tk.Label(row_f, text="[" + str(direction) + "]", bg=bg,
                     font=("Segoe UI", 8, "italic"), fg=col).pack(
                     side="right", padx=8)
        self._upd_cnt()

    def _upd_cnt(self):
        sel   = sum(v.get() for v in self.method_vars.values())
        total = len(self.method_vars)
        self.cnt_lbl.configure(text=f"Выбрано {sel} из {total} методик")

    def _sel_all(self):
        for v in self.method_vars.values(): v.set(True)
        self._upd_cnt()

    def _desel_all(self):
        for v in self.method_vars.values(): v.set(False)
        self._upd_cnt()

    def _run(self):
        if self.methods_df is None:
            messagebox.showwarning("Нет данных",
                "Загрузите файл с показателями (Вкладка 1).")
            return
        if self.indices_df is None:
            messagebox.showwarning("Нет данных",
                "Загрузите файл с индексами (Вкладка 1).")
            return
        if not self.out_path.get():
            messagebox.showwarning("Путь не указан",
                "Укажите путь выходного файла (Вкладка 1).")
            return
        selected = [m for m, v in self.method_vars.items() if v.get()]
        if not selected:
            messagebox.showwarning("Методики не выбраны",
                "Выберите хотя бы одну методику (Вкладка 2).")
            return

        self.run_btn.configure(state="disabled")
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")
        self.pb["value"] = 0

        self._log_add("Запуск анализа...")
        self._log_add(f"  Субъектов: {len(self.methods_df)}")
        self._log_add(f"  Методик выбрано: {len(selected)}")
        self._log_add(f"  Индексов/показателей: {len(self.indices_df.columns)}")
        self.update_idletasks()

        def pb_cb(val, msg):
            self.pb["value"] = val
            self.pb_lbl.configure(text=msg)
            self._log_add(f"  [{val:3d}%] {msg}")
            self.update_idletasks()

        try:
            ind_sel  = self.methods_df[selected]
            dirs_sel = {m: self.directions[m] for m in selected}
            idx_re   = self.indices_df.reindex(ind_sel.index)
            build_report(ind_sel, dirs_sel, selected, idx_re,
                         self.out_path.get(), pb_cb)
            self._log_add("\nОтчёт успешно сохранён:\n" + self.out_path.get())
            messagebox.showinfo("Готово",
                "Отчёт сформирован и сохранён:\n" + self.out_path.get())
        except Exception as e:
            import traceback
            self._log_add("\nОШИБКА: " + str(e))
            self._log_add(traceback.format_exc())
            messagebox.showerror("Ошибка", str(e))
        finally:
            self.run_btn.configure(state="normal")


if __name__ == "__main__":
    App().mainloop()
