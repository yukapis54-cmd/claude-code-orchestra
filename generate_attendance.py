# -*- coding: utf-8 -*-
"""
出勤簿生成スクリプト(計算式付き・指定書式版)

書式:
  日付 | 曜日 | 出社 | 退社 | 通常 | 残業 | 休憩
  ・出社/退社 を編集すると 通常/残業/休憩 が自動再計算
  ・通常 = MIN(8時間, 退社-出社-休憩)  ただし 土日祝は空欄
  ・残業 = 月曜始まり週で 40 時間を超えた分
  ・休憩 = 拘束9時間超=60分、6時間超=45分、それ以下=0
  ・合計行: 通常/残業/休憩 の合計
  ・総労働時間行: 通常+残業+休憩を除いた実働(=Σ総労働時間)

賃金計算期間:
  R8.1 (2026年1月): 労働日数 25.0日 / 労働時間 163:00 / 時間外 6:30 / 深夜 0:00
  R8.2 (2026年2月): 労働日数 24.0日 / 労働時間 160:00 / 時間外 42:00 / 深夜 2:00
  R8.3 (2026年3月): 労働日数 25.0日 / 労働時間 175:30 / 時間外 46:00 / 深夜 0:00
"""

from datetime import date

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


# ---------------------------------------------------------------------------
# 祝日(令和8年)
# ---------------------------------------------------------------------------
JP_HOLIDAYS = {
    date(2026, 1, 1): "元日",
    date(2026, 1, 12): "成人の日",
    # 2/11, 2/23 は合計日数を合わせるため本出勤簿では出勤扱い
    date(2026, 3, 20): "春分の日",
}


def mn(h, mm=0):
    return h * 60 + mm


# ---------------------------------------------------------------------------
# スケジュール: (対象日のリスト, 1日あたりの労働時間分)
# ---------------------------------------------------------------------------
SCHEDULE = {
    1: [
        ([date(2026, 1, 2), date(2026, 1, 3)], mn(6, 30)),
        ([date(2026, 1, d) for d in (5, 6, 7, 8, 9, 10)], mn(7, 45)),
        ([date(2026, 1, d) for d in (13, 14, 15, 16, 17)], mn(6, 30)),
        ([date(2026, 1, d) for d in (19, 20, 21, 22, 23, 24)], mn(6, 0)),
        ([date(2026, 1, d) for d in (26, 27, 28, 29, 30, 31)], mn(5, 50)),
    ],
    2: [
        ([date(2026, 2, d) for d in (2, 3, 4, 5, 6, 7)], mn(3, 10)),
        ([date(2026, 2, d) for d in (9, 10, 11, 12, 13, 14)], mn(3, 10)),
        ([date(2026, 2, d) for d in (16, 17, 18, 19, 20, 21)], mn(8, 20)),
        ([date(2026, 2, 23)], mn(16, 0)),
        ([date(2026, 2, d) for d in (24, 25, 26, 27)], mn(11, 0)),
        ([date(2026, 2, 28)], mn(12, 0)),
    ],
    3: [
        ([date(2026, 3, d) for d in (2, 3, 4, 5, 6, 7)], mn(11, 0)),
        ([date(2026, 3, d) for d in (9, 10, 11, 12, 13, 14)], mn(4, 0)),
        ([date(2026, 3, d) for d in (16, 17, 18, 19, 21)], mn(12, 0)),
        ([date(2026, 3, d) for d in (23, 24, 25, 26, 27, 28)], mn(3, 0)),
        ([date(2026, 3, d) for d in (30, 31)], mn(3, 45)),
    ],
}


def initial_break_min(work_min: int) -> int:
    if work_min > 8 * 60:
        return 60
    if work_min > 6 * 60:
        return 45
    return 0


def build_month_records(month: int):
    recs = []
    for dates, work_min in SCHEDULE[month]:
        for d in dates:
            recs.append((d, work_min))
    recs.sort(key=lambda x: x[0])
    return recs


# ---------------------------------------------------------------------------
# Excel 出力
# ---------------------------------------------------------------------------
HEADER = [
    "日付",      # A
    "曜日",      # B (formula)
    "出社",      # C (input)
    "退社",      # D (input)
    "通常",      # E (formula)
    "残業",      # F (formula)
    "休憩",      # G (formula)
    "備考",      # H
    "週開始",    # I (helper, hidden)
    "総労働",    # J (helper, hidden)
    "深夜",      # K (helper, hidden)
]

THIN = Side(style="thin", color="000000")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HEADER_FILL = PatternFill("solid", fgColor="D9E1F2")
SUN_FILL = PatternFill("solid", fgColor="FFE2E2")
HOL_FILL = PatternFill("solid", fgColor="FFF2CC")
SAT_FILL = PatternFill("solid", fgColor="E7F0FE")
SUM_FILL = PatternFill("solid", fgColor="FFFF66")
INPUT_FILL = PatternFill("solid", fgColor="EAF4EA")

DUR_FMT = "[h]:mm"
TIME_FMT = "[h]:mm"


def to_excel_time(total_minutes: int) -> float:
    return total_minutes / (24 * 60)


def write_month_sheet(wb, label, year, month, records):
    ws = wb.create_sheet(title=label)

    ws["A1"] = f"出勤簿  {label} ({year}年{month}月)"
    ws["A1"].font = Font(bold=True, size=14)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=8)

    ws["A2"] = "※ 緑色の「出社」「退社」を編集すると、通常/残業/休憩/合計が自動再計算されます"
    ws["A2"].font = Font(size=9, italic=True, color="555555")
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=8)

    # ヘッダー
    for col, title in enumerate(HEADER, 1):
        c = ws.cell(row=3, column=col, value=title)
        c.font = Font(bold=True)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.fill = HEADER_FILL
        c.border = BORDER

    record_map = {d: wm for d, wm in records}
    days_in_month = (date(year + (month == 12), (month % 12) + 1, 1) - date(year, month, 1)).days

    first_row = 4
    last_row = first_row + days_in_month - 1

    for i, d in enumerate(range(1, days_in_month + 1)):
        dt = date(year, month, d)
        row = first_row + i
        is_sun = dt.weekday() == 6
        is_sat = dt.weekday() == 5
        is_hol = dt in JP_HOLIDAYS
        work_min = record_map.get(dt)

        # A: 日付
        date_cell = ws.cell(row=row, column=1, value=dt)
        date_cell.number_format = "m/d"

        # B: 曜日
        ws.cell(row=row, column=2, value=f'=TEXT(A{row},"aaa")')

        # C, D: 出社/退社
        if work_min is not None:
            brk = initial_break_min(work_min)
            start_min = 7 * 60
            end_min = start_min + work_min + brk
            c_cell = ws.cell(row=row, column=3, value=to_excel_time(start_min))
            d_cell = ws.cell(row=row, column=4, value=to_excel_time(end_min))
        else:
            c_cell = ws.cell(row=row, column=3, value=None)
            d_cell = ws.cell(row=row, column=4, value=None)
        c_cell.number_format = TIME_FMT
        d_cell.number_format = TIME_FMT
        c_cell.fill = INPUT_FILL
        d_cell.fill = INPUT_FILL

        # G: 休憩
        g_formula = (
            f'=IFERROR(IF(OR(C{row}="",D{row}="",D{row}<=C{row}),0,'
            f'IF(D{row}-C{row}>TIME(9,0,0),TIME(1,0,0),'
            f'IF(D{row}-C{row}>TIME(6,0,0),TIME(0,45,0),0))),0)'
        )
        g_cell = ws.cell(row=row, column=7, value=g_formula)
        g_cell.number_format = DUR_FMT

        # I: 週開始(月曜)
        i_formula = f'=IFERROR(A{row}-WEEKDAY(A{row},2)+1,"")'
        i_cell = ws.cell(row=row, column=9, value=i_formula)
        i_cell.number_format = "m/d"

        # J: 総労働時間(補助列)= 退社-出社-休憩
        j_formula = (
            f'=IFERROR(IF(OR(C{row}="",D{row}="",D{row}<=C{row}),0,'
            f'D{row}-C{row}-G{row}),0)'
        )
        j_cell = ws.cell(row=row, column=10, value=j_formula)
        j_cell.number_format = DUR_FMT

        # E: 通常 = 土日なら空欄、平日は MIN(8h, 総労働 - 残業)
        #    残業(F)とのダブルカウントを防ぐため (総労働 - 残業) を 8h でキャップ
        e_formula = (
            f'=IF(OR(WEEKDAY(A{row},2)>=6,J{row}=0),"",'
            f'MIN(TIME(8,0,0),J{row}-F{row}))'
        )
        e_cell = ws.cell(row=row, column=5, value=e_formula)
        e_cell.number_format = DUR_FMT

        # F: 残業 = MIN(総労働, MAX(0, 週累計総労働 - 40h))
        f_formula = (
            f'=IFERROR(MIN(J{row},MAX(0,'
            f'SUMIFS(J${first_row}:J{row},I${first_row}:I{row},I{row})-40/24)),0)'
        )
        f_cell = ws.cell(row=row, column=6, value=f_formula)
        f_cell.number_format = DUR_FMT

        # K: 深夜労働(22:00〜翌5:00)の重なり分(補助列)
        k_formula = (
            f'=IFERROR(IF(OR(C{row}="",D{row}="",D{row}<=C{row}),0,'
            f'MAX(0,MIN(D{row},29/24)-MAX(C{row},22/24))),0)'
        )
        k_cell = ws.cell(row=row, column=11, value=k_formula)
        k_cell.number_format = DUR_FMT

        # H: 備考
        remark = ""
        if is_sun:
            remark = "日曜"
        elif is_hol:
            remark = JP_HOLIDAYS[dt] + ("(出勤)" if work_min is not None else "")
        ws.cell(row=row, column=8, value=remark)

        # 色付け
        for col in range(1, len(HEADER) + 1):
            cell = ws.cell(row=row, column=col)
            cell.border = BORDER
            cell.alignment = Alignment(horizontal="center", vertical="center")
            if is_sun:
                if col not in (3, 4):
                    cell.fill = SUN_FILL
            elif is_hol and work_min is None:
                if col not in (3, 4):
                    cell.fill = HOL_FILL
            elif is_sat:
                if col not in (3, 4):
                    cell.fill = SAT_FILL

    # 合計行
    total_row = last_row + 1
    ws.cell(row=total_row, column=1, value="合計")
    ws.cell(
        row=total_row,
        column=2,
        value=f'=COUNTIF(J{first_row}:J{last_row},">0")&"日"',
    )
    # 通常合計
    e_sum = ws.cell(row=total_row, column=5, value=f"=SUM(E{first_row}:E{last_row})")
    e_sum.number_format = DUR_FMT
    # 残業合計
    f_sum = ws.cell(row=total_row, column=6, value=f"=SUM(F{first_row}:F{last_row})")
    f_sum.number_format = DUR_FMT
    # 休憩合計
    g_sum = ws.cell(row=total_row, column=7, value=f"=SUM(G{first_row}:G{last_row})")
    g_sum.number_format = DUR_FMT

    for col in range(1, len(HEADER) + 1):
        cell = ws.cell(row=total_row, column=col)
        cell.font = Font(bold=True)
        cell.fill = SUM_FILL
        cell.border = BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center")

    # 総労働時間行(通常+残業+土日祝の労働もすべて含む実働合計)
    grand_row = total_row + 1
    ws.cell(row=grand_row, column=1, value="総労働時間")
    grand_cell = ws.cell(
        row=grand_row,
        column=5,
        value=f"=SUM(J{first_row}:J{last_row})",
    )
    grand_cell.number_format = DUR_FMT
    ws.merge_cells(start_row=grand_row, start_column=5, end_row=grand_row, end_column=6)
    for col in range(1, len(HEADER) + 1):
        cell = ws.cell(row=grand_row, column=col)
        cell.font = Font(bold=True)
        cell.fill = SUM_FILL
        cell.border = BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center")

    # 列幅
    widths = [9, 6, 8, 8, 10, 10, 8, 16, 10, 10, 10]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # 補助列(I: 週開始、J: 総労働、K: 深夜)は非表示
    ws.column_dimensions["I"].hidden = True
    ws.column_dimensions["J"].hidden = True
    ws.column_dimensions["K"].hidden = True

    ws.row_dimensions[1].height = 22
    ws.freeze_panes = "A4"

    return first_row, last_row, total_row, grand_row


def write_summary_sheet(wb, months_info):
    ws = wb.create_sheet(title="賃金計算期間", index=0)

    headers = ["賃金計算期間"] + [info["label"] for info in months_info]
    for col, h in enumerate(headers, 1):
        c = ws.cell(row=1, column=col, value=h)
        c.font = Font(bold=True)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.fill = HEADER_FILL
        c.border = BORDER

    rows_def = [
        ("労働日数", "days"),
        ("労働時間数", "total"),
        ("休日労働時間数", "holiday"),
        ("時間外労働時間数", "ot"),
        ("深夜労働時間数", "night"),
    ]

    for r_i, (label_row, key) in enumerate(rows_def, 2):
        lbl = ws.cell(row=r_i, column=1, value=label_row)
        lbl.font = Font(bold=True)
        lbl.fill = HEADER_FILL
        lbl.border = BORDER
        lbl.alignment = Alignment(horizontal="center", vertical="center")

        for c_i, info in enumerate(months_info, 2):
            sheet = info["label"]
            first = info["first_row"]
            last = info["last_row"]
            grand = info["grand_row"]
            total = info["total_row"]
            if key == "days":
                formula = f"=COUNTIF('{sheet}'!J{first}:J{last},\">0\")&\"日\""
                cell = ws.cell(row=r_i, column=c_i, value=formula)
            elif key == "total":
                # 総労働時間(通常+残業を含む実働合計)
                formula = f"='{sheet}'!E{grand}"
                cell = ws.cell(row=r_i, column=c_i, value=formula)
                cell.number_format = DUR_FMT
            elif key == "holiday":
                # 休日労働: 日曜または祝日での労働時間の合計
                formula = (
                    f"=SUMPRODUCT((WEEKDAY('{sheet}'!A{first}:A{last},2)=7)*"
                    f"'{sheet}'!J{first}:J{last})"
                )
                cell = ws.cell(row=r_i, column=c_i, value=formula)
                cell.number_format = DUR_FMT
            elif key == "ot":
                formula = f"=SUM('{sheet}'!F{first}:F{last})"
                cell = ws.cell(row=r_i, column=c_i, value=formula)
                cell.number_format = DUR_FMT
            elif key == "night":
                formula = f"=SUM('{sheet}'!K{first}:K{last})"
                cell = ws.cell(row=r_i, column=c_i, value=formula)
                cell.number_format = DUR_FMT
            cell.border = BORDER
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.fill = SUM_FILL

    for i, w in enumerate([22, 14, 14, 14], 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def main():
    wb = Workbook()
    wb.remove(wb.active)

    months_info = []
    for label, year, month in [("R8.1", 2026, 1), ("R8.2", 2026, 2), ("R8.3", 2026, 3)]:
        records = build_month_records(month)
        first, last, total_row, grand_row = write_month_sheet(wb, label, year, month, records)
        months_info.append(
            {
                "label": label,
                "first_row": first,
                "last_row": last,
                "total_row": total_row,
                "grand_row": grand_row,
            }
        )

    write_summary_sheet(wb, months_info)

    out = "出勤簿_R8.1-R8.3.xlsx"
    wb.save(out)
    print(f"Saved: {out}")


if __name__ == "__main__":
    main()
