# -*- coding: utf-8 -*-
"""
出勤簿生成スクリプト(計算式付きバージョン)

賃金計算期間:
  R8.1 (2026年1月): 労働日数 25.0日 / 労働時間 163:00 / 時間外 6:30 / 深夜 0:00
  R8.2 (2026年2月): 労働日数 24.0日 / 労働時間 160:00 / 時間外 42:00 / 深夜 2:00
  R8.3 (2026年3月): 労働日数 25.0日 / 労働時間 175:30 / 時間外 46:00 / 深夜 0:00

条件:
  - 始業は 7:00、終業時刻は編集可能
  - 休憩・労働時間・時間外・深夜・合計はすべて Excel の数式で自動計算
    -> 始業/終業を変更すれば集計が動的に更新される
  - 日曜・祝日は原則休み(R8.2のみ合計日数を合わせるため祝日出勤あり)
  - 休憩は 9時間超(拘束)で60分、6時間超(拘束)で45分、それ以下は0
  - 残業は 月曜始まり週 で 40 時間を超えた分
  - 深夜労働は 22:00〜翌5:00(終業時刻が 24:00 を超える場合は "25:00" 形式で入力可)
"""

from datetime import date, datetime, time, timedelta

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


# ---------------------------------------------------------------------------
# 祝日定義(令和8年)
# ---------------------------------------------------------------------------
JP_HOLIDAYS = {
    date(2026, 1, 1): "元日",
    date(2026, 1, 12): "成人の日",
    # 2/11, 2/23 は合計日数を合わせるため本出勤簿では出勤扱い
    date(2026, 3, 20): "春分の日",
}


# ---------------------------------------------------------------------------
# 各月のスケジュール(指定の合計時間に合うよう事前設計)
#   (対象日のリスト, 所定労働時間(分))
# ---------------------------------------------------------------------------
def mn(h, mm=0):
    return h * 60 + mm


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


# ---------------------------------------------------------------------------
# 開始終了時刻の計算(編集前の初期値用)
# ---------------------------------------------------------------------------
def initial_break(work_min: int) -> int:
    if work_min > 8 * 60:
        return 60
    if work_min > 6 * 60:
        return 45
    return 0


def build_initial_times(work_min: int):
    """始業7:00固定で、所定労働時間を満たす終業時刻を分単位で返す(24:00 超は 24 以上)"""
    brk = initial_break(work_min)
    start_min = 7 * 60  # 7:00 in minutes from midnight
    end_min = start_min + work_min + brk
    return start_min, end_min, brk


def build_month_records(month: int):
    """各月の(日付, 労働分)のリスト化"""
    recs = []
    for dates, work_min in SCHEDULE[month]:
        for d in dates:
            recs.append((d, work_min))
    recs.sort(key=lambda x: x[0])
    return recs


# ---------------------------------------------------------------------------
# Excel 書き出し(数式ベース)
# ---------------------------------------------------------------------------
HEADER = [
    "日付",            # A
    "曜日",            # B (formula)
    "始業",            # C (input)
    "終業",            # D (input)
    "休憩",            # E (formula)
    "労働時間",        # F (formula)
    "休日労働時間",    # G (formula)
    "時間外労働時間",  # H (formula)
    "深夜労働時間",    # I (formula)
    "備考",            # J
    "週開始(月)",    # K (helper, hidden)
]

THIN = Side(style="thin", color="000000")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HEADER_FILL = PatternFill("solid", fgColor="D9E1F2")
SUN_FILL = PatternFill("solid", fgColor="FFE2E2")
HOL_FILL = PatternFill("solid", fgColor="FFF2CC")
SUM_FILL = PatternFill("solid", fgColor="FFFF66")
INPUT_FILL = PatternFill("solid", fgColor="EAF4EA")  # 編集セル(緑系)

DUR_FMT = "[h]:mm"
TIME_FMT = "[h]:mm"


def minutes_to_excel_time(total_minutes: int) -> float:
    """分単位の値を Excel のシリアル時間(1日=1)に変換"""
    return total_minutes / (24 * 60)


def write_month_sheet(wb, label, year, month, records):
    ws = wb.create_sheet(title=label)

    ws["A1"] = f"出勤簿  {label} ({year}年{month}月)"
    ws["A1"].font = Font(bold=True, size=14)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(HEADER) - 1)

    ws["A2"] = "※ 緑色セル(始業・終業)を編集すると 休憩・労働時間・残業・深夜・合計が自動再計算されます"
    ws["A2"].font = Font(size=9, italic=True, color="555555")
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(HEADER) - 1)

    # ヘッダー
    for col, title in enumerate(HEADER, 1):
        c = ws.cell(row=3, column=col, value=title)
        c.font = Font(bold=True)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.fill = HEADER_FILL
        c.border = BORDER

    record_map = {d: wm for d, wm in records}
    days_in_month = (date(year + (month == 12), (month % 12) + 1, 1) - date(year, month, 1)).days

    first_data_row = 4
    last_data_row = first_data_row + days_in_month - 1

    for i, d in enumerate(range(1, days_in_month + 1)):
        dt = date(year, month, d)
        row = first_data_row + i
        is_sun = dt.weekday() == 6
        is_hol = dt in JP_HOLIDAYS
        work_min = record_map.get(dt)

        # A: 日付(実際の日付値として)
        date_cell = ws.cell(row=row, column=1, value=dt)
        date_cell.number_format = "m/d"

        # B: 曜日(TEXT 関数で日付から)
        ws.cell(row=row, column=2, value=f'=TEXT(A{row},"aaa")')

        # C, D: 始業/終業(数値で入力、[h]:mm 形式)
        if work_min is not None:
            s_min, e_min, _ = build_initial_times(work_min)
            c_cell = ws.cell(row=row, column=3, value=minutes_to_excel_time(s_min))
            d_cell = ws.cell(row=row, column=4, value=minutes_to_excel_time(e_min))
            c_cell.fill = INPUT_FILL
            d_cell.fill = INPUT_FILL
        else:
            c_cell = ws.cell(row=row, column=3, value=None)
            d_cell = ws.cell(row=row, column=4, value=None)
            c_cell.fill = INPUT_FILL
            d_cell.fill = INPUT_FILL
        c_cell.number_format = TIME_FMT
        d_cell.number_format = TIME_FMT

        # E: 休憩 = IF(拘束>9h,60m, IF(拘束>6h,45m, 0))
        #    IFERROR で空欄のときは 0
        e_formula = (
            f'=IFERROR(IF(D{row}-C{row}>TIME(9,0,0),TIME(1,0,0),'
            f'IF(D{row}-C{row}>TIME(6,0,0),TIME(0,45,0),0)),0)'
        )
        e_cell = ws.cell(row=row, column=5, value=e_formula)
        e_cell.number_format = DUR_FMT

        # F: 労働時間 = IF(終業<=始業 or 空白, 0, 終業-始業-休憩)
        f_formula = (
            f'=IFERROR(IF(OR(C{row}="",D{row}="",D{row}<=C{row}),0,D{row}-C{row}-E{row}),0)'
        )
        f_cell = ws.cell(row=row, column=6, value=f_formula)
        f_cell.number_format = DUR_FMT

        # G: 休日労働時間 = IF(日曜 or 祝日, 労働時間, 0)
        #    A列の日付から WEEKDAY で日曜判定。祝日は備考列(J)が "元日/成人の日/..." のときと見なす
        #    簡便のためここでは 日曜 または J列が休日名なら休日労働扱い
        g_formula = (
            f'=IF(OR(WEEKDAY(A{row},2)=7, '
            f'AND(J{row}<>"",J{row}<>"振替")), F{row}, 0)'
        )
        # 仕様書の通り休日労働は 0 固定にするため、本ファイルでは 0 を出しておき、
        # 編集者が実際に休日出勤したら手で書き換える運用とする
        g_cell = ws.cell(row=row, column=7, value=0)
        g_cell.number_format = DUR_FMT

        # K: 週開始(月曜)= A - (WEEKDAY(A,2)-1)
        k_formula = f'=IFERROR(A{row}-WEEKDAY(A{row},2)+1,"")'
        k_cell = ws.cell(row=row, column=11, value=k_formula)
        k_cell.number_format = "m/d"

        # H: 時間外労働時間(週40h超)
        #   = MIN(F, MAX(0, 週間累計F - 40h))
        #   週間累計 = SUMIFS(F$first:Frow, K$first:Krow, K)
        h_formula = (
            f'=IFERROR(MIN(F{row},MAX(0,'
            f'SUMIFS(F${first_data_row}:F{row},K${first_data_row}:K{row},K{row})'
            f'-TIME(0,0,0)-40/24)),0)'
        )
        h_cell = ws.cell(row=row, column=8, value=h_formula)
        h_cell.number_format = DUR_FMT

        # I: 深夜労働時間(22:00〜翌5:00 = 22/24〜29/24)
        #   = MAX(0, MIN(終業, 29/24) - MAX(始業, 22/24))  (始業が日中であることを前提)
        i_formula = (
            f'=IFERROR(IF(F{row}=0,0,'
            f'MAX(0,MIN(D{row},29/24)-MAX(C{row},22/24))),0)'
        )
        i_cell = ws.cell(row=row, column=9, value=i_formula)
        i_cell.number_format = DUR_FMT

        # J: 備考
        remark = ""
        if is_sun:
            remark = "日曜"
        elif is_hol:
            if work_min is not None:
                remark = JP_HOLIDAYS[dt] + "(出勤)"
            else:
                remark = JP_HOLIDAYS[dt]
        ws.cell(row=row, column=10, value=remark)

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

    # 合計行
    total_row = last_data_row + 1
    ws.cell(row=total_row, column=1, value="合計")
    # 労働日数 = COUNTIF(F列 > 0)
    ws.cell(
        row=total_row,
        column=2,
        value=f'=COUNTIF(F{first_data_row}:F{last_data_row},">0")&"日"',
    )
    # 労働時間合計
    f_sum = ws.cell(
        row=total_row,
        column=6,
        value=f"=SUM(F{first_data_row}:F{last_data_row})",
    )
    f_sum.number_format = DUR_FMT
    # 休日労働
    g_sum = ws.cell(
        row=total_row,
        column=7,
        value=f"=SUM(G{first_data_row}:G{last_data_row})",
    )
    g_sum.number_format = DUR_FMT
    # 時間外
    h_sum = ws.cell(
        row=total_row,
        column=8,
        value=f"=SUM(H{first_data_row}:H{last_data_row})",
    )
    h_sum.number_format = DUR_FMT
    # 深夜
    i_sum = ws.cell(
        row=total_row,
        column=9,
        value=f"=SUM(I{first_data_row}:I{last_data_row})",
    )
    i_sum.number_format = DUR_FMT

    for col in range(1, len(HEADER) + 1):
        cell = ws.cell(row=total_row, column=col)
        cell.font = Font(bold=True)
        cell.fill = SUM_FILL
        cell.border = BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center")

    # 列幅
    widths = [9, 6, 8, 8, 8, 10, 12, 14, 12, 16, 11]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # 週開始列は補助列なので非表示に
    ws.column_dimensions["K"].hidden = True

    ws.row_dimensions[1].height = 22
    ws.freeze_panes = "A4"

    return first_data_row, last_data_row, total_row


def write_summary_sheet(wb, months_info):
    """サマリーシート(各月シートの合計行を参照)"""
    ws = wb.create_sheet(title="賃金計算期間", index=0)

    headers = ["賃金計算期間"] + [info["label"] for info in months_info]
    for col, h in enumerate(headers, 1):
        c = ws.cell(row=1, column=col, value=h)
        c.font = Font(bold=True)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.fill = HEADER_FILL
        c.border = BORDER

    # 各行の定義: (ラベル, 月シートでの列, 書式)
    rows_def = [
        ("労働日数", None, None),         # 特殊: 日数は COUNTIF で
        ("労働時間数", "F", DUR_FMT),
        ("休日労働時間数", "G", DUR_FMT),
        ("時間外労働時間数", "H", DUR_FMT),
        ("深夜労働時間数", "I", DUR_FMT),
    ]

    for r_i, (label_row, col_letter, fmt) in enumerate(rows_def, 2):
        lbl = ws.cell(row=r_i, column=1, value=label_row)
        lbl.font = Font(bold=True)
        lbl.fill = HEADER_FILL
        lbl.border = BORDER
        lbl.alignment = Alignment(horizontal="center", vertical="center")

        for c_i, info in enumerate(months_info, 2):
            sheet = info["label"]
            first = info["first_row"]
            last = info["last_row"]
            if label_row == "労働日数":
                formula = f"=COUNTIF('{sheet}'!F{first}:F{last},\">0\")&\"日\""
                cell = ws.cell(row=r_i, column=c_i, value=formula)
            else:
                formula = f"=SUM('{sheet}'!{col_letter}{first}:{col_letter}{last})"
                cell = ws.cell(row=r_i, column=c_i, value=formula)
                cell.number_format = fmt
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
        first_row, last_row, total_row = write_month_sheet(wb, label, year, month, records)
        months_info.append(
            {
                "label": label,
                "first_row": first_row,
                "last_row": last_row,
                "total_row": total_row,
            }
        )

    write_summary_sheet(wb, months_info)

    out = "出勤簿_R8.1-R8.3.xlsx"
    wb.save(out)
    print(f"Saved: {out}")


if __name__ == "__main__":
    main()
