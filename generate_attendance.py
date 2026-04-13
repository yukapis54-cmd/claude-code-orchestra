# -*- coding: utf-8 -*-
"""
出勤簿生成スクリプト

賃金計算期間:
  R8.1 (2026年1月): 労働日数 25.0日 / 労働時間 163:00 / 時間外 6:30 / 深夜 0:00
  R8.2 (2026年2月): 労働日数 24.0日 / 労働時間 160:00 / 時間外 42:00 / 深夜 2:00
  R8.3 (2026年3月): 労働日数 25.0日 / 労働時間 175:30 / 時間外 46:00 / 深夜 0:00

条件:
  - 始業は 7:00
  - 終業時刻は各日の労働時間に合わせて設定
  - 日曜・祝日は出勤なし(R8.2のみ祝日出勤あり ※合計日数を合わせるため)
  - 法定休憩時間(6時間超:45分、8時間超:60分)を考慮
  - 残業は 月曜〜土曜 で週40時間を超えた分
  - 深夜労働は 22:00〜翌5:00
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
    # 2/11 建国記念の日、2/23 天皇誕生日 は本出勤簿では出勤扱い(合計日数を合わせるため)
    date(2026, 3, 20): "春分の日",
}

WEEKDAY_JP = ["月", "火", "水", "木", "金", "土", "日"]


# ---------------------------------------------------------------------------
# 各月のスケジュール定義
#  (週の労働日ごとに 所定労働時間(分) と 終業前の付加労働(分) を設定)
# ---------------------------------------------------------------------------
def m(h, mm=0):
    return h * 60 + mm


# schedule[月] = [(対象日のリスト, 1日あたりの労働時間分), ...]
SCHEDULE = {
    1: [
        # W1 (1/2, 1/3): 6:30/日
        ([date(2026, 1, 2), date(2026, 1, 3)], m(6, 30)),
        # W2 (1/5-1/10): 7:45/日  → 週46:30 (法定超過 6:30)
        ([date(2026, 1, d) for d in (5, 6, 7, 8, 9, 10)], m(7, 45)),
        # W3 (1/13-1/17): 6:30/日
        ([date(2026, 1, d) for d in (13, 14, 15, 16, 17)], m(6, 30)),
        # W4 (1/19-1/24): 6:00/日
        ([date(2026, 1, d) for d in (19, 20, 21, 22, 23, 24)], m(6, 0)),
        # W5 (1/26-1/31): 5:50/日
        ([date(2026, 1, d) for d in (26, 27, 28, 29, 30, 31)], m(5, 50)),
    ],
    2: [
        # W1 (2/2-2/7): 3:10/日
        ([date(2026, 2, d) for d in (2, 3, 4, 5, 6, 7)], m(3, 10)),
        # W2 (2/9-2/14): 3:10/日
        ([date(2026, 2, d) for d in (9, 10, 11, 12, 13, 14)], m(3, 10)),
        # W3 (2/16-2/21): 8:20/日  → 週50:00 (超過 10:00)
        ([date(2026, 2, d) for d in (16, 17, 18, 19, 20, 21)], m(8, 20)),
        # W4 (2/23-2/28): 可変
        #   2/23(月): 16:00 (7:00-24:00, 深夜 22-24 の 2時間)
        #   2/24-2/27: 11:00/日
        #   2/28(土): 12:00
        ([date(2026, 2, 23)], m(16, 0)),
        ([date(2026, 2, d) for d in (24, 25, 26, 27)], m(11, 0)),
        ([date(2026, 2, 28)], m(12, 0)),
    ],
    3: [
        # W1 (3/2-3/7): 11:00/日 → 週66:00 (超過 26:00)
        ([date(2026, 3, d) for d in (2, 3, 4, 5, 6, 7)], m(11, 0)),
        # W2 (3/9-3/14): 4:00/日
        ([date(2026, 3, d) for d in (9, 10, 11, 12, 13, 14)], m(4, 0)),
        # W3 (3/16-3/21 ※3/20祝日): 12:00/日 × 5日 → 週60:00 (超過 20:00)
        ([date(2026, 3, d) for d in (16, 17, 18, 19, 21)], m(12, 0)),
        # W4 (3/23-3/28): 3:00/日
        ([date(2026, 3, d) for d in (23, 24, 25, 26, 27, 28)], m(3, 0)),
        # W5 (3/30-3/31): 3:45/日
        ([date(2026, 3, d) for d in (30, 31)], m(3, 45)),
    ],
}


# ---------------------------------------------------------------------------
# 計算ユーティリティ
# ---------------------------------------------------------------------------
def legal_break(work_minutes: int) -> int:
    """労働基準法34条 に基づく休憩時間(分)"""
    if work_minutes > 8 * 60:
        return 60
    if work_minutes > 6 * 60:
        return 45
    return 0


def fmt_hhmm(total_minutes: int) -> str:
    if total_minutes == 0:
        return "0:00"
    h, mm = divmod(total_minutes, 60)
    return f"{h}:{mm:02d}"


def night_minutes(start_dt: datetime, end_dt: datetime) -> int:
    """深夜(22:00〜翌5:00)に含まれる分を計算"""
    total = 0
    cur = start_dt
    while cur < end_dt:
        next_boundary = cur + timedelta(minutes=1)
        # その分の始点の時刻で深夜判定
        t = cur.time()
        if t >= time(22, 0) or t < time(5, 0):
            total += 1
        cur = next_boundary
    return total


def build_day_rows(target_date: date, work_min: int):
    """1日分のレコードを生成"""
    brk = legal_break(work_min)
    start_dt = datetime.combine(target_date, time(7, 0))
    end_dt = start_dt + timedelta(minutes=work_min + brk)

    # 深夜労働(休憩を除いた実労働時間のうち 22:00-5:00 の部分)
    # 休憩を取る時刻を「労働が6時間に達した時点」で 45分/60分 連続で取得すると仮定
    # 深夜帯は start 以降の連続帯で判定(休憩時間は労働ではないので除外)
    if brk == 0:
        work_intervals = [(start_dt, end_dt)]
    else:
        # 6時間勤務後に休憩を挿入
        six_h = start_dt + timedelta(hours=6)
        # 休憩がその日の勤務時間内に収まるように
        if six_h + timedelta(minutes=brk) <= end_dt:
            work_intervals = [
                (start_dt, six_h),
                (six_h + timedelta(minutes=brk), end_dt),
            ]
        else:
            # 想定されない(break 発生=6時間超労働)だが念のため
            work_intervals = [(start_dt, end_dt - timedelta(minutes=brk))]

    night_min = sum(night_minutes(s, e) for s, e in work_intervals)
    return {
        "date": target_date,
        "weekday": WEEKDAY_JP[target_date.weekday()],
        "start": start_dt,
        "end": end_dt,
        "break_min": brk,
        "work_min": work_min,
        "night_min": night_min,
    }


def build_month_records(month: int):
    records = []
    for dates, work_min in SCHEDULE[month]:
        for d in dates:
            records.append(build_day_rows(d, work_min))
    return records


def calc_weekly_overtime(records):
    """週(月曜始まり)ごとに40時間を超えた分を残業として算出し、
    その週の出勤日に按分して返す。
    返り値: dict{date: overtime_minutes}"""
    from collections import defaultdict

    weeks = defaultdict(list)
    for r in records:
        d = r["date"]
        monday = d - timedelta(days=d.weekday())
        weeks[monday].append(r)

    overtime_per_day = {}
    for monday, days in weeks.items():
        total = sum(r["work_min"] for r in days)
        ot = max(0, total - 40 * 60)
        # 残業分は後ろの日から配分(遅い日に残業があったイメージ)
        remaining = ot
        # 日付でソート
        days_sorted = sorted(days, key=lambda r: r["date"], reverse=True)
        for r in days_sorted:
            take = min(remaining, r["work_min"])
            overtime_per_day[r["date"]] = take
            remaining -= take
            if remaining <= 0:
                break
        for r in days_sorted:
            overtime_per_day.setdefault(r["date"], 0)
    return overtime_per_day


# ---------------------------------------------------------------------------
# Excel 出力
# ---------------------------------------------------------------------------
HEADER = [
    "日付",
    "曜日",
    "始業",
    "終業",
    "休憩",
    "労働時間",
    "休日労働時間",
    "時間外労働時間",
    "深夜労働時間",
    "備考",
]

THIN = Side(style="thin", color="000000")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HEADER_FILL = PatternFill("solid", fgColor="D9E1F2")
SUN_FILL = PatternFill("solid", fgColor="FFE2E2")
HOL_FILL = PatternFill("solid", fgColor="FFF2CC")
SUM_FILL = PatternFill("solid", fgColor="FFFF66")


def write_month_sheet(wb, month_label, year, month, records, ot_by_date):
    ws = wb.create_sheet(title=month_label)

    ws["A1"] = f"出勤簿  {month_label} ({year}年{month}月)"
    ws["A1"].font = Font(bold=True, size=14)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(HEADER))

    # ヘッダー行
    for col, title in enumerate(HEADER, 1):
        c = ws.cell(row=3, column=col, value=title)
        c.font = Font(bold=True)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.fill = HEADER_FILL
        c.border = BORDER

    # 日付ごとの行(月内全日)
    days_in_month = (date(year + (month == 12), (month % 12) + 1, 1) - date(year, month, 1)).days
    record_map = {r["date"]: r for r in records}

    row = 4
    sum_work = 0
    sum_ot = 0
    sum_night = 0
    sum_days = 0
    for d in range(1, days_in_month + 1):
        dt = date(year, month, d)
        is_sun = dt.weekday() == 6
        is_hol = dt in JP_HOLIDAYS
        rec = record_map.get(dt)

        ws.cell(row=row, column=1, value=f"{month}/{d}")
        ws.cell(row=row, column=2, value=WEEKDAY_JP[dt.weekday()])
        if rec:
            ws.cell(row=row, column=3, value=rec["start"].strftime("%H:%M"))
            ws.cell(row=row, column=4, value=rec["end"].strftime("%H:%M"))
            ws.cell(row=row, column=5, value=fmt_hhmm(rec["break_min"]))
            ws.cell(row=row, column=6, value=fmt_hhmm(rec["work_min"]))
            ws.cell(row=row, column=7, value="")  # 休日労働なし
            ws.cell(row=row, column=8, value=fmt_hhmm(ot_by_date.get(dt, 0)))
            ws.cell(row=row, column=9, value=fmt_hhmm(rec["night_min"]))
            ws.cell(row=row, column=10, value=JP_HOLIDAYS.get(dt, ""))
            sum_work += rec["work_min"]
            sum_ot += ot_by_date.get(dt, 0)
            sum_night += rec["night_min"]
            sum_days += 1
        else:
            ws.cell(row=row, column=10, value=("日曜" if is_sun else JP_HOLIDAYS.get(dt, "休")))

        # 書式
        for col in range(1, len(HEADER) + 1):
            cell = ws.cell(row=row, column=col)
            cell.border = BORDER
            cell.alignment = Alignment(horizontal="center", vertical="center")
            if is_sun:
                cell.fill = SUN_FILL
            elif is_hol and not rec:
                cell.fill = HOL_FILL
        row += 1

    # 合計行
    total_row = row
    ws.cell(row=total_row, column=1, value="合計")
    ws.cell(row=total_row, column=2, value=f"{sum_days:.1f}日")
    ws.cell(row=total_row, column=6, value=fmt_hhmm(sum_work))
    ws.cell(row=total_row, column=7, value="")
    ws.cell(row=total_row, column=8, value=fmt_hhmm(sum_ot))
    ws.cell(row=total_row, column=9, value=fmt_hhmm(sum_night))
    for col in range(1, len(HEADER) + 1):
        cell = ws.cell(row=total_row, column=col)
        cell.font = Font(bold=True)
        cell.fill = SUM_FILL
        cell.border = BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center")

    # 列幅
    widths = [8, 6, 8, 8, 8, 10, 12, 14, 12, 14]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.row_dimensions[1].height = 22
    ws.freeze_panes = "A4"

    return sum_days, sum_work, sum_ot, sum_night


def write_summary_sheet(wb, summary):
    ws = wb.create_sheet(title="賃金計算期間", index=0)
    headers = ["賃金計算期間", "R8.1", "R8.2", "R8.3"]
    rows = [
        ["労働日数"] + [f"{s['days']:.1f}日" for s in summary],
        ["労働時間数"] + [fmt_hhmm(s["work"]) for s in summary],
        ["休日労働時間数"] + ["" for _ in summary],
        ["時間外労働時間数"] + [fmt_hhmm(s["ot"]) for s in summary],
        ["深夜労働時間数"] + [fmt_hhmm(s["night"]) for s in summary],
    ]

    for col, h in enumerate(headers, 1):
        c = ws.cell(row=1, column=col, value=h)
        c.font = Font(bold=True)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.fill = HEADER_FILL
        c.border = BORDER

    for r_i, row in enumerate(rows, 2):
        for c_i, val in enumerate(row, 1):
            cell = ws.cell(row=r_i, column=c_i, value=val)
            cell.border = BORDER
            cell.alignment = Alignment(horizontal="center", vertical="center")
            if c_i == 1:
                cell.font = Font(bold=True)
                cell.fill = HEADER_FILL
            else:
                cell.fill = SUM_FILL

    for i, w in enumerate([22, 12, 12, 12], 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def main():
    wb = Workbook()
    wb.remove(wb.active)

    summary = []
    for label, year, month in [("R8.1", 2026, 1), ("R8.2", 2026, 2), ("R8.3", 2026, 3)]:
        records = build_month_records(month)
        ot = calc_weekly_overtime(records)
        days, work, ot_total, night = write_month_sheet(wb, label, year, month, records, ot)
        summary.append({"days": days, "work": work, "ot": ot_total, "night": night})

    write_summary_sheet(wb, summary)

    out = "出勤簿_R8.1-R8.3.xlsx"
    wb.save(out)
    print(f"Saved: {out}")
    for label, s in zip(["R8.1", "R8.2", "R8.3"], summary):
        print(
            f"  {label}: 日数={s['days']:.1f}  労働={fmt_hhmm(s['work'])}  "
            f"時間外={fmt_hhmm(s['ot'])}  深夜={fmt_hhmm(s['night'])}"
        )


if __name__ == "__main__":
    main()
