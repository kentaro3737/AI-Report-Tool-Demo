"""Local sales import, validation, aggregation and Japanese PDF generation."""
from io import BytesIO
from decimal import Decimal, InvalidOperation
from pathlib import Path
import unicodedata

import pandas as pd
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.graphics.shapes import Drawing, Rect, Line, String

PDF_FONT_PATH = Path(__file__).resolve().parent / "assets" / "fonts" / "ipaexg.ttf"


def load_data(content: bytes, filename: str, sheet_name=0) -> pd.DataFrame:
    if len(content) > 20 * 1024 * 1024:
        raise ValueError("ファイルは20MB以下にしてください。")
    if filename.lower().endswith(".xlsx"):
        frame = pd.read_excel(BytesIO(content), sheet_name=sheet_name, engine="openpyxl")
    elif filename.lower().endswith(".csv"):
        for encoding in ("utf-8-sig", "cp932"):
            try:
                frame = pd.read_csv(BytesIO(content), encoding=encoding)
                break
            except UnicodeDecodeError:
                continue
        else:
            raise ValueError("CSVの文字コードはUTF-8またはShift_JISにしてください。")
    else:
        raise ValueError(".xlsx または .csv を選んでください。")
    if frame.empty:
        raise ValueError("データ行がありません。1行目を見出し、2行目以降を売上データにしてください。")
    if len(frame) > 100_000:
        raise ValueError("試作品では10万行以下のデータをご利用ください。")
    frame.columns = frame.columns.map(str)
    if frame.columns.duplicated().any():
        raise ValueError("列名が重複しています。列名を一意にしてください。")
    return frame


def _yen(value):
    text = unicodedata.normalize("NFKC", str(value)).strip()
    text = text.replace(",", "").replace("¥", "").replace("円", "").strip()
    try:
        amount = Decimal(text)
        if not amount.is_finite() or amount != amount.to_integral_value() or abs(amount) > 10**12:
            return None
        return int(amount)
    except (InvalidOperation, ValueError):
        return None


def prepare_data(frame: pd.DataFrame, date_column: str, amount_column: str) -> pd.DataFrame:
    if date_column == amount_column:
        raise ValueError("日付列と売上金額列は別々の列を選んでください。")
    # Bare numbers (including Excel serial dates) are deliberately rejected.
    def parse_date(value):
        if isinstance(value, (int, float)) or pd.isna(value):
            return pd.NaT
        try:
            parsed = pd.Timestamp(value)
            if parsed.tzinfo is not None or not 1900 <= parsed.year <= 2100:
                return pd.NaT
            return parsed
        except (ValueError, TypeError, OverflowError):
            return pd.NaT

    dates = frame[date_column].map(parse_date)
    amounts = frame[amount_column].map(_yen)
    invalid = dates.isna() | amounts.isna()
    if invalid.any():
        rows = ", ".join(str(i + 2) for i, bad in enumerate(invalid) if bad)[:200]
        raise ValueError(
            f"日付または金額が不正な行が{int(invalid.sum())}件あります（見出しを1行目とした行番号: {rows}）。"
            "日付は2026-01-31など、金額は1兆円以下の整数円で入力してください。空欄・小数・無限大は使用できません。"
            "不正な行は自動除外せず、集計を停止しています。"
        )
    result = pd.DataFrame({"日付": pd.to_datetime(dates), "売上金額": amounts.astype("int64")})
    result["月"] = result["日付"].dt.to_period("M")
    if result["月"].max().ordinal - result["月"].min().ordinal > 119:
        raise ValueError("試作品ではデータ期間を10年以内にしてください。")
    return result


def aggregate(data: pd.DataFrame) -> pd.DataFrame:
    monthly = data.groupby("月").agg(売上金額=("売上金額", "sum"), 件数=("売上金額", "size"))
    # Missing months remain missing, so March is never compared against January.
    months = pd.period_range(monthly.index.min(), monthly.index.max(), freq="M")
    # Nullable integers retain exact yen amounts even when a month is missing.
    return monthly.astype("Int64").reindex(months).rename_axis("月")


def comparison(monthly: pd.DataFrame, month: pd.Period):
    previous = month - 1
    if previous not in monthly.index or pd.isna(monthly.loc[previous, "売上金額"]):
        return None, "前月データなし"
    prior = int(monthly.loc[previous, "売上金額"])
    if prior == 0:
        return None, "前月売上が0円のため計算不可"
    rate = (int(monthly.loc[month, "売上金額"]) - prior) / abs(prior) * 100
    return rate, f"{rate:+.1f}%"


def summary_text(monthly: pd.DataFrame, month: pd.Period) -> str:
    row = monthly.loc[month]
    _, change = comparison(monthly, month)
    return (f"{month.year}年{month.month}月の売上は{int(row['売上金額']):,}円、"
            f"売上件数は{int(row['件数']):,}件でした。前月比は{change}です。")


def _pdf_font() -> str:
    """Embed the bundled font instead of relying on a PDF viewer's CJK fonts."""
    font = "IPAexGothic"
    # Check on every generation, even when ReportLab has cached a registered font.
    if not PDF_FONT_PATH.is_file():
        raise FileNotFoundError("日本語フォント assets/fonts/ipaexg.ttf がありません。フォントを含むフォルダ全体を配置してください。")
    if font not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont(font, str(PDF_FONT_PATH)))
    return font


def create_pdf(monthly: pd.DataFrame, month: pd.Period) -> bytes:
    font = _pdf_font()
    output = BytesIO()
    doc = SimpleDocTemplate(output, pagesize=(595.28, 841.89), rightMargin=42, leftMargin=42,
                            topMargin=40, bottomMargin=40, title=f"月次売上報告書 {month}", author="AI月次報告書ツール")
    body = ParagraphStyle("Japanese", fontName=font, fontSize=10, leading=17, textColor=colors.HexColor("#172b4d"))
    title = ParagraphStyle("Title", parent=body, fontSize=22, leading=30, spaceAfter=12)
    story = [Paragraph("月次売上報告書", title), Paragraph(f"対象月：{month.year}年{month.month}月", body), Spacer(1, 16),
             Paragraph(summary_text(monthly, month), body), Spacer(1, 12),
             Paragraph("集計方法：1行を1件として集計。通貨は日本円。返品等の負数も合算します。", body),
             Paragraph("前月比 =（当月売上 - 前月売上）÷ 前月売上の絶対値 × 100", body), Spacer(1, 16)]
    history = monthly.loc[:month].tail(12)
    story.append(Paragraph("月別売上の推移（対象月までの最大12か月）", body))
    chart = Drawing(505, 175)
    values = [0 if pd.isna(v) else int(v) for v in history["売上金額"]]
    low, high = min(0, min(values)), max(0, max(values))
    span = high - low or 1
    baseline = 35 + (0 - low) / span * 105
    chart.add(Line(40, baseline, 495, baseline, strokeColor=colors.HexColor("#cad4df")))
    for i, (period, value) in enumerate(zip(history.index, values)):
        x = 48 + i * (440 / len(values))
        height = value / span * 105
        if not pd.isna(history.loc[period, "売上金額"]):
            chart.add(Rect(x, min(baseline, baseline + height), 440 / len(values) * .65,
                           abs(height), fillColor=colors.HexColor("#167d8d"), strokeColor=None))
        chart.add(String(x, 17, str(period)[2:], fontName=font, fontSize=8))
    chart.add(String(40, 153, f"最大 {high:,} 円 / 最小 {low:,} 円", fontName=font, fontSize=9))
    story.extend([chart, Spacer(1, 8)])
    rows = [["月", "売上金額（円）", "件数", "前月比"]]
    for period, row in history.iterrows():
        if pd.isna(row["売上金額"]):
            rows.append([str(period), "データなし", "-", "-"])
        else:
            rows.append([str(period), f"{int(row['売上金額']):,}", f"{int(row['件数']):,}", comparison(monthly, period)[1]])
    table = Table(rows, colWidths=[75, 135, 65, 230], repeatRows=1)
    table.setStyle(TableStyle([
        ("FONTNAME", (0, 0), (-1, -1), font), ("FONTSIZE", (0, 0), (-1, -1), 9),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#e5f1f3")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f5f7fa")]),
        ("ALIGN", (1, 1), (2, -1), "RIGHT"), ("TOPPADDING", (0, 0), (-1, -1), 7),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
    ]))
    story.extend([table, Spacer(1, 14), Paragraph("データなしの月は0円とみなさず、前月比を計算しません。", body),
                  Paragraph("この報告書の文章は集計値に基づく定型文です。外部AI APIは使用していません。", body)])
    def footer(canvas, document):
        canvas.setFont(font, 8)
        canvas.drawRightString(553, 24, f"{document.page} ページ")
    doc.build(story, onFirstPage=footer, onLaterPages=footer)
    return output.getvalue()
