"""
GraphRAG 테스트 PPT 생성기
- 도메인별 12~15 슬라이드
- 각 PPT에 표, 차트(bar/pie/line), 다이어그램(플로우차트/조직도/프로세스) 포함
"""
from __future__ import annotations

import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

from pathlib import Path
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR_TYPE
from pptx.chart.data import CategoryChartData, ChartData

OUT_DIR = Path(__file__).parent / "documents"

# ═══════════════════════════════════════════════════════════════
#  COLOR PALETTE
# ═══════════════════════════════════════════════════════════════
C_BLUE    = RGBColor(0x2E, 0x86, 0xC1)
C_DKBLUE  = RGBColor(0x1A, 0x52, 0x76)
C_LTBLUE  = RGBColor(0xD6, 0xEA, 0xF8)
C_GREEN   = RGBColor(0x27, 0xAE, 0x60)
C_LTGREEN = RGBColor(0xE8, 0xF8, 0xF5)
C_RED     = RGBColor(0xE7, 0x4C, 0x3C)
C_ORANGE  = RGBColor(0xF3, 0x9C, 0x12)
C_PURPLE  = RGBColor(0x8E, 0x44, 0xAD)
C_LTPURP  = RGBColor(0xF5, 0xEE, 0xF8)
C_DARK    = RGBColor(0x2C, 0x3E, 0x50)
C_GRAY    = RGBColor(0x7F, 0x8C, 0x8D)
C_LTGRAY  = RGBColor(0xEC, 0xF0, 0xF1)
C_WHITE   = RGBColor(0xFF, 0xFF, 0xFF)
C_BLACK   = RGBColor(0x00, 0x00, 0x00)


# ═══════════════════════════════════════════════════════════════
#  HELPERS
# ═══════════════════════════════════════════════════════════════

def new_prs():
    """Create a 16:9 presentation."""
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    return prs


def add_blank(prs):
    """Add a blank slide."""
    layout = prs.slide_layouts[6]  # blank
    return prs.slides.add_slide(layout)


def add_textbox(slide, left, top, width, height, text, font_size=14,
                bold=False, color=C_DARK, alignment=PP_ALIGN.LEFT, font_name="맑은 고딕"):
    """Add a text box."""
    txBox = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(font_size)
    p.font.bold = bold
    p.font.color.rgb = color
    p.font.name = font_name
    p.alignment = alignment
    return txBox


def add_multiline_textbox(slide, left, top, width, height, lines, font_size=13,
                          color=C_DARK, line_spacing=1.3, font_name="맑은 고딕"):
    """Add text box with multiple lines."""
    txBox = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = txBox.text_frame
    tf.word_wrap = True
    for i, line in enumerate(lines):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.text = line
        p.font.size = Pt(font_size)
        p.font.color.rgb = color
        p.font.name = font_name
        p.space_after = Pt(font_size * (line_spacing - 1))
    return txBox


def add_title_slide(prs, title, subtitle, bg_color=C_DKBLUE):
    """Add a styled title slide."""
    slide = add_blank(prs)
    # Background
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    bg.fill.solid()
    bg.fill.fore_color.rgb = bg_color
    bg.line.fill.background()
    # Accent bar
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, Inches(5.0), prs.slide_width, Inches(0.08))
    bar.fill.solid()
    bar.fill.fore_color.rgb = C_ORANGE
    bar.line.fill.background()
    # Title
    add_textbox(slide, 1.0, 2.0, 11, 1.5, title, font_size=40, bold=True, color=C_WHITE, alignment=PP_ALIGN.CENTER)
    # Subtitle
    add_textbox(slide, 1.0, 3.8, 11, 1.0, subtitle, font_size=20, color=C_LTGRAY, alignment=PP_ALIGN.CENTER)
    return slide


def add_section_slide(prs, section_num, section_title, bg_color=C_BLUE):
    """Add a section divider slide."""
    slide = add_blank(prs)
    # Left bar
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(0.4), prs.slide_height)
    bar.fill.solid()
    bar.fill.fore_color.rgb = bg_color
    bar.line.fill.background()
    # Number circle
    circle = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(1.5), Inches(2.5), Inches(1.2), Inches(1.2))
    circle.fill.solid()
    circle.fill.fore_color.rgb = bg_color
    circle.line.fill.background()
    tf = circle.text_frame
    tf.paragraphs[0].text = str(section_num)
    tf.paragraphs[0].font.size = Pt(36)
    tf.paragraphs[0].font.bold = True
    tf.paragraphs[0].font.color.rgb = C_WHITE
    tf.paragraphs[0].alignment = PP_ALIGN.CENTER
    tf.word_wrap = False
    circle.text_frame.paragraphs[0].font.name = "맑은 고딕"
    # Title
    add_textbox(slide, 3.2, 2.6, 8, 1.0, section_title, font_size=32, bold=True, color=C_DARK)
    return slide


def add_kpi_boxes(slide, kpis, top=1.8):
    """Add KPI metric boxes. kpis = [(label, value, color), ...]"""
    n = len(kpis)
    total_w = 11.0
    box_w = total_w / n - 0.3
    start_x = (13.333 - (box_w * n + 0.3 * (n - 1))) / 2
    for i, (label, value, color) in enumerate(kpis):
        x = start_x + i * (box_w + 0.3)
        shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(top), Inches(box_w), Inches(1.6))
        shape.fill.solid()
        shape.fill.fore_color.rgb = color
        shape.line.fill.background()
        tf = shape.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = value
        p.font.size = Pt(28)
        p.font.bold = True
        p.font.color.rgb = C_WHITE
        p.font.name = "맑은 고딕"
        p.alignment = PP_ALIGN.CENTER
        p2 = tf.add_paragraph()
        p2.text = label
        p2.font.size = Pt(13)
        p2.font.color.rgb = C_WHITE
        p2.font.name = "맑은 고딕"
        p2.alignment = PP_ALIGN.CENTER


def add_table(slide, left, top, width, height, headers, rows, header_color=C_DKBLUE):
    """Add a styled table."""
    num_rows = len(rows) + 1
    num_cols = len(headers)
    tbl_shape = slide.shapes.add_table(num_rows, num_cols, Inches(left), Inches(top), Inches(width), Inches(height))
    tbl = tbl_shape.table

    col_w = width / num_cols
    for i in range(num_cols):
        tbl.columns[i].width = Inches(col_w)

    # Header row
    for j, h in enumerate(headers):
        cell = tbl.cell(0, j)
        cell.text = h
        cell.fill.solid()
        cell.fill.fore_color.rgb = header_color
        for p in cell.text_frame.paragraphs:
            p.font.size = Pt(12)
            p.font.bold = True
            p.font.color.rgb = C_WHITE
            p.font.name = "맑은 고딕"
            p.alignment = PP_ALIGN.CENTER

    # Data rows
    for i, row in enumerate(rows):
        for j, val in enumerate(row):
            cell = tbl.cell(i + 1, j)
            cell.text = str(val)
            if i % 2 == 0:
                cell.fill.solid()
                cell.fill.fore_color.rgb = C_LTGRAY
            for p in cell.text_frame.paragraphs:
                p.font.size = Pt(11)
                p.font.name = "맑은 고딕"
                p.alignment = PP_ALIGN.CENTER

    return tbl_shape


def add_chart(slide, chart_type, left, top, width, height, categories, series_list, title=None):
    """Add a chart. series_list = [(name, values), ...]"""
    chart_data = CategoryChartData()
    chart_data.categories = categories
    for name, values in series_list:
        chart_data.add_series(name, values)

    chart_frame = slide.shapes.add_chart(
        chart_type, Inches(left), Inches(top), Inches(width), Inches(height), chart_data
    )
    chart = chart_frame.chart
    chart.has_legend = len(series_list) > 1
    if chart.has_legend:
        chart.legend.position = XL_LEGEND_POSITION.BOTTOM
        chart.legend.include_in_layout = False
        chart.legend.font.size = Pt(10)
        chart.legend.font.name = "맑은 고딕"

    if title:
        chart.has_title = True
        chart.chart_title.text_frame.paragraphs[0].text = title
        chart.chart_title.text_frame.paragraphs[0].font.size = Pt(13)
        chart.chart_title.text_frame.paragraphs[0].font.name = "맑은 고딕"

    # Style plots
    plot = chart.plots[0]
    if chart_type in (XL_CHART_TYPE.PIE, XL_CHART_TYPE.DOUGHNUT):
        plot.has_data_labels = True
        data_labels = plot.data_labels
        data_labels.font.size = Pt(10)
        data_labels.font.name = "맑은 고딕"
        data_labels.number_format = '0.0"%"'
        data_labels.show_percentage = True
        data_labels.show_category_name = True

    return chart_frame


def add_flowchart(slide, steps, left=1.0, top=2.0, box_w=2.0, box_h=0.9, gap=0.4,
                  colors=None, direction="horizontal"):
    """Add a flowchart with connected boxes."""
    if colors is None:
        colors = [C_BLUE, C_GREEN, C_ORANGE, C_PURPLE, C_RED, C_DKBLUE, C_GRAY]

    shapes = []
    for i, step_text in enumerate(steps):
        color = colors[i % len(colors)]
        if direction == "horizontal":
            x = left + i * (box_w + gap + 0.5)
            y = top
        else:
            x = left
            y = top + i * (box_h + gap + 0.3)

        box = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(box_w), Inches(box_h)
        )
        box.fill.solid()
        box.fill.fore_color.rgb = color
        box.line.fill.background()
        tf = box.text_frame
        tf.word_wrap = True
        tf.paragraphs[0].alignment = PP_ALIGN.CENTER
        p = tf.paragraphs[0]
        p.text = step_text
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = C_WHITE
        p.font.name = "맑은 고딕"
        shapes.append(box)

        # Arrow between boxes
        if i > 0:
            if direction == "horizontal":
                ax = Inches(x - gap - 0.25)
                ay = Inches(y + box_h / 2)
                ax2 = Inches(x - 0.05)
                ay2 = ay
            else:
                ax = Inches(x + box_w / 2)
                ay = Inches(y - gap + 0.05)
                ax2 = ax
                ay2 = Inches(y - 0.05)

            arrow = slide.shapes.add_shape(
                MSO_SHAPE.RIGHT_ARROW if direction == "horizontal" else MSO_SHAPE.DOWN_ARROW,
                ax, ay, Inches(0.35) if direction == "horizontal" else Inches(0.35),
                Inches(0.25) if direction == "horizontal" else Inches(0.3)
            )
            arrow.fill.solid()
            arrow.fill.fore_color.rgb = C_GRAY
            arrow.line.fill.background()

    return shapes


def add_org_chart(slide, org_data, left=1.5, top=1.5):
    """Add an org chart. org_data = [(level, name, title, color), ...]
    level 0 = top, 1 = direct reports, 2 = sub-reports"""
    level_positions = {}
    level_counts = {}
    for level, _, _, _ in org_data:
        level_counts[level] = level_counts.get(level, 0) + 1

    level_idx = {}
    for level, name, title, color in org_data:
        idx = level_idx.get(level, 0)
        count = level_counts[level]

        box_w = 2.2
        box_h = 0.85
        total_w = count * box_w + (count - 1) * 0.3
        start_x = left + (10.333 - total_w) / 2
        x = start_x + idx * (box_w + 0.3)
        y = top + level * 1.6

        box = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(box_w), Inches(box_h)
        )
        box.fill.solid()
        box.fill.fore_color.rgb = color
        box.line.fill.background()
        tf = box.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = name
        p.font.size = Pt(12)
        p.font.bold = True
        p.font.color.rgb = C_WHITE
        p.font.name = "맑은 고딕"
        p.alignment = PP_ALIGN.CENTER
        p2 = tf.add_paragraph()
        p2.text = title
        p2.font.size = Pt(9)
        p2.font.color.rgb = RGBColor(0xDD, 0xDD, 0xDD)
        p2.font.name = "맑은 고딕"
        p2.alignment = PP_ALIGN.CENTER

        level_idx[level] = idx + 1


def add_matrix_diagram(slide, quadrants, title="", left=2.0, top=1.8, size=4.5):
    """Add a 2x2 matrix diagram. quadrants = [(label, items_text, color), ...] (TL, TR, BL, BR)"""
    half = size / 2
    gap = 0.15
    positions = [
        (left, top),
        (left + half + gap, top),
        (left, top + half + gap),
        (left + half + gap, top + half + gap),
    ]
    for i, (label, items, color) in enumerate(quadrants):
        x, y = positions[i]
        box = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(half), Inches(half)
        )
        box.fill.solid()
        box.fill.fore_color.rgb = color
        box.line.fill.background()
        tf = box.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = label
        p.font.size = Pt(14)
        p.font.bold = True
        p.font.color.rgb = C_WHITE
        p.font.name = "맑은 고딕"
        p.alignment = PP_ALIGN.CENTER
        p2 = tf.add_paragraph()
        p2.text = items
        p2.font.size = Pt(10)
        p2.font.color.rgb = C_WHITE
        p2.font.name = "맑은 고딕"
        p2.alignment = PP_ALIGN.CENTER


def add_timeline(slide, events, left=0.8, top=2.5, total_width=11.5):
    """Add a horizontal timeline. events = [(date, label, color), ...]"""
    n = len(events)
    # Horizontal line
    line = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(left), Inches(top + 0.4), Inches(total_width), Inches(0.06)
    )
    line.fill.solid()
    line.fill.fore_color.rgb = C_GRAY
    line.line.fill.background()

    seg_w = total_width / (n - 1) if n > 1 else total_width
    for i, (date, label, color) in enumerate(events):
        x = left + i * seg_w if n > 1 else left + total_width / 2
        # Dot
        dot = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x - 0.12), Inches(top + 0.28), Inches(0.3), Inches(0.3))
        dot.fill.solid()
        dot.fill.fore_color.rgb = color
        dot.line.fill.background()
        # Date above
        add_textbox(slide, x - 0.6, top - 0.5, 1.3, 0.4, date, font_size=10, bold=True, color=color, alignment=PP_ALIGN.CENTER)
        # Label below
        add_textbox(slide, x - 0.8, top + 0.8, 1.8, 0.6, label, font_size=9, color=C_DARK, alignment=PP_ALIGN.CENTER)


def add_slide_header(slide, title, subtitle=None):
    """Add a consistent header to content slides."""
    # Top bar
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(0.06))
    bar.fill.solid()
    bar.fill.fore_color.rgb = C_BLUE
    bar.line.fill.background()
    add_textbox(slide, 0.5, 0.25, 12, 0.6, title, font_size=24, bold=True, color=C_DKBLUE)
    if subtitle:
        add_textbox(slide, 0.5, 0.85, 12, 0.4, subtitle, font_size=13, color=C_GRAY)


def add_page_number(slide, num, total):
    add_textbox(slide, 12.0, 7.0, 1.0, 0.4, f"{num}/{total}", font_size=9, color=C_GRAY, alignment=PP_ALIGN.RIGHT)


# ═══════════════════════════════════════════════════════════════
#  DOMAIN: E-COMMERCE
# ═══════════════════════════════════════════════════════════════

def gen_ecommerce_ppt():
    prs = new_prs()
    TOTAL = 14

    # 1. Title
    add_title_slide(prs, "이커머스 사업 현황 보고서", "2026년 1분기 | 주식회사 샘플쇼핑\n경영기획팀")

    # 2. KPI Overview
    slide = add_blank(prs)
    add_slide_header(slide, "1분기 핵심 성과 지표 (KPI)")
    add_kpi_boxes(slide, [
        ("총 매출액", "45.2억원", C_BLUE),
        ("전년비 성장률", "+23%", C_GREEN),
        ("MAU", "285만명", C_PURPLE),
        ("주문 건수", "482,000건", C_ORANGE),
        ("객단가", "93,800원", C_RED),
    ], top=2.0)
    add_multiline_textbox(slide, 1.0, 4.5, 11, 2.5, [
        "- 1분기 매출은 전년 동기(36.7억원) 대비 23% 성장하여 사상 최고치 달성",
        "- 모바일 채널 비중 65% 돌파 (전년 52%), 앱 다운로드 150만 건 돌파",
        "- 신규 회원 가입 42만명 (+35%), 재구매율 42%(+5%p) 기록",
        "- 3월 봄 시즌 프로모션 효과로 월별 최고 매출 17.2억원 달성",
    ], font_size=13)
    add_page_number(slide, 2, TOTAL)

    # 3. Monthly Revenue Chart
    slide = add_blank(prs)
    add_slide_header(slide, "월별 매출 추이", "2025.01 ~ 2026.03 (단위: 억원)")
    add_chart(slide, XL_CHART_TYPE.LINE, 0.5, 1.5, 7.5, 5.5,
        ["1월", "2월", "3월", "4월", "5월", "6월", "7월", "8월", "9월", "10월", "11월", "12월", "1월", "2월", "3월"],
        [
            ("2025년", (9.5, 8.8, 10.2, 11.0, 10.5, 12.3, 11.8, 9.2, 10.8, 11.5, 13.0, 15.2, None, None, None)),
            ("2026년", (None, None, None, None, None, None, None, None, None, None, None, None, 13.8, 14.2, 17.2)),
        ],
        title="월별 매출 추이 (억원)")
    add_multiline_textbox(slide, 8.5, 1.8, 4.3, 5.0, [
        "[분석 요약]",
        "",
        "1. 12월 연말 특수로 최고 매출 15.2억 기록",
        "2. 2026년 1분기 역대 최고 분기 매출",
        "3. 3월 봄 시즌 프로모션 효과 극대화",
        "",
        "[성장 요인]",
        "- AI 추천 엔진 도입 (2025.09)",
        "- 당일배송 서비스 확대",
        "- 멤버십 프로그램 개편",
    ], font_size=11)
    add_page_number(slide, 3, TOTAL)

    # 4. Category Breakdown (pie)
    slide = add_blank(prs)
    add_slide_header(slide, "카테고리별 매출 구성", "2026년 1분기 기준")
    add_chart(slide, XL_CHART_TYPE.PIE, 0.5, 1.5, 6.0, 5.5,
        ["전자기기", "패션/의류", "식품", "생활용품", "도서/문구", "기타"],
        [("매출비중", (40.9, 27.2, 17.9, 8.0, 3.5, 2.5))],
        title="카테고리별 매출 비중")
    add_table(slide, 7.0, 1.8, 5.8, 4.0,
        ["카테고리", "매출(억원)", "전년비", "마진율"],
        [
            ["전자기기", "18.5", "+35%", "12%"],
            ["패션/의류", "12.3", "+15%", "38%"],
            ["식품", "8.1", "+18%", "22%"],
            ["생활용품", "3.6", "+12%", "28%"],
            ["도서/문구", "1.6", "+8%", "25%"],
            ["기타", "1.1", "+5%", "15%"],
        ])
    add_page_number(slide, 4, TOTAL)

    # 5. Top Products Table
    slide = add_blank(prs)
    add_slide_header(slide, "인기 상품 TOP 10", "2026년 1분기 판매 수량 기준")
    add_table(slide, 0.5, 1.5, 12.3, 5.2,
        ["순위", "상품명", "카테고리", "판매가", "판매 수량", "매출액", "리뷰 평점"],
        [
            ["1", "Galaxy S26 Ultra", "전자기기", "1,199,000원", "12,500대", "149.9억", "4.8"],
            ["2", "ProBook X15 노트북", "전자기기", "1,890,000원", "4,200대", "79.4억", "4.6"],
            ["3", "에어프라이어 XL", "생활용품", "189,000원", "15,800대", "29.9억", "4.7"],
            ["4", "봄 린넨 셔츠 (남성)", "패션", "49,000원", "28,300장", "13.9억", "4.5"],
            ["5", "유기농 그래놀라 세트", "식품", "32,000원", "32,100개", "10.3억", "4.9"],
            ["6", "블루투스 이어폰 Pro", "전자기기", "299,000원", "8,900개", "26.6억", "4.4"],
            ["7", "플라워 원피스 (여성)", "패션", "79,000원", "18,500장", "14.6억", "4.6"],
            ["8", "프리미엄 커피 원두 1kg", "식품", "28,000원", "22,400개", "6.3억", "4.8"],
            ["9", "스마트워치 Band 5", "전자기기", "349,000원", "5,600개", "19.5억", "4.3"],
            ["10", "초슬림 보조배터리", "전자기기", "39,000원", "41,200개", "16.1억", "4.5"],
        ])
    add_page_number(slide, 5, TOTAL)

    # 6. Customer Segmentation (bar chart)
    slide = add_blank(prs)
    add_slide_header(slide, "고객 세분화 분석", "연령대 x 구매 빈도 기준")
    add_chart(slide, XL_CHART_TYPE.COLUMN_CLUSTERED, 0.5, 1.5, 7.0, 5.5,
        ["10대", "20대", "30대", "40대", "50대", "60대+"],
        [
            ("신규 고객", (8, 35, 28, 18, 9, 2)),
            ("재구매 고객", (3, 22, 42, 38, 20, 8)),
            ("VIP 고객", (0, 5, 18, 25, 15, 5)),
        ],
        title="연령대별 고객 분포 (천명)")
    add_multiline_textbox(slide, 8.0, 1.8, 4.8, 5.0, [
        "[고객 세그먼트 정의]",
        "",
        "신규 고객:",
        "  최근 3개월 내 첫 구매",
        "",
        "재구매 고객:",
        "  2회 이상 구매, 월 1회 미만",
        "",
        "VIP 고객:",
        "  월 2회 이상 구매 또는",
        "  분기 30만원 이상 지출",
        "",
        "[인사이트]",
        "30~40대 재구매/VIP 비중 높음",
        "20대 신규 유입 최대 → 전환 필요",
    ], font_size=11)
    add_page_number(slide, 6, TOTAL)

    # 7. Order Process Flowchart
    slide = add_blank(prs)
    add_slide_header(slide, "주문 처리 프로세스", "고객 주문 ~ 배송 완료까지 전체 흐름")
    add_flowchart(slide, [
        "주문 접수", "결제 확인", "재고 확인", "상품 포장", "배송 출발", "배송 완료"
    ], left=0.5, top=2.2, box_w=1.7, box_h=0.85, gap=0.15)
    add_multiline_textbox(slide, 0.5, 3.8, 12, 3.0, [
        "[단계별 SLA]",
        "  주문접수 → 결제확인: 즉시 (자동)",
        "  결제확인 → 재고확인: 5분 이내",
        "  재고확인 → 상품포장: 2시간 이내 (오후2시 이전 주문)",
        "  상품포장 → 배송출발: 당일 (수도권) / 익일 (지방)",
        "  배송출발 → 배송완료: 1~2일 (수도권) / 2~3일 (지방)",
        "",
        "[자동화 현황]",
        "  주문접수~재고확인: 100% 자동 | 포장: 60% 자동(로봇) | 배송추적: 실시간 API 연동"
    ], font_size=12)
    add_page_number(slide, 7, TOTAL)

    # 8. Customer Journey Map (matrix)
    slide = add_blank(prs)
    add_slide_header(slide, "고객 여정 맵 (Customer Journey)", "단계별 접점 및 개선 포인트")
    add_matrix_diagram(slide, [
        ("인지 (Awareness)", "SNS 광고\n검색엔진\n인플루언서\n지인 추천", C_BLUE),
        ("탐색 (Consideration)", "상품 검색\n리뷰 확인\n가격 비교\n위시리스트", C_GREEN),
        ("구매 (Purchase)", "장바구니\n쿠폰 적용\n결제\n주문 확인", C_ORANGE),
        ("유지 (Retention)", "배송 추적\n리뷰 작성\n포인트 적립\n재구매", C_PURPLE),
    ], left=2.0, top=1.8, size=4.2)
    add_multiline_textbox(slide, 7.0, 1.8, 5.8, 5.0, [
        "[단계별 전환율]",
        "",
        "인지 → 탐색: 12% (CTR)",
        "탐색 → 장바구니: 35%",
        "장바구니 → 결제: 68%",
        "결제 → 재구매: 42%",
        "",
        "[개선 포인트]",
        "1. 장바구니 이탈율 32% 감소 필요",
        "   → 이탈 시 푸시 알림 도입",
        "2. 리뷰 작성률 15% → 25% 목표",
        "   → 리뷰 포인트 2배 이벤트",
    ], font_size=11)
    add_page_number(slide, 8, TOTAL)

    # 9. Marketing Performance (chart + table)
    slide = add_blank(prs)
    add_slide_header(slide, "마케팅 채널별 성과", "2026년 1분기 광고비 대비 효율")
    add_chart(slide, XL_CHART_TYPE.BAR_CLUSTERED, 0.5, 1.5, 6.0, 5.5,
        ["네이버 검색", "인스타그램", "유튜브", "카카오톡", "구글 검색", "페이스북"],
        [
            ("광고비(백만원)", (250, 180, 150, 120, 100, 80)),
            ("매출기여(백만원)", (1200, 850, 620, 480, 380, 190)),
        ],
        title="채널별 투자 vs 매출 기여")
    add_table(slide, 7.0, 1.8, 5.8, 4.2,
        ["채널", "ROAS", "CPA", "전환율"],
        [
            ["네이버 검색", "480%", "12,500원", "3.2%"],
            ["인스타그램", "472%", "8,200원", "2.8%"],
            ["유튜브", "413%", "15,300원", "1.9%"],
            ["카카오톡", "400%", "9,800원", "2.5%"],
            ["구글 검색", "380%", "11,000원", "2.1%"],
            ["페이스북", "238%", "18,500원", "1.2%"],
        ])
    add_page_number(slide, 9, TOTAL)

    # 10. Competitive Analysis Table
    slide = add_blank(prs)
    add_slide_header(slide, "경쟁사 비교 분석", "주요 KPI 벤치마크")
    add_table(slide, 0.5, 1.5, 12.3, 4.5,
        ["항목", "자사 (샘플쇼핑)", "경쟁사 A", "경쟁사 B", "경쟁사 C", "업계 평균"],
        [
            ["분기 매출", "45.2억", "120.5억", "38.7억", "22.1억", "42.3억"],
            ["MAU", "285만", "850만", "210만", "95만", "260만"],
            ["객단가", "93,800원", "78,200원", "102,500원", "85,400원", "88,000원"],
            ["재구매율", "42%", "48%", "38%", "35%", "40%"],
            ["NPS", "62점", "58점", "55점", "48점", "52점"],
            ["배송 만족도", "4.5/5.0", "4.3/5.0", "4.1/5.0", "3.8/5.0", "4.0/5.0"],
            ["앱 평점", "4.6", "4.4", "4.3", "4.1", "4.2"],
        ])
    add_multiline_textbox(slide, 0.5, 6.2, 12, 1.0, [
        "* 자사 강점: 객단가, NPS, 배송 만족도에서 업계 상위. 약점: MAU 규모 — 신규 유저 확보 전략 필요"
    ], font_size=12, color=C_GRAY)
    add_page_number(slide, 10, TOTAL)

    # 11. Technology Roadmap (timeline)
    slide = add_blank(prs)
    add_slide_header(slide, "기술 로드맵", "2026년 주요 기술 투자 계획")
    add_timeline(slide, [
        ("2026.Q1", "AI 추천 v2\n개인화 강화", C_BLUE),
        ("2026.Q2", "음성 검색\n도입", C_GREEN),
        ("2026.Q3", "AR 피팅룸\n패션 카테고리", C_ORANGE),
        ("2026.Q4", "자율배송 로봇\n수도권 시범", C_PURPLE),
    ], left=1.5, top=2.5, total_width=10.0)
    add_multiline_textbox(slide, 0.8, 4.5, 11.5, 2.5, [
        "[기술 투자 예산: 연간 15억원]",
        "",
        "Q1: AI 추천 엔진 v2 — 실시간 행동 기반 추천, 예상 매출 +8% (투자 3억원)",
        "Q2: 음성 검색 — 한국어 음성인식 통합, 모바일 앱 우선 (투자 2억원)",
        "Q3: AR 피팅룸 — 의류 가상 착용, 반품률 20% 감소 목표 (투자 5억원)",
        "Q4: 자율배송 시범 — 수도권 3개 거점, 배송비 30% 절감 기대 (투자 5억원)",
    ], font_size=12)
    add_page_number(slide, 11, TOTAL)

    # 12. Regional Sales (chart)
    slide = add_blank(prs)
    add_slide_header(slide, "지역별 매출 분포", "2026년 1분기")
    add_chart(slide, XL_CHART_TYPE.DOUGHNUT, 0.5, 1.5, 6.0, 5.5,
        ["서울", "경기", "인천", "부산", "대구", "대전", "광주", "기타"],
        [("매출비중", (35, 25, 8, 7, 5, 4, 3, 13))],
        title="지역별 매출 비중")
    add_table(slide, 7.0, 1.8, 5.8, 4.8,
        ["지역", "매출(억원)", "비중", "성장률", "당일배송"],
        [
            ["서울", "15.8", "35%", "+22%", "O"],
            ["경기", "11.3", "25%", "+28%", "O"],
            ["인천", "3.6", "8%", "+25%", "O"],
            ["부산", "3.2", "7%", "+18%", "X"],
            ["대구", "2.3", "5%", "+15%", "X"],
            ["대전", "1.8", "4%", "+20%", "X"],
            ["광주", "1.4", "3%", "+12%", "X"],
            ["기타", "5.8", "13%", "+10%", "X"],
        ])
    add_page_number(slide, 12, TOTAL)

    # 13. Risk & Mitigation
    slide = add_blank(prs)
    add_slide_header(slide, "리스크 분석 및 대응 방안")
    add_table(slide, 0.5, 1.5, 12.3, 5.0,
        ["리스크", "영향도", "발생 가능성", "현재 상태", "대응 방안"],
        [
            ["물류비 상승", "높음", "높음", "주시중", "자체 물류센터 확보 (2026.Q3)"],
            ["경쟁사 가격 공세", "높음", "중간", "대응중", "PB 상품 라인 확대 + 멤버십 강화"],
            ["개인정보 규제 강화", "중간", "높음", "준비중", "CDP 도입 + 1st party 데이터 전략"],
            ["환율 변동", "중간", "중간", "모니터링", "환헤지 + 국내 소싱 비중 확대"],
            ["기술 인력 이탈", "높음", "낮음", "안정", "RSU 보상 + 기술 교육 프로그램"],
            ["서버 장애", "높음", "낮음", "안정", "멀티 AZ + Auto Scaling + DR 계획"],
        ])
    add_page_number(slide, 13, TOTAL)

    # 14. Next Steps
    slide = add_blank(prs)
    add_slide_header(slide, "2분기 핵심 과제")
    add_flowchart(slide, [
        "신규 고객\n확보 50만명", "객단가 회복\n10만원 목표", "모바일 비중\n70% 돌파", "AI 추천 v2\n정식 출시", "당일배송\n전국 확대"
    ], left=0.4, top=2.0, box_w=2.1, box_h=1.1, gap=0.1,
    colors=[C_BLUE, C_GREEN, C_PURPLE, C_ORANGE, C_RED])
    add_multiline_textbox(slide, 0.5, 4.2, 12, 3.0, [
        "[분기 목표]",
        "  매출: 52억원 (+15% QoQ) | MAU: 340만명 | 재구매율: 45%",
        "",
        "[핵심 액션 아이템]",
        "  1. 여름 시즌 프로모션 기획 (6월 론칭, 마케팅비 5억원)",
        "  2. AI 추천 v2 A/B 테스트 → 정식 출시 (4월)",
        "  3. 부산/대전 당일배송 센터 오픈 (5월)",
        "  4. PB 브랜드 '샘플 에센셜' 론칭 (30개 SKU)",
        "  5. VIP 전용 라운지 서비스 시범 운영",
    ], font_size=13)
    add_page_number(slide, 14, TOTAL)

    fp = OUT_DIR / "ecommerce" / "이커머스_사업_현황_보고서.pptx"
    fp.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(fp))
    print(f"  ✓ E-commerce PPT — {TOTAL} slides")


# ═══════════════════════════════════════════════════════════════
#  DOMAIN: HR
# ═══════════════════════════════════════════════════════════════

def gen_hr_ppt():
    prs = new_prs()
    TOTAL = 14

    add_title_slide(prs, "인사관리 현황 보고서", "2026년 1분기 | 인사팀\n기밀 - 내부 공유용", bg_color=RGBColor(0x1B, 0x4F, 0x72))

    # 2. KPI
    slide = add_blank(prs)
    add_slide_header(slide, "인사 핵심 지표")
    add_kpi_boxes(slide, [
        ("총 인원", "127명", C_BLUE),
        ("신규 채용", "18명", C_GREEN),
        ("퇴사율", "4.7%", C_ORANGE),
        ("교육 이수율", "89%", C_PURPLE),
        ("직원 만족도", "4.2/5.0", C_DKBLUE),
    ])
    add_page_number(slide, 2, TOTAL)

    # 3. Org Chart
    slide = add_blank(prs)
    add_slide_header(slide, "조직도", "2026년 3월 기준")
    add_org_chart(slide, [
        (0, "김대표", "대표이사", C_DKBLUE),
        (1, "경영지원본부", "박인사 본부장", C_BLUE),
        (1, "기술본부", "최기술 본부장", C_GREEN),
        (1, "영업본부", "임영업 본부장", C_ORANGE),
        (2, "인사팀 (8명)", "박인사 팀장", C_BLUE),
        (2, "재무팀 (6명)", "이재무 팀장", C_BLUE),
        (2, "총무팀 (5명)", "김총무 팀장", C_BLUE),
        (2, "개발1팀 (15명)", "정개발 팀장", C_GREEN),
        (2, "개발2팀 (12명)", "서개발 팀장", C_GREEN),
        (2, "QA팀 (8명)", "송QA 팀장", C_GREEN),
        (2, "국내영업 (12명)", "임영업 팀장", C_ORANGE),
        (2, "해외영업 (8명)", "오해외 팀장", C_ORANGE),
    ], top=1.4)
    add_page_number(slide, 3, TOTAL)

    # 4. Department headcount (bar)
    slide = add_blank(prs)
    add_slide_header(slide, "부서별 인원 현황")
    add_chart(slide, XL_CHART_TYPE.COLUMN_CLUSTERED, 0.5, 1.5, 7.5, 5.5,
        ["인사팀", "재무팀", "총무팀", "개발1팀", "개발2팀", "QA팀", "데이터팀", "국내영업", "해외영업", "마케팅"],
        [
            ("정원", (10, 8, 6, 18, 15, 10, 8, 15, 10, 8)),
            ("현원", (8, 6, 5, 15, 12, 8, 6, 12, 8, 7)),
        ],
        title="부서별 정원 vs 현원")
    add_table(slide, 8.5, 1.8, 4.3, 4.5,
        ["부서", "충원율"],
        [
            ["인사팀", "80%"],
            ["재무팀", "75%"],
            ["개발1팀", "83%"],
            ["개발2팀", "80%"],
            ["QA팀", "80%"],
            ["국내영업", "80%"],
            ["해외영업", "80%"],
        ])
    add_page_number(slide, 4, TOTAL)

    # 5. Salary by position (table)
    slide = add_blank(prs)
    add_slide_header(slide, "직급별 연봉 현황", "2026년 기준 (단위: 만원)")
    add_table(slide, 0.5, 1.5, 12.3, 5.0,
        ["직급", "인원", "평균 연봉", "최저", "최고", "전년비", "업계 대비"],
        [
            ["임원", "5명", "12,000", "10,000", "15,000", "+5%", "상위 30%"],
            ["본부장/팀장", "12명", "9,200", "8,200", "12,000", "+6%", "상위 25%"],
            ["시니어", "28명", "7,500", "6,500", "8,500", "+7%", "상위 20%"],
            ["대리/과장", "35명", "5,800", "5,000", "6,800", "+5%", "평균"],
            ["주니어", "32명", "4,800", "4,200", "5,500", "+8%", "상위 15%"],
            ["인턴", "15명", "2,800", "2,400", "3,200", "+10%", "상위 10%"],
        ])
    add_multiline_textbox(slide, 0.5, 6.2, 12, 1.0, [
        "* 주니어/인턴 인상률이 높은 것은 시장 경쟁력 확보를 위한 전략적 판단. 전체 인건비 예산 전년비 +6.5% 증가."
    ], font_size=11, color=C_GRAY)
    add_page_number(slide, 5, TOTAL)

    # 6. Recruitment Process (flowchart)
    slide = add_blank(prs)
    add_slide_header(slide, "채용 프로세스", "전체 리드타임: 평균 28일")
    add_flowchart(slide, [
        "채용 요청\n(부서장)", "서류 심사\n(인사팀)", "1차 면접\n(실무진)", "2차 면접\n(임원)", "처우 협상\n(인사팀)", "입사\n(온보딩)"
    ], left=0.3, top=2.0, box_w=1.8, box_h=0.9, gap=0.1)
    add_table(slide, 0.5, 3.8, 12.3, 3.0,
        ["단계", "소요일", "통과율", "담당자", "주요 평가 항목"],
        [
            ["서류 심사", "3일", "35%", "인사팀", "학력, 경력, 포트폴리오"],
            ["1차 면접", "7일", "50%", "실무 팀장", "기술 역량, 문제 해결력"],
            ["2차 면접", "7일", "70%", "본부장/임원", "조직 적합성, 리더십"],
            ["처우 협상", "5일", "90%", "인사팀", "연봉, 복리후생, 입사일"],
            ["온보딩", "5일", "100%", "멘토", "장비, 계정, OJT"],
        ])
    add_page_number(slide, 6, TOTAL)

    # 7. Recruitment trend (line)
    slide = add_blank(prs)
    add_slide_header(slide, "연간 채용/퇴사 추이", "2022~2026 (1분기 기준)")
    add_chart(slide, XL_CHART_TYPE.LINE, 0.5, 1.5, 7.5, 5.5,
        ["2022", "2023", "2024", "2025", "2026"],
        [
            ("채용 (명)", (12, 18, 25, 22, 18)),
            ("퇴사 (명)", (5, 8, 10, 7, 6)),
            ("순증 (명)", (7, 10, 15, 15, 12)),
        ],
        title="연간 채용/퇴사 추이")
    add_multiline_textbox(slide, 8.5, 1.8, 4.5, 5.0, [
        "[채용 현황]",
        "",
        "2026년 1분기:",
        "  신규 채용: 18명",
        "  퇴사: 6명",
        "  순증: 12명",
        "",
        "[퇴사 사유 분석]",
        "  이직: 3명 (50%)",
        "  개인사유: 2명 (33%)",
        "  계약만료: 1명 (17%)",
        "",
        "자발적 퇴사율: 4.7%",
        "(업계 평균 8.2%)",
    ], font_size=11)
    add_page_number(slide, 7, TOTAL)

    # 8. Training stats (table + chart)
    slide = add_blank(prs)
    add_slide_header(slide, "교육 이수 현황", "2026년 1분기")
    add_chart(slide, XL_CHART_TYPE.BAR_CLUSTERED, 0.5, 1.5, 6.5, 5.5,
        ["리더십", "기술(Python)", "기술(AWS)", "기술(React)", "데이터분석", "커뮤니케이션", "보안"],
        [
            ("이수율(%)", (92, 85, 78, 88, 72, 95, 100)),
        ],
        title="과정별 이수율 (%)")
    add_table(slide, 7.5, 1.8, 5.5, 4.5,
        ["교육 과정", "대상", "이수/대상", "평균 점수"],
        [
            ["리더십 워크숍", "팀장급", "11/12", "85점"],
            ["Python 고급", "개발자", "22/26", "88점"],
            ["AWS 자격증", "인프라팀", "6/8", "92점"],
            ["React 심화", "프론트엔드", "7/8", "90점"],
            ["데이터 분석", "전직원", "82/114", "78점"],
            ["보안 교육", "전직원", "127/127", "95점"],
        ])
    add_page_number(slide, 8, TOTAL)

    # 9. Performance Review Distribution (pie)
    slide = add_blank(prs)
    add_slide_header(slide, "2025년 하반기 인사평가 결과", "등급 분포")
    add_chart(slide, XL_CHART_TYPE.PIE, 0.5, 1.5, 6.0, 5.5,
        ["S등급", "A등급", "B+등급", "B등급", "C등급"],
        [("인원비율", (5, 20, 32, 35, 8))],
        title="평가 등급 분포 (%)")
    add_table(slide, 7.0, 1.8, 5.8, 4.0,
        ["등급", "인원", "연봉인상", "인센티브"],
        [
            ["S (상위 5%)", "6명", "+12%", "300%"],
            ["A (상위 25%)", "25명", "+8%", "200%"],
            ["B+ (상위 55%)", "40명", "+5%", "100%"],
            ["B (상위 90%)", "44명", "+3%", "-"],
            ["C (하위 10%)", "10명", "동결", "PIP 대상"],
        ])
    add_page_number(slide, 9, TOTAL)

    # 10. Employee Satisfaction (bar)
    slide = add_blank(prs)
    add_slide_header(slide, "직원 만족도 조사 결과", "2026년 3월 실시 (응답률 91%)")
    add_chart(slide, XL_CHART_TYPE.BAR_CLUSTERED, 0.5, 1.5, 7.5, 5.5,
        ["급여/보상", "복리후생", "업무환경", "성장기회", "조직문화", "리더십", "워라밸"],
        [
            ("2025년", (3.8, 4.0, 4.1, 3.9, 4.0, 3.7, 4.2)),
            ("2026년", (4.0, 4.3, 4.2, 4.1, 4.3, 3.9, 4.5)),
        ],
        title="항목별 만족도 (5점 만점)")
    add_multiline_textbox(slide, 8.5, 1.8, 4.3, 5.0, [
        "[YoY 개선 항목]",
        "복리후생: +0.3점",
        "  (식대/교육비 인상 효과)",
        "워라밸: +0.3점",
        "  (재택근무 확대 효과)",
        "",
        "[개선 필요]",
        "리더십: 3.9점 (최저)",
        "  → 리더십 코칭 프로그램",
        "     2분기 도입 예정",
    ], font_size=11)
    add_page_number(slide, 10, TOTAL)

    # 11. Diversity metrics
    slide = add_blank(prs)
    add_slide_header(slide, "다양성 지표", "2026년 3월 기준")
    add_chart(slide, XL_CHART_TYPE.PIE, 0.5, 1.5, 4.0, 4.0,
        ["남성", "여성"],
        [("성별비율", (62, 38))],
        title="성별 비율")
    add_chart(slide, XL_CHART_TYPE.PIE, 4.8, 1.5, 4.0, 4.0,
        ["20대", "30대", "40대", "50대+"],
        [("연령대", (25, 42, 25, 8))],
        title="연령대 분포")
    add_table(slide, 0.5, 5.5, 12.3, 1.5,
        ["지표", "현재", "목표 (2026말)", "업계 평균"],
        [
            ["여성 관리자 비율", "22%", "30%", "25%"],
            ["장애인 고용률", "3.2%", "3.5%", "3.1%"],
        ])
    add_page_number(slide, 11, TOTAL)

    # 12. Retention Strategy (matrix)
    slide = add_blank(prs)
    add_slide_header(slide, "핵심 인재 리텐션 전략", "성과-이탈 위험 매트릭스")
    add_matrix_diagram(slide, [
        ("고성과 / 저위험\n(Star)", "보상 강화\nRSU 부여\n핵심 프로젝트 배치", C_GREEN),
        ("고성과 / 고위험\n(Risk)", "즉시 면담\n처우 개선\n경력 경로 제시", C_RED),
        ("저성과 / 저위험\n(Steady)", "역량 개발\n멘토링\n직무 전환 검토", C_BLUE),
        ("저성과 / 고위험\n(Exit)", "PIP 진행\n직무 재배치\n원만한 전환 지원", C_GRAY),
    ], left=1.5, top=1.6, size=4.5)
    add_multiline_textbox(slide, 7.2, 1.8, 5.5, 5.0, [
        "[현황 분석]",
        "",
        "Star (42명, 33%):",
        "  핵심 개발자 및 영업 에이스",
        "",
        "Risk (8명, 6%):",
        "  시니어 개발자 3명 이직 의향",
        "  → 긴급 처우 개선 진행 중",
        "",
        "Steady (62명, 49%):",
        "  안정적, 교육 투자 효과 높음",
        "",
        "Exit (15명, 12%):",
        "  PIP 5명 진행 중",
    ], font_size=11)
    add_page_number(slide, 12, TOTAL)

    # 13. HR Calendar Timeline
    slide = add_blank(prs)
    add_slide_header(slide, "2026년 인사 일정", "주요 이벤트 타임라인")
    add_timeline(slide, [
        ("1월", "연봉 협상\n인사평가 피드백", C_BLUE),
        ("3월", "상반기 채용\n교육 계획", C_GREEN),
        ("6월", "상반기 평가\n인턴 채용", C_ORANGE),
        ("9월", "하반기 채용\n리더십 워크숍", C_PURPLE),
        ("12월", "하반기 평가\n승진 심사", C_RED),
    ], left=1.0, top=2.0, total_width=11.0)
    add_multiline_textbox(slide, 0.5, 4.5, 12, 2.5, [
        "[2분기 중점 과제]",
        "  1. 상반기 인사평가 준비 (5월 평가, 6월 피드백)",
        "  2. 하계 인턴십 프로그램 (6~8월, 15명 선발)",
        "  3. 리더십 코칭 프로그램 론칭 (4월, 팀장급 대상)",
        "  4. 개발팀 충원 (시니어 3명, 주니어 5명)",
        "  5. 보안 교육 리프레시 (전직원 대상, 5월)",
    ], font_size=12)
    add_page_number(slide, 13, TOTAL)

    # 14. Summary
    slide = add_blank(prs)
    add_slide_header(slide, "요약 및 의사결정 사항")
    add_table(slide, 0.5, 1.5, 12.3, 5.0,
        ["안건", "현황", "제안", "예산", "의사결정"],
        [
            ["시니어 개발자 처우 개선", "3명 이탈 위험", "연봉 10% 인상 + RSU", "1.5억원", "승인 요청"],
            ["하계 인턴십 확대", "작년 10명 → 15명", "15명 선발, 8주 과정", "0.8억원", "승인 요청"],
            ["리더십 코칭 프로그램", "만족도 최저 항목", "외부 코치 6개월 계약", "0.5억원", "승인 요청"],
            ["신규 복리후생 (육아지원)", "워킹맘 만족도 낮음", "어린이집 제휴 + 유연근무", "0.3억원", "검토 요청"],
            ["사무실 확장 (8층 추가)", "좌석 95% 사용률", "8층 임차 + 인테리어", "2.0억원", "검토 요청"],
        ])
    add_page_number(slide, 14, TOTAL)

    fp = OUT_DIR / "hr" / "인사관리_현황_보고서.pptx"
    fp.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(fp))
    print(f"  ✓ HR PPT — {TOTAL} slides")


# ═══════════════════════════════════════════════════════════════
#  DOMAIN: HOSPITAL
# ═══════════════════════════════════════════════════════════════

def gen_hospital_ppt():
    prs = new_prs()
    TOTAL = 13

    add_title_slide(prs, "병원 운영 현황 보고서", "2026년 1분기 | 샘플종합병원\n경영기획실", bg_color=RGBColor(0xC0, 0x39, 0x2B))

    # 2. KPI
    slide = add_blank(prs)
    add_slide_header(slide, "병원 핵심 지표")
    add_kpi_boxes(slide, [
        ("병상 가동률", "87.3%", C_RED),
        ("외래 환자 수", "12,850명", C_BLUE),
        ("입원 환자 수", "2,340명", C_GREEN),
        ("평균 재원일수", "5.2일", C_ORANGE),
        ("환자 만족도", "4.3/5.0", C_PURPLE),
    ])
    add_page_number(slide, 2, TOTAL)

    # 3. Department stats (bar)
    slide = add_blank(prs)
    add_slide_header(slide, "진료과별 환자 수", "2026년 1분기 (외래 + 입원)")
    add_chart(slide, XL_CHART_TYPE.COLUMN_CLUSTERED, 0.5, 1.5, 8.0, 5.5,
        ["내과", "외과", "소아과", "정형외과", "응급의학과", "산부인과", "신경과"],
        [
            ("외래", (3200, 2100, 1800, 2500, 1500, 1200, 550)),
            ("입원", (450, 380, 120, 520, 280, 310, 280)),
        ],
        title="진료과별 환자 수 (명)")
    add_table(slide, 9.0, 2.0, 3.8, 4.0,
        ["진료과", "의사 수"],
        [
            ["내과", "12명"],
            ["외과", "8명"],
            ["소아과", "5명"],
            ["정형외과", "6명"],
            ["응급의학과", "8명"],
            ["산부인과", "4명"],
            ["신경과", "3명"],
        ])
    add_page_number(slide, 3, TOTAL)

    # 4. Monthly trend (line)
    slide = add_blank(prs)
    add_slide_header(slide, "월별 외래/입원 환자 추이")
    add_chart(slide, XL_CHART_TYPE.LINE, 0.5, 1.5, 12.3, 5.5,
        ["1월", "2월", "3월", "4월", "5월", "6월", "7월", "8월", "9월", "10월", "11월", "12월"],
        [
            ("외래환자(명)", (4100, 3800, 4950, 4200, 4500, 4800, 4300, 3900, 4600, 4700, 4400, 4200)),
            ("입원환자(명)", (720, 680, 940, 780, 820, 850, 790, 710, 830, 810, 780, 750)),
        ],
        title="2025년 월별 환자 수 추이")
    add_page_number(slide, 4, TOTAL)

    # 5. Disease classification (pie + table)
    slide = add_blank(prs)
    add_slide_header(slide, "주요 질병 분류별 통계", "ICD-10 기준 상위 질환")
    add_chart(slide, XL_CHART_TYPE.PIE, 0.5, 1.5, 5.5, 5.0,
        ["소화기계", "근골격계", "호흡기계", "순환기계", "내분비계", "기타"],
        [("비율", (22, 20, 18, 15, 10, 15))],
        title="질환 대분류별 비율")
    add_table(slide, 6.5, 1.5, 6.3, 5.0,
        ["순위", "ICD코드", "질환명", "환자수", "비율"],
        [
            ["1", "K29", "위염", "1,250명", "8.2%"],
            ["2", "M17", "무릎 골관절염", "980명", "6.4%"],
            ["3", "J06", "급성 상기도감염", "920명", "6.0%"],
            ["4", "I10", "고혈압", "850명", "5.6%"],
            ["5", "E11", "제2형 당뇨", "780명", "5.1%"],
            ["6", "G43", "편두통", "650명", "4.3%"],
            ["7", "M54", "요통", "620명", "4.1%"],
            ["8", "K21", "위식도역류", "580명", "3.8%"],
        ])
    add_page_number(slide, 5, TOTAL)

    # 6. Medication usage (table)
    slide = add_blank(prs)
    add_slide_header(slide, "의약품 사용 현황", "2026년 1분기 주요 처방 약품")
    add_table(slide, 0.5, 1.5, 12.3, 5.5,
        ["순위", "약품명", "분류", "처방 건수", "단가", "총 비용", "재고 현황"],
        [
            ["1", "오메프라졸 20mg", "소화기", "3,200건", "250원", "80만원", "충분"],
            ["2", "아목시실린 500mg", "항생제", "2,800건", "300원", "84만원", "충분"],
            ["3", "이부프로펜 400mg", "진통제", "2,500건", "150원", "37.5만원", "충분"],
            ["4", "아토르바스타틴 10mg", "고지혈증", "1,800건", "800원", "144만원", "주의"],
            ["5", "메트포르민 500mg", "당뇨", "1,650건", "200원", "33만원", "충분"],
            ["6", "수마트립탄 50mg", "두통", "980건", "1,200원", "117.6만원", "충분"],
            ["7", "세티리진 10mg", "항히스타민", "850건", "200원", "17만원", "충분"],
            ["8", "암로디핀 5mg", "고혈압", "1,500건", "350원", "52.5만원", "충분"],
            ["9", "프레드니솔론 5mg", "스테로이드", "720건", "500원", "36만원", "주의"],
            ["10", "아세트아미노펜 시럽", "해열제(소아)", "1,200건", "30원/ml", "36만원", "충분"],
        ])
    add_page_number(slide, 6, TOTAL)

    # 7. Patient care process (flowchart)
    slide = add_blank(prs)
    add_slide_header(slide, "외래 진료 프로세스", "환자 내원 ~ 귀가 전체 흐름")
    add_flowchart(slide, [
        "접수\n(원무과)", "대기\n(진료과)", "진료\n(담당의)", "검사\n(해당시)", "결과확인\n(담당의)", "수납/처방\n(원무과)"
    ], left=0.3, top=2.0, box_w=1.8, box_h=0.9, gap=0.1,
    colors=[C_BLUE, C_GRAY, C_GREEN, C_ORANGE, C_GREEN, C_PURPLE])
    add_table(slide, 0.5, 3.8, 12.3, 3.0,
        ["단계", "평균 소요시간", "목표", "개선방안"],
        [
            ["접수", "5분", "3분", "무인접수기 확대 (현재 4대 → 8대)"],
            ["대기", "25분", "15분", "예약제 강화 + 실시간 대기현황 앱"],
            ["진료", "15분", "15분", "적정 진료시간 유지"],
            ["검사", "30분~2시간", "20분~1시간", "검사장비 추가 도입 (CT 1대)"],
            ["수납", "10분", "5분", "무인수납기 + 모바일 결제"],
        ])
    add_page_number(slide, 7, TOTAL)

    # 8. ER stats (chart + table)
    slide = add_blank(prs)
    add_slide_header(slide, "응급실 운영 통계", "2026년 1분기")
    add_chart(slide, XL_CHART_TYPE.COLUMN_CLUSTERED, 0.5, 1.5, 6.5, 5.0,
        ["1등급\n(소생)", "2등급\n(긴급)", "3등급\n(응급)", "4등급\n(준응급)", "5등급\n(비응급)"],
        [
            ("환자수", (45, 180, 520, 450, 305)),
        ],
        title="KTAS 중증도별 응급 환자 수")
    add_table(slide, 7.5, 1.5, 5.3, 5.0,
        ["지표", "실적", "목표"],
        [
            ["총 내원 환자", "1,500명", "-"],
            ["평균 대기시간", "18분", "15분"],
            ["입원 전환율", "32%", "-"],
            ["72시간 재방문율", "4.2%", "5% 이하"],
            ["사망률 (1등급)", "8.9%", "10% 이하"],
            ["평균 체류시간", "4.5시간", "4시간"],
            ["이송율", "3.5%", "5% 이하"],
        ])
    add_page_number(slide, 8, TOTAL)

    # 9. Patient satisfaction (bar)
    slide = add_blank(prs)
    add_slide_header(slide, "환자 만족도 조사", "2026년 3월 (응답: 1,250명)")
    add_chart(slide, XL_CHART_TYPE.BAR_CLUSTERED, 0.5, 1.5, 7.5, 5.5,
        ["의료진 친절도", "진료 전문성", "시설 청결도", "대기 시간", "설명 충분성", "식사 품질", "주차 편의"],
        [
            ("2025년", (4.1, 4.3, 4.0, 3.2, 3.8, 3.5, 3.0)),
            ("2026년", (4.4, 4.5, 4.2, 3.5, 4.1, 3.8, 3.3)),
        ],
        title="항목별 만족도 (5점 만점)")
    add_multiline_textbox(slide, 8.5, 2.0, 4.3, 4.5, [
        "[개선 사항]",
        "",
        "대기시간 (3.5점):",
        "  예약 시스템 개선 중",
        "  무인접수기 추가 도입",
        "",
        "주차 (3.3점):",
        "  지하 주차장 증축 계획",
        "  2026년 하반기 완공",
        "",
        "식사 (3.8점):",
        "  환자식 메뉴 다양화",
        "  영양사 상담 서비스 추가",
    ], font_size=11)
    add_page_number(slide, 9, TOTAL)

    # 10. Surgery stats
    slide = add_blank(prs)
    add_slide_header(slide, "수술 실적", "2026년 1분기")
    add_table(slide, 0.5, 1.5, 12.3, 4.5,
        ["수술 종류", "건수", "평균 수술시간", "합병증률", "재원일수", "담당과"],
        [
            ["무릎 관절경", "85건", "1.5시간", "2.4%", "3일", "정형외과"],
            ["맹장 수술(복강경)", "62건", "1시간", "1.6%", "3일", "외과"],
            ["담낭 절제술", "48건", "1.5시간", "2.1%", "4일", "외과"],
            ["척추 디스크", "35건", "2시간", "3.4%", "7일", "정형외과"],
            ["제왕절개", "42건", "1시간", "1.2%", "5일", "산부인과"],
            ["유방암 수술", "18건", "2.5시간", "2.8%", "7일", "외과"],
            ["갑상선 수술", "22건", "2시간", "1.8%", "3일", "외과"],
        ])
    add_multiline_textbox(slide, 0.5, 6.2, 12, 1.0, [
        "총 수술 건수: 312건 (전년비 +8%). 합병증률 평균 2.2% (목표 3% 이하 달성). 수술 로봇 도입 후 복강경 합병증 30% 감소."
    ], font_size=11, color=C_GRAY)
    add_page_number(slide, 10, TOTAL)

    # 11. Infection control (line)
    slide = add_blank(prs)
    add_slide_header(slide, "감염 관리 지표", "2025.01 ~ 2026.03")
    add_chart(slide, XL_CHART_TYPE.LINE, 0.5, 1.5, 8.0, 5.5,
        ["1월", "2월", "3월", "4월", "5월", "6월", "7월", "8월", "9월", "10월", "11월", "12월", "1월", "2월", "3월"],
        [
            ("병원감염률(%)", (2.8, 2.5, 2.3, 2.4, 2.1, 2.0, 2.2, 2.3, 1.9, 1.8, 1.7, 1.6, 1.5, 1.4, 1.3)),
            ("손위생 이행률(%)", (82, 84, 85, 86, 88, 89, 90, 91, 92, 93, 94, 95, 96, 96, 97)),
        ],
        title="감염률 vs 손위생 이행률 추이")
    add_multiline_textbox(slide, 9.0, 2.0, 4.0, 4.5, [
        "[성과]",
        "감염률 2.8% → 1.3%",
        "(15개월간 54% 감소)",
        "",
        "손위생 이행률",
        "82% → 97%",
        "",
        "[핵심 요인]",
        "전자 모니터링 도입",
        "월간 피드백 리포트",
        "부서별 경쟁 인센티브",
    ], font_size=11)
    add_page_number(slide, 11, TOTAL)

    # 12. Financial summary
    slide = add_blank(prs)
    add_slide_header(slide, "재무 현황", "2026년 1분기 (단위: 억원)")
    add_chart(slide, XL_CHART_TYPE.COLUMN_CLUSTERED, 0.5, 1.5, 6.5, 5.5,
        ["진료수익", "입원수익", "건강검진", "기타수익"],
        [
            ("수익", (85, 62, 18, 5)),
            ("비용", (65, 50, 12, 3)),
        ],
        title="수익/비용 구조")
    add_table(slide, 7.5, 1.5, 5.3, 5.0,
        ["항목", "금액", "전년비"],
        [
            ["총 수익", "170억원", "+12%"],
            ["총 비용", "130억원", "+8%"],
            ["영업이익", "40억원", "+25%"],
            ["인건비", "78억원", "+10%"],
            ["재료비", "32억원", "+5%"],
            ["감가상각", "12억원", "+15%"],
            ["기타 경비", "8억원", "+3%"],
        ])
    add_page_number(slide, 12, TOTAL)

    # 13. Improvement plan
    slide = add_blank(prs)
    add_slide_header(slide, "2분기 개선 계획")
    add_flowchart(slide, [
        "CT 장비\n도입 (1대)", "무인접수기\n확대 (4대)", "전자처방\n시스템 개선", "주차장\n증축 착공"
    ], left=0.8, top=2.0, box_w=2.5, box_h=1.0, gap=0.2,
    colors=[C_RED, C_BLUE, C_GREEN, C_ORANGE])
    add_table(slide, 0.5, 4.0, 12.3, 3.0,
        ["과제", "예산", "기대효과", "완료 예정"],
        [
            ["CT 장비 도입", "8억원", "검사 대기 50% 단축", "2026.Q2"],
            ["무인접수기 확대", "0.5억원", "접수 대기 3분 이내", "2026.04"],
            ["전자처방 개선", "1.2억원", "처방 오류 80% 감소", "2026.Q2"],
            ["주차장 증축", "15억원", "주차 200대 → 350대", "2026.Q4"],
        ])
    add_page_number(slide, 13, TOTAL)

    fp = OUT_DIR / "hospital" / "병원_운영_현황_보고서.pptx"
    fp.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(fp))
    print(f"  ✓ Hospital PPT — {TOTAL} slides")


# ═══════════════════════════════════════════════════════════════
#  DOMAIN: INSURANCE
# ═══════════════════════════════════════════════════════════════

def gen_insurance_ppt():
    prs = new_prs()
    TOTAL = 12

    add_title_slide(prs, "보험 사업 현황 보고서", "2026년 1분기 | 샘플생명보험\n경영전략팀", bg_color=C_PURPLE)

    # 2. KPI
    slide = add_blank(prs)
    add_slide_header(slide, "1분기 핵심 성과")
    add_kpi_boxes(slide, [
        ("수입 보험료", "1,250억원", C_PURPLE),
        ("신계약 건수", "18,500건", C_BLUE),
        ("보험금 지급", "820억원", C_RED),
        ("손해율", "65.6%", C_ORANGE),
        ("13회차 유지율", "89.2%", C_GREEN),
    ])
    add_page_number(slide, 2, TOTAL)

    # 3. Product mix (pie)
    slide = add_blank(prs)
    add_slide_header(slide, "상품별 수입 보험료 구성")
    add_chart(slide, XL_CHART_TYPE.PIE, 0.5, 1.5, 5.5, 5.5,
        ["종합보험", "암보험", "실손의료", "운전자보험", "연금보험", "기타"],
        [("보험료비중", (28, 22, 20, 12, 10, 8))],
        title="상품별 수입보험료 비중 (%)")
    add_table(slide, 6.5, 1.5, 6.3, 5.0,
        ["상품", "보험료(억원)", "전년비", "신계약비율"],
        [
            ["종합보험", "350", "+8%", "25%"],
            ["암보험", "275", "+15%", "30%"],
            ["실손의료", "250", "+5%", "20%"],
            ["운전자보험", "150", "+3%", "12%"],
            ["연금보험", "125", "-2%", "8%"],
            ["기타", "100", "+10%", "5%"],
        ])
    add_page_number(slide, 3, TOTAL)

    # 4. Claims process (flowchart)
    slide = add_blank(prs)
    add_slide_header(slide, "보험금 청구 프로세스", "접수 ~ 지급 전체 흐름 (목표: 3영업일)")
    add_flowchart(slide, [
        "사고 접수\n(고객)", "서류 접수\n(콜센터/앱)", "심사\n(손해사정)", "의료자문\n(필요시)", "지급 결정\n(심사팀)", "보험금 지급\n(재무팀)"
    ], left=0.3, top=2.0, box_w=1.8, box_h=0.9, gap=0.1,
    colors=[C_BLUE, C_GREEN, C_ORANGE, C_RED, C_PURPLE, C_DKBLUE])
    add_table(slide, 0.5, 3.8, 12.3, 3.0,
        ["단계", "소요시간", "자동화율", "개선계획"],
        [
            ["사고 접수", "즉시", "80% (앱)", "AI 챗봇 접수 도입"],
            ["서류 접수", "1~2일", "60% (OCR)", "OCR 정확도 95% 목표"],
            ["심사", "1~2일", "40%", "AI 자동심사 확대 (소액)"],
            ["의료 자문", "2~3일", "0%", "텔레메디슨 자문 도입"],
            ["지급", "0.5일", "90%", "실시간 이체 시스템"],
        ])
    add_page_number(slide, 4, TOTAL)

    # 5. Loss ratio trend (line)
    slide = add_blank(prs)
    add_slide_header(slide, "손해율 추이", "2022 ~ 2026.Q1")
    add_chart(slide, XL_CHART_TYPE.LINE, 0.5, 1.5, 7.5, 5.5,
        ["2022.Q1", "Q2", "Q3", "Q4", "2023.Q1", "Q2", "Q3", "Q4", "2024.Q1", "Q2", "Q3", "Q4", "2025.Q1", "Q2", "Q3", "Q4", "2026.Q1"],
        [
            ("손해율(%)", (72, 70, 68, 71, 69, 67, 66, 68, 67, 65, 64, 66, 65, 64, 63, 65, 65.6)),
            ("업계평균(%)", (75, 74, 72, 74, 73, 71, 70, 72, 71, 69, 68, 70, 69, 68, 67, 69, 68)),
        ],
        title="분기별 손해율 추이 (자사 vs 업계)")
    add_multiline_textbox(slide, 8.5, 2.0, 4.5, 4.5, [
        "[분석]",
        "",
        "자사 손해율: 65.6%",
        "업계 평균: 68.0%",
        "차이: -2.4%p (양호)",
        "",
        "[요인]",
        "1. 언더라이팅 강화",
        "2. 보험사기 탐지 AI",
        "3. 건강관리 프로그램",
    ], font_size=11)
    add_page_number(slide, 5, TOTAL)

    # 6. Agent performance (bar + table)
    slide = add_blank(prs)
    add_slide_header(slide, "설계사 채널 실적", "2026년 1분기")
    add_chart(slide, XL_CHART_TYPE.COLUMN_CLUSTERED, 0.5, 1.5, 6.5, 5.5,
        ["전속 설계사", "GA (법인대리점)", "방카슈랑스", "온라인 직판", "TM (텔레마케팅)"],
        [
            ("신계약(건)", (8500, 4200, 2800, 2000, 1000)),
            ("수입보험료(억원)", (520, 310, 220, 120, 80)),
        ],
        title="채널별 실적")
    add_table(slide, 7.5, 1.5, 5.3, 5.0,
        ["채널", "설계사수", "1인당생산성", "유지율"],
        [
            ["전속", "1,200명", "4,333만원", "91%"],
            ["GA", "800명", "3,875만원", "85%"],
            ["방카", "350명", "6,286만원", "88%"],
            ["온라인", "-", "-", "92%"],
            ["TM", "150명", "5,333만원", "82%"],
        ])
    add_page_number(slide, 6, TOTAL)

    # 7. Claim analysis (doughnut + table)
    slide = add_blank(prs)
    add_slide_header(slide, "보험금 지급 분석", "2026년 1분기 총 820억원")
    add_chart(slide, XL_CHART_TYPE.DOUGHNUT, 0.5, 1.5, 5.5, 5.0,
        ["입원비", "수술비", "진단비", "통원비", "사망보험금", "기타"],
        [("지급비중", (30, 25, 20, 12, 8, 5))],
        title="보험금 유형별 비중")
    add_table(slide, 6.5, 1.5, 6.3, 5.0,
        ["유형", "지급액(억원)", "건수", "건당평균"],
        [
            ["입원비", "246", "12,500건", "197만원"],
            ["수술비", "205", "3,800건", "539만원"],
            ["진단비", "164", "820건", "2,000만원"],
            ["통원비", "98", "45,000건", "2.2만원"],
            ["사망보험금", "66", "65건", "1.02억원"],
            ["기타", "41", "2,100건", "195만원"],
        ])
    add_page_number(slide, 7, TOTAL)

    # 8. Fraud detection
    slide = add_blank(prs)
    add_slide_header(slide, "보험사기 탐지 현황", "AI 기반 이상탐지 시스템")
    add_matrix_diagram(slide, [
        ("자동 승인\n(정상)", "소액 청구\n서류 완비\n패턴 정상\n→ 즉시 지급", C_GREEN),
        ("수동 심사\n(주의)", "고액 청구\n복합 질환\n빈번 청구\n→ 심사팀 배정", C_ORANGE),
        ("AI 정상\n(모니터링)", "정상 패턴\n신규 계약\n→ 데이터 축적", C_BLUE),
        ("사기 의심\n(조사)", "패턴 이상\n허위 서류\n공모 의심\n→ SIU 조사", C_RED),
    ], left=1.5, top=1.6, size=4.5)
    add_multiline_textbox(slide, 7.2, 1.8, 5.5, 5.0, [
        "[1분기 실적]",
        "",
        "AI 탐지 건수: 342건",
        "실제 사기 확인: 85건",
        "적중률: 24.9%",
        "절감 금액: 42억원",
        "",
        "[사기 유형]",
        "허위 입원: 35건 (41%)",
        "진단서 위조: 22건 (26%)",
        "사고 조작: 15건 (18%)",
        "공모 청구: 13건 (15%)",
    ], font_size=11)
    add_page_number(slide, 8, TOTAL)

    # 9. Retention analysis
    slide = add_blank(prs)
    add_slide_header(slide, "계약 유지율 분석")
    add_chart(slide, XL_CHART_TYPE.LINE, 0.5, 1.5, 12.3, 5.5,
        ["1회차", "3회차", "6회차", "12회차", "13회차", "24회차", "25회차"],
        [
            ("자사", (99.5, 97.8, 95.2, 91.5, 89.2, 85.3, 83.8)),
            ("업계평균", (99.2, 96.5, 93.0, 88.0, 85.5, 80.0, 78.0)),
        ],
        title="회차별 유지율 추이 (%)")
    add_page_number(slide, 9, TOTAL)

    # 10. Digital transformation
    slide = add_blank(prs)
    add_slide_header(slide, "디지털 전환 현황")
    add_chart(slide, XL_CHART_TYPE.BAR_CLUSTERED, 0.5, 1.5, 6.5, 5.5,
        ["모바일 청약", "AI 심사", "챗봇 상담", "전자 서명", "OCR 접수", "마이데이터"],
        [
            ("도입률(%)", (85, 40, 65, 92, 60, 30)),
            ("목표(%)", (95, 70, 80, 98, 85, 60)),
        ],
        title="디지털 전환 진행률 (%)")
    add_table(slide, 7.5, 1.5, 5.3, 5.0,
        ["항목", "투자액", "ROI"],
        [
            ["모바일 청약", "5억원", "320%"],
            ["AI 심사", "15억원", "180%"],
            ["챗봇 상담", "3억원", "250%"],
            ["OCR 접수", "8억원", "210%"],
            ["마이데이터", "12억원", "진행중"],
        ])
    add_page_number(slide, 10, TOTAL)

    # 11-12. Timeline + Summary
    slide = add_blank(prs)
    add_slide_header(slide, "2026년 사업 로드맵")
    add_timeline(slide, [
        ("Q1", "AI 심사 v2\n사기탐지 강화", C_BLUE),
        ("Q2", "마이데이터\n연동 확대", C_GREEN),
        ("Q3", "건강관리 앱\n출시", C_ORANGE),
        ("Q4", "해외 재보험\n포트폴리오 최적화", C_PURPLE),
    ], left=1.5, top=2.0, total_width=10.0)
    add_multiline_textbox(slide, 0.5, 4.2, 12, 3.0, [
        "[연간 목표]",
        "  수입보험료: 5,200억원 (+8%) | 신계약: 80,000건 | 손해율: 64% 이하 | 13회차 유지율: 90%",
        "",
        "[핵심 전략]",
        "  1. AI 언더라이팅 확대 — 심사 자동화율 40% → 70%",
        "  2. 건강관리 연계 상품 — 운동/식단 관리 앱 + 보험료 할인",
        "  3. MZ세대 타겟 미니보험 — 월 1만원대 소액 상품 5종 출시",
        "  4. ESG 경영 — 그린본드 투자 + 탄소중립 보험 상품",
    ], font_size=13)
    add_page_number(slide, 11, TOTAL)

    slide = add_blank(prs)
    add_slide_header(slide, "의사결정 요청 사항")
    add_table(slide, 0.5, 1.5, 12.3, 5.0,
        ["안건", "내용", "예산", "기대효과", "결정"],
        [
            ["AI 심사 고도화", "심사 자동화 70% 확대", "15억원", "인건비 20억 절감/년", "승인요청"],
            ["건강관리 앱 개발", "헬스케어 연계 서비스", "8억원", "신계약 +15%", "승인요청"],
            ["미니보험 5종 출시", "MZ세대 신시장 개척", "3억원", "신규 가입 5만건/년", "승인요청"],
            ["해외 재보험 조정", "리스크 분산 최적화", "비용 중립", "손해율 -2%p", "검토요청"],
        ])
    add_page_number(slide, 12, TOTAL)

    fp = OUT_DIR / "insurance" / "보험_사업_현황_보고서.pptx"
    fp.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(fp))
    print(f"  ✓ Insurance PPT — {TOTAL} slides")


# ═══════════════════════════════════════════════════════════════
#  DOMAIN: EDUCATION
# ═══════════════════════════════════════════════════════════════

def gen_education_ppt():
    prs = new_prs()
    TOTAL = 12

    add_title_slide(prs, "대학 현황 보고서", "2026학년도 1학기 | 샘플대학교\n기획처", bg_color=RGBColor(0x1A, 0x52, 0x76))

    # 2. KPI
    slide = add_blank(prs)
    add_slide_header(slide, "주요 현황 지표")
    add_kpi_boxes(slide, [
        ("재학생 수", "12,350명", C_BLUE),
        ("취업률", "78.5%", C_GREEN),
        ("충원율", "98.2%", C_PURPLE),
        ("등록금 수입", "425억원", C_ORANGE),
        ("연구비 수주", "85억원", C_DKBLUE),
    ])
    add_page_number(slide, 2, TOTAL)

    # 3. Enrollment by dept (bar)
    slide = add_blank(prs)
    add_slide_header(slide, "학과별 재학생 현황")
    add_chart(slide, XL_CHART_TYPE.COLUMN_CLUSTERED, 0.5, 1.5, 8.0, 5.5,
        ["컴퓨터공학", "경영학", "디자인", "수학", "영어영문", "전자공학", "기계공학", "간호학", "심리학", "화학"],
        [
            ("정원", (400, 350, 200, 150, 180, 300, 250, 200, 150, 120)),
            ("재학생", (395, 340, 198, 145, 170, 290, 240, 200, 148, 115)),
        ],
        title="학과별 정원 vs 재학생")
    add_multiline_textbox(slide, 9.0, 2.0, 3.8, 4.0, [
        "[충원율 TOP 3]",
        "1. 간호학과 100%",
        "2. 컴퓨터공학 98.8%",
        "3. 디자인학과 99.0%",
        "",
        "[충원율 하위]",
        "화학과 95.8%",
        "수학과 96.7%",
    ], font_size=11)
    add_page_number(slide, 3, TOTAL)

    # 4. GPA distribution (pie)
    slide = add_blank(prs)
    add_slide_header(slide, "학생 성적 분포", "2025학년도 2학기 전체 평점 분포")
    add_chart(slide, XL_CHART_TYPE.PIE, 0.5, 1.5, 5.5, 5.5,
        ["4.0~4.5", "3.5~3.99", "3.0~3.49", "2.5~2.99", "2.0~2.49", "2.0 미만"],
        [("학생비율", (15, 28, 30, 18, 7, 2))],
        title="평점 분포 (%)")
    add_table(slide, 6.5, 1.5, 6.3, 5.0,
        ["학과", "평균 GPA", "장학금 수혜율", "학사경고율"],
        [
            ["컴퓨터공학", "3.42", "35%", "3.2%"],
            ["경영학", "3.28", "28%", "4.5%"],
            ["디자인", "3.55", "40%", "2.1%"],
            ["수학", "3.15", "45%", "5.8%"],
            ["영어영문", "3.38", "32%", "3.0%"],
            ["간호학", "3.62", "50%", "1.5%"],
        ])
    add_page_number(slide, 4, TOTAL)

    # 5. Curriculum structure (table)
    slide = add_blank(prs)
    add_slide_header(slide, "컴퓨터공학과 교과 과정 체계", "2026학년도 커리큘럼 맵")
    add_table(slide, 0.5, 1.5, 12.3, 5.5,
        ["학년", "1학기", "2학기", "학점", "비고"],
        [
            ["1학년", "프로그래밍기초, 이산수학, 컴퓨터개론", "C프로그래밍, 선형대수, 웹기초", "36", "기초 필수"],
            ["2학년", "자료구조, 컴퓨터구조, 확률통계", "알고리즘, 운영체제, OOP", "36", "전공 심화"],
            ["3학년", "AI개론, 데이터베이스, 네트워크", "머신러닝, 소프트웨어공학, 보안", "36", "전공 심화"],
            ["4학년", "졸업프로젝트I, 전공선택2과목", "졸업프로젝트II, 전공선택2과목", "24", "캡스톤"],
        ])
    add_flowchart(slide, [
        "프로그래밍\n기초", "자료구조", "알고리즘", "AI 개론", "졸업\n프로젝트"
    ], left=1.5, top=5.5, box_w=1.8, box_h=0.7, gap=0.2,
    colors=[C_BLUE, C_GREEN, C_ORANGE, C_PURPLE, C_RED])
    add_page_number(slide, 5, TOTAL)

    # 6. Employment stats (bar)
    slide = add_blank(prs)
    add_slide_header(slide, "취업 현황", "2025년 졸업생 기준")
    add_chart(slide, XL_CHART_TYPE.COLUMN_CLUSTERED, 0.5, 1.5, 7.0, 5.5,
        ["컴퓨터공학", "경영학", "디자인", "수학", "영어영문", "전자공학", "기계공학", "간호학"],
        [
            ("취업률(%)", (92, 75, 70, 65, 68, 85, 80, 95)),
            ("대학원(%)", (15, 8, 5, 25, 10, 18, 12, 3)),
        ],
        title="학과별 취업률 / 대학원 진학률")
    add_table(slide, 8.0, 1.5, 4.8, 5.0,
        ["주요 취업처", "인원"],
        [
            ["삼성전자", "45명"],
            ["네이버", "28명"],
            ["카카오", "22명"],
            ["LG전자", "18명"],
            ["현대자동차", "15명"],
            ["SK하이닉스", "12명"],
            ["스타트업", "85명"],
            ["공기업/공무원", "35명"],
        ])
    add_page_number(slide, 6, TOTAL)

    # 7. Research output
    slide = add_blank(prs)
    add_slide_header(slide, "연구 성과", "2025년 기준")
    add_chart(slide, XL_CHART_TYPE.BAR_CLUSTERED, 0.5, 1.5, 6.5, 5.5,
        ["SCI 논문", "KCI 논문", "특허 출원", "특허 등록", "기술이전", "정부과제"],
        [
            ("2024년", (120, 85, 35, 18, 8, 22)),
            ("2025년", (145, 92, 42, 25, 12, 28)),
        ],
        title="연구 성과 지표")
    add_table(slide, 7.5, 1.5, 5.3, 5.0,
        ["연구 분야", "논문수", "피인용"],
        [
            ["인공지능/ML", "38편", "520회"],
            ["데이터베이스", "22편", "180회"],
            ["컴퓨터비전", "18편", "310회"],
            ["자연어처리", "15편", "250회"],
            ["보안/암호학", "12편", "95회"],
        ])
    add_page_number(slide, 7, TOTAL)

    # 8. Facilities usage
    slide = add_blank(prs)
    add_slide_header(slide, "시설 이용 현황")
    add_table(slide, 0.5, 1.5, 12.3, 5.5,
        ["시설", "수용인원", "이용률", "만족도", "주요 개선 요청"],
        [
            ["중앙도서관", "500석", "92%", "4.3/5.0", "24시간 열람실 좌석 확대"],
            ["공학관 실습실", "30석x10실", "88%", "4.1/5.0", "PC 교체 (노후 장비)"],
            ["학생식당 (3개소)", "800석", "95%", "3.5/5.0", "메뉴 다양화, 가격 인하"],
            ["기숙사", "2,000명", "100%", "3.8/5.0", "에어컨 설치 (구관)"],
            ["체육관", "실내 3개, 운동장 1개", "78%", "4.0/5.0", "샤워실 리모델링"],
            ["창업보육센터", "20팀", "100%", "4.5/5.0", "입주 공간 확대"],
        ])
    add_page_number(slide, 8, TOTAL)

    # 9. Budget (doughnut + table)
    slide = add_blank(prs)
    add_slide_header(slide, "예산 현황", "2026학년도 (단위: 억원)")
    add_chart(slide, XL_CHART_TYPE.DOUGHNUT, 0.5, 1.5, 5.5, 5.5,
        ["인건비", "시설비", "장학금", "연구비", "운영비", "기타"],
        [("예산비중", (40, 15, 20, 12, 8, 5))],
        title="예산 구성 비율")
    add_table(slide, 6.5, 1.5, 6.3, 5.0,
        ["항목", "예산", "집행", "집행률"],
        [
            ["인건비", "320억", "240억", "75%"],
            ["시설비", "120억", "45억", "38%"],
            ["장학금", "160억", "85억", "53%"],
            ["연구비", "96억", "48억", "50%"],
            ["운영비", "64억", "35억", "55%"],
            ["기타", "40억", "12억", "30%"],
            ["합계", "800억", "465억", "58%"],
        ])
    add_page_number(slide, 9, TOTAL)

    # 10. Satisfaction
    slide = add_blank(prs)
    add_slide_header(slide, "학생 만족도 조사", "2026년 3월 (응답: 5,800명)")
    add_chart(slide, XL_CHART_TYPE.BAR_CLUSTERED, 0.5, 1.5, 12.3, 5.5,
        ["수업 품질", "교수 상담", "취업 지원", "시설/환경", "행정 서비스", "동아리/활동", "기숙사", "식당"],
        [
            ("2024년", (3.8, 3.5, 3.6, 3.4, 3.2, 4.0, 3.3, 3.0)),
            ("2025년", (4.0, 3.7, 3.9, 3.6, 3.5, 4.1, 3.5, 3.2)),
            ("2026년", (4.1, 3.9, 4.0, 3.7, 3.6, 4.2, 3.6, 3.4)),
        ],
        title="항목별 만족도 추이 (5점 만점)")
    add_page_number(slide, 10, TOTAL)

    # 11. Academic calendar timeline
    slide = add_blank(prs)
    add_slide_header(slide, "2026학년도 학사 일정")
    add_timeline(slide, [
        ("3월 2일", "1학기 개강", C_BLUE),
        ("4월 18일", "중간고사", C_GREEN),
        ("6월 12일", "기말고사", C_ORANGE),
        ("6월 19일", "1학기 종강", C_RED),
        ("9월 1일", "2학기 개강", C_PURPLE),
    ], left=1.0, top=2.0, total_width=11.0)
    add_table(slide, 0.5, 4.5, 12.3, 2.5,
        ["일정", "날짜", "대상", "비고"],
        [
            ["수강신청", "2.23~2.27", "전학년", "포털 > 학사정보"],
            ["수강정정", "3.2~3.6", "전학년", "변경 가능"],
            ["수강철회", "4.13~4.17", "전학년", "W 표기"],
            ["계절학기", "6.22~7.17", "희망자", "최대 6학점"],
        ])
    add_page_number(slide, 11, TOTAL)

    # 12. Strategy
    slide = add_blank(prs)
    add_slide_header(slide, "중장기 발전 전략", "VISION 2030")
    add_matrix_diagram(slide, [
        ("교육 혁신", "AI 융합 교육\n플립러닝 확대\n마이크로디그리\n산학 연계 강화", C_BLUE),
        ("연구 경쟁력", "BK21 사업 확대\nSCI 논문 200편\n산학협력 100건\n연구비 120억", C_GREEN),
        ("글로벌화", "교환학생 500명\n해외 대학 MOU 50건\n영어 강의 30%\n해외 캠퍼스 1개", C_ORANGE),
        ("캠퍼스 혁신", "스마트 캠퍼스\n기숙사 증축\n친환경 건물\n장애인 편의 확대", C_PURPLE),
    ], left=1.5, top=1.6, size=4.5)
    add_multiline_textbox(slide, 7.2, 1.8, 5.5, 5.0, [
        "[VISION 2030 목표]",
        "",
        "국내 대학 순위: 15위 → 10위",
        "QS 아시아: 200위 → 150위",
        "",
        "재학생: 12,350 → 15,000명",
        "외국인 학생: 5% → 15%",
        "취업률: 78% → 85%",
        "",
        "[투자 계획]",
        "5개년 총 1,500억원",
        "교육 500 / 연구 400",
        "시설 400 / 글로벌 200",
    ], font_size=11)
    add_page_number(slide, 12, TOTAL)

    fp = OUT_DIR / "education" / "대학_현황_보고서.pptx"
    fp.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(fp))
    print(f"  ✓ Education PPT — {TOTAL} slides")


# ═══════════════════════════════════════════════════════════════
#  DOMAIN: ACCOUNTING
# ═══════════════════════════════════════════════════════════════

def gen_accounting_ppt():
    prs = new_prs()
    TOTAL = 12

    add_title_slide(prs, "재무 보고서", "2026년 1분기 | 주식회사 샘플테크\n재무팀", bg_color=RGBColor(0x22, 0x99, 0x54))

    # 2. Financial KPIs
    slide = add_blank(prs)
    add_slide_header(slide, "재무 핵심 지표")
    add_kpi_boxes(slide, [
        ("매출액", "152억원", C_GREEN),
        ("영업이익", "23억원", C_BLUE),
        ("영업이익률", "15.1%", C_PURPLE),
        ("당기순이익", "18억원", C_ORANGE),
        ("부채비율", "42%", C_DKBLUE),
    ])
    add_page_number(slide, 2, TOTAL)

    # 3. Income Statement (table)
    slide = add_blank(prs)
    add_slide_header(slide, "손익계산서 요약", "2026년 1분기 (단위: 백만원)")
    add_table(slide, 1.5, 1.5, 10.3, 5.5,
        ["항목", "2026.Q1", "2025.Q1", "증감", "증감률"],
        [
            ["매출액", "15,200", "12,800", "+2,400", "+18.8%"],
            ["매출원가", "(9,120)", "(7,936)", "(+1,184)", "+14.9%"],
            ["매출총이익", "6,080", "4,864", "+1,216", "+25.0%"],
            ["판매관리비", "(3,780)", "(3,200)", "(+580)", "+18.1%"],
            ["  인건비", "(2,100)", "(1,800)", "(+300)", "+16.7%"],
            ["  마케팅비", "(850)", "(700)", "(+150)", "+21.4%"],
            ["  감가상각비", "(380)", "(320)", "(+60)", "+18.8%"],
            ["  기타", "(450)", "(380)", "(+70)", "+18.4%"],
            ["영업이익", "2,300", "1,664", "+636", "+38.2%"],
            ["영업외수익", "180", "120", "+60", "+50.0%"],
            ["영업외비용", "(280)", "(250)", "(+30)", "+12.0%"],
            ["법인세차감전이익", "2,200", "1,534", "+666", "+43.4%"],
            ["법인세비용", "(400)", "(280)", "(+120)", "+42.9%"],
            ["당기순이익", "1,800", "1,254", "+546", "+43.5%"],
        ],
        header_color=RGBColor(0x22, 0x99, 0x54))
    add_page_number(slide, 3, TOTAL)

    # 4. Balance Sheet (table)
    slide = add_blank(prs)
    add_slide_header(slide, "재무상태표 요약", "2026.03.31 기준 (단위: 백만원)")
    add_table(slide, 0.3, 1.5, 6.2, 5.5,
        ["자산", "금액", "비중"],
        [
            ["유동자산", "18,500", "42%"],
            ["  현금및현금성자산", "5,200", "12%"],
            ["  매출채권", "8,300", "19%"],
            ["  재고자산", "3,800", "9%"],
            ["  기타유동자산", "1,200", "3%"],
            ["비유동자산", "25,500", "58%"],
            ["  유형자산", "18,000", "41%"],
            ["  무형자산", "4,500", "10%"],
            ["  투자자산", "3,000", "7%"],
            ["자산총계", "44,000", "100%"],
        ],
        header_color=RGBColor(0x22, 0x99, 0x54))
    add_table(slide, 6.8, 1.5, 6.2, 5.5,
        ["부채 및 자본", "금액", "비중"],
        [
            ["유동부채", "8,200", "19%"],
            ["  매입채무", "4,500", "10%"],
            ["  단기차입금", "2,000", "5%"],
            ["  기타유동부채", "1,700", "4%"],
            ["비유동부채", "5,300", "12%"],
            ["  장기차입금", "4,000", "9%"],
            ["  기타비유동부채", "1,300", "3%"],
            ["부채총계", "13,500", "31%"],
            ["자본총계", "30,500", "69%"],
            ["부채+자본", "44,000", "100%"],
        ],
        header_color=RGBColor(0x22, 0x99, 0x54))
    add_page_number(slide, 4, TOTAL)

    # 5. Revenue trend (line)
    slide = add_blank(prs)
    add_slide_header(slide, "매출 및 이익 추이", "2023 ~ 2026 분기별")
    add_chart(slide, XL_CHART_TYPE.COLUMN_CLUSTERED, 0.5, 1.5, 12.3, 5.5,
        ["23Q1", "Q2", "Q3", "Q4", "24Q1", "Q2", "Q3", "Q4", "25Q1", "Q2", "Q3", "Q4", "26Q1"],
        [
            ("매출(억원)", (80, 85, 90, 95, 95, 100, 108, 115, 128, 132, 140, 148, 152)),
            ("영업이익(억원)", (8, 9, 10, 11, 11, 12, 14, 16, 17, 18, 20, 22, 23)),
        ],
        title="분기별 매출/영업이익 추이")
    add_page_number(slide, 5, TOTAL)

    # 6. Cost breakdown (pie)
    slide = add_blank(prs)
    add_slide_header(slide, "비용 구조 분석", "2026년 1분기")
    add_chart(slide, XL_CHART_TYPE.PIE, 0.5, 1.5, 6.0, 5.5,
        ["인건비", "재료비", "마케팅비", "임차료", "감가상각", "외주비", "기타"],
        [("비용비중", (33, 25, 13, 8, 6, 10, 5))],
        title="비용 구성 비율")
    add_multiline_textbox(slide, 7.0, 1.8, 5.8, 5.0, [
        "[비용 분석]",
        "",
        "총 비용: 129억원 (전년비 +15%)",
        "",
        "인건비 (33%, 42.6억원):",
        "  직원 127명, 평균연봉 6,400만원",
        "  신규 채용 18명 반영",
        "",
        "재료비 (25%, 32.3억원):",
        "  원자재 가격 상승 영향",
        "  대체 공급처 발굴 진행중",
        "",
        "마케팅비 (13%, 16.8억원):",
        "  봄 시즌 프로모션 집중 집행",
    ], font_size=11)
    add_page_number(slide, 6, TOTAL)

    # 7. Cash Flow (flowchart-style)
    slide = add_blank(prs)
    add_slide_header(slide, "현금흐름표 요약", "2026년 1분기")
    add_table(slide, 1.5, 1.5, 10.3, 5.5,
        ["항목", "금액(백만원)", "주요 내용"],
        [
            ["기초 현금", "4,500", "2025.12.31 기준"],
            ["", "", ""],
            ["영업활동 CF", "+2,800", ""],
            ["  당기순이익", "+1,800", ""],
            ["  감가상각비", "+380", "비현금 비용"],
            ["  매출채권 증가", "-500", "외상 매출 증가"],
            ["  매입채무 증가", "+320", "결제 조건 연장"],
            ["  기타", "+800", ""],
            ["", "", ""],
            ["투자활동 CF", "-1,500", ""],
            ["  설비 투자", "-1,200", "생산설비 증설"],
            ["  무형자산 취득", "-300", "소프트웨어 라이선스"],
            ["", "", ""],
            ["재무활동 CF", "-600", ""],
            ["  차입금 상환", "-500", "단기 차입금"],
            ["  배당금 지급", "-100", ""],
            ["", "", ""],
            ["기말 현금", "5,200", "+700 (순증)"],
        ],
        header_color=RGBColor(0x22, 0x99, 0x54))
    add_page_number(slide, 7, TOTAL)

    # 8. AR/AP aging (bar)
    slide = add_blank(prs)
    add_slide_header(slide, "매출채권/매입채무 연령 분석")
    add_chart(slide, XL_CHART_TYPE.COLUMN_CLUSTERED, 0.5, 1.5, 6.0, 5.5,
        ["30일 이내", "31~60일", "61~90일", "91~120일", "120일 초과"],
        [
            ("매출채권(백만원)", (4800, 2100, 850, 350, 200)),
            ("매입채무(백만원)", (2800, 1200, 350, 100, 50)),
        ],
        title="채권/채무 연령별 잔액")
    add_table(slide, 7.0, 1.5, 5.8, 5.0,
        ["지표", "금액", "평가"],
        [
            ["매출채권 총액", "83억원", "-"],
            ["매출채권회전율", "7.3회", "양호"],
            ["평균회수일", "50일", "개선필요"],
            ["대손충당금", "1.5억원", "-"],
            ["대손율", "1.8%", "양호"],
            ["매입채무 총액", "45억원", "-"],
            ["평균지급일", "38일", "양호"],
        ])
    add_page_number(slide, 8, TOTAL)

    # 9. Financial ratios (table)
    slide = add_blank(prs)
    add_slide_header(slide, "주요 재무 비율", "업계 평균 대비")
    add_table(slide, 0.5, 1.5, 12.3, 5.5,
        ["구분", "비율", "2025.Q1", "2026.Q1", "변동", "업계평균", "평가"],
        [
            ["수익성", "매출총이익률", "38.0%", "40.0%", "+2.0%p", "35%", "우수"],
            ["수익성", "영업이익률", "13.0%", "15.1%", "+2.1%p", "12%", "우수"],
            ["수익성", "순이익률", "9.8%", "11.8%", "+2.0%p", "8%", "우수"],
            ["수익성", "ROE", "16.4%", "23.6%", "+7.2%p", "15%", "우수"],
            ["안정성", "부채비율", "48%", "42%", "-6%p", "60%", "양호"],
            ["안정성", "유동비율", "215%", "226%", "+11%p", "150%", "양호"],
            ["활동성", "총자산회전율", "1.4회", "1.5회", "+0.1회", "1.2회", "양호"],
            ["활동성", "재고자산회전율", "12회", "13회", "+1회", "10회", "양호"],
        ],
        header_color=RGBColor(0x22, 0x99, 0x54))
    add_page_number(slide, 9, TOTAL)

    # 10. Budget vs Actual
    slide = add_blank(prs)
    add_slide_header(slide, "예산 대비 실적", "2026년 1분기")
    add_chart(slide, XL_CHART_TYPE.COLUMN_CLUSTERED, 0.5, 1.5, 12.3, 5.5,
        ["매출", "매출원가", "인건비", "마케팅", "설비투자", "연구개발", "관리비"],
        [
            ("예산(억원)", (140, 84, 40, 15, 10, 8, 5)),
            ("실적(억원)", (152, 91, 43, 17, 12, 9, 5)),
        ],
        title="예산 vs 실적 비교")
    add_page_number(slide, 10, TOTAL)

    # 11. Monthly closing process (flowchart)
    slide = add_blank(prs)
    add_slide_header(slide, "월말 결산 프로세스")
    add_flowchart(slide, [
        "전표 마감\n(D-5)", "은행조정\n(D-2)", "수정분개\n(D-1)", "시산표\n검토", "재무제표\n확정", "경영진\n보고"
    ], left=0.3, top=2.0, box_w=1.8, box_h=0.9, gap=0.1,
    colors=[C_BLUE, C_GREEN, C_ORANGE, C_PURPLE, C_RED, C_DKBLUE])
    add_table(slide, 0.5, 3.8, 12.3, 3.0,
        ["단계", "담당", "소요일", "자동화", "체크포인트"],
        [
            ["전표 마감", "영업/구매팀", "D-5~D-3", "80%", "미처리 전표 제로"],
            ["은행 조정", "재무팀", "D-2", "60%", "차이 ±10만원 이내"],
            ["수정 분개", "재무팀", "D-1", "40%", "감가상각/충당금"],
            ["시산표 검토", "재무팀장", "D-Day", "수동", "차대 균형 확인"],
            ["재무제표 확정", "CFO", "D-Day", "자동생성", "최종 승인"],
        ])
    add_page_number(slide, 11, TOTAL)

    # 12. Outlook
    slide = add_blank(prs)
    add_slide_header(slide, "2분기 전망 및 이슈")
    add_table(slide, 0.5, 1.5, 12.3, 5.0,
        ["항목", "전망", "리스크", "대응"],
        [
            ["매출", "165억원 (+8.6%)", "경기 둔화", "신규 고객 확보 강화"],
            ["영업이익", "26억원 (15.8%)", "원자재 가격", "대체 공급처 확보"],
            ["설비투자", "15억원", "공급 지연", "조기 발주"],
            ["인건비", "45억원", "인력 이탈", "처우 개선 (시니어)"],
            ["세금", "법인세 중간예납 (8월)", "세율 변경", "세무사 자문"],
        ])
    add_multiline_textbox(slide, 0.5, 5.8, 12, 1.5, [
        "[CFO Comment] 1분기 영업이익률 15.1%로 사상 최고치 달성. 2분기에도 성장 모멘텀 유지하되, 원자재 및 인건비 상승에 대비한 비용 효율화 필요."
    ], font_size=13, color=C_GRAY)
    add_page_number(slide, 12, TOTAL)

    fp = OUT_DIR / "accounting" / "재무_보고서.pptx"
    fp.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(fp))
    print(f"  ✓ Accounting PPT — {TOTAL} slides")


# ═══════════════════════════════════════════════════════════════
#  DOMAIN: KOREAN BOOKSTORE
# ═══════════════════════════════════════════════════════════════

def gen_bookstore_ppt():
    prs = new_prs()
    TOTAL = 12

    add_title_slide(prs, "온라인 서점 운영 보고서", "2026년 1분기 | 샘플북스\n운영팀", bg_color=RGBColor(0x8B, 0x45, 0x13))

    # 2. KPI
    slide = add_blank(prs)
    add_slide_header(slide, "1분기 핵심 지표")
    add_kpi_boxes(slide, [
        ("매출액", "38.5억원", RGBColor(0x8B, 0x45, 0x13)),
        ("판매 권수", "285,000권", C_BLUE),
        ("회원 수", "125,000명", C_GREEN),
        ("재구매율", "52%", C_ORANGE),
        ("평균 리뷰 평점", "4.3/5.0", C_PURPLE),
    ])
    add_page_number(slide, 2, TOTAL)

    # 3. Genre breakdown (pie)
    slide = add_blank(prs)
    add_slide_header(slide, "장르별 판매 현황")
    add_chart(slide, XL_CHART_TYPE.PIE, 0.5, 1.5, 5.5, 5.5,
        ["소설", "자기계발", "에세이", "과학/기술", "역사", "어린이", "만화", "기타"],
        [("판매비중", (28, 18, 15, 12, 8, 8, 6, 5))],
        title="장르별 판매 비중 (%)")
    add_table(slide, 6.5, 1.5, 6.3, 5.0,
        ["장르", "판매(억원)", "전년비", "평균 단가"],
        [
            ["소설", "10.8", "+35%", "14,200원"],
            ["자기계발", "6.9", "+8%", "16,500원"],
            ["에세이", "5.8", "+22%", "13,800원"],
            ["과학/기술", "4.6", "+18%", "19,200원"],
            ["역사", "3.1", "+12%", "18,500원"],
            ["어린이", "3.1", "+5%", "12,000원"],
            ["만화", "2.3", "+25%", "6,500원"],
        ])
    add_page_number(slide, 3, TOTAL)

    # 4. Bestseller (table)
    slide = add_blank(prs)
    add_slide_header(slide, "분기 베스트셀러 TOP 10")
    add_table(slide, 0.5, 1.5, 12.3, 5.5,
        ["순위", "도서명", "작가", "장르", "가격", "판매 부수", "평점", "리뷰 수"],
        [
            ["1", "채식주의자", "한강", "소설", "13,000원", "18,500부", "4.8", "2,350"],
            ["2", "사피엔스", "유발 하라리", "역사/과학", "22,000원", "12,800부", "4.7", "1,890"],
            ["3", "작별하지 않는다", "한강", "소설", "14,800원", "11,200부", "4.6", "1,520"],
            ["4", "1984", "조지 오웰", "소설", "11,000원", "9,800부", "4.5", "3,200"],
            ["5", "소년이 온다", "한강", "소설", "12,000원", "8,500부", "4.9", "980"],
            ["6", "밤의 언어", "김영하", "소설", "15,800원", "7,200부", "4.4", "620"],
            ["7", "AI 시대의 인간", "김대식", "과학", "18,000원", "6,800부", "4.3", "450"],
            ["8", "노르웨이의 숲", "하루키", "소설", "14,800원", "6,500부", "4.6", "2,800"],
            ["9", "28", "정유정", "소설/추리", "14,000원", "5,200부", "4.5", "1,100"],
            ["10", "살인자의 기억법", "김영하", "소설/추리", "13,500원", "4,800부", "4.4", "890"],
        ])
    add_page_number(slide, 4, TOTAL)

    # 5. Monthly sales trend (line)
    slide = add_blank(prs)
    add_slide_header(slide, "월별 매출 추이")
    add_chart(slide, XL_CHART_TYPE.LINE, 0.5, 1.5, 7.5, 5.5,
        ["1월", "2월", "3월", "4월", "5월", "6월", "7월", "8월", "9월", "10월", "11월", "12월"],
        [
            ("2025년(억원)", (8.5, 7.8, 9.2, 8.0, 8.5, 9.0, 7.5, 8.0, 9.5, 10.0, 11.5, 14.0)),
            ("2026년(억원)", (11.5, 12.0, 15.0, None, None, None, None, None, None, None, None, None)),
        ],
        title="월별 매출 추이 (억원)")
    add_multiline_textbox(slide, 8.5, 1.8, 4.3, 5.0, [
        "[1분기 분석]",
        "",
        "1월: 11.5억원",
        "  (설 선물 세트 특수)",
        "",
        "2월: 12.0억원",
        "  (발렌타인 이벤트)",
        "",
        "3월: 15.0억원",
        "  (한강 노벨문학상 효과",
        "   지속 + 봄 독서 캠페인)",
        "",
        "전년 동기 대비 +51%",
    ], font_size=11)
    add_page_number(slide, 5, TOTAL)

    # 6. Author analysis (bar)
    slide = add_blank(prs)
    add_slide_header(slide, "작가별 판매 실적 분석")
    add_chart(slide, XL_CHART_TYPE.BAR_CLUSTERED, 0.5, 1.5, 7.0, 5.5,
        ["한강", "김영하", "유발 하라리", "정유정", "무라카미 하루키", "조지 오웰", "김대식"],
        [
            ("판매부수(천부)", (38.2, 12.0, 12.8, 5.2, 6.5, 9.8, 6.8)),
            ("매출(억원)", (5.3, 1.8, 2.8, 0.7, 1.0, 1.1, 1.2)),
        ],
        title="TOP 7 작가 실적")
    add_multiline_textbox(slide, 8.0, 1.8, 4.8, 5.0, [
        "[한강 작가 효과]",
        "",
        "노벨문학상 수상 이후",
        "한강 작품 전체 매출 +280%",
        "",
        "1분기 한강 작품 매출: 5.3억",
        "(전체 소설 매출의 49%)",
        "",
        "[특이사항]",
        "- 채식주의자 특별판 품절",
        "  → 2쇄 인쇄 진행중",
        "- 작별하지않는다 영화화",
        "  → 4월 추가 매출 기대",
    ], font_size=11)
    add_page_number(slide, 6, TOTAL)

    # 7. Member grade distribution (pie + table)
    slide = add_blank(prs)
    add_slide_header(slide, "회원 등급 분포")
    add_chart(slide, XL_CHART_TYPE.DOUGHNUT, 0.5, 1.5, 5.5, 5.5,
        ["일반", "실버", "골드", "VIP"],
        [("회원비율", (55, 25, 15, 5))],
        title="회원 등급 비율")
    add_table(slide, 6.5, 1.5, 6.3, 5.0,
        ["등급", "인원", "평균 구매액/월", "포인트 적립률", "연 매출 기여"],
        [
            ["일반", "68,750명", "12,000원", "1%", "8.3억원"],
            ["실버", "31,250명", "28,000원", "2%", "10.5억원"],
            ["골드", "18,750명", "55,000원", "3%", "12.4억원"],
            ["VIP", "6,250명", "95,000원", "5%", "7.1억원"],
        ])
    add_page_number(slide, 7, TOTAL)

    # 8. Purchase process (flowchart)
    slide = add_blank(prs)
    add_slide_header(slide, "도서 구매 여정", "검색 ~ 리뷰 작성까지")
    add_flowchart(slide, [
        "도서 검색\n/추천", "상세 정보\n리뷰 확인", "장바구니\n담기", "결제\n(쿠폰적용)", "배송\n(1~2일)", "수령 후\n리뷰 작성"
    ], left=0.3, top=2.0, box_w=1.8, box_h=0.9, gap=0.1,
    colors=[C_BLUE, C_GREEN, C_ORANGE, C_PURPLE, RGBColor(0x8B, 0x45, 0x13), C_RED])
    add_table(slide, 0.5, 3.8, 12.3, 3.0,
        ["단계", "전환율", "이탈 원인", "개선 방안"],
        [
            ["검색 → 상세", "45%", "검색 결과 부정확", "AI 추천 + 검색 고도화"],
            ["상세 → 장바구니", "35%", "가격 비교", "최저가 보장 배지"],
            ["장바구니 → 결제", "72%", "배송비 부담", "1.5만원 이상 무료배송"],
            ["수령 → 리뷰", "15%", "리뷰 작성 귀찮음", "포인트 2배 + 간편 리뷰"],
        ])
    add_page_number(slide, 8, TOTAL)

    # 9. Reading club & events
    slide = add_blank(prs)
    add_slide_header(slide, "독서 모임 및 이벤트 현황")
    add_table(slide, 0.5, 1.5, 12.3, 3.5,
        ["이벤트", "기간", "참여자", "매출 기여", "효과"],
        [
            ["월간 독서 모임", "매월 마지막 토요일", "평균 120명", "+1,200만원/월", "재구매율 +8%p"],
            ["한강 특별전", "1.15~2.28", "온라인 85,000명", "+3.5억원", "소설 카테고리 +45%"],
            ["봄 독서 캠페인", "3.1~3.31", "42,000명 참여", "+2.1억원", "신규 회원 +5,200명"],
            ["저자 사인회 (김영하)", "3.15", "오프라인 200명", "+800만원", "SNS 노출 50만"],
        ])
    add_chart(slide, XL_CHART_TYPE.BAR_CLUSTERED, 0.5, 5.0, 12.3, 2.3,
        ["1월", "2월", "3월"],
        [
            ("이벤트매출(백만원)", (120, 180, 350)),
            ("일반매출(백만원)", (1030, 1020, 1150)),
        ],
        title="이벤트 vs 일반 매출 비교")
    add_page_number(slide, 9, TOTAL)

    # 10. e-book vs paper
    slide = add_blank(prs)
    add_slide_header(slide, "전자책 vs 종이책 비교")
    add_chart(slide, XL_CHART_TYPE.COLUMN_CLUSTERED, 0.5, 1.5, 6.5, 5.5,
        ["2022", "2023", "2024", "2025", "2026.Q1"],
        [
            ("종이책(%)", (85, 80, 75, 70, 65)),
            ("전자책(%)", (12, 16, 20, 25, 30)),
            ("오디오북(%)", (3, 4, 5, 5, 5)),
        ],
        title="매체별 매출 비중 추이 (%)")
    add_table(slide, 7.5, 1.5, 5.3, 5.0,
        ["매체", "매출", "성장률", "마진율"],
        [
            ["종이책", "25.0억", "+12%", "25%"],
            ["전자책", "11.6억", "+42%", "55%"],
            ["오디오북", "1.9억", "+35%", "60%"],
        ])
    add_multiline_textbox(slide, 7.5, 5.5, 5.3, 1.5, [
        "전자책/오디오북 마진율이 높아",
        "디지털 채널 확대가 수익성 핵심"
    ], font_size=11, color=C_GRAY)
    add_page_number(slide, 10, TOTAL)

    # 11. Inventory management
    slide = add_blank(prs)
    add_slide_header(slide, "재고 관리 현황")
    add_table(slide, 0.5, 1.5, 12.3, 5.5,
        ["장르", "보유 종수", "총 재고", "회전율", "반품률", "절판 위험", "발주 상태"],
        [
            ["소설", "2,500종", "85,000권", "3.2회/월", "5%", "12종", "정상"],
            ["자기계발", "1,800종", "42,000권", "2.8회/월", "8%", "5종", "정상"],
            ["에세이", "1,200종", "35,000권", "2.5회/월", "6%", "3종", "정상"],
            ["과학/기술", "900종", "18,000권", "2.0회/월", "10%", "8종", "주의"],
            ["역사", "800종", "15,000권", "1.8회/월", "12%", "15종", "주의"],
            ["어린이", "1,500종", "55,000권", "2.2회/월", "4%", "2종", "정상"],
            ["만화", "3,000종", "90,000권", "3.5회/월", "3%", "20종", "정상"],
            ["합계", "11,700종", "340,000권", "2.7회/월", "6%", "65종", "-"],
        ])
    add_page_number(slide, 11, TOTAL)

    # 12. Strategy
    slide = add_blank(prs)
    add_slide_header(slide, "2분기 전략 및 계획")
    add_flowchart(slide, [
        "AI 추천\n엔진 도입", "오디오북\n라인업 확대", "작가 팬미팅\n시리즈", "구독 서비스\n베타 출시"
    ], left=0.8, top=2.0, box_w=2.5, box_h=1.0, gap=0.2,
    colors=[C_BLUE, C_GREEN, C_ORANGE, C_PURPLE])
    add_multiline_textbox(slide, 0.5, 4.0, 12, 3.0, [
        "[2분기 목표] 매출 42억원 (+9% QoQ) | 회원 135,000명 | 전자책 비중 33%",
        "",
        "[핵심 과제]",
        "1. AI 도서 추천 — 구매/검색 이력 기반 개인화 (5월 출시, 투자 1.5억원)",
        "2. 오디오북 50종 추가 제작 — 베스트셀러 중심, 성우 녹음 (투자 0.8억원)",
        "3. 작가 팬미팅 시리즈 — 월 2회, 온/오프라인 병행 (매출 기여 +5%)",
        "4. 월간 구독 서비스 베타 — 월 9,900원, 전자책 무제한 (6월 오픈)",
        "5. 한강 노벨문학상 1주년 특별 기획전 (10월, 매출 기여 +8억 예상)",
    ], font_size=12)
    add_page_number(slide, 12, TOTAL)

    fp = OUT_DIR / "korean_bookstore" / "온라인_서점_운영_보고서.pptx"
    fp.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(fp))
    print(f"  ✓ Korean Bookstore PPT — {TOTAL} slides")


# ═══════════════════════════════════════════════════════════════
#  MAIN
# ═══════════════════════════════════════════════════════════════

if __name__ == "__main__":
    print("=" * 60)
    print("GraphRAG 테스트 PPT 생성기")
    print("=" * 60)
    print()

    gen_ecommerce_ppt()
    gen_hr_ppt()
    gen_hospital_ppt()
    gen_insurance_ppt()
    gen_education_ppt()
    gen_accounting_ppt()
    gen_bookstore_ppt()

    print()
    total = sum(1 for d in OUT_DIR.rglob("*.pptx"))
    print(f"총 {total}개 PPT 생성 완료!")
    print("=" * 60)
