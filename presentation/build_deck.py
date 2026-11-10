"""Сборка верхнеуровневой презентации по проекту «Карта и приоритеты процессов бухгалтерии».

    python3 presentation/build_deck.py [выходной.pptx]
"""
import sys
from pathlib import Path

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION, XL_LEGEND_POSITION
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Emu, Inches, Pt

GREEN = RGBColor(0x16, 0x62, 0x4D)
GREEN_L = RGBColor(0xD5, 0xEB, 0xE3)
INK = RGBColor(0x17, 0x20, 0x1D)
MUTED = RGBColor(0x5B, 0x69, 0x64)
BG = RGBColor(0xF3, 0xF5, 0xF4)
LINE = RGBColor(0xD3, 0xDB, 0xD8)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
ORANGE = RGBColor(0xC4, 0x6A, 0x12)
ORANGE_L = RGBColor(0xFB, 0xE7, 0xD1)
BLUE = RGBColor(0x2C, 0x62, 0xA8)
BLUE_L = RGBColor(0xDC, 0xE7, 0xF6)
RED = RGBColor(0x9B, 0x3B, 0x3B)
RED_L = RGBColor(0xF4, 0xDC, 0xDC)
FONT = "Arial"

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
BLANK = prs.slide_layouts[6]
SW, SH = 13.333, 7.5
MX = 0.6


def rect(slide, x, y, w, h, fill=None, line=None, shape=MSO_SHAPE.RECTANGLE, lw=0.75):
    s = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    if fill is None:
        s.fill.background()
    else:
        s.fill.solid()
        s.fill.fore_color.rgb = fill
    if line is None:
        s.line.fill.background()
    else:
        s.line.color.rgb = line
        s.line.width = Pt(lw)
    s.shadow.inherit = False
    return s


def text(slide, x, y, w, h, content, size=14, bold=False, color=INK, align=PP_ALIGN.LEFT,
         anchor=MSO_ANCHOR.TOP, spacing=1.1, space_after=0):
    """content: строка или список абзацев. Абзац - строка или (строка, {опции})."""
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Inches(0.05)
    tf.margin_top = tf.margin_bottom = Inches(0.03)
    tf.vertical_anchor = anchor
    paras = content if isinstance(content, list) else [content]
    for i, p in enumerate(paras):
        opts = {}
        if isinstance(p, tuple):
            p, opts = p
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        para.alignment = opts.get("align", align)
        para.line_spacing = spacing
        para.space_after = Pt(opts.get("space_after", space_after))
        run = para.add_run()
        run.text = p
        f = run.font
        f.name = FONT
        f.size = Pt(opts.get("size", size))
        f.bold = opts.get("bold", bold)
        f.color.rgb = opts.get("color", color)
    return tb


def bullets(slide, x, y, w, h, items, size=14, color=INK, gap=6, marker="•"):
    paras = [(f"{marker} {t}", {"space_after": gap}) for t in items]
    return text(slide, x, y, w, h, paras, size=size, color=color)


def header(slide, title, kicker=None, n=None):
    rect(slide, 0, 0, SW, SH, fill=BG)
    rect(slide, 0, 0, 0.18, SH, fill=GREEN)
    if kicker:
        text(slide, MX, 0.38, 9, 0.3, kicker.upper(), size=11, bold=True, color=GREEN)
    text(slide, MX, 0.65, SW - 2 * MX, 0.7, title, size=24, bold=True, color=INK, anchor=MSO_ANCHOR.TOP)
    if n:
        text(slide, SW - 1.2, SH - 0.5, 0.7, 0.3, str(n), size=10, color=MUTED, align=PP_ALIGN.RIGHT)
    text(slide, MX, SH - 0.5, 9, 0.3, "Проект: карта и приоритеты процессов бухгалтерии · черновик для обсуждения",
         size=10, color=MUTED)


def card(slide, x, y, w, h, title, body=None, accent=GREEN, fill=WHITE, tsize=16, bsize=14, items=None):
    rect(slide, x, y, w, h, fill=fill, line=LINE)
    rect(slide, x, y, 0.08, h, fill=accent)
    text(slide, x + 0.2, y + 0.12, w - 0.35, 0.5, title, size=tsize, bold=True, color=INK)
    if items:
        bullets(slide, x + 0.2, y + 0.12 + (0.42 if len(title) < 38 else 0.7), w - 0.35, h - 0.7, items, size=bsize, gap=4)
    elif body:
        text(slide, x + 0.2, y + 0.12 + (0.42 if len(title) < 38 else 0.7), w - 0.35, h - 0.7, body, size=bsize, color=MUTED)


def stat(slide, x, y, w, value, label, color=GREEN, h=1.6):
    rect(slide, x, y, w, h, fill=WHITE, line=LINE)
    text(slide, x + 0.1, y + 0.1, w - 0.2, 0.65, value, size=30, bold=True, color=color)
    text(slide, x + 0.1, y + 0.78, w - 0.2, h - 0.8, label, size=12, color=MUTED)


def notes(slide, s):
    slide.notes_slide.notes_text_frame.text = s


def table(slide, x, y, w, col_w, rows, header_row=True, size=11.5, row_h=0.4, head_fill=GREEN):
    nrows, ncols = len(rows), len(rows[0])
    gt = slide.shapes.add_table(nrows, ncols, Inches(x), Inches(y), Inches(w), Inches(row_h * nrows))
    tbl = gt.table
    tblPr = gt._element.graphic.graphicData.tbl.tblPr
    tblPr.set("bandRow", "0")
    tblPr.set("firstRow", "0")
    for i, cw in enumerate(col_w):
        tbl.columns[i].width = Inches(cw)
    for r in range(nrows):
        tbl.rows[r].height = Inches(0.45 if (header_row and r == 0) else row_h)
        for c in range(ncols):
            cell = tbl.cell(r, c)
            cell.margin_left = cell.margin_right = Inches(0.08)
            cell.margin_top = cell.margin_bottom = Inches(0.04)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            val = rows[r][c]
            tf = cell.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            run = p.add_run()
            run.text = val
            run.font.name = FONT
            run.font.size = Pt(size)
            is_head = header_row and r == 0
            run.font.bold = is_head or c == 0
            run.font.color.rgb = WHITE if is_head else INK
            cell.fill.solid()
            cell.fill.fore_color.rgb = head_fill if is_head else (WHITE if r % 2 else RGBColor(0xEC, 0xF1, 0xEF))
    return tbl


# ---------------------------------------------------------------- 0. Титул
s = prs.slides.add_slide(BLANK)
rect(s, 0, 0, SW, SH, fill=GREEN)
rect(s, 0.6, 2.2, 0.12, 1.9, fill=ORANGE)
text(s, 0.95, 1.2, 11, 0.4, "ПРОЦЕССНЫЙ ОФИС · ДБиНУ TS И БУХГАЛТЕРИЯ КТ", size=13, bold=True, color=GREEN_L)
text(s, 0.95, 2.0, 11.5, 2.8, ["Карта и приоритеты процессов бухгалтерии",
                                ("От карты процессов к плану оптимизации и автоматизации", {"size": 22, "bold": False, "color": GREEN_L})],
     size=40, bold=True, color=WHITE, spacing=1.15, space_after=10)
text(s, 0.95, 5.3, 11, 0.9, ["Верхнеуровневая презентация проекта · черновик для обсуждения",
                             "Основа: карта процессов бухгалтерии v0.1 (05.11.2026), методика KT_METHOD_v1"],
     size=14, color=GREEN_L, space_after=4)
notes(s, "Титульный слайд. Данные дальше взяты из карты процессов бухгалтерии v0.1, методики KT_METHOD_v1 и расчёта приоритетов (prioritization/).")

# ---------------------------------------------------------------- 1. Что за проект
s = prs.slides.add_slide(BLANK)
header(s, "Что за проект: единая модель процессов и план улучшения", "1 · Что за проект", 2)
text(s, MX, 1.65, 12.1, 0.8,
     "Описать, как на самом деле устроена бухгалтерия группы, по единой методике, и на этой основе выбрать "
     "процессы, которые стоит оптимизировать и автоматизировать в первую очередь.", size=16, color=INK)
cw, gap, y0, ch = 3.9, 0.2, 2.65, 2.85
card(s, MX, y0, cw, ch, "1. Карта процессов",
     items=["10 процессов, 29 подпроцессов, 84 этапа (уровни по методике: процесс - подпроцесс - этап - функция)",
            "Собрана снизу вверх из должностных инструкций и матрицы компетенций"], accent=GREEN)
card(s, MX + cw + gap, y0, cw, ch, "2. Приоритеты",
     items=["Оценка 81 этапа по 10 критериям: метрики методики и дополнительные",
            "Итог: 12 направлений, из них 6 первоочередных"], accent=BLUE)
card(s, MX + 2 * (cw + gap), y0, cw, ch, "3. План и TO BE",
     items=["Целевая модель: описание, измерение, автоматизация",
            "Дорожная карта волн работ с контрольными точками"], accent=ORANGE)
rect(s, MX, 5.7, 12.1, 1.05, fill=GREEN_L)
text(s, MX + 0.2, 5.85, 11.7, 0.85,
     [("Границы: ", {"bold": True}),
      ],
     size=12.5)
text(s, MX + 1.1, 5.88, 10.9, 0.85,
     "бухгалтерия группы: КТ (ЦА и филиалы), ST, SP, IPT, TS, OL, ZRT и малые ТОО. Вне границ: бюджетирование, закуп, "
     "финансово-экономические модели (бухгалтерия там только согласует). Реестр бизнес-процессов группы в эту версию не входит.",
     size=12.5, color=INK)
notes(s, "Источник: карта v0.1, вкладка «Метод». Состав компаний - вкладка «Функционал × КП», граница модели - вкладки «SimBase × КП» и «Confluence × КП».")

# ---------------------------------------------------------------- 2. Почему взялись
s = prs.slides.add_slide(BLANK)
header(s, "Почему взялись: единой картины нет, а нагрузка и риски растут", "2 · Почему взялись", 3)
rows = [
    ("Нет единой модели",
     "Должностные инструкции утверждены в 2024 году под прежнюю структуру. У казначейства, отдела оплаты труда и ДВОУ и отдела ST инструкций нет."),
    ("Две модели исполнения",
     "В центре участки по 9-17 этапов, в филиалах «бухгалтер в одном лице»: 9 универсальных позиций ведут 21-29 этапов каждая."),
    ("Регламентов мало там, где риск",
     "Регламент или ВНД есть на 20 из 81 этапа. Налоги, зарплата, инвентаризация, сверки, касса и учётная политика регламентов не имеют."),
    ("Методика требует больше, чем есть",
     "Для процесса нужны паспорт, SIPOC, KPI, SLA, LOA. Для бухгалтерии этого пока нет, а реестр группы не содержит учёт, налоги и закрытие периода."),
    ("Зависимость от людей",
     "Около 29,5 FTE обслуживают 17 активных ТОО и филиалов. 7 из 30 позиций значатся вакансиями или совмещениями."),
    ("Большой ручной поток",
     "В SimBase 9 844 документа согласований по бухгалтерии. Основные потоки проходят через 1-2 согласующих, на договорах их больше."),
]
cw, ch = 5.95, 1.45
for i, (t, b) in enumerate(rows):
    x = MX + (i % 2) * (cw + 0.2)
    y = 1.75 + (i // 2) * (ch + 0.15)
    rect(s, x, y, cw, ch, fill=WHITE, line=LINE)
    rect(s, x, y, 0.08, ch, fill=ORANGE if i in (2, 4) else GREEN)
    text(s, x + 0.2, y + 0.08, cw - 0.3, 0.35, t, size=14.5, bold=True)
    text(s, x + 0.2, y + 0.46, cw - 0.3, ch - 0.5, b, size=13.5, color=MUTED)
notes(s, "Причины восстановлены по материалам карты: вкладки «Метод», «Должности и респонденты», «Функционал × КП», «Confluence × КП» (пробелы), «SimBase × КП». "
         "Формулировку мотивации проекта стоит уточнить у заказчика.")

# ---------------------------------------------------------------- 3. Ситуация сейчас (цифры)
s = prs.slides.add_slide(BLANK)
header(s, "Ситуация: процессы видны, но описаны и измерены неравномерно", "3 · Какая ситуация сейчас", 4)
sx, sw_ = MX, 2.85
stat(s, sx, 1.7, sw_, "84", "этапа в 10 процессах карты v0.1 (81 без блока управления)")
stat(s, sx + (sw_ + 0.23), 1.7, sw_, "29,5", "FTE: 30 позиций по функционалу на 01.07.2026, из них 7 вакансий и совмещений", color=ORANGE)
stat(s, sx + 2 * (sw_ + 0.23), 1.7, sw_, "17", "активных ТОО и филиалов, 63 формы ФНО", color=BLUE)
stat(s, sx + 3 * (sw_ + 0.23), 1.7, sw_, "9 844", "документа согласований SimBase, 89% легли на этапы карты")

cd = CategoryChartData()
cd.categories = ["Регламент / ВНД", "Только инструкции", "Только матрицы", "Нет документов"]
cd.add_series("Этапов", (20, 19, 3, 39))
gf = s.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED, Inches(MX), Inches(3.5), Inches(6.2), Inches(3.3), cd)
ch_ = gf.chart
ch_.has_legend = False
ch_.has_title = True
ch_.chart_title.text_frame.text = "Чем регламентированы 81 этап (Confluence)"
ch_.chart_title.text_frame.paragraphs[0].runs[0].font.size = Pt(13)
ch_.chart_title.text_frame.paragraphs[0].runs[0].font.bold = True
ch_.chart_title.text_frame.paragraphs[0].runs[0].font.name = FONT
pl = ch_.plots[0]
pl.gap_width = 60
pl.has_data_labels = True
pl.data_labels.font.size = Pt(12)
pl.data_labels.font.name = FONT
pl.data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
ser = pl.series[0]
ser.format.fill.solid()
ser.format.fill.fore_color.rgb = GREEN
for idx, col in enumerate([GREEN, BLUE, ORANGE, RED]):
    pt = ser.points[idx]
    pt.format.fill.solid()
    pt.format.fill.fore_color.rgb = col
ch_.category_axis.tick_labels.font.size = Pt(12)
ch_.category_axis.tick_labels.font.name = FONT
ch_.category_axis.reverse_order = True
ch_.category_axis.format.line.color.rgb = LINE
ch_.value_axis.visible = False
ch_.value_axis.has_major_gridlines = False

rect(s, 7.1, 3.5, 5.6, 3.3, fill=WHITE, line=LINE)
text(s, 7.3, 3.6, 5.2, 0.35, "Как этапы подтверждены источниками", size=13, bold=True)
bullets(s, 7.3, 4.0, 5.2, 2.75, [
    "56 этапов подтверждают и инструкции, и матрица компетенций",
    "23 этапа есть только в матрице: работа идёт, в инструкциях её нет",
    "5 этапов есть только в инструкциях: респонденты их не описали",
    "Матрицу заполнили 17 человек из примерно 33 позиций, поэтому часть этапов может быть недоописана",
    "Метрик (время цикла, дефекты, стоимость, LOA) в карте пока нет",
], size=12, gap=5)
notes(s, "Источники: карта v0.1 (шапка, вкладки «Функционал × КП», «Confluence × КП», «SimBase × КП», «Метод»). Регламентация считается по тем же правилам, что и в карте, без блока 99 (81 этап).")

# ---------------------------------------------------------------- 4. Ситуация: главные проблемы
s = prs.slides.add_slide(BLANK)
header(s, "Ситуация сейчас: где основные пробелы", "3 · Какая ситуация сейчас", 5)
rows = [
    ["Область", "Что видно по данным карты", "Чем это грозит"],
    ["Налоговая отчётность", "В Confluence один документ по налогам (алгоритм строки ФНО 100.00); нет налогового календаря и регламентов по НДС, ИПН, СН", "Штрафы, зависимость от отдельных людей; позиция налоговой группы значится вакансией"],
    ["Зарплата и ФОТ", "Нет регламента расчёта по юрлицам; у отдела оплаты труда и ДВОУ нет ДИ; в SimBase 481 выплата ЗП и 440 отпускных", "Ошибки в расчётах, ручной стык HR - ЗУП - банк"],
    ["Документы и платежи", "1 800 подтверждений АВР и 2 826 заявок на оплату; положение о платежах 2021 года, инструкция по СНТ 2022 года не открывается", "Рутинная проверка вручную, устаревшие правила"],
    ["Копии регламентов", "Регламенты продублированы для КТ, SP и ST; копии 2023-2024, версии КТ обновлены в 2026", "Расхождение правил внутри группы"],
    ["Согласования", "Договоры: 934 документа, в среднем 1,7 согласующего, заместитель директора в 72-96%; два входа в один этап (АО, договоры)", "Очереди, перегрузка руководителей"],
    ["Данные для управления", "233 заявки СЗ и 144 ЗО без темы; метрик процессов нет", "Нельзя измерять потоки и считать эффект автоматизации"],
]
table(s, MX, 1.5, 12.1, [2.2, 6.2, 3.7], rows, size=12, row_h=0.78)
notes(s, "Источники: вкладки «Confluence × КП» (пробелы и качество данных), «SimBase × КП», «Функционал × КП». Цифры SimBase - выгрузка по бухгалтерии в карте, период указан в самой карте.")

# ---------------------------------------------------------------- 5. Что уже сделано
s = prs.slides.add_slide(BLANK)
header(s, "Что уже сделано: от методики до приоритетных направлений", "4 · Что уже сделано", 6)
steps = [
    ("Методика", "Изучена KT_METHOD_v1: уровни модели, метрики, LOA, зрелость", GREEN),
    ("Карта v0.1", "10 процессов, 29 подпроцессов, 84 этапа; 703 функции матрицы, 17 респондентов, 19 ДИ", GREEN),
    ("Наложение данных", "SimBase (9 844 док.), функционал (30 позиций), Confluence (91 документ)", GREEN),
    ("Сверка с реестром v0.8", "78 этапов реестра сопоставлены с картой, найдено 18 новых этапов", GREEN),
    ("Приоритизация", "81 этап, 10 критериев, проверка устойчивости (2 000 прогонов весов)", BLUE),
    ("Результат", "12 направлений, 6 первоочередных, быстрые улучшения, исключения", BLUE),
]
cw = 1.9
for i, (t, b, col) in enumerate(steps):
    x = MX + i * (cw + 0.14)
    rect(s, x, 1.7, cw, 0.75, fill=col, shape=MSO_SHAPE.PENTAGON)
    text(s, x + 0.08, 1.72, cw - 0.4, 0.7, f"{i + 1}. {t}", size=11.5, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
    rect(s, x, 2.55, cw, 2.15, fill=WHITE, line=LINE)
    text(s, x + 0.08, 2.63, cw - 0.16, 2.05, b, size=12.5, color=MUTED)
rect(s, MX, 4.9, 12.1, 1.75, fill=GREEN_L)
text(s, MX + 0.2, 4.98, 11.7, 0.35, "Первые выводы", size=14, bold=True, color=GREEN)
bullets(s, MX + 0.2, 5.36, 11.7, 1.3, [
    "Первые направления: налоги, зарплата и налоги с ФОТ, входящие первички и ЭСФ, платёжный контур, реализация, согласование договоров",
    "Автоматизировать можно 33 этапа сразу, 23 этапа сначала нужно стандартизировать; 25 этапов требуют оптимизации маршрута или наблюдения",
    "Статус: карта v0.1 и приоритеты - черновик, владельцы и состав этапов не утверждены, баллы по 5 критериям экспертные",
], size=12, gap=3)
notes(s, "Источники: карта v0.1, PR #1 (prioritization/priority_processes.md, priority_scoring.xlsx, build_priorities.py).")

# ---------------------------------------------------------------- 6. Приоритеты (выход)
s = prs.slides.add_slide(BLANK)
header(s, "Результат приоритизации: шесть направлений в первую очередь", "4 · Что уже сделано", 7)
rows = [
    ["№", "Направление", "Действие", "Главная причина"],
    ["1", "Налоговая отчётность", "Стандартизировать, затем автоматизировать", "Максимальный регуляторный риск, минимум регламентов"],
    ["2", "Зарплата и налоги с ФОТ", "Автоматизировать цепочку кадры - ЗУП - банк - ФНО", "Регулярный большой поток, регламента расчёта нет"],
    ["3", "Входящие первички и ЭСФ/ЭАВР", "Автоматизировать приём и сопоставление", "Самый крупный поток: 1 800 документов"],
    ["4", "Платёжный контур", "Автопроверки, автоодобрение регулярных платежей", "2 826 заявок, 29% потока SimBase"],
    ["5", "Реализация и выписка ЭСФ", "Автоматизировать формирование АВР и ЭСФ", "Один из двух максимальных индексов (78 из 100)"],
    ["6", "Согласование договоров", "Оптимизировать маршрут", "1,7 согласующего на документ против 1,0-1,3 в учётных потоках"],
]
table(s, MX, 1.5, 12.1, [0.5, 3.3, 4.2, 4.1], rows, size=12, row_h=0.62)
text(s, MX, 6.05, 12.1, 0.8,
     "Далее: подотчёт, закрытие периода и ФО, сверки, инвентаризация, статистика, компенсация по ГЧП. "
     "Ключевой этап каждого из шести направлений остаётся в топ-20 минимум в 95% вариантов весов (±40%).",
     size=12, color=MUTED)
notes(s, "Источник: prioritization/priority_processes.md, раздел 1 и 4. Индекс считается как сумма вес x (балл - 1) / 4 по 10 критериям.")

# ---------------------------------------------------------------- 7. Что будет сделано
s = prs.slides.add_slide(BLANK)
header(s, "Что будет сделано: пять направлений работ", "5 · Что будет сделано", 8)
ws = [
    ("A. Утвердить карту", GREEN, [
        "Интервью с владельцами этапов и ИТ-владельцем 1С",
        "Назначить бизнес-владельцев процессов и подпроцессов, RACI",
        "Закрыть открытые вопросы карты, сверить с реестром группы"]),
    ("B. Измерить базу", BLUE, [
        "Для топ-15 этапов: время цикла и ожидание, возвраты, стоимость, LOA и охват автоматизации",
        "Данные из дат SimBase и учётных систем",
        "Заменить экспертные баллы фактом"]),
    ("C. Стандартизировать", ORANGE, [
        "Налоговый календарь и регламенты по налогам и ФОТ",
        "Регламенты инвентаризации, сверок, расчёта ЗП, кассы",
        "ДИ для казначейства, оплаты труда и ST; типовые регламенты группы"]),
    ("D. Оптимизировать и автоматизировать", GREEN, [
        "Быстрые улучшения: обходной лист, два входа АО и договоров, обязательные темы",
        "Автоматизация шести направлений по волнам, с оценкой ROI",
        "Паспорт процесса и SIPOC для каждого направления"]),
    ("E. Управлять результатом", BLUE, [
        "Ежеквартальная оценка зрелости (методика гл. 15)",
        "Отчётность владельцев о метриках",
        "Актуализация карты и приоритетов"]),
]
cw = 2.3
for i, (t, col, items) in enumerate(ws):
    x = MX + i * (cw + 0.15)
    rect(s, x, 1.75, cw, 0.8, fill=col)
    text(s, x + 0.1, 1.78, cw - 0.2, 0.75, t, size=13.5, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
    rect(s, x, 2.55, cw, 4.1, fill=WHITE, line=LINE)
    bullets(s, x + 0.1, 2.68, cw - 0.2, 3.9, items, size=12.5, gap=8)
notes(s, "План собран из методики (цикл управления процессами, гл. 15), открытых вопросов карты и выводов приоритизации. Сроки не заданы: их нужно согласовать после интервью и замеров.")

# ---------------------------------------------------------------- 8. TO BE (модель)
s = prs.slides.add_slide(BLANK)
header(s, "TO BE: бухгалтерия работает по единой модели и измеряется", "6 · Каким будет TO BE", 9)
rows = [
    ["Аспект", "Сейчас (AS IS)", "Целевое состояние (TO BE)"],
    ["Модель процессов", "Карта v0.1, владельцы по гипотезе, состав этапов не утверждён", "Утверждённые процессы, подпроцессы и этапы, назначенные бизнес-владельцы, RACI"],
    ["Описание", "Регламент на 20 из 81 этапа; ДИ 2024 года", "Паспорт процесса, SIPOC и актуальные ВНД по приоритетным направлениям; типовые регламенты группы"],
    ["Метрики", "Замеров нет", "KPI и SLA измеряются: срок ≥ 95% в срок, дефекты < 2%, возвраты < 5%, выход с первого раза ≥ 90%"],
    ["Автоматизация", "Уровень автоматизации (охват и LOA) не измерен", "Охват автоматизации ≥ 70%, средний LOA ≥ 7 из 10 на приоритетных процессах"],
    ["Исполнение", "Две модели: участки в ЦА и универсалы в филиалах", "Единые стандарты этапов с вариантами для ЦА и филиалов; меньше зависимости от отдельных людей"],
    ["Управление", "Разовая оценка", "Ежеквартальная оценка зрелости (цель не ниже 3 «Определённый» для приоритетных процессов), план улучшений"],
]
table(s, MX, 1.5, 12.1, [2.0, 4.4, 5.7], rows, size=11.5, row_h=0.72)
text(s, MX, 6.6, 12.1, 0.4, "Целевые значения метрик взяты из методики KT_METHOD_v1 (гл. 9-13, 15); значения по каждому процессу уточняются после замеров.",
     size=10.5, color=MUTED)
notes(s, "Целевые значения: On-Time/SLA ≥ 95%, Defect Rate < 2%, Rework Rate < 5%, FPY ≥ 90%, Automation Rate ≥ 70%, средний LOA ≥ 7, зрелость по модели из 5 уровней (табл. 19).")

# ---------------------------------------------------------------- 9. TO BE (по направлениям)
s = prs.slides.add_slide(BLANK)
header(s, "TO BE по первоочередным направлениям: как будет работать", "6 · Каким будет TO BE", 10)
rows = [
    ["Направление", "Как будет работать"],
    ["Налоги", "Налоговый календарь с напоминаниями; ФНО собирается из регистров 1С, контрольные соотношения проверяются автоматически; заявка на оплату налога формируется системой"],
    ["Зарплата и ФОТ", "Кадровые данные идут из HR в ЗУП без ручного ввода; ЗО и платёжные поручения формируются автоматически; человек подтверждает выплату"],
    ["Первички и ЭСФ", "ЭСФ/ЭАВР загружаются из ИС ЭСФ в 1С, сопоставляются с договором и заявкой, при совпадении проводятся автоматически; входящий НДС сверяется с ФНО"],
    ["Платежи", "Правила автопроверки по виду платежа; регулярные платежи в пределах лимита одобряются автоматически с правом вето; выписки разносятся без ручного ввода"],
    ["Реализация и ЭСФ", "АВР и ЭСФ формируются по договору и графику услуг; единый стандарт для КТ, SP и ST"],
    ["Договоры", "Маршрут по сумме и типу договора, чек-листы для типовых; заместитель директора согласует только нетиповые"],
]
table(s, MX, 1.5, 12.1, [2.4, 9.7], rows, size=12, row_h=0.72)
text(s, MX, 6.6, 12.1, 0.4, "Решения по интеграциям (1С, ИС ЭСФ, ЗУП, клиент-банк, SimBase) требуют проверки с ИТ-владельцем.", size=10.5, color=MUTED)
notes(s, "Источник: раздел 4 отчёта prioritization/priority_processes.md. Описание - целевая гипотеза, не проектное решение.")

# ---------------------------------------------------------------- 10. Дорожная карта
s = prs.slides.add_slide(BLANK)
header(s, "Дорожная карта: пять волн с контрольными точками", "7 · Дорожная карта", 11)
waves = ["Волна 0\nОснова", "Волна 1\nБыстрые улучшения", "Волна 2\nСтандарты", "Волна 3\nАвтоматизация ядра", "Волна 4\nМасштаб"]
x0, lw, cw_ = MX + 2.3, 2.3, 1.95
for i, w_ in enumerate(waves):
    rect(s, x0 + i * (cw_ + 0.06), 1.7, cw_, 0.7, fill=GREEN, shape=MSO_SHAPE.PENTAGON)
    text(s, x0 + i * (cw_ + 0.06) + 0.08, 1.72, cw_ - 0.3, 0.66, w_.split("\n"), size=11.5, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE, spacing=1.0)
lanes = [
    ("A. Карта", [(0, 1, "Интервью, владельцы, RACI", GREEN), (3, 4, "Актуализация карты", GREEN)]),
    ("B. Замеры", [(0, 2, "База метрик по топ-15 этапам", BLUE), (3, 4, "Мониторинг KPI и SLA", BLUE)]),
    ("C. Стандарты", [(1, 3, "Налоговый календарь, регламенты ЗП, инвентаризации, сверок; ДИ", ORANGE)]),
    ("D. Улучшения", [(1, 1, "Быстрые улучшения", GREEN), (3, 3, "Налоги, ЗП, первички, платежи, реализация", GREEN),
                      (4, 4, "Закрытие периода, инвентаризация, сверки, статистика", GREEN)]),
    ("E. Управление", [(0, 0, "Критерии и веса", BLUE), (2, 4, "Ежеквартальная оценка зрелости", BLUE)]),
]
ly = 2.55
lh = 0.55
for r, (name, bars) in enumerate(lanes):
    y = ly + r * (lh + 0.1)
    rect(s, MX, y, 2.2, lh, fill=WHITE, line=LINE)
    text(s, MX + 0.1, y, 2.0, lh, name, size=12, bold=True, anchor=MSO_ANCHOR.MIDDLE)
    for (a, b_, label, col) in bars:
        bx = x0 + a * (cw_ + 0.06)
        bw = (b_ - a + 1) * cw_ + (b_ - a) * 0.06
        rect(s, bx, y, bw, lh, fill=col, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        text(s, bx + 0.05, y, bw - 0.1, lh, label, size=10, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER, spacing=1.0)
gy = ly + 5 * (lh + 0.1) + 0.05
rect(s, MX, gy, 12.1, 0.95, fill=GREEN_L)
text(s, MX + 0.15, gy + 0.04, 11.8, 0.9, [
    ("Контрольные точки (переход к следующей волне):", {"bold": True, "size": 11.5, "color": GREEN}),
    ("0 → 1: утверждены владельцы и критерии.   1 → 2: подтверждена база метрик.   2 → 3: есть регламент и паспорт процесса.   "
     "3 → 4: достигнуты целевые AR и LOA на ядре.", {"size": 11.5}),
], spacing=1.05)
notes(s, "Сроки волн не заданы намеренно: их нужно согласовать после интервью и замеров. Квартальный ритм оценки зрелости задан методикой (гл. 15). "
         "Принцип: автоматизировать только описанный и стабильный процесс (зрелость не ниже 2).")

# ---------------------------------------------------------------- 11. Решения и риски
s = prs.slides.add_slide(BLANK)
header(s, "Что нужно решить сейчас и какие есть риски", "Следующие шаги", 12)
card(s, MX, 1.75, 5.95, 4.9, "Решения, которые нужны от заказчика", accent=GREEN, items=[
    "Утвердить критерии и веса приоритизации",
    "Назначить владельцев процессов и ИТ-владельца 1С",
    "Выбрать направления для замеров и пилота (рекомендация: налоги, зарплата, первички)",
    "Дать доступ к данным: даты SimBase, отчёты 1С и ЗУП, содержимое Confluence",
    "Определить формат и ритм отчётности владельцев",
], bsize=15)
card(s, MX + 6.15, 1.75, 5.95, 4.9, "Риски", accent=ORANGE, items=[
    "Нет замеров: экспертные баллы по 5 критериям могут сместить порядок",
    "Матрицу заполнили 17 из примерно 33 позиций: карта может быть неполной",
    "Вакансии в налоговой группе, оплате труда и учёте проекта ST замедляют интервью и внедрение",
    "Копии регламентов по юрлицам расходятся: унификация потребует согласований",
    "Интеграции 1С, ИС ЭСФ, ЗУП и клиент-банка не проверены с ИТ",
], bsize=15)
notes(s, "Риски взяты из раздела «Ограничения» отчёта prioritization/priority_processes.md и из ограничений карты v0.1.")

out = sys.argv[1] if len(sys.argv) > 1 else str(Path(__file__).parent / "accounting_processes_deck.pptx")
prs.save(out)
print("saved", out)
