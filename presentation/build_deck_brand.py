"""Сборка (в брендбуке Sergek Group) верхнеуровневой презентации по проекту «Карта и приоритеты процессов бухгалтерии».

    python3 presentation/build_deck_brand.py [выходной.pptx]
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

NAVY = RGBColor(0x01, 0x3D, 0x85)
SLATE = RGBColor(0x1C, 0x33, 0x41)
INK = RGBColor(0x1E, 0x1E, 0x1E)
MUTED = RGBColor(0x4E, 0x5F, 0x69)
GRAY = RGBColor(0xA7, 0xB6, 0xBD)
LINE = RGBColor(0xD4, 0xD8, 0xE4)
PANEL = RGBColor(0xF3, 0xF5, 0xF9)
TINT = RGBColor(0xE6, 0xEB, 0xF4)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
ACCENT = RGBColor(0x28, 0xAB, 0x6A)
RED = RGBColor(0xE9, 0x42, 0x35)
LBLUE = RGBColor(0x43, 0x84, 0xF4)

F_REG, F_MED, F_SEMI = "Golos Text", "Golos Text Medium", "Golos Text SemiBold"
FOOTER = "Все права защищены Sergek Group ©"
ASSETS = Path(__file__).parent / "assets"

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
BLANK = prs.slide_layouts[6]
SW, SH = 13.333, 7.5
MX = 0.6
LOGO_RATIO = 1400 / 211


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


def style_run(run, size, bold, color, face=None):
    f = run.font
    f.name = face or (F_SEMI if bold else F_REG)
    f.size = Pt(size)
    f.bold = False
    f.color.rgb = color


def text(slide, x, y, w, h, content, size=14, bold=False, color=INK, align=PP_ALIGN.LEFT,
         anchor=MSO_ANCHOR.TOP, spacing=1.1, space_after=0, face=None):
    """content: строка или список абзацев. Абзац - строка или (строка, {опции}).
    Строка абзаца может быть списком сегментов [(текст, цвет), ...] для выделения слов."""
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
        segs = p if isinstance(p, list) else [(p, None)]
        for seg, seg_color in segs:
            run = para.add_run()
            run.text = seg
            style_run(run, opts.get("size", size), opts.get("bold", bold), seg_color or opts.get("color", color),
                      opts.get("face", face))
    return tb


def bullets(slide, x, y, w, h, items, size=14, color=INK, gap=6, marker="•"):
    paras = [(f"{marker} {t}", {"space_after": gap}) for t in items]
    return text(slide, x, y, w, h, paras, size=size, color=color)


def logo(slide, x, y, h, variant="blue"):
    name = "logo_sg_white.png" if variant == "white" else "logo_sg_blue.png"
    slide.shapes.add_picture(str(ASSETS / name), Inches(x), Inches(y), Inches(h * LOGO_RATIO), Inches(h))


def header(slide, title, kicker=None, n=None):
    rect(slide, 0, 0, SW, SH, fill=WHITE)
    logo(slide, MX, 0.3, 0.27)
    if kicker:
        text(slide, SW - MX - 6, 0.3, 6, 0.3, kicker, size=11, color=MUTED, align=PP_ALIGN.RIGHT, face=F_MED)
    text(slide, MX, 0.78, SW - 2 * MX, 0.6, title, size=26, bold=True, color=INK, anchor=MSO_ANCHOR.TOP, spacing=1.0)
    text(slide, MX, SH - 0.45, 8, 0.3, FOOTER, size=10, color=GRAY)
    if n:
        text(slide, SW - 1.2, SH - 0.45, 0.6, 0.3, str(n), size=10, color=GRAY, align=PP_ALIGN.RIGHT)


def card(slide, x, y, w, h, title, body=None, accent=NAVY, fill=PANEL, tsize=16, bsize=14, items=None):
    rect(slide, x, y, w, h, fill=fill)
    rect(slide, x, y, 0.08, h, fill=accent)
    text(slide, x + 0.2, y + 0.12, w - 0.35, 0.5, title, size=tsize, bold=True, color=INK)
    if items:
        bullets(slide, x + 0.2, y + 0.12 + (0.42 if len(title) < 38 else 0.7), w - 0.35, h - 0.7, items, size=bsize, gap=4)
    elif body:
        text(slide, x + 0.2, y + 0.12 + (0.42 if len(title) < 38 else 0.7), w - 0.35, h - 0.7, body, size=bsize, color=MUTED)


def stat(slide, x, y, w, value, label, color=NAVY, h=1.6):
    rect(slide, x, y, w, h, fill=PANEL)
    rect(slide, x, y, w, 0.06, fill=color)
    text(slide, x + 0.1, y + 0.14, w - 0.2, 0.65, value, size=30, bold=True, color=color)
    text(slide, x + 0.1, y + 0.8, w - 0.2, h - 0.82, label, size=12, color=MUTED)


def notes(slide, s):
    slide.notes_slide.notes_text_frame.text = s


def table(slide, x, y, w, col_w, rows, header_row=True, size=11.5, row_h=0.4, head_fill=NAVY):
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
            tf = cell.text_frame
            tf.word_wrap = True
            run = tf.paragraphs[0].add_run()
            run.text = rows[r][c]
            is_head = header_row and r == 0
            style_run(run, size, is_head or c == 0, WHITE if is_head else INK, F_MED if is_head else None)
            cell.fill.solid()
            cell.fill.fore_color.rgb = head_fill if is_head else (WHITE if r % 2 else PANEL)
    return tbl


# ---------------------------------------------------------------- 0. Титул
s = prs.slides.add_slide(BLANK)
rect(s, 0, 0, SW, SH, fill=NAVY)
s.shapes.add_picture(str(ASSETS / "wave_on_blue.png"), Inches(6.6), Inches(3.4), Inches(7.8), Inches(3.834))
logo(s, MX, 0.5, 0.42, "white")
text(s, MX, 1.75, 11, 0.4, "Процессный офис · ДБиНУ TS и бухгалтерия КТ", size=14, color=LINE, face=F_MED)
text(s, MX, 2.3, 10.5, 2.4, [[("Карта и приоритеты процессов ", WHITE), ("бухгалтерии", ACCENT)]],
     size=46, bold=True, spacing=1.05)
text(s, MX, 4.35, 9.5, 0.9, "От карты процессов к плану оптимизации и автоматизации", size=22, color=LINE)
text(s, MX, 5.55, 9.5, 0.9, ["Верхнеуровневая презентация проекта · черновик для обсуждения",
                             "Основа: карта процессов бухгалтерии v0.1 (05.11.2026), методика KT_METHOD_v1"],
     size=13, color=LINE, space_after=4)
text(s, MX, SH - 0.5, 8, 0.3, FOOTER, size=10, color=GRAY)
notes(s, "Титульный слайд. Данные дальше взяты из карты процессов бухгалтерии v0.1, методики KT_METHOD_v1 и расчёта приоритетов (prioritization/). "
         "Оформление по брендбуку Sergek Group: цвета black sea / dark slate, шрифт Golos Text, графический элемент «волна».")

# ---------------------------------------------------------------- 1. Что за проект
s = prs.slides.add_slide(BLANK)
header(s, "Что за проект: единая модель процессов и план улучшения", "1 · Что за проект", 2)
text(s, MX, 1.65, 12.1, 0.8,
     "Описать, как на самом деле устроена бухгалтерия группы, по единой методике, и на этой основе выбрать "
     "процессы, которые стоит оптимизировать и автоматизировать в первую очередь.", size=16, color=INK)
cw, gap, y0, ch = 3.9, 0.2, 2.65, 2.85
card(s, MX, y0, cw, ch, "1. Карта процессов",
     items=["10 процессов, 29 подпроцессов, 84 этапа (уровни по методике: процесс - подпроцесс - этап - функция)",
            "Собрана снизу вверх из должностных инструкций и матрицы компетенций"], accent=NAVY)
card(s, MX + cw + gap, y0, cw, ch, "2. Приоритеты",
     items=["Оценка 81 этапа по 10 критериям: метрики методики и дополнительные",
            "Итог: 12 направлений, из них 6 первоочередных"], accent=SLATE)
card(s, MX + 2 * (cw + gap), y0, cw, ch, "3. План и TO BE",
     items=["Целевая модель: описание, измерение, автоматизация",
            "Дорожная карта волн работ с контрольными точками"], accent=ACCENT)
rect(s, MX, 5.7, 12.1, 1.05, fill=TINT)
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
    rect(s, x, y, cw, ch, fill=PANEL)
    rect(s, x, y, 0.08, ch, fill=ACCENT if i in (2, 4) else NAVY)
    text(s, x + 0.2, y + 0.08, cw - 0.3, 0.35, t, size=14.5, bold=True)
    text(s, x + 0.2, y + 0.46, cw - 0.3, ch - 0.5, b, size=13.5, color=MUTED)
notes(s, "Причины восстановлены по материалам карты: вкладки «Метод», «Должности и респонденты», «Функционал × КП», «Confluence × КП» (пробелы), «SimBase × КП». "
         "Формулировку мотивации проекта стоит уточнить у заказчика.")

# ---------------------------------------------------------------- 3. Ситуация сейчас (цифры)
s = prs.slides.add_slide(BLANK)
header(s, "Ситуация: процессы видны, но описаны неравномерно", "3 · Какая ситуация сейчас", 4)
sx, sw_ = MX, 2.85
stat(s, sx, 1.7, sw_, "84", "этапа в 10 процессах карты v0.1 (81 без блока управления)")
stat(s, sx + (sw_ + 0.23), 1.7, sw_, "29,5", "FTE: 30 позиций на 01.07.2026, из них 7 вакансий и совмещений", color=ACCENT)
stat(s, sx + 2 * (sw_ + 0.23), 1.7, sw_, "17", "активных ТОО и филиалов, 63 формы ФНО", color=SLATE)
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
ch_.chart_title.text_frame.paragraphs[0].runs[0].font.bold = False
ch_.chart_title.text_frame.paragraphs[0].runs[0].font.name = F_REG
pl = ch_.plots[0]
pl.gap_width = 60
pl.has_data_labels = True
pl.data_labels.font.size = Pt(12)
pl.data_labels.font.name = F_REG
pl.data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
ser = pl.series[0]
ser.format.fill.solid()
ser.format.fill.fore_color.rgb = NAVY
for idx, col in enumerate([NAVY, LBLUE, ACCENT, RED]):
    pt = ser.points[idx]
    pt.format.fill.solid()
    pt.format.fill.fore_color.rgb = col
ch_.category_axis.tick_labels.font.size = Pt(12)
ch_.category_axis.tick_labels.font.name = F_REG
ch_.category_axis.reverse_order = True
ch_.category_axis.format.line.color.rgb = LINE
ch_.value_axis.visible = False
ch_.value_axis.has_major_gridlines = False

rect(s, 7.1, 3.5, 5.6, 3.3, fill=PANEL)
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
    ("Методика", "Изучена KT_METHOD_v1: уровни модели, метрики, LOA, зрелость", NAVY),
    ("Карта v0.1", "10 процессов, 29 подпроцессов, 84 этапа; 703 функции матрицы, 17 респондентов, 19 ДИ", NAVY),
    ("Наложение данных", "SimBase (9 844 док.), функционал (30 позиций), Confluence (91 документ)", NAVY),
    ("Сверка с реестром v0.8", "78 этапов реестра сопоставлены с картой, найдено 18 новых этапов", NAVY),
    ("Приоритизация", "81 этап, 10 критериев, проверка устойчивости (2 000 прогонов весов)", SLATE),
    ("Результат", "12 направлений, 6 первоочередных, быстрые улучшения, исключения", SLATE),
]
cw = 1.9
for i, (t, b, col) in enumerate(steps):
    x = MX + i * (cw + 0.14)
    rect(s, x, 1.7, cw, 0.75, fill=col, shape=MSO_SHAPE.PENTAGON)
    text(s, x + 0.08, 1.72, cw - 0.4, 0.7, f"{i + 1}. {t}", size=11.5, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
    rect(s, x, 2.55, cw, 2.15, fill=PANEL)
    text(s, x + 0.08, 2.63, cw - 0.16, 2.05, b, size=12.5, color=MUTED)
rect(s, MX, 4.9, 12.1, 1.75, fill=TINT)
text(s, MX + 0.2, 4.98, 11.7, 0.35, "Первые выводы", size=14, bold=True, color=NAVY)
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
    ("A. Утвердить карту", NAVY, [
        "Интервью с владельцами этапов и ИТ-владельцем 1С",
        "Назначить бизнес-владельцев процессов и подпроцессов, RACI",
        "Закрыть открытые вопросы карты, сверить с реестром группы"]),
    ("B. Измерить базу", SLATE, [
        "Для топ-15 этапов: время цикла и ожидание, возвраты, стоимость, LOA и охват автоматизации",
        "Данные из дат SimBase и учётных систем",
        "Заменить экспертные баллы фактом"]),
    ("C. Стандартизировать", ACCENT, [
        "Налоговый календарь и регламенты по налогам и ФОТ",
        "Регламенты инвентаризации, сверок, расчёта ЗП, кассы",
        "ДИ для казначейства, оплаты труда и ST; типовые регламенты группы"]),
    ("D. Оптимизировать и автоматизировать", NAVY, [
        "Быстрые улучшения: обходной лист, два входа АО и договоров, обязательные темы",
        "Автоматизация шести направлений по волнам, с оценкой ROI",
        "Паспорт процесса и SIPOC для каждого направления"]),
    ("E. Управлять результатом", SLATE, [
        "Ежеквартальная оценка зрелости (методика гл. 15)",
        "Отчётность владельцев о метриках",
        "Актуализация карты и приоритетов"]),
]
cw = 2.3
for i, (t, col, items) in enumerate(ws):
    x = MX + i * (cw + 0.15)
    rect(s, x, 1.75, cw, 0.8, fill=col)
    text(s, x + 0.1, 1.78, cw - 0.2, 0.75, t, size=13.5, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
    rect(s, x, 2.55, cw, 4.1, fill=PANEL)
    bullets(s, x + 0.1, 2.68, cw - 0.2, 3.9, items, size=12.5, gap=8)
notes(s, "План собран из методики (цикл управления процессами, гл. 15), открытых вопросов карты и выводов приоритизации. Сроки не заданы: их нужно согласовать после интервью и замеров.")

# ---------------------------------------------------------------- 8. TO BE (схема)
s = prs.slides.add_slide(BLANK)
header(s, "TO BE: единая измеряемая модель и как она заработает", "6 · Каким будет TO BE", 9)
pillars = [
    ("Описание", "Регламент на 20 из 81 этапов, ДИ 2024 года",
     "Паспорт, SIPOC и ВНД на приоритетных процессах; владельцы и RACI"),
    ("Измерение", "Замеров нет",
     "KPI и SLA считаются: в срок ≥ 95%, дефекты < 2%, возвраты < 5%"),
    ("Автоматизация", "Охват и LOA не измерены",
     "Охват ≥ 70%, средний LOA ≥ 7 из 10 на приоритетных процессах"),
    ("Управление", "Разовая оценка, две модели исполнения",
     "Зрелость ≥ 3, оценка каждый квартал; единые стандарты ЦА и филиалов"),
]
pw, pg = 2.9, 0.167
for i, (name, now, goal) in enumerate(pillars):
    x = MX + i * (pw + pg)
    rect(s, x, 1.5, pw, 0.42, fill=NAVY)
    text(s, x + 0.1, 1.5, pw - 0.2, 0.42, name, size=14, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
    rect(s, x, 1.92, pw, 0.78, fill=PANEL)
    text(s, x + 0.1, 1.95, pw - 0.2, 0.25, "СЕЙЧАС", size=9.5, color=MUTED, face=F_MED)
    text(s, x + 0.1, 2.18, pw - 0.2, 0.5, now, size=12, color=MUTED, spacing=1.0)
    arrow = rect(s, x + pw / 2 - 0.17, 2.74, 0.34, 0.26, fill=ACCENT, shape=MSO_SHAPE.DOWN_ARROW)
    rect(s, x, 3.04, pw, 0.98, fill=TINT)
    rect(s, x, 3.04, 0.06, 0.98, fill=ACCENT)
    text(s, x + 0.16, 3.07, pw - 0.26, 0.25, "ЦЕЛЬ", size=9.5, color=NAVY, face=F_MED)
    text(s, x + 0.16, 3.3, pw - 0.26, 0.72, goal, size=12, color=INK, spacing=1.0)

text(s, MX, 4.2, 12.1, 0.35, "Как заработают шесть первоочередных направлений", size=15, bold=True, color=INK)
dirs = [
    ("Налоги", "Налоговый календарь; ФНО собирается из 1С и проверяется автоматически"),
    ("Зарплата и ФОТ", "Данные HR идут в ЗУП без ручного ввода; ЗО и платёжки формируются сами"),
    ("Первички и ЭСФ", "ЭСФ загружаются в 1С, сверяются с договором и заявкой и проводятся сами"),
    ("Платежи", "Автопроверка по виду платежа; регулярные платежи в лимите одобряются сами"),
    ("Реализация и ЭСФ", "АВР и ЭСФ формируются по договору и графику; единый стандарт КТ, SP, ST"),
    ("Договоры", "Маршрут по сумме и типу; заместитель директора согласует только нетиповые"),
]
dw, dg = 1.9, 0.14
for i, (name, body) in enumerate(dirs):
    x = MX + i * (dw + dg)
    rect(s, x, 4.65, dw, 1.82, fill=PANEL)
    rect(s, x, 4.65, dw, 0.06, fill=SLATE)
    rect(s, x + 0.12, 4.83, 0.34, 0.34, fill=NAVY, shape=MSO_SHAPE.OVAL)
    text(s, x + 0.12, 4.83, 0.34, 0.34, str(i + 1), size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
         anchor=MSO_ANCHOR.MIDDLE, spacing=1.0)
    text(s, x + 0.5, 4.78, dw - 0.55, 0.5, name, size=12, bold=True, color=INK, anchor=MSO_ANCHOR.MIDDLE, spacing=1.0)
    text(s, x + 0.1, 5.3, dw - 0.2, 1.17, body, size=11, color=MUTED, spacing=1.05)
text(s, MX, 6.55, 12.1, 0.3,
     "Целевые значения - из методики KT_METHOD_v1; решения по интеграциям (1С, ИС ЭСФ, ЗУП, клиент-банк, SimBase) нужно проверить с ИТ-владельцем.",
     size=10, color=MUTED)
notes(s, "Целевые значения: On-Time/SLA ≥ 95%, Defect Rate < 2%, Rework Rate < 5%, FPY ≥ 90%, Automation Rate ≥ 70%, средний LOA ≥ 7, зрелость по модели из 5 уровней (табл. 19; цель - не ниже 3 «Определённый» для приоритетных процессов). "
         "Описание по направлениям - целевая гипотеза, не проектное решение (раздел 4 отчёта prioritization/priority_processes.md).")

# ---------------------------------------------------------------- 10. Дорожная карта
s = prs.slides.add_slide(BLANK)
header(s, "Дорожная карта: пять волн с контрольными точками", "7 · Дорожная карта", 10)
waves = ["Волна 0\nОснова", "Волна 1\nБыстрые улучшения", "Волна 2\nСтандарты", "Волна 3\nАвтоматизация ядра", "Волна 4\nМасштаб"]
x0, lw, cw_ = MX + 2.3, 2.3, 1.95
for i, w_ in enumerate(waves):
    rect(s, x0 + i * (cw_ + 0.06), 1.7, cw_, 0.7, fill=NAVY, shape=MSO_SHAPE.PENTAGON)
    text(s, x0 + i * (cw_ + 0.06) + 0.08, 1.72, cw_ - 0.3, 0.66, w_.split("\n"), size=11.5, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE, spacing=1.0)
lanes = [
    ("A. Карта", [(0, 1, "Интервью, владельцы, RACI", NAVY), (3, 4, "Актуализация карты", NAVY)]),
    ("B. Замеры", [(0, 2, "База метрик по топ-15 этапам", SLATE), (3, 4, "Мониторинг KPI и SLA", SLATE)]),
    ("C. Стандарты", [(1, 3, "Налоговый календарь, регламенты ЗП, инвентаризации, сверок; ДИ", ACCENT)]),
    ("D. Улучшения", [(1, 1, "Быстрые улучшения", NAVY), (3, 3, "Налоги, ЗП, первички, платежи, реализация", NAVY),
                      (4, 4, "Закрытие периода, сверки, статистика", NAVY)]),
    ("E. Управление", [(0, 0, "Критерии и веса", SLATE), (2, 4, "Ежеквартальная оценка зрелости", SLATE)]),
]
ly = 2.55
lh = 0.55
for r, (name, bars) in enumerate(lanes):
    y = ly + r * (lh + 0.1)
    rect(s, MX, y, 2.2, lh, fill=PANEL)
    text(s, MX + 0.1, y, 2.0, lh, name, size=12, bold=True, anchor=MSO_ANCHOR.MIDDLE)
    for (a, b_, label, col) in bars:
        bx = x0 + a * (cw_ + 0.06)
        bw = (b_ - a + 1) * cw_ + (b_ - a) * 0.06
        rect(s, bx, y, bw, lh, fill=col, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        text(s, bx + 0.05, y, bw - 0.1, lh, label, size=10, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER, spacing=1.0)
gy = ly + 5 * (lh + 0.1) + 0.05
rect(s, MX, gy, 12.1, 0.95, fill=TINT)
text(s, MX + 0.15, gy + 0.04, 11.8, 0.9, [
    ("Контрольные точки (переход к следующей волне):", {"bold": True, "size": 11.5, "color": NAVY}),
    ("0 → 1: утверждены владельцы и критерии.   1 → 2: подтверждена база метрик.   2 → 3: есть регламент и паспорт процесса.   "
     "3 → 4: достигнуты целевые AR и LOA на ядре.", {"size": 11.5}),
], spacing=1.05)
notes(s, "Сроки волн не заданы намеренно: их нужно согласовать после интервью и замеров. Квартальный ритм оценки зрелости задан методикой (гл. 15). "
         "Принцип: автоматизировать только описанный и стабильный процесс (зрелость не ниже 2).")

# ---------------------------------------------------------------- 11. Решения и риски
s = prs.slides.add_slide(BLANK)
header(s, "Что нужно решить сейчас и какие есть риски", "Следующие шаги", 11)
card(s, MX, 1.6, 5.95, 3.75, "Решения, которые нужны от заказчика", accent=NAVY, items=[
    "Утвердить критерии и веса приоритизации",
    "Назначить владельцев процессов и ИТ-владельца 1С",
    "Выбрать направления для замеров и пилота (рекомендация: налоги, зарплата, первички)",
    "Дать доступ к данным: даты SimBase, отчёты 1С и ЗУП, содержимое Confluence",
    "Определить формат и ритм отчётности владельцев",
], bsize=14)
card(s, MX + 6.15, 1.6, 5.95, 3.75, "Риски", accent=RED, items=[
    "Нет замеров: экспертные баллы по 5 критериям могут сместить порядок",
    "Матрицу заполнили 17 из примерно 33 позиций: карта может быть неполной",
    "Вакансии в налоговой группе, оплате труда и учёте проекта ST замедляют интервью и внедрение",
    "Копии регламентов по юрлицам расходятся: унификация потребует согласований",
    "Интеграции 1С, ИС ЭСФ, ЗУП и клиент-банка не проверены с ИТ",
], bsize=14)
rect(s, MX, 5.6, 12.1, 1.15, fill=NAVY)
s.shapes.add_picture(str(ASSETS / "wave_on_blue.png"), Inches(10.3), Inches(5.6), Inches(2.34), Inches(1.15))
text(s, MX + 0.3, 5.6, 8.6, 1.15, [[("Предложение: ", ACCENT), ("начать с интервью с владельцами этапов и замеров по налогам, зарплате и первичкам (волна 0 и 1)", WHITE)]],
     size=16, anchor=MSO_ANCHOR.MIDDLE, spacing=1.1)
notes(s, "Риски взяты из раздела «Ограничения» отчёта prioritization/priority_processes.md и из ограничений карты v0.1.")

FONT_FILES = [
    ("Golos Text", "GolosText-regular", "GolosText-bold"),
    ("Golos Text Medium", "GolosTextMedium-regular", "GolosTextMedium-bold"),
    ("Golos Text SemiBold", "GolosTextSemiBold-regular", "GolosTextSemiBold-bold"),
]


def embed_fonts(pptx_path):
    """Встраивает Golos Text (файлы из брендированного шаблона) в готовый pptx."""
    import re
    import zipfile

    fonts_dir = ASSETS / "fonts"
    tmp = pptx_path + ".tmp"
    with zipfile.ZipFile(pptx_path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        rels_name = "ppt/_rels/presentation.xml.rels"
        rels = zin.read(rels_name).decode("utf8")
        pres = zin.read("ppt/presentation.xml").decode("utf8")
        ctypes = zin.read("[Content_Types].xml").decode("utf8")
        ids = [int(i) for i in re.findall(r'Id="rId(\d+)"', rels)]
        nxt = max(ids) + 1
        lst = "<p:embeddedFontLst>"
        new_rels = ""
        for face, reg, bold in FONT_FILES:
            lst += f'<p:embeddedFont><p:font typeface="{face}"/>'
            for tag, fn in (("regular", reg), ("bold", bold)):
                rid = f"rId{nxt}"
                nxt += 1
                new_rels += ('<Relationship Id="%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
                             'relationships/font" Target="fonts/%s.fntdata"/>' % (rid, fn))
                lst += f'<p:{tag} r:id="{rid}"/>'
            lst += "</p:embeddedFont>"
        lst += "</p:embeddedFontLst>"
        rels = rels.replace("</Relationships>", new_rels + "</Relationships>")
        pres = re.sub(r"(<p:notesSz[^>]*/>)", r"\1" + lst, pres, count=1)
        if 'Extension="fntdata"' not in ctypes:
            ctypes = ctypes.replace("<Default ", '<Default Extension="fntdata" ContentType="application/x-fontdata"/><Default ', 1)
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == rels_name:
                data = rels.encode("utf8")
            elif item.filename == "ppt/presentation.xml":
                data = pres.encode("utf8")
            elif item.filename == "[Content_Types].xml":
                data = ctypes.encode("utf8")
            zout.writestr(item, data)
        for _, reg, bold in FONT_FILES:
            for fn in (reg, bold):
                zout.write(fonts_dir / f"{fn}.fntdata", f"ppt/fonts/{fn}.fntdata")
    Path(tmp).replace(pptx_path)


out = sys.argv[1] if len(sys.argv) > 1 else str(Path(__file__).parent / "accounting_processes_deck_brand.pptx")
prs.save(out)
embed_fonts(out)
print("saved", out)
