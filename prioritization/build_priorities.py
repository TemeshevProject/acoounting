"""Расчёт индекса приоритета этапов карты процессов бухгалтерии v0.1.

Использование:
    python3 prioritization/build_priorities.py <карта.html> [выходной.xlsx]

Данные берутся из блока `const D = {...}` HTML-карты (этапы, потоки SimBase,
функционал и документы Confluence). Экспертные оценки - в expert_scores.py.
"""
import collections
import json
import random
import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

sys.path.insert(0, str(Path(__file__).parent))
from expert_scores import EXCLUDED, EXPERT  # noqa: E402

WEIGHTS = {
    "vol": 10,
    "fte": 12,
    "appr": 8,
    "E": 10,
    "R": 12,
    "K": 8,
    "A": 10,
    "ready": 10,
    "S": 10,
    "repl": 10,
}
SEGMENT_CUT = 55
VALUE_KEYS = ["vol", "fte", "appr", "E", "R", "K", "A"]
FEAS_KEYS = ["ready", "S", "repl"]


def load_map(path):
    text = Path(path).read_text(encoding="utf-8")
    start = text.index("const D = ") + len("const D = ")
    data, _ = json.JSONDecoder().raw_decode(text[start:])
    return data


def bucket(value, thresholds):
    """thresholds: возрастающие границы для баллов 2..5."""
    score = 1
    for t in thresholds:
        if value >= t:
            score += 1
    return score


def collect(D):
    L3 = {x["id"]: x for x in D["L3"]}
    L2 = {x["id"]: x for x in D["L2"]}
    L1 = {x["id"]: x for x in D["L1"]}

    docs_main, docs_rel = collections.Counter(), collections.Counter()
    appr_w, appr_n = collections.Counter(), collections.Counter()
    share_w = collections.defaultdict(lambda: [0.0] * 5)
    share_n = collections.Counter()
    for r in D["sb"]:
        n, cnt, stage, rels, shares, avg = r[0], r[3], r[4], r[5], r[7], r[8]
        if stage:
            docs_main[stage] += cnt
            if avg is not None:
                appr_w[stage] += avg * cnt
                appr_n[stage] += cnt
            if all(v is not None for v in shares):
                for i, v in enumerate(shares):
                    share_w[stage][i] += v * cnt
                share_n[stage] += cnt
        for q in rels:
            docs_rel[q] += cnt

    fte, npos, comps = collections.Counter(), collections.Counter(), collections.defaultdict(set)
    for p in D["fd"]["people"]:
        stages = list(p["st"])
        for s in stages:
            fte[s] += p["fte"] / len(stages)
            npos[s] += 1
            comps[s].update(p["co"].keys())

    resp = collections.defaultdict(set)
    nfn = collections.Counter()
    for r in D["fn"]:
        resp[r[2]].add(r[1])
        nfn[r[2]] += 1

    cf_docs = collections.defaultdict(list)
    for d in D["cf"]["docs"]:
        for s in d[10]:
            cf_docs[s].append(d)

    def cf_level(stage):
        ds = [d for d in cf_docs[stage] if not (d[13] and "чужие" in d[13])]
        if any(d[3] == "Регламент / ВНД" for d in ds):
            return "reg"
        if any(("Инструкция" in d[3] or "Чек-лист" in d[3]) for d in ds):
            return "ins"
        return "mx" if ds else "no"

    return dict(L1=L1, L2=L2, L3=L3, docs_main=docs_main, docs_rel=docs_rel, appr_w=appr_w,
                appr_n=appr_n, share_w=share_w, share_n=share_n, fte=fte, npos=npos,
                comps=comps, resp=resp, nfn=nfn, cf_docs=cf_docs, cf_level=cf_level)


LEVEL_NAMES = {"reg": "регламент/ВНД", "ins": "инструкции", "mx": "только матрицы/шаблоны", "no": "нет документов"}
LEVEL_SCORE = {"reg": 4.5, "ins": 3.5, "mx": 2.5, "no": 1.5}
FLAG_SHORT = {"Оба источника": "ДИ+матрица", "Только матрица": "только матрица", "Только ДИ": "только ДИ", "Нет источников": "нет"}


def score_stage(sid, C):
    st = C["L3"][sid]
    main, rel = C["docs_main"][sid], C["docs_rel"][sid]
    vol_docs = main + 0.5 * rel
    vol = 2 if vol_docs == 0 else bucket(vol_docs, [50, 200, 700, 1500])
    vol = max(vol, 2) if vol_docs else 2

    f = C["fte"][sid]
    fte = bucket(f, [0.2, 0.4, 0.6, 0.9]) if f else 1

    if C["appr_n"][sid]:
        avg = C["appr_w"][sid] / C["appr_n"][sid]
        appr = bucket(avg, [1.1, 1.2, 1.4, 1.7])
        if C["share_n"][sid]:
            sh = [v / C["share_n"][sid] for v in C["share_w"][sid]]
            if sh[1] >= 0.5 or sh[0] >= 0.4:
                appr = min(5, appr + 1)
        else:
            sh = None
    else:
        avg, sh, appr = None, None, 2

    lvl = C["cf_level"](sid)
    ready = LEVEL_SCORE[lvl] + (0.5 if st["flag"] == "Оба источника" else 0)
    ready = min(5.0, ready)

    npos = C["npos"][sid]
    repl = bucket(npos, [3, 5, 8, 12]) if npos else 1

    E, R, K, A, S = EXPERT[sid]
    scores = dict(vol=vol, fte=fte, appr=appr, E=E, R=R, K=K, A=A, ready=ready, S=S, repl=repl)

    def idx(keys):
        w = sum(WEIGHTS[k] for k in keys)
        return sum(WEIGHTS[k] * (scores[k] - 1) / 4 for k in keys) / w * 100

    total = sum(WEIGHTS[k] * (scores[k] - 1) / 4 for k in WEIGHTS)
    value, feas = idx(VALUE_KEYS), idx(FEAS_KEYS)

    if A >= 4 and ready >= 3.5:
        action = "Автоматизировать"
    elif A >= 4:
        action = "Сначала стандартизировать, затем автоматизировать"
    elif E >= 4 or appr >= 4 or vol >= 4:
        action = "Оптимизировать (упростить, убрать лишние согласования)"
    else:
        action = "Оптимизировать по мере ресурсов / наблюдать"

    return dict(id=sid, name=st["n"], flag=st["flag"], docs_main=main, docs_rel=rel, fte=f, npos=npos,
                n_comp=len(C["comps"][sid]), n_resp=len(C["resp"][sid]), n_fn=C["nfn"][sid],
                avg_appr=avg, shares=sh, level=lvl, n_docs_cf=len(C["cf_docs"][sid]),
                scores=scores, total=total, value=value, feas=feas, action=action)


def segment(r, top_value, top_feas):
    hi_v, hi_f = r["value"] >= top_value, r["feas"] >= top_feas
    if hi_v and hi_f:
        return "A. Быстрые победы"
    if hi_v:
        return "B. Стратегические (сначала подготовка)"
    if hi_f:
        return "C. Малые улучшения"
    return "D. Низкий приоритет сейчас"


def build(D):
    C = collect(D)
    rows = [score_stage(s, C) for s in C["L3"] if s not in EXCLUDED]
    rows.sort(key=lambda r: -r["total"])
    top_value, top_feas = SEGMENT_CUT, SEGMENT_CUT
    for i, r in enumerate(rows, 1):
        r["rank"] = i
        r["segment"] = segment(r, top_value, top_feas)
        r["l2"] = C["L2"][r["id"][:5]]["n"]
        r["l1"] = C["L1"][r["id"][:2]]["n"]
    return rows, C, (top_value, top_feas)


INITIATIVES = [
    ("И1", "Платёжный контур: заявка на оплату - платёжное поручение - выписка банка",
     ["04-01-01", "04-01-02", "04-01-03", "04-02-01", "04-02-02"]),
    ("И2", "Входящие первичные документы и ЭСФ/ЭАВР: приём, проверка, оприходование, входящий НДС",
     ["01-01-01", "01-01-02", "01-01-03", "01-01-04", "01-02-01", "01-02-02", "01-02-03"]),
    ("И3", "Расчёт зарплаты и налоги с ФОТ: от кадровых данных до перечисления и ФНО 200",
     ["05-01-01", "05-01-02", "05-01-03", "05-01-04", "05-01-05", "05-02-01", "05-03-01", "05-03-02"]),
    ("И4", "Налоговая отчётность: база, расчёт, ФНО, сдача, уплата, сверка с ОГД",
     ["06-01-01", "06-01-02", "06-01-03", "06-02-01", "06-02-02", "06-02-03", "06-02-04"]),
    ("И5", "Реализация, выписка ЭСФ и расчёты с покупателями",
     ["02-01-01", "02-01-02", "02-01-03", "02-01-04", "02-02-02"]),
    ("И6", "Закрытие периода и финансовая отчётность",
     ["07-01-01", "07-01-02", "07-01-03", "07-02-01", "07-02-02", "07-02-03", "07-04-01"]),
    ("И7", "Подотчёт, командировки и авансовые отчёты",
     ["04-03-01", "04-03-02", "04-03-03", "04-03-04"]),
    ("И8", "Согласование договоров на предмет финансового контроля", ["09-03-01"]),
    ("И9", "Сверки взаиморасчётов с поставщиками и покупателями", ["01-03-01", "01-03-02", "02-02-01"]),
    ("И10", "Инвентаризация и учёт ОС и ТМЗ",
     ["03-01-01", "03-01-02", "03-01-04", "03-02-01", "03-03-01", "03-03-02", "03-03-03"]),
    ("И11", "Статистическая и валютная отчётность", ["08-01-01", "08-01-02", "08-02-01"]),
    ("И12", "Компенсация затрат по проектам ГЧП", ["02-03-01", "02-03-02"]),
]


def robustness(rows, runs=2000, top=20, spread=0.4, seed=7):
    rnd = random.Random(seed)
    hits = collections.Counter()
    for _ in range(runs):
        w = {k: v * rnd.uniform(1 - spread, 1 + spread) for k, v in WEIGHTS.items()}
        ranked = sorted(rows, key=lambda r: -sum(w[k] * (r["scores"][k] - 1) / 4 for k in w))
        for r in ranked[:top]:
            hits[r["id"]] += 1
    for r in rows:
        r["stable"] = hits[r["id"]] / runs * 100


def initiative_stats(rows):
    by_id = {r["id"]: r for r in rows}
    out = []
    for iid, name, stages in INITIATIVES:
        rs = [by_id[s] for s in stages]
        out.append(dict(
            id=iid, name=name, stages=stages, n=len(rs),
            mean_total=sum(r["total"] for r in rs) / len(rs),
            top_total=max(r["total"] for r in rs),
            fte=sum(r["fte"] for r in rs),
            docs=sum(r["docs_main"] for r in rs),
            in_top20=sum(1 for r in rs if r["rank"] <= 20),
            mean_value=sum(r["value"] for r in rs) / len(rs),
            mean_feas=sum(r["feas"] for r in rs) / len(rs),
        ))
    out.sort(key=lambda x: -(x["mean_total"] * 0.5 + x["top_total"] * 0.5))
    return out


HEAD = [
    ("rank", "№", 5), ("id", "ID этапа", 10), ("l1", "Процесс", 30), ("l2", "Подпроцесс", 30), ("name", "Этап", 55),
    ("total", "Индекс приоритета (0-100)", 12), ("value", "Ценность (0-100)", 10), ("feas", "Готовность и простота (0-100)", 12),
    ("segment", "Сегмент", 26), ("action", "Рекомендуемое действие", 36), ("stable", "Устойчивость: доля прогонов в топ-20, %", 14),
]
CRIT = [
    ("vol", "Объём потока (SimBase)", "данные"), ("fte", "Трудозатраты, FTE", "данные"),
    ("appr", "Согласования (число согласующих, доля ГБ/Зам)", "данные"), ("E", "Риск ошибок и переработок", "эксперт"),
    ("R", "Регуляторный и финансовый риск", "эксперт"), ("K", "Влияние на закрытие периода и смежные процессы", "эксперт"),
    ("A", "Потенциал автоматизации", "эксперт"), ("ready", "Регламентированность", "данные"),
    ("S", "Простота внедрения", "эксперт"), ("repl", "Тиражируемость (число исполнителей)", "данные"),
]


def write_xlsx(rows, C, thresholds, path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Рейтинг этапов"
    hdr = [h[1] for h in HEAD] + [c[1] for c in CRIT] + [
        "SimBase: осн. док.", "SimBase: связ. док.", "FTE (оценка)", "Позиций в функционале", "Респондентов матрицы",
        "Ср. число согласующих", "Регламентация (Confluence)", "Покрытие карты"]
    ws.append(hdr)
    for r in rows:
        line = [r["rank"], r["id"], r["l1"], r["l2"], r["name"], round(r["total"], 1), round(r["value"], 1),
                round(r["feas"], 1), r["segment"], r["action"], round(r["stable"])]
        line += [r["scores"][c[0]] for c in CRIT]
        line += [r["docs_main"], r["docs_rel"], round(r["fte"], 2), r["npos"], r["n_resp"],
                 round(r["avg_appr"], 2) if r["avg_appr"] else None, LEVEL_NAMES[r["level"]], FLAG_SHORT[r["flag"]]]
        ws.append(line)
    fill = PatternFill("solid", fgColor="16624D")
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = fill
        c.alignment = Alignment(wrap_text=True, vertical="center")
    widths = [h[2] for h in HEAD] + [14] * len(CRIT) + [12, 12, 10, 12, 12, 12, 22, 16]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "F2"
    ws.auto_filter.ref = ws.dimensions
    ws.row_dimensions[1].height = 60

    ini = wb.create_sheet("Инициативы")
    ini.append(["№", "Инициатива (группа этапов)", "Этапов", "Средний индекс", "Макс. индекс", "Этапов в топ-20",
                "Документов SimBase (осн.)", "FTE (оценка)", "Этапы"])
    for x in initiative_stats(rows):
        ini.append([x["id"], x["name"], x["n"], round(x["mean_total"], 1), round(x["top_total"], 1), x["in_top20"],
                    x["docs"], round(x["fte"], 1), ", ".join(x["stages"])])
    for c in ini[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = fill
        c.alignment = Alignment(wrap_text=True, vertical="center")
    for i, w in enumerate([6, 70, 8, 12, 12, 12, 14, 10, 70], 1):
        ini.column_dimensions[get_column_letter(i)].width = w
    for row in ini.iter_rows(min_row=2):
        for c in row:
            c.alignment = Alignment(wrap_text=True, vertical="top")

    m = wb.create_sheet("Критерии и веса")
    m.append(["Критерий", "Вес, %", "Блок", "Источник оценки", "Откуда критерий", "Как оценивается (1-5)"])
    meta = {
        "vol": ("Ценность", "Методика гл.12 (пропускная способность, объём транзакций)",
                "Документов SimBase за период: основных + 0,5 связанных. Нет данных - 2. Границы: 50 / 200 / 700 / 1500"),
        "fte": ("Ценность", "Методика гл.9, 11, 12 (трудоёмкость, CPT, утилизация)",
                "Оценка FTE по функционалу (доля позиции делится поровну между её этапами). Границы: 0,2 / 0,4 / 0,6 / 0,9"),
        "appr": ("Ценность", "Методика гл.9, 12 (Wait Time %, узкое место BNI, SLA)",
                 "Среднее число согласующих в потоках SimBase; +1, если Зам. директора участвует в 50% и более документов или ГБ КТ в 40% и более"),
        "E": ("Ценность", "Методика гл.10, 11 (Defect Rate, Rework Rate, FPY, COPQ)",
              "Эксперт: ручной ввод, число стыков и сверок, склонность к исправлениям"),
        "R": ("Ценность", "Добавлено: комплаенс", "Эксперт: штрафы, налоги, аудит, существенность ошибки для отчётности"),
        "K": ("Ценность", "Добавлено: влияние на другие процессы",
              "Эксперт: влияние на закрытие периода, ФО, платежи, смежные этапы"),
        "A": ("Ценность", "Методика гл.13 (разрыв AR и LOA до целевых 70% / 7 из 10)",
              "Эксперт: формализуемость правил, цифровые входы (ИС ЭСФ, 1С, банк, СЭД), потенциал LOA 7 и выше"),
        "ready": ("Готовность", "Методика гл.15 (зрелость: описание и стандартизация)",
                  "Confluence: регламент/ВНД - 4,5; инструкции - 3,5; матрицы - 2,5; нет - 1,5; плюс 0,5, если этап есть и в ДИ, и в матрице"),
        "S": ("Готовность", "Добавлено: сложность и срок внедрения (Методика гл.11: ROI)",
              "Эксперт: нужны ли интеграции, судебные и экспертные решения; 5 - просто"),
        "repl": ("Готовность", "Добавлено: тиражируемость",
                 "Число позиций, выполняющих этап (по функционалу). Границы: 3 / 5 / 8 / 12; чем больше, тем выгоднее единый стандарт"),
    }
    for k, name, src in CRIT:
        b, origin, how = meta[k]
        m.append([name, WEIGHTS[k], b, src, origin, how])
    m.append([])
    m.append(["Индекс приоритета = сумма вес x (балл - 1) / 4, диапазон 0-100"])
    m.append([f"Сегмент: порог {thresholds[0]:.0f} баллов по ценности и по готовности "
              f"(выше обоих порогов - быстрые победы)"])
    for c in m[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = fill
    for i, w in enumerate([48, 8, 12, 12, 52, 90], 1):
        m.column_dimensions[get_column_letter(i)].width = w
    for row in m.iter_rows(min_row=2):
        for c in row:
            c.alignment = Alignment(wrap_text=True, vertical="top")
    wb.save(path)


if __name__ == "__main__":
    src = sys.argv[1]
    out = sys.argv[2] if len(sys.argv) > 2 else str(Path(__file__).parent / "priority_scoring.xlsx")
    D = load_map(src)
    rows, C, th = build(D)
    robustness(rows)
    write_xlsx(rows, C, th, out)
    for r in rows:
        print(r["rank"], r["id"], round(r["total"], 1), round(r["value"]), round(r["feas"]), r["segment"][:1],
              round(r["stable"]), r["action"][:22], r["name"][:60], sep=" | ")
    print()
    for x in initiative_stats(rows):
        print(x["id"], round(x["mean_total"], 1), round(x["top_total"], 1), x["in_top20"], "/", x["n"], x["docs"],
              round(x["fte"], 1), x["name"], sep=" | ")
