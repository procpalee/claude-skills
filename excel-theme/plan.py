# -*- coding: utf-8 -*-
"""
plan.py — 순수 format_plan: 테마 spec(dict) → Excel 서식 연산 리스트.

의존성 없음(표준 라이브러리만) → 그대로 웹 MCP 서버에 복사해 쓸 수 있다.
excel-theme 스킬(theme_plan.py)과 excel-theme-mcp(웹) 가 **이 한 파일**을 공유한다.
(export_theme.py 가 이 파일을 MCP 폴더로 복사 + theme.json 생성)
"""


def _letter(n):
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def _rng(c1, r1, c2, r2):
    return "%s%d:%s%d" % (_letter(c1), r1, _letter(c2), r2)


def _coll(start_col, key):
    return key if isinstance(key, str) else _letter(start_col + key)


def format_plan(spec, *, header_row=4, start_col=2, n_cols=1, n_rows=1, title_cell="B2",
                number_format_cols=None, currency_cols=(), percent_cols=(),
                section_rows=(), subtotal_rows=(), total_rows=(),
                input_cells=(), linked_cells=()):
    """테마 spec + 표 레이아웃 → 결정론적 Excel 서식 연산 리스트.

    인덱스 규약: 헤더=header_row, 데이터=그 아래, 표는 start_col(기본 B=2)부터.
    section/subtotal/total_rows = 데이터 1-based 인덱스. number_format_cols 키=열문자 또는 0-based 데이터열.
    """
    col = spec["colors"]
    sizes = spec["sizes"]
    fmts = spec["number_formats"]
    end_col = start_col + n_cols - 1
    data_start = header_row + 1
    last_row = data_start + n_rows - 1
    ops = [{"op": "column_width", "col": "A", "width": 2.0}]

    # 제목(B2) 헤더 — 색을 데이터 끝열까지 확장
    if title_cell and spec.get("title_fill"):
        trow = int("".join(ch for ch in title_cell if ch.isdigit()))
        ops.append({"op": "fill", "range": _rng(start_col, trow, end_col, trow),
                    "color": col["title_bg"]})
        ops.append({"op": "font", "range": title_cell, "bold": True,
                    "size": sizes["title"], "color": col["title_font"]})

    # 본문 기본 헤더
    hdr = _rng(start_col, header_row, end_col, header_row)
    ops.append({"op": "fill", "range": hdr, "color": col["header_bg"]})
    ops.append({"op": "font", "range": hdr, "bold": True, "align": "center",
                "size": sizes["header"], "color": col["header_font"]})

    # 숫자/통화/퍼센트 서식 (열 단위)
    cf = {}
    for k in currency_cols:
        cf[_coll(start_col, k)] = spec["currency_format_key"]
    for k in percent_cols:
        cf[_coll(start_col, k)] = "percent_acct"
    for k, v in (number_format_cols or {}).items():
        cf[_coll(start_col, k)] = v
    for cl, fk in cf.items():
        ops.append({"op": "number_format",
                    "range": "%s%d:%s%d" % (cl, data_start, cl, last_row),
                    "format": fmts.get(fk, fk)})

    # 본문 글꼴·정렬 (숫자=오른쪽, 텍스트=왼쪽, 세로 가운데)
    body = _rng(start_col, data_start, end_col, last_row)
    ops.append({"op": "font", "range": body, "color": col["text"], "size": sizes["body"]})
    ops.append({"op": "align", "range": body, "vertical": "center",
                "numbers": "right", "text": "left"})

    # 테두리
    table = _rng(start_col, header_row, end_col, last_row)
    mode = spec["border_mode"]
    if mode == "frame":              # 삼선표: 좌/우 외곽선 없음·상하 medium·내부 thin
        ops.append({"op": "border", "range": table, "style": "frame", "sides": "top_bottom",
                    "outer_color": col["border_outer"], "outer_weight": "medium",
                    "inner_color": col["border_inner"], "inner_weight": "thin"})
    elif mode == "grid":
        ops.append({"op": "border", "range": table, "style": "grid",
                    "outer_color": col["border_outer"], "outer_weight": "medium",
                    "inner_color": col["border_inner"], "inner_weight": "thin"})
    elif mode == "statement":
        ops.append({"op": "border", "range": table, "style": "statement", "sides": "top_bottom",
                    "outer_color": col["border_outer"], "outer_weight": "medium"})
    # rules: 격자 없음 → 합계/소계 가로선이 구분

    # 섹션행(subheader): 배경=subheader_fill(테마에 따라 흰색=없음) + 굵게
    for idx in section_rows:
        r = data_start + idx - 1
        rg = _rng(start_col, r, end_col, r)
        ops.append({"op": "fill", "range": rg, "color": col["subheader_fill"]})
        ops.append({"op": "font", "range": rg, "bold": True})

    # 소계행: 윗선(medium)+굵게, 배경 없음
    for idx in subtotal_rows:
        r = data_start + idx - 1
        rg = _rng(start_col, r, end_col, r)
        ops.append({"op": "border_top", "range": rg, "weight": "medium",
                    "color": col["border_outer"]})
        ops.append({"op": "font", "range": rg, "bold": True})

    # 합계행: 윗선(medium)+굵게+배경(total_fill)
    for idx in total_rows:
        r = data_start + idx - 1
        rg = _rng(start_col, r, end_col, r)
        ops.append({"op": "fill", "range": rg, "color": col["total_fill"]})
        ops.append({"op": "border_top", "range": rg, "weight": "medium",
                    "color": col["border_outer"]})
        ops.append({"op": "font", "range": rg, "bold": True})

    # 입력/참조 셀 배경
    for rg in (input_cells or []):
        ops.append({"op": "fill", "range": rg, "color": col["input_fill"]})
    for rg in (linked_cells or []):
        ops.append({"op": "fill", "range": rg, "color": col["linked_fill"]})

    # 행높이 자동
    ops.append({"op": "row_height", "rows": "%d:%d" % (header_row, last_row),
                "height": "auto"})

    # 음수: 테마가 굵게 강조면 조건부서식, 아니면 숫자서식 [Red]만(추가 op 없음)
    if spec.get("highlight_neg_bold"):
        ops.append({"op": "conditional_negative", "range": body,
                    "color": spec["accent"], "bold": True})

    return {"theme": spec["name"], "ops": ops}
