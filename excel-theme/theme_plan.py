# -*- coding: utf-8 -*-
"""
theme_plan.py — 테마 → "Excel 실행 스펙/연산 리스트" (스킬 측 진입점).

목적: Claude for Excel(애드인)은 로컬 파이썬/openpyxl을 못 쓰고 **열린 워크북을 직접 서식**한다.
그래서 openpyxl로 '그리는' 대신, **무엇을 어디에 적용할지의 연산 리스트**를 돌려주면 Claude가 실행한다.

★ 단일 소스: 색·규칙은 excel_theme.PALETTES 하나.
   - theme_spec() : PALETTES 토큰 → JSON 안전 스펙(색/서식/규칙)  ← 여기서만 excel_theme 의존
   - format_plan(): 스펙 + 표 레이아웃 → 연산 리스트  ← 순수 모듈 plan.py 가 담당(웹 MCP와 공유)
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from excel_theme import (PALETTES, get_theme, get_palette,  # noqa: E402
                         DEFAULT_THEME, ALIASES, _canon)
from plan import format_plan as _format_plan  # noqa: E402


def list_themes():
    """테마 목록(이름·라벨·표스타일) + 기본/별칭."""
    themes = [{"name": n, "label": get_palette(n)["label"],
               "border_mode": get_theme(n)["border_mode"]} for n in PALETTES]
    return {"default": DEFAULT_THEME, "aliases": ALIASES, "themes": themes}


def theme_spec(name=None):
    """테마 1종의 색·서식·규칙 스펙(JSON 안전). Claude가 직접 서식하거나 plan 생성에 사용."""
    cname = _canon(name) if name else DEFAULT_THEME
    th = get_theme(cname)
    return {
        "name": cname,
        "label": get_palette(cname)["label"],
        "font": th["font_name"],
        "sizes": {"title": th["font_size_title"], "header": th["font_size_header"],
                  "body": th["font_size_body"]},
        "border_mode": th["border_mode"],            # frame / rules / statement / grid
        "title_fill": bool(th.get("title_fill")),
        "row_height": "auto",
        "highlight_neg_bold": th["excel_highlight_neg"],   # False면 [Red] 숫자서식만(굵게 X)
        "currency_format_key": th["currency_format"],
        "accent": th["accent_color"],
        "colors": {
            "title_bg": th.get("title_bg"),
            "title_font": th.get("title_font_color"),
            "header_bg": th["header_bg"],              # 본문 기본 헤더
            "header_font": th["header_font_color"],
            "header2_bg": th["header2_bg"],            # 본문 세컨더리 헤더
            "header2_font": th["header2_font"],
            "total_fill": th["total_fill"],           # 소계/합계 배경
            "subheader_fill": th["subheader_fill"],   # 섹션행 배경(흰색=없음)
            "input_fill": th["note_fill"],            # 하드코딩(입력) 셀
            "linked_fill": th["linked_fill"],         # 참조값(타시트 연결) 셀
            "border_outer": th["border_color_outer"],
            "border_inner": th["border_color"],
            "text": th.get("text_color"),
        },
        "number_formats": th["number_formats"],       # accounting / percent_acct / …
    }


def format_plan(name=None, **layout):
    """테마명 + 표 레이아웃 → 연산 리스트 (plan.format_plan 위임, 공유 알고리즘)."""
    return _format_plan(theme_spec(name), **layout)


if __name__ == "__main__":
    import json
    print("THEMES:", [t["name"] for t in list_themes()["themes"]])
    plan = format_plan("charcoal", header_row=4, start_col=2, n_cols=4, n_rows=5,
                       number_format_cols={1: "accounting", 2: "accounting", 3: "accounting"},
                       subtotal_rows=[3], total_rows=[5], input_cells=["C5:C6"])
    print(json.dumps(plan, ensure_ascii=False, indent=2)[:900])
