---
name: word-theme
description: 워드(.docx) 문서를 새로 만들거나 편집할 때 일관된 색·폰트·제목/표/머리글 서식 테마를 적용한다. python-docx로 보고서·메모·제안서 등을 만들 때, 또는 사용자가 "워드/문서/보고서 만들어줘 / 양식 통일" 등을 요청할 때 사용한다.
---

# Word Theme

Claude Code로 워드(.docx)를 **생성하거나 편집할 때마다** 이 스킬의 헬퍼를 거쳐
일관된 색·폰트·제목부/표 서식을 적용한다. 워드 전용 테마는 `word_themes.py`에서 정의하며
(엑셀과 색 일치 불필요), 없으면 `..\excel-theme\excel_theme.py`(PALETTES)로 폴백한다.

## 테마 (2종)
- **default** (기본): 모노톤 그레이(헤딩 `#333333`), 제목단계 동일색,
  **라이트 헤더 표 + 삼선표**(좌우 외곽선 없음·상하 굵게), 에어리한 톤. 보고서·메모·조서 공용.
- **procpa**: 브랜드 블루 — default 와 같은 레이아웃에 procpa.co.kr 팔레트
  (딥네이비 `#0B1C4A` 헤딩 + 블루 `#2563EB` 악센트 + 블루 틴트 `#EFF5FF` 헤더), 폰트 Pretendard.
- (별칭) **workpaper → default** 로 통합됨. 조서는 `doc_type="workpaper"` 로 구분.
- (폴백) 엑셀 팔레트 이름(dcf-valuation 등)도 받을 수 있다.
- 미지정 시 **기본 `default`**.

## 사용 방법
```python
import sys
sys.path.insert(0, r"C:\Users\PC\.claude\skills\word-theme")
from word_theme import (new_doc, add_title, add_heading, add_paragraph,
                        add_callout, add_table)

doc = new_doc("procpa", doc_type="report")       # 테마·문서유형은 문서에 기억됨
add_title(doc, "FY25 손상평가 보고서", "OO사 · 2026.03")   # doc_type 프리셋 → bar 스타일
add_heading(doc, "1. 평가 개요", level=1)                  # 이후 헬퍼는 테마 자동 상속
add_paragraph(doc, "본 보고서는 ...")
add_callout(doc, "핵심 결론: ...")                          # note색 강조 박스
add_table(doc,
    headers=["계정과목", "전기", "당기", "증감"],
    rows=[["매출", 1200000, 1450000, 250000],
          ["합계", 1420000, 1740000, 320000]],
    number_cols={1: "accounting", 2: "accounting", 3: "accounting"},
    total_last_row=True)                       # 마지막 행 합계 강조
doc.save("output.docx")
```

## 제목부 스타일 5종 + 문서유형 프리셋
`add_title(style=...)` 로 직접 고르거나, `new_doc(doc_type=...)` 프리셋에 맡긴다:

| style | 모양 | 용도 | doc_type 프리셋 |
|---|---|---|---|
| `bar` | 제목 + 하단 강조 바 | 기본 — 일반 보고서 | (미지정 시), `report` |
| `side` | 좌측 세로 악센트 바 | 간단한 메모·검토 | `memo` |
| `band` | 전폭 색 밴드(짙은 배경+흰 제목) | 제안서·트렌디 산출물 | `proposal` |
| `center` | 중앙 정렬 + 상하 가는 선(88% 폭) | 공식 의견서·공문·회신문 | `opinion` |
| `meta` | 제목(좌) + 문서정보 표(우) | 조서·품질관리 문서 | `workpaper` |

제목이 두 줄로 꺾이면 고정 줄높이(폰트×1.15)로 압축된다.
`meta=` 인자로 부가정보를 넘긴다: `{"caption": 문서유형 라벨(band),
"작성": ..., "검토": ..., "문서번호": ...}` — center 는 메타 라인,
meta 스타일은 우측 문서정보 표로 표기된다. 프리셋 정의는 `word_themes.py` 의 `DOC_TYPES`.

## 헬퍼 (`word_theme.py`)
- `new_doc(theme, doc_type)` — Normal/Heading 스타일에 팔레트 글꼴·색 적용 + 바닥글 페이지번호.
  **테마를 문서에 기억**시키므로 이후 헬퍼에 theme= 반복 불필요(명시 시 그 값 우선).
- `add_title(style=, meta=)` — 위 5종 제목부. `add_heading(level=1~3)` — H1=primary, H2=primary_dark, H3=secondary (default/procpa는 셋 동일색).
- `add_paragraph`, `add_bullets(items, numbered=False)`(글머리표/번호 목록, `(level,text)`로 들여쓰기), `add_callout`(강조 박스).
- `add_table` — 헤더 배경+글자, 숫자 우측·합계행 강조. **숫자는 엑셀 회계서식 규약을 따른다:
  음수 = negative색 괄호 `(1,200)`, 0 = 대시 `-`** (`accounting`/`percent`/`percent_acct`/`decimal`/`multiple`).
  `autofit`으로 표 너비 모드 제어(모두 행 높이는 압축): `True`/`"content"`(기본)=열을 내용폭에 맞춤, `"window"`=페이지 폭 100%로 채우고 열은 내용 비율 분배, `False`=고정폭 균등 표. **테마 옵션**으로 표 모양 제어:
  `table_header_fill`/`table_header_text`(라이트 헤더), `table_sides`(False=좌우 외곽선 없음=삼선표),
  `table_edge_sz`(상/하 외곽 굵기), `table_header_sep`(헤더 하단 구분선), `zebra`(줄무늬).

테마 색을 바꾸려면 `word_themes.py`(WORD_THEMES)를 수정한다. 상세 규격은 `WORD_GUIDE.md`.
공통 디자인 규칙은 `..\excel-theme\GUIDE.md`(C. 디자인 원칙) 를 따른다.

## 검증
`python examples/build_preview.py` → 테마별 보고서/메모 `.docx` 생성.
`python examples/title_samples.py` → 제목부 6종 비교 샘플.
전체(엑셀·워드·PPT) 일괄 미리보기는 `..\excel-theme\build_all.py`.
