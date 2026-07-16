---
name: excel-theme
description: 엑셀(.xlsx) 파일을 새로 생성하거나 편집할 때 일관된 색상·폰트·셀 서식 테마를 적용한다. openpyxl로 스프레드시트를 만들거나 표를 스타일링할 때, 또는 사용자가 "엑셀 만들어줘 / 표 정리해줘 / 양식 통일" 등을 요청할 때 사용한다.
---

# Excel Theme

Claude Code로 엑셀(.xlsx)을 **생성하거나 편집할 때마다** 이 스킬의 헬퍼를 거쳐
일관된 색상·폰트·셀 서식을 적용한다. 매번 스타일을 새로 짜지 말고 헬퍼를 호출한다.

색/폰트는 이 파일 `excel_theme.py` 상단 **편집 영역(PALETTES)** 이 단일 출처다.
공식 테마는 3종: **`default`**(기본, 정산표·결산 회색 삼선표) /
**`procpa`**(브랜드 블루 — 딥네이비 `#0B1C4A` + 블루 `#2563EB`, procpa.co.kr 팔레트) /
**`dcf-valuation`**(DCF·평가, Ocean Blue 딥블루 `#004889`).
(`navy`·`charcoal` 은 legacy — 기존 감사조서 하위호환용으로만 유지, 신규 사용 금지.
구 이름 `closing`·`valuation`·`audit_charcoal`·`audit_navy` 는 별칭으로 계속 동작)

## 언제 적용하나
- openpyxl 등으로 새 .xlsx 를 만들 때 → 표를 채운 뒤 마지막에 `apply_theme(ws, ...)` 호출.
- 기존 .xlsx 의 표 서식을 다듬을 때도 동일하게 `apply_theme` 적용.
- 미지정 시 **기본 `default`**(정산표·결산 회색). 브랜드 산출물은 `procpa`, DCF·평가는 `dcf-valuation` 명시.

## 사용 방법

이 스킬 폴더(`excel-theme/`)를 import 경로에 추가하고 `apply_theme` 를 호출한다.

```python
import sys
sys.path.insert(0, r"C:\Users\PC\.claude\skills\excel-theme")

from openpyxl import Workbook
from excel_theme import apply_theme, style_total_row, THEMES

wb = Workbook()
ws = wb.active
# ... 헤더와 데이터를 먼저 채운다 ...

apply_theme(
    ws,
    theme="default",            # THEMES 키 중 하나(default/procpa/dcf-valuation). 생략 시 기본값.
    header_row=4,               # 헤더가 있는 행 번호
    data_range="B4:F9",         # 표 영역. 생략하면 사용 영역 자동 감지.
    title_cell="B2",            # (선택) 시트 제목 셀 — 큰 글씨로 스타일
    number_format_cols={        # (선택) 열별 숫자 서식 (회계서식 권장)
        "C": "accounting",      #   숫자: 천단위 콤마, 음수 빨강 괄호, 0→대시
        "F": "percent_acct",    #   퍼센트: 음수 빨강 괄호, 0→대시
    },
)
# 소계/합계 행은 별도 강조 (medium 상단선 + 굵게)
style_total_row(ws, THEMES["default"], row=9, min_col=2, max_col=6)

wb.save("output.xlsx")
```

`apply_theme` 는 아래 4종 셀 규칙을 한 번에 적용한다. 개별 제어가 필요하면
인자 `zebra=`, `freeze=`, `autofit=`, `highlight_neg=` 로 끄거나, 개별 함수를 직접 호출한다.

### 더 간단히: `write_table` (권장 — 데이터+옵션 한 번에)
셀을 일일이 채우는 대신, 데이터와 옵션만 주면 채우기·서식·소계/합계·입력색까지 끝난다:
```python
from excel_theme import write_table
write_table(ws, ["계정","당기","전기"], data, theme="default",
            title="재무상태표",
            currency_cols=[1, 2],                 # 통화 열 → 테마 단위(default=백만원)
            section_rows=[1], subtotal_rows=[4], total_rows=[5],   # rows의 1-based 인덱스
            input_cells=["C6:D7"], linked_cells=["C10:D10"])       # 입력=크림 / 연결=파랑
```
`mark_cells(ws, theme, input=…, linked=…)` 로 입력(#FFFBEF)·연결(#CCECFF) 셀만 따로 칠할 수도 있다.

## 사용 가능한 테마 (공식 3종 + legacy 2종)
| 이름 | 분위기 | 본문 헤더 | 표 스타일 |
|---|---|---|---|
| `default` (기본) | 정산표·결산 — 회색 | #F2F2F2/검정 | frame(삼선표) |
| `procpa` | 브랜드 블루 — 딥네이비 & 블루 | #EFF5FF/딥네이비 | frame(삼선표) |
| `dcf-valuation` | DCF·평가 — Ocean Blue | #004889/흰색 | rules |
| `navy` *(legacy)* | 네이비 & 골드 | #E1E8F1/네이비 | frame(삼선표) |
| `charcoal` *(legacy)* | 차콜 & 틸 | #D6DEE6/차콜 | frame(삼선표) |

> **구 이름 `valuation`→`dcf-valuation`, `closing`→`default`, `audit_charcoal`→`charcoal`, `audit_navy`→`navy` 는 별칭(ALIASES)으로 계속 동작** — 감사조서 빌더 `frame.py`·valuation-tools 등 하위호환. legacy 2종은 신규 산출물에 쓰지 않는다(미리보기 빌드에서도 제외).
> **default·procpa(+legacy navy·charcoal) 공통 규칙(dcf-valuation 제외):**
> - 표 = **frame(삼선표)**: 좌/우 외곽선 없음 · 상/하 **굵게(medium)** · 내부 **얇게(thin)**.
> - 셀 배경: 하드코딩(입력)=`note`(#FFFBEF) · 수식=흰배경 · 참조값=`linked`(#F2F2F2).
> - 숫자=`accounting`(`#,###,##0;[Red](#,###,##0);-`) · 비율=`percent_acct`(`0.00%_);[Red](0.00%);"-"_);@_)`).
> - 헤더 3색: **제목**(B2, `title_bg`) / **본문 기본**(`excel_header_fill`) / **본문 세컨더리**(`excel_header2_fill`, 2차 표용).
>   세컨더리 헤더는 `style_header_row(..., secondary=True)` 로 적용.
> - **섹션행 배경색 없음**(`subheader=#FFFFFF`, 굵게만). **소계·합계 배경**: `default`=**헤더색 #F2F2F2**, `procpa`=**#EFF5FF**(`excel_total_fill`), legacy `charcoal`·`navy`=없음(윗선+굵게만). dcf-valuation=band 유지.
> - **음수**: 테마 `excel_highlight_neg=False` → 굵게/별색 오버레이 없이 **숫자서식 `[Red]`(기본 빨강)만**. `apply_theme`/`write_table` 가 테마값을 따르므로 실제 사용에서도 동일.

색상·폰트·서식·디자인 원칙 상세는 `GUIDE.md` 참고.
**테마 색을 바꾸려면 `excel_theme.py` 상단 PALETTES 만 수정**하면 엑셀·워드·PPT가 함께 바뀐다.

## 셀 규칙 → 함수 매핑 (`excel_theme.py`)
1. **숫자/통화 서식** — `set_number_formats` : 회계서식 `accounting`(#,###,##0;[Red](#,###,##0);-, **0은 대시**), `percent_acct`, 원화/백만원, 소수 등.
2. **테두리·정렬·행높이** — `apply_borders`(외곽 굵게+내부 얇게), `style_body`(숫자 우측·텍스트 좌측, 세로 가운데, 행높이).
3. **헤더·틀고정·줄무늬** — `style_header_row`(배경색+글씨), `freeze_below_header`, `zebra_stripes`.
4. **자동 열너비·조건부서식** — `autofit_columns`(한글 폭 보정), `highlight_negatives`/`highlight_threshold`.
5. **정산표 관행** — `style_total_row`(소계/합계 행: medium 상단선+thin 하단선+굵게+배경), `style_subheader_row`(섹션 구분 행).

## 엑셀 생성 기본 규칙 (apply_theme 가 자동 적용)
1. **틀 고정 안 함** — `freeze` 기본값 False. 필요할 때만 `freeze=True`.
2. **A열·1행 여백** — `margin=True`(기본). A열 너비 2.0, 1행 높이 ~27px 로 비워 둠. 표는 B2 제목/B열부터.
3. **제목 색 확장** — `title_cell="B2"` 의 색(채움)을 **데이터가 있는 마지막 열까지** 가로로 확장.

## 폰트 (앱별 변형)
엑셀 기본 폰트는 **맑은 고딕**이다(procpa·dcf-valuation 의 워드/PPT는 Pretendard).
엑셀 폰트를 바꾸려면 `excel_theme.py` 의 해당 테마 `font_excel` 값을 수정한다.

## 표 유형별 샘플 + 규칙 커스터마이징
`examples/table_samples.py` 의 `SAMPLES` 리스트에 표 유형이 선언돼 있다(단순표 / 합계만 /
소계+합계 / 섹션+소계+합계 / 묶음헤더 / 와이드+퍼센트 / 줄무늬OFF / 2열목록 / 입력·연결셀 등).
각 표의 규칙(`subheader_rows`, `subtotal_rows`, `total_rows`, `number_format_cols`,
`zebra`, `freeze` 등)을 dict 값만 고쳐 직접 커스터마이징한다.

## 디자인 원칙
공통 디자인 규칙(절제된 악센트·헤어라인·60-30-10·anti-slop)은 `GUIDE.md`(C. 디자인 원칙) 를 따른다.

## 검증
`python examples/build_preview.py` → 테마마다 `preview_<theme>.xlsx`(샘플마다 한 시트) 생성.
전체(엑셀·워드·PPT) 일괄 미리보기는 `build_all.py`(이 폴더).

> 참고: `tools/` 폴더는 테마 *개발용* 도구(샘플 피드백 diff·실무파일 분석기)다. 엑셀 생성 런타임에서는 사용하지 않는다 — 개발 워크플로우는 프로젝트 폴더의 DEV.md 참조.
