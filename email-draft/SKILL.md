---
name: email-draft
description: 업무 메일 초안 작성. "메일 써줘 / 초안 만들어줘 / 보내는 글 작성해줘" 요청 시 표준 본문 구조로 초안을 만들고, 확정되면 .eml 파일(더블클릭 → Outlook 보내지 않은 새 메일)로 저장한다.
---

# email-draft — 메일 초안 표준

업무 메일 초안을 **항상 같은 구조·톤**으로 만들기 위한 스킬. 산출 형식은 **.eml Outlook 초안**
(X-Unsent 헤더 — 더블클릭하면 작성창으로 열림, 사용자 확정 2026-07). 본문 디자인은 **①안 미니멀**
(굵게 강조만·얇은 테두리 표 — 수신자 클라이언트 어디서든 안 깨짐).

## 절차

1. `templates/`에서 메일 유형 골격을 고른다: `자료요청` / `결과송부` / `검토의견`.
   없는 유형은 아래 '본문 표준 구조'만 지켜 작성.
2. **본문을 채팅에 먼저 보여주고 사용자 확인**을 받는다(제목·수신자·첨부 포함).
3. 확정되면 `email_draft.py`로 .eml 생성:
   ```python
   import sys; sys.path.insert(0, r"C:\Users\PC\.claude\skills\email-draft")
   from email_draft import p, bullets, table, section, assemble_body, build_email
   body = assemble_body(greeting="안녕하세요, OOO 님. 이재현입니다.", blocks=[...])
   build_email(subject="[회사명] 주제 — ...의 건", body_html=body,
               attachments=[r"...\파일.xlsx"], out_path=r"...\메일초안_주제_YYYYMMDD.eml")
   ```
4. 저장 위치는 **해당 프로젝트 폴더**(사용자가 지정한 곳). 파일명: `메일초안_{주제}_{YYYYMMDD}.eml`.
5. 메일 발송은 항상 **사용자가 직접** 한다(초안까지만 만든다).

## 본문 표준 구조 (assemble_body가 이 순서를 강제)

1. **인사** — "안녕하세요, {수신자} 님. 이재현입니다."
2. **배경** 1~2문장 — 무슨 건인지
3. **핵심 내용** — 번호 소제목(`section`) + 불릿(`bullets`)/요약 표(`table`)
4. **요청·다음 단계** — 상대가 할 일을 명확히
5. **맺음** — "감사합니다." + "이재현 드림"
- 제목 형식: `[회사명] 주제 — …의 건`
- 첨부가 있으면 본문에 `attachments_note`로 파일명 표기

## 스타일 규칙 (①안 미니멀)

- 인라인 스타일만. 외부 CSS·웹폰트·이미지 금지 (Outlook/Gmail 호환)
- 강조는 `<b>`만, 글자색 지정 안 함(수신자 기본값 상속)
- 표: 테두리 1px `#999`, 헤더만 `#F2F2F2` — `table()` 헬퍼가 처리
- 폰트 스택: `'Noto Sans KR','맑은 고딕',sans-serif`

## 파일

- `email_draft.py` — 본문 블록 헬퍼(`p/section/bullets/table/assemble_body`) + `build_email`(.eml)
- `templates/자료요청.md` · `결과송부.md` · `검토의견.md` — 유형별 본문 골격
- 정리 페이지가 필요하면 → `html-theme` 스킬 (별도)
