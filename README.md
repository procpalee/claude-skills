# claude-skills

회계·재무 실무자를 위한 **Claude 스킬 모음**입니다. [procpa.co.kr](https://procpa.co.kr/tools)에서 소개·배포합니다.

## 스킬 목록

| 스킬 | 설명 | 다운로드 |
|---|---|---|
| [excel-theme](excel-theme/) | 엑셀(.xlsx) 생성·편집 시 일관된 색·폰트·표 서식 테마 적용 (default·procpa·dcf-valuation 3종) | [excel-theme.zip](https://github.com/procpalee/claude-skills/releases/latest/download/excel-theme.zip) |

## 설치

1. 위 표의 다운로드 링크(항상 최신 릴리스)에서 zip을 받아 압축을 풉니다.
2. 스킬 폴더를 Claude Code 스킬 디렉터리에 넣습니다:
   - 개인용: `~/.claude/skills/<스킬이름>/`
   - 프로젝트용: `<프로젝트>/.claude/skills/<스킬이름>/`

각 스킬의 자세한 사용법은 [procpa.co.kr/tools](https://procpa.co.kr/tools)의 상세 페이지를 참고하세요.
스킬 개념이 처음이라면 → [스킬이란?](https://procpa.co.kr/guide/what-is-skill)

## 릴리스 방식

`v*` 태그를 push하면 GitHub Actions가 **모든 스킬을 각각 zip으로 묶어 하나의 릴리스에 첨부**합니다.
따라서 `releases/latest/download/<스킬이름>.zip` URL은 항상 모든 스킬의 최신 버전을 가리킵니다.

```bash
git tag v2026.07.16 && git push origin v2026.07.16
```

## 문의

버그 제보·기능 제안: [Issues](https://github.com/procpalee/claude-skills/issues) 또는 [procpa.co.kr/contact](https://procpa.co.kr/contact?type=tools)

## License

MIT
