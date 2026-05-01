# Changelog

## 2026-05-01

### SaveEmail.bas
- 첨부파일 및 HTML을 메일 제목별 서브폴더(`attachments/{메일명}/`)로 분리 저장
- YAML frontmatter 필드 정리 (한글 필드명 → 영문: `tags`, `title`, `date`, `from`, `ITSM`, `ITSM_URL`)
- HTML 임베드를 YAML 바로 아래로 위치 이동 (빠른 미리보기)
- HTML → Markdown 본문 변환 추가 (인라인 이미지 포함)
- YAML `tags` 형식을 리스트 형식으로 변경 (`tags: SOURCE/MAIL today` → `- SOURCE/MAIL`)
- `today` 태그 제거, `SOURCE/MAIL` 단일 태그로 정리

### SaveUtilities.bas
- `ConvertHTMLToMarkdown` 함수 신규 추가
  - HTML noise 제거 (head/style/script/주석/VML)
  - `<table>` → Markdown 파이프 테이블 변환
  - HTML 소스 줄바꿈 공백 처리 (불필요한 빈줄 방지)
  - `<img>` → Obsidian wikilink(`![[...]]`) 변환
  - `&nbsp;` 조기 디코딩
  - heading 변환 (`h1~h6` → `##`, `###`, `####`)
  - bold/italic 태그 제거 (Outlook 중첩 `<b>` 태그로 인한 `**` 노이즈 방지)
  - block 태그 → 줄바꿈 변환
  - HTML entity 디코딩
  - 다중 공백 → 단일 공백 정리
  - 줄 앞뒤 공백 trim
  - 회신 헤더 필드 줄바꿈 분리 (`보낸 날짜:`, `받는 사람:`, `참조:`, `제목:`)
  - 서명 블록 앞 `---` 구분선 자동 삽입
  - 연속 빈줄 최대 2줄로 정규화
- `ReplaceImageTagsWithMarkdownLinks` 함수 신규 추가
- `ConvertHTMLTableToMarkdown` / `CleanCellText` 함수 신규 추가
- `SaveAsUTF8` / `ReadFileContent` 함수 신규 추가
