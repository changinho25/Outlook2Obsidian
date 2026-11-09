# Outlook2Obsidian - Claude 작업 규칙

## 커밋 워크플로

1. **수정 후 커밋 전에 반드시 확인**: 코드 수정 완료 후 바로 커밋하지 않고, 커밋할지 사용자에게 먼저 물어볼 것
2. **"커밋해줘" 요청 시 순서대로 진행**:
   1. `git commit` 실행
   2. 작업 내용을 Obsidian 작업 세션 폴더에 md 파일로 저장 (`D:\GoogleDrive\Obsidian\NOTE\01. AI 관련\작업 세션\Outlook2Obsidian\`)
   3. `git push origin main` 실행

## 인코딩 주의사항

- `USER_CONFIG.bas`는 한글 경로가 포함되어 있어 파일로 직접 수정하면 인코딩이 깨짐
- `USER_CONFIG.bas` 수정이 필요할 경우 사용자에게 **VBA 에디터에서 직접 수정**하도록 안내할 것
- `.bas` 파일 내에 한글 문자열 리터럴 사용 금지 (ChrW() 사용)

## 파일 구성과 적용

| 파일 | 역할 |
|---|---|
| `SaveEmail.bas` | 메일을 추출·변환하는 메인 로직 |
| `SaveUtilities.bas` | 메일 처리·변환 유틸리티 |
| `USER_CONFIG.bas`, `vault-path.txt` | 개인 설정. git 추적 대상이 아니며 `*.example*` 이 견본이다 |

코드 적용은 사용자가 VBA 편집기의 `File` > `Import File` 로 한다. Claude 는 파일 수정과 안내까지만 하며, 동작 확인 절차는 `README.md` 의 사용 방법과 `C:\CC\outlook\CLAUDE.md` 의 매크로 적용 절차를 참고한다.

## 저장소 정보

- GitHub: https://github.com/changinho25/Outlook2Obsidian
- Obsidian 작업 세션: `D:\GoogleDrive\Obsidian\NOTE\01. AI 관련\작업 세션\Outlook2Obsidian\`
