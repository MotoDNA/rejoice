# 리조이스 챌린지 (9월 Special Month)

감사 인증샷과 리조이스 사행시로 함께하는 9월 스페셜 챌린지 페이지입니다.

- 참여 페이지: https://motodna.github.io/rejoice/
- 데이터 시트: https://docs.google.com/spreadsheets/d/1iUKIQJSwh0d5HBXZ5C67OBUUMUTyCpnjnmi89nXizEw/edit
- 참여 기간: 2026년 9월 14일(월) ~ 9월 25일(금)

## 구성

- `index.html` — 참여자용 웹페이지 (GitHub Pages: `main` / `/(root)`)
- `code.gs` — Google Apps Script 웹 앱 서버. 독립 실행형 프로젝트(`리조이스 챌린지 서버`)로
  배포되어 있으며, `SHEET_ID`로 스프레드시트를 직접 엽니다.

## 동작 방식

`index.html`의 `API_URL`이 Apps Script 웹 앱의 `/exec` 주소를 가리킵니다.
글/좋아요/점수는 시트의 `entries`·`config` 탭에 저장되고, 사진은 드라이브
`리조이스 챌린지 사진` 폴더에 올라간 뒤 썸네일 링크로 표시됩니다.
두 시트 탭은 첫 요청이 들어올 때 자동으로 생성됩니다.

## 설정값

| 값 | 위치 | 현재 |
|---|---|---|
| 임원진 코드 | `code.gs`의 `ADMIN_CODE` | `0914` |
| 사진 폴더 | `code.gs`의 `FOLDER_NAME` | `리조이스 챌린지 사진` |
| 시트 ID | `code.gs`의 `SHEET_ID` | 위 데이터 시트 |
| 서버 주소 | `index.html`의 `API_URL` | 배포된 `/exec` 주소 |

> 이 저장소는 공개(public)이므로 `ADMIN_CODE`도 공개됩니다. 바꾸려면 Apps Script
> 편집기에서 값을 수정하고 **배포 → 배포 관리 → 수정 → 새 버전**으로 다시 배포하세요.

## code.gs를 고친 뒤 반영하는 법

1. https://script.google.com 에서 `리조이스 챌린지 서버` 열기
2. 코드 수정 후 저장
3. **배포 → 배포 관리 → (연필) 수정 → 버전: 새 버전 → 배포**
   — 기존 `/exec` 주소가 그대로 유지되므로 `index.html`은 건드릴 필요 없습니다.
