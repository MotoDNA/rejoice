# 리조이스 챌린지 (9월 Special Month)

감사 인증샷과 리조이스 사행시로 함께하는 9월 스페셜 챌린지 페이지입니다.

## 구성

- `index.html` — 참여자용 웹페이지 (GitHub Pages로 배포)
- `code.gs` — Google Apps Script 서버 (스프레드시트 저장 + 드라이브 사진 업로드)

## 배포 순서

1. **Apps Script 배포**
   - Google 스프레드시트 → 확장 프로그램 → Apps Script
   - `code.gs` 내용을 붙여넣고 저장
   - 배포 → 새 배포 → 유형: **웹 앱**
   - 액세스 권한: **모든 사용자**
   - 배포 후 나오는 `/exec` URL 복사

2. **index.html에 URL 연결**
   - `index.html`의 `const API_URL = 'PASTE_URL_HERE';` 부분을 위 URL로 교체
   - 커밋 & 푸시

3. **GitHub Pages 켜기**
   - 저장소 → Settings → Pages
   - Source: `Deploy from a branch` / Branch: `main` / 폴더: `/ (root)`
   - 몇 분 뒤 https://motodna.github.io/rejoice/ 에서 접속

## 설정값

- 임원진 코드: `code.gs`의 `ADMIN_CODE`
- 사진 저장 폴더: `code.gs`의 `FOLDER_NAME`
