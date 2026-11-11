# 리베이트 추천 프로그램

정적 사이트입니다. 데이터는 Firebase Firestore(프로젝트 `eruda-sale-ver2`)에 그대로 저장됩니다.

## 파일
- index.html: 앱 본체
- support.js, ver2-seed-data.js, xlsx-lite.js: 앱 구동용 스크립트
- vercel.json: cleanUrls 설정

## 배포 (Vercel)
Framework Preset = Other, Build Command 비움, Output Directory 비움(루트).

## 참고
Firebase 웹 설정값(apiKey 등)은 공개되어도 되는 값이며, 접근 제어는 Firestore 보안 규칙으로 관리합니다.
