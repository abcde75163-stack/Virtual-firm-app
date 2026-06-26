# VF TBD 보고서 생성기

부산대학교 산학협력단 TLO — Virtual Firm Tech Business Development 보고서 자동 생성 시스템

## 📋 개요

특허 명세서(PDF/DOCX)를 업로드하면 Claude AI가 VF TBD v3.0 기준의
기술사업화 분석 보고서를 자동 생성하고 Word 파일로 다운로드 제공합니다.

## 🚀 Streamlit Cloud 배포 방법

### 1단계: GitHub 저장소 생성
```
이 폴더 전체를 GitHub 저장소에 업로드
```

### 2단계: Streamlit Cloud 연결
1. https://share.streamlit.io 접속
2. "New app" → GitHub 저장소 선택
3. Main file path: `app.py`

### 3단계: API 키 설정 (필수)
Streamlit Cloud → App Settings → Secrets에 아래 내용 추가:
```toml
ANTHROPIC_API_KEY = "sk-ant-xxxxxxxxxxxxxxxx"
```

## 📁 파일 구조

```
vf_app/
├── app.py                    ← Streamlit 메인 앱
├── requirements.txt          ← 패키지 의존성
├── .streamlit/
│   └── config.toml           ← Streamlit 설정 (업로드 크기, 테마)
└── core/
    ├── __init__.py
    ├── patent_parser.py      ← PDF/DOCX 텍스트 추출
    ├── report_generator.py   ← Claude API 스트리밍 생성
    └── docx_builder.py       ← Word 문서 변환
```

## ⚙️ 주요 기능

| 기능 | 설명 |
|---|---|
| 특허 명세서 업로드 | PDF, DOCX 지원 (최대 50MB) |
| 수요기업 정보 입력 | 기업명 + IR자료 업로드 (선택) |
| 실시간 생성 진행 표시 | 4섹션 분할 생성 + 로그 스트리밍 |
| 에러 상세 정보 | traceback 포함 에러 박스 |
| Word 보고서 다운로드 | VF TBD v3.0 기준 .docx |

## 🔧 로컬 실행

```bash
pip install -r requirements.txt
export ANTHROPIC_API_KEY="sk-ant-..."
streamlit run app.py
```

## 📌 버전 정보

- 보고서 기준: VF TBD v3.0 (27개 섹션)
- AI 모델: claude-sonnet-4-6
- 인증: 미구현 (v2에서 추가 예정)
- 생성 이력: 미구현 (v2에서 추가 예정)

## ⚠️ 주의사항

- 본 시스템이 생성하는 보고서는 **참고용 자료**입니다
- 기술이전 협상 전 반드시 전문가 검토를 받으시기 바랍니다
- API 키는 반드시 Streamlit Secrets로 관리하고, 코드에 직접 입력하지 마세요
