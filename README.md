# XBRL Fact & Presentation 통합 추출기

XBRL 파일에서 **Arelle 라이브러리**를 사용해 Fact 정보와 Presentation 계층 구조를 함께 추출하여 Excel 파일로 저장하는 프로그램입니다.

---

## 📁 파일 준비

추출 전에 `.xbrl` 파일과 같은 폴더에 아래 파일들이 모두 있어야 합니다.

```
XBRL 폴더/
├── entity_2025-xx-xx.xbrl       ← 필수 | 인스턴스 문서 (실제 데이터)
├── entity_2025-xx-xx.xsd        ← 필수 | 스키마 파일 (Role 한글명 정의)
└── entity_2025-xx-xx_pre.xml    ← 필수 | Presentation Linkbase (계층 구조)
```

| 파일 | 역할 |
|------|------|
| `.xbrl` | 재무 데이터(Fact 값)가 담긴 인스턴스 문서 |
| `.xsd` | Taxonomy 스키마, Role 한글명 정의 포함 |
| `_pre.xml` | Presentation Linkbase, 항목 간 계층 구조 정의 |

---

## ▶ 실행 방법

1. `Xbrl_Extractor.py` 파일을 실행
2. 파일 선택 창이 열리면 `.xbrl` 파일 선택
3. 추출이 완료되면 **Downloads 폴더**에 Excel 파일 자동 생성

> `.xsd` 및 `_pre.xml` 파일은 직접 선택할 필요 없이 Arelle이 `.xbrl` 파일 기준으로 자동 참조합니다.

---

## 결과 파일

| 항목 | 내용 |
|------|------|
| 저장 위치 | `C:\Users\사용자명\Downloads\` |
| 파일명 형식 | `XBRL_Complete_{원본파일명}_{날짜시간}.xlsx` |
| 시트명 | `Complete_Data` |
| Excel 설정 | 줌 70%, 헤더 폰트 10pt, 컬럼 너비 자동 조정 |

---

## 출력 컬럼 설명

| 컬럼명 | 설명 |
|--------|------|
| `Label_Korean` | 항목 한글 레이블 |
| `Label_English` | 항목 영문 레이블 |
| `Seq` | 전체 행 순번 (1부터 시작) |
| `ContextRef` | Context 참조 ID (기간/시점 구분) |
| `Decimals` | 소수점 자릿수 |
| `Value` | Fact 실제 값 (없으면 빈 값) |
| `Pres_Level` | Presentation 계층 레벨 (루트=0) |
| `Pres_Role_Name` | Role 한글명 (예: `[D861310] 33. 배당금`) |
| `Pres_Role_URI` | Role URI |
| `Pres_Parent_Label_KR` | 부모 항목 한글 레이블 |
| `Pres_Parent_Name` | 부모 항목 Concept 이름 |
| `Concept_Name` | Concept 전체 이름 (namespace 포함) |
| `Concept_ID` | Concept ID |
| `Abstract` | 추상 항목 여부 (`True` / `False`) |
| `Substitution_Group` | 대체 그룹 (예: `xbrli:item`) |
| `Type` | 데이터 타입 (예: `xbrli:monetaryItemType`) |
| `Period_Type` | 기간 유형 (`instant` / `duration`) |
| `Balance` | 잔액 유형 (`debit` / `credit`) |

---

## ⚙️ 동작 방식

```
.xbrl 파일 선택
    │
    ▼
Arelle로 XBRL 파일 로드 (validate=False)
    │
    ├─ Role 정보 추출          → role URI → 한글명 딕셔너리 생성
    │
    ├─ Presentation 구조 분석  → BFS로 계층 레벨 계산
    │                            부모-자식 관계 매핑
    │
    ├─ Fact 그룹화             → concept별 Fact 리스트 인덱싱
    │
    └─ Fact × Presentation 조인
          │
          ├─ Fact 있음 → (Fact 수) × (Pres 수) 만큼 행 생성
          └─ Fact 없음 → (Pres 수) 만큼 행 생성 (Value = None)
                │
                ▼
          Excel 파일 저장 → Downloads 폴더
```

> Fact가 없는 항목도 Presentation 구조에 포함된 경우 누락 없이 출력됩니다.

---

## 개발 환경 및 의존성

| 라이브러리 | 용도 |
|-----------|------|
| `arelle` | XBRL 파싱 엔진 |
| `pandas` | 데이터 처리 및 Excel 출력 |
| `openpyxl` | Excel 스타일 설정 |
| `tkinter` | 파일 선택 GUI (Python 기본 내장) |
| `pathlib` | 경로 처리 |
| `collections` | `defaultdict` 사용 |

---

## 참고 사항

- `.xsd` 또는 `_pre.xml` 파일이 누락된 경우 Role 한글명 또는 Presentation 계층 정보가 비어 있을 수 있습니다.
- 동일한 Concept이 여러 Role에 속하는 경우, Role 수만큼 행이 반복 생성됩니다.
- 컬럼 너비는 상위 100개 셀 기준으로 자동 조정되며 최대 너비는 50으로 고정됩니다.