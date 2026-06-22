# F.I.T 카카오뱅크 장부 자동화

카카오뱅크 모임통장 거래내역 엑셀과 기존 F.I.T 장부 엑셀을 읽어서 개인 입금 내역과 입출금 내역을 자동 반영하는 로컬 프로그램입니다.

## 핵심 원칙

- 정확한 입금만 자동 반영합니다.
- 기준 회비와 금액이 다른 입금은 기본적으로 `자동화_검토결과` 시트로 분리합니다.
- 비밀번호는 코드에 저장하지 않습니다.
- 장부 양식이 바뀌면 `config/sheet_mapping.yml`을 먼저 수정합니다.

## 설치

```bash
python -m venv .venv
source .venv/bin/activate  # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

## 실행 예시

```bash
python app.py \
  --kakao "input/카카오뱅크_거래내역.xlsx" \
  --ledger "input/F.I.T_장부_Vol_3.2.xlsm" \
  --members "config/members_template.csv" \
  --output "output/FIT_장부_자동작성.xlsm"
```

실행하면 카카오뱅크 파일 비밀번호와 장부 파일 비밀번호를 프롬프트에서 입력합니다.

비밀번호를 인자로 넘길 수도 있습니다. 다만 쉘 히스토리에 남을 수 있으므로 개인 PC에서만 사용하세요.

```bash
python app.py \
  --kakao "input/카카오뱅크_거래내역.xlsx" \
  --kakao-password "<카카오_비번>" \
  --ledger "input/F.I.T_장부_Vol_3.2.xlsm" \
  --ledger-password "<장부_비번>" \
  --members "config/members_template.csv" \
  --output "output/FIT_장부_자동작성.xlsm"
```

## 검토필요 입금까지 반영하고 싶을 때

기본값은 기준 회비와 금액이 다른 입금(예: 회비 배수가 아닌 금액)을 장부에 바로 쓰지 않습니다.

검토필요 입금도 반영하려면 아래 옵션을 추가합니다.

```bash
--auto-write-review
```

## 회원명단 파일

`config/members_template.csv`를 수정해서 사용합니다.

```csv
name,monthly_due,status,aliases,active
김준섭,20000,취직,준섭,true
홍길동,10000,미취직,길동,true
```

- `name`: 장부에 있는 회원명
- `monthly_due`: 월 회비
- `aliases`: 카카오뱅크 입금자명 별칭. 여러 개면 쉼표 대신 세미콜론 권장
- `active`: false면 자동 매칭에서 제외

## 파일별 역할

| 파일 | 역할 |
|---|---|
| `app.py` | CLI 실행 진입점 |
| `fit_ledger_automation/decryptor.py` | 비밀번호 걸린 엑셀 복호화 |
| `fit_ledger_automation/kakao_parser.py` | 카카오뱅크 거래내역 파싱 |
| `fit_ledger_automation/member_loader.py` | 회원명단 로딩 |
| `fit_ledger_automation/name_matcher.py` | 입금자명/회원 자동 매칭 |
| `fit_ledger_automation/ledger_writer.py` | 기존 장부 시트 입력 |
| `fit_ledger_automation/validator.py` | 잔액 검증 |
| `fit_ledger_automation/report_writer.py` | 검토결과 시트 생성 |
| `fit_ledger_automation/models.py` | 데이터 모델 |
| `config/sheet_mapping.yml` | 장부 시트/열 위치 설정 |
| `config/category_rules.yml` | 출금내역 자동 분류 규칙 |

## 현재 장부 기준 매핑

- 개인 입금 내역 시트
  - 회원명: 15행
  - 입금일/입금액: 16행
  - 월 행 시작: 17행
- 입출금 내역 시트
  - 시작 행: 17행
  - 일자: C열
  - 구분: E열
  - 내역: G열
  - 세부 내역: K열
  - 금액: S열
  - 잔액: W열

## 주의

- 매크로가 있는 `.xlsm`은 `openpyxl`의 `keep_vba=True`로 보존합니다. VBA 코드를 수정하지는 않습니다.
- 엑셀 수식 재계산은 openpyxl이 직접 수행하지 않습니다. 수식 계산이 필요하면 결과 파일을 Excel로 열어 저장하거나, 2차 버전에서 `xlwings`를 붙이면 됩니다.
- 같은 파일을 여러 번 실행하면 자동화 행이 중복될 수 있으므로, 생성된 결과 파일이 아니라 원본 템플릿을 기준으로 다시 실행하는 방식을 권장합니다.
