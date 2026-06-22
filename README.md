# SageLedger — 모임 회비 장부 자동화

카카오뱅크 모임통장 거래내역 엑셀을 읽어, 기존 모임 장부(개인 입금 내역 + 입출금 내역)에
이번 달 입금/이자/지출을 자동 반영합니다. 여러 모임(FIT, SCM, …)을 모임 이름만으로 처리합니다.

## 동작 요약

- **개인 입금 내역**: 카카오뱅크 `일반입금` 을 회원과 매칭해, 그 회원의 **첫 빈칸부터** 입금일/입금액을 채웁니다.
  몰아서 낸 입금은 월 회비 단위로 **등분**하여 같은 입금일로 여러 칸에 채웁니다(예: 220,000원 → 2만원×11칸).
- **입출금 내역**: 이번에 새로 채운 개인입금을 **입금월별 회비 정산 한 줄**로, 예금이자/출금(지출)을 각 행으로 추가합니다.
- **중복 방지**: 같은 `날짜+금액` 이 이미 있으면 건너뜁니다. 같은 파일을 다시 실행해도 중복되지 않습니다.
- **검토 분리**: 월 회비 배수가 아닌 입금, 미매칭 입금 등은 `자동화_검토결과` 시트로 분리합니다.
- **저장**: xlwings(Excel)로 열고 저장하여 **수식 재계산 + 비밀번호 보존**을 모두 처리합니다.
- **잔액 검증**: 카카오뱅크 최종 잔액과 장부 최종 잔액이 일치하는지 확인합니다.

## 요구사항

- Windows + **Microsoft Excel 설치** (xlwings 가 Excel 을 구동합니다)
- Python 3.10+

```bash
pip install -r requirements.txt
```

## 설정 (한 번만)

비밀번호와 회원 명단은 **git 에 올라가지 않는 로컬 파일**에만 둡니다.

1. `.env` (비밀번호)

   ```bash
   cp .env.example .env   # 그리고 실제 비밀번호 입력
   ```

   ```
   KAKAO_PASSWORD=...
   FIT_LEDGER_READ_PW=...
   FIT_LEDGER_WRITE_PW=...
   SCM_LEDGER_READ_PW=...
   SCM_LEDGER_WRITE_PW=...
   ```

2. `config/groups.yml` (모임/회원 명단)

   ```bash
   cp config/groups.example.yml config/groups.yml   # 그리고 실제 회원으로 수정
   ```

   - `dues`: status 별 월 회비 (취직 20000 / 미취직 10000)
   - 회원별 `status` 를 `미취직` 으로 두면 그 회원만 1만원으로 계산됩니다.

## 입력 파일

장부와 카카오뱅크 거래내역을 모임별 입력 폴더에 둡니다(`config/groups.yml` 의 glob 패턴 기준).

```
input/fit/F.I.T_장부_*.xlsm
input/fit/카카오뱅크_거래내역_*.xlsx
input/scm/...
```

## 실행

```bash
python app.py fit      # FIT 모임 처리
python app.py scm      # SCM 모임 처리
```

결과는 `output/<group>_장부_자동작성_<시각><확장자>` 로 저장됩니다(원본은 건드리지 않음).

## 주의

- **공개 저장소입니다.** 회원 실명·비밀번호·실제 엑셀 파일은 절대 커밋하지 마세요
  (`.env`, `config/groups.yml`, `input/`, `output/`, `*.xlsx`, `*.xlsm` 은 `.gitignore` 처리됨).
- 입출금 내역의 회비 정산은 **이번에 새로 채운 입금**만 합산합니다. 기존 수기 월별 정산과
  방식이 다를 수 있으니, 잔액 불일치 경고가 뜨면 중복/누락을 확인하세요.

## 구조

| 파일 | 역할 |
|---|---|
| `app.py` | CLI 진입점 (`python app.py <group>`) |
| `SageLedger/config.py` | `.env` + `groups.yml` 로딩, 모임 설정 해석 |
| `SageLedger/decryptor.py` | 비밀번호 걸린 엑셀 복호화(읽기) |
| `SageLedger/kakao_parser.py` | 카카오뱅크 거래내역 파싱 |
| `SageLedger/name_matcher.py` | 입금자명 ↔ 회원 매칭 |
| `SageLedger/planner.py` | 등분·빈칸순차·중복방지·회비정산 등 쓰기 계획 수립(openpyxl 읽기) |
| `SageLedger/excel_writer.py` | xlwings 로 계획 적용·재계산·저장(비번 보존) |
| `SageLedger/validator.py` | 잔액 검증 메시지 |
| `config/sheet_mapping.yml` | 장부 시트/열 위치 |
| `config/category_rules.yml` | 지출 자동 분류 규칙 |
