# AutoGraph Public Data

AutoGraph의 외부 시연과 기능 검증에 사용하는 공개 샘플 데이터 저장소입니다. 고객 데이터, 운영 데이터, 자격 증명과 환경 설정은 포함하지 않습니다.

## 폴더 구조

| 경로 | 설명 |
|---|---|
| `datasets` | 스키마, CSV와 문서를 도메인별로 함께 보관합니다. |
| `schema-gallery` | 출처와 라이선스가 확인된 DDL 전용 예제를 추가하기 위한 공간입니다. |
| `tools` | 샘플 CSV와 문서를 다시 생성하는 도구를 보관합니다. |
| `catalog` | 데이터셋의 구성과 공개 검토 상태를 기록합니다. |

## 데이터셋 사용 방법

각 데이터셋은 같은 구조를 사용합니다.

```text
datasets/<dataset-id>/
├── schema.sql
├── tables/
└── documents/
```

회계 데모를 실행할 때에는 다음 순서로 파일을 사용합니다.

1. `datasets/accounting/schema.sql`을 업로드합니다.
2. `datasets/accounting/tables`의 CSV 파일을 업로드합니다.
3. Knowledge Graph를 생성합니다.
4. 문서 검색이 필요하면 `datasets/accounting/documents`의 파일을 업로드합니다.

데이터셋의 목록과 구성은 `catalog/datasets.yaml`에서 확인합니다.

## 공개와 라이선스 정책

`datasets`에는 AutoGraph가 공개용으로 만든 샘플만 둡니다. 새 파일을 추가하기 전에는 실제 개인정보, 고객명, 내부 주소와 자격 증명이 없는지 확인합니다. 현재 데이터셋은 구조 정리를 마쳤지만, 회사 GitHub에 게시하기 전에는 개인정보와 상표 사용 여부를 한 번 더 검토해야 합니다.

출처나 라이선스가 확인되지 않은 DDL은 공개 저장소에 두지 않습니다. 이번 정리에서 발견한 외부 SQL 덤프는 내부 검토 영역으로 옮겼으며, 처리 내역은 `THIRD_PARTY_NOTICES.md`에 기록했습니다.

## 데이터 재생성

다음 도구는 저장소 루트의 `datasets` 구조에 맞추어 결과를 생성합니다.

```bash
python tools/generate_test_data.py
python tools/generate_test_ppts.py
python tools/generate_stress_data.py
```

생성 후에는 스키마의 테이블 이름과 CSV 파일 이름이 일치하는지 확인하고, 공개 검토를 거친 뒤 커밋합니다.
