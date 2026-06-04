# AutoGraph Public Data

## 프로젝트 소개

AutoGraph Public Data는 GraphRAG Ontology Builder 외부 시연을 위한 공개 샘플 데이터 저장소입니다.

GraphRAG Ontology Builder는 데이터베이스 스키마(DDL)와 CSV 데이터를 기반으로 온톨로지와 Knowledge Graph를 자동 생성하고, 자연어 질문에 대해 그래프 근거 경로와 함께 답변하는 GraphRAG 데모 서비스입니다.

이 저장소에는 외부 공개가 가능한 샘플 DDL, CSV, 문서 데이터만 포함되어 있습니다. 내부 데이터, 고객 데이터, 운영 환경 설정값, API Key는 포함하지 않습니다.

- Demo: 준비 중
- Public Data: [moinji/AutoGraphPublicData](https://github.com/moinji/AutoGraphPublicData)

## 사용법

### 1. 공개 데이터 다운로드

```bash
git clone https://github.com/moinji/AutoGraphPublicData.git
cd AutoGraphPublicData
```

### 2. 데모 서비스 접속

배포 URL이 준비되면 브라우저에서 데모 서비스에 접속합니다.

```text
Demo URL: 준비 중
```

### 3. 추천 시연 데이터

다중 홉 질의 시연을 위해 `accounting` 데이터를 권장합니다.

```text
DDL: examples/schemas/demo_accounting.sql
CSV: examples/data/accounting/*.csv
Documents: examples/documents/accounting/*
```

`accounting` 데이터는 거래처, 청구서, 결제, 회계전표, 전표 라인, 계정과목, 부서, 직원, 예산, 회계기간 관계를 포함합니다. 특히 청구서에서 회계전표, 작성자, 승인자, 결제 처리자까지 이어지는 다중 홉 질의를 보여주기에 적합합니다.

### 4. DDL 업로드

Upload 화면에서 아래 DDL 파일을 업로드합니다.

```text
examples/schemas/demo_accounting.sql
```

업로드 후 파싱된 테이블, 컬럼, 관계를 확인하고 `Generate Ontology`를 실행합니다.

### 5. 온톨로지 확인 및 승인

Review 화면에서 자동 생성된 노드와 관계를 확인합니다.

필요하면 노드명, 관계명, 관계 방향을 수정한 뒤 저장하고 승인합니다.

### 6. CSV 데이터 업로드

승인된 온톨로지에 맞는 공개 CSV 파일을 업로드합니다.

```text
examples/data/accounting/account.csv
examples/data/accounting/budget.csv
examples/data/accounting/department.csv
examples/data/accounting/employee.csv
examples/data/accounting/fiscal_period.csv
examples/data/accounting/invoice.csv
examples/data/accounting/journal_entry.csv
examples/data/accounting/journal_line.csv
examples/data/accounting/payment.csv
examples/data/accounting/vendor.csv
```

CSV 업로드 후 `Build KG`를 실행하면 Knowledge Graph가 생성됩니다.

### 7. 자연어 질의 실행

Query 화면에서 질문을 입력합니다. 시연에서는 `Hybrid` 또는 `GraphRAG` 모드를 권장합니다.

추천 질문:

```text
삼성전자 거래처의 사업자등록번호가 어떻게 돼?
한국오피스는 무슨 품목 거래처야?
박소영 과장은 어느 부서 소속이야?
삼성전자 청구서를 전표로 기록한 작성자는 누구야?
삼성전자 청구서를 전표로 기록한 작성자와 승인자는 누구야?
김재현 부장이 승인한 삼성전자 관련 전표는 누가 작성했어?
부분 지급 상태인 삼성전자 청구서는 어떤 전표로 기록됐고 누가 작성했어?
부분 지급 상태인 삼성전자 청구서의 결제는 누가 처리했어?
한국오피스 청구서를 전표로 기록한 직원은 누구야?
```

답변과 함께 그래프 경로, 관련 엔티티, 근거 정보를 확인할 수 있습니다.

### 8. 문서 기반 검색 선택 사용

문서 기반 검색까지 함께 시연하려면 Documents 화면에서 accounting 문서를 업로드합니다.

```text
examples/documents/accounting/결산_및_재무보고_절차.md
examples/documents/accounting/매입_및_지급_관리.md
examples/documents/accounting/세금계산서_발행_안내.txt
examples/documents/accounting/전표_처리_및_내부통제.md
examples/documents/accounting/회계_기준_안내.md
examples/documents/accounting/예산_관리_및_원가_통제.md
examples/documents/accounting/결산_절차_매뉴얼.html
examples/documents/accounting/계정과목_체계_및_사용_지침.md
examples/documents/accounting/재무_보고서.pptx
examples/documents/accounting/감사 체크리스트.docx
```

문서 업로드 후 Hybrid 모드에서 구조화 데이터와 문서 근거를 함께 활용한 질의를 실행할 수 있습니다.

참고: `examples/documents/accounting/~$재무_보고서.pptx`는 Office 임시 파일이므로 업로드하지 않아도 됩니다.

### 9. 그래프 탐색

Explore 화면에서 생성된 Knowledge Graph를 시각적으로 탐색할 수 있습니다.

노드 검색, 관계 필터링, 이웃 노드 확장, Radial Map 보기를 사용해 회계 데이터의 연결 구조를 확인합니다.
