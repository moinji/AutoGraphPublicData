-- ═══════════════════════════════════════════════════════════════
-- 임상 연구 병원 DDL (Clinical Research Hospital)
--
-- 설계 의도:
--   FK가 의도적으로 불완전하여 LLM 보강(Stage 2)이 차이를 만든다.
--   - FK 있는 관계: 11개 (규칙 엔진이 잡음)
--   - FK 없는 암묵적 관계: 5~7개 (LLM이 추가해야 함)
--
-- 암묵적 관계 예시:
--   1. drug_interaction: drug_a_id, drug_b_id → FK 없음 (INTERACTS_WITH)
--   2. referral: referring_doctor_id → FK 없음 (REFERRED_BY)
--   3. clinical_trial_enrollment: patient_id, trial_id → FK 없음 (ENROLLED_IN)
--   4. research_publication: lead_doctor_id → FK 없음 (AUTHORED_BY)
--   5. department.head_doctor_id → FK 없음 (HEADS)
-- ═══════════════════════════════════════════════════════════════

-- 진료과
CREATE TABLE department (
    id INT PRIMARY KEY,
    name VARCHAR(100) NOT NULL,
    floor INT,
    head_doctor_id INT,          -- ★ FK 없음! LLM이 Doctor→Department HEADS 관계 추론 필요
    established_year INT
);

-- 의사
CREATE TABLE doctor (
    id INT PRIMARY KEY,
    name VARCHAR(100) NOT NULL,
    specialty VARCHAR(100),
    license_no VARCHAR(50),
    phone VARCHAR(20),
    department_id INT,
    FOREIGN KEY (department_id) REFERENCES department(id)
);

-- 환자
CREATE TABLE patient (
    id INT PRIMARY KEY,
    name VARCHAR(100) NOT NULL,
    birth_date DATE,
    gender VARCHAR(10),
    phone VARCHAR(20),
    blood_type VARCHAR(5),
    emergency_contact VARCHAR(100),
    insurance_type VARCHAR(50)
);

-- 진료 예약
CREATE TABLE appointment (
    id INT PRIMARY KEY,
    patient_id INT NOT NULL,
    doctor_id INT NOT NULL,
    appointment_date DATE,
    status VARCHAR(20) DEFAULT 'scheduled',
    chief_complaint TEXT,
    FOREIGN KEY (patient_id) REFERENCES patient(id),
    FOREIGN KEY (doctor_id) REFERENCES doctor(id)
);

-- 진단
CREATE TABLE diagnosis (
    id INT PRIMARY KEY,
    appointment_id INT NOT NULL,
    icd_code VARCHAR(20),
    disease_name VARCHAR(200),
    severity VARCHAR(20),
    notes TEXT,
    FOREIGN KEY (appointment_id) REFERENCES appointment(id)
);

-- 약물
CREATE TABLE drug (
    id INT PRIMARY KEY,
    name VARCHAR(200) NOT NULL,
    generic_name VARCHAR(200),
    category VARCHAR(100),
    manufacturer VARCHAR(200),
    unit_price DECIMAL(10,2),
    requires_prescription BOOLEAN DEFAULT TRUE
);

-- 처방
CREATE TABLE prescription (
    id INT PRIMARY KEY,
    appointment_id INT NOT NULL,
    drug_id INT NOT NULL,
    dosage VARCHAR(100),
    frequency VARCHAR(50),
    duration_days INT,
    FOREIGN KEY (appointment_id) REFERENCES appointment(id),
    FOREIGN KEY (drug_id) REFERENCES drug(id)
);

-- ★ 약물 상호작용 — FK 없음! LLM이 Drug-[INTERACTS_WITH]->Drug 관계 추론 필요
CREATE TABLE drug_interaction (
    id INT PRIMARY KEY,
    drug_a_id INT NOT NULL,      -- FK 의도적 누락
    drug_b_id INT NOT NULL,      -- FK 의도적 누락
    severity VARCHAR(20),
    description TEXT
);

-- 검사
CREATE TABLE lab_test (
    id INT PRIMARY KEY,
    name VARCHAR(200) NOT NULL,
    category VARCHAR(100),
    normal_range VARCHAR(100),
    unit VARCHAR(50)
);

-- 검사 오더
CREATE TABLE lab_order (
    id INT PRIMARY KEY,
    appointment_id INT NOT NULL,
    lab_test_id INT NOT NULL,
    order_date DATE,
    result_value VARCHAR(100),
    result_status VARCHAR(20),
    FOREIGN KEY (appointment_id) REFERENCES appointment(id),
    FOREIGN KEY (lab_test_id) REFERENCES lab_test(id)
);

-- 병동
CREATE TABLE ward (
    id INT PRIMARY KEY,
    name VARCHAR(100),
    department_id INT,
    capacity INT,
    ward_type VARCHAR(50),
    FOREIGN KEY (department_id) REFERENCES department(id)
);

-- 입원
CREATE TABLE admission (
    id INT PRIMARY KEY,
    patient_id INT NOT NULL,
    ward_id INT NOT NULL,
    attending_doctor_id INT NOT NULL,
    admit_date DATE,
    discharge_date DATE,
    admission_reason TEXT,
    FOREIGN KEY (patient_id) REFERENCES patient(id),
    FOREIGN KEY (ward_id) REFERENCES ward(id),
    FOREIGN KEY (attending_doctor_id) REFERENCES doctor(id)
);

-- ★ 의뢰/전원 — referring_doctor_id에 FK 없음! LLM이 Doctor-[REFERRED_BY]->Referral 추론 필요
CREATE TABLE referral (
    id INT PRIMARY KEY,
    patient_id INT NOT NULL,
    referring_doctor_id INT NOT NULL,    -- FK 의도적 누락
    receiving_doctor_id INT NOT NULL,    -- FK 의도적 누락
    referral_date DATE,
    reason TEXT,
    status VARCHAR(20) DEFAULT 'pending',
    FOREIGN KEY (patient_id) REFERENCES patient(id)
);

-- ★ 임상시험
CREATE TABLE clinical_trial (
    id INT PRIMARY KEY,
    title VARCHAR(300) NOT NULL,
    phase VARCHAR(20),
    start_date DATE,
    end_date DATE,
    status VARCHAR(30),
    target_enrollment INT,
    sponsor VARCHAR(200)
);

-- ★ 임상시험 등록 — patient_id, trial_id에 FK 없음! LLM이 ENROLLED_IN 추론 필요
CREATE TABLE trial_enrollment (
    id INT PRIMARY KEY,
    patient_id INT NOT NULL,             -- FK 의도적 누락
    trial_id INT NOT NULL,               -- FK 의도적 누락
    enrollment_date DATE,
    consent_signed BOOLEAN DEFAULT FALSE,
    status VARCHAR(20) DEFAULT 'screening'
);

-- ★ 연구 논문 — lead_doctor_id에 FK 없음! LLM이 AUTHORED_BY 추론 필요
CREATE TABLE research_publication (
    id INT PRIMARY KEY,
    title VARCHAR(500) NOT NULL,
    journal VARCHAR(200),
    publish_date DATE,
    lead_doctor_id INT NOT NULL,         -- FK 의도적 누락
    trial_id INT,                        -- FK 의도적 누락
    doi VARCHAR(100),
    citation_count INT DEFAULT 0
);
