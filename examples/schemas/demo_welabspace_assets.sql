-- WeLabSpace 공식 교육 도메인 DDL
-- PostgreSQL 15+

DROP TABLE IF EXISTS certificates CASCADE;
DROP TABLE IF EXISTS feedback CASCADE;
DROP TABLE IF EXISTS evaluation_results CASCADE;
DROP TABLE IF EXISTS evaluations CASCADE;
DROP TABLE IF EXISTS submissions CASCADE;
DROP TABLE IF EXISTS assignments CASCADE;
DROP TABLE IF EXISTS enrollments CASCADE;
DROP TABLE IF EXISTS students CASCADE;
DROP TABLE IF EXISTS bootcamp_cohorts CASCADE;
DROP TABLE IF EXISTS courses CASCADE;
DROP TABLE IF EXISTS employees CASCADE;
DROP TABLE IF EXISTS departments CASCADE;
DROP TABLE IF EXISTS campuses CASCADE;

CREATE TABLE campuses (
    id BIGINT PRIMARY KEY,
    name VARCHAR(100) NOT NULL UNIQUE,
    city VARCHAR(50) NOT NULL,
    address VARCHAR(200) NOT NULL
);

CREATE TABLE departments (
    id BIGINT PRIMARY KEY,
    name VARCHAR(100) NOT NULL UNIQUE,
    description TEXT
);

CREATE TABLE employees (
    id BIGINT PRIMARY KEY,
    name VARCHAR(100) NOT NULL,
    email VARCHAR(150) NOT NULL UNIQUE,
    phone VARCHAR(20),
    department_id BIGINT NOT NULL REFERENCES departments(id),
    role VARCHAR(50) NOT NULL,
    hire_date DATE NOT NULL
);

CREATE TABLE courses (
    id BIGINT PRIMARY KEY,
    name VARCHAR(200) NOT NULL UNIQUE,
    description TEXT,
    category VARCHAR(50) NOT NULL,
    duration_weeks INTEGER NOT NULL CHECK (duration_weeks > 0),
    campus_id BIGINT NOT NULL REFERENCES campuses(id),
    lead_instructor_id BIGINT NOT NULL REFERENCES employees(id),
    max_students INTEGER NOT NULL CHECK (max_students > 0)
);

CREATE TABLE bootcamp_cohorts (
    id BIGINT PRIMARY KEY,
    course_id BIGINT NOT NULL REFERENCES courses(id),
    cohort_no INTEGER NOT NULL CHECK (cohort_no > 0),
    code VARCHAR(30) NOT NULL UNIQUE,
    name VARCHAR(120) NOT NULL,
    campus_id BIGINT NOT NULL REFERENCES campuses(id),
    instructor_id BIGINT NOT NULL REFERENCES employees(id),
    mentor_id BIGINT REFERENCES employees(id),
    start_date DATE NOT NULL,
    end_date DATE NOT NULL,
    status VARCHAR(20) NOT NULL CHECK (status IN ('planned', 'active', 'completed', 'cancelled')),
    UNIQUE (course_id, cohort_no),
    CHECK (start_date < end_date)
);

CREATE TABLE students (
    id BIGINT PRIMARY KEY,
    name VARCHAR(100) NOT NULL,
    email VARCHAR(150) NOT NULL UNIQUE,
    phone VARCHAR(20),
    birth_date DATE,
    registered_at DATE NOT NULL,
    current_status VARCHAR(20) NOT NULL CHECK (current_status IN ('active', 'alumni', 'leave'))
);

CREATE TABLE enrollments (
    id BIGINT PRIMARY KEY,
    student_id BIGINT NOT NULL REFERENCES students(id),
    cohort_id BIGINT NOT NULL REFERENCES bootcamp_cohorts(id),
    enrolled_at DATE NOT NULL,
    status VARCHAR(20) NOT NULL CHECK (status IN ('active', 'completed', 'dropped')),
    final_outcome VARCHAR(30) NOT NULL CHECK (final_outcome IN ('ongoing', 'certified', 'completed_not_certified', 'failed')),
    completion_notes TEXT,
    UNIQUE (student_id, cohort_id)
);

CREATE TABLE assignments (
    id BIGINT PRIMARY KEY,
    cohort_id BIGINT NOT NULL REFERENCES bootcamp_cohorts(id),
    title VARCHAR(200) NOT NULL,
    assignment_type VARCHAR(30) NOT NULL CHECK (assignment_type IN ('project', 'presentation', 'report')),
    due_date DATE NOT NULL,
    max_score INTEGER NOT NULL CHECK (max_score > 0),
    weight_percent NUMERIC(5,2) NOT NULL CHECK (weight_percent >= 0 AND weight_percent <= 100)
);

CREATE TABLE submissions (
    id BIGINT PRIMARY KEY,
    assignment_id BIGINT NOT NULL REFERENCES assignments(id),
    student_id BIGINT NOT NULL REFERENCES students(id),
    submitted_at DATE NOT NULL,
    score INTEGER CHECK (score >= 0 AND score <= 100),
    status VARCHAR(20) NOT NULL CHECK (status IN ('on_time', 'late')),
    feedback TEXT,
    UNIQUE (assignment_id, student_id)
);

CREATE TABLE evaluations (
    id BIGINT PRIMARY KEY,
    cohort_id BIGINT NOT NULL REFERENCES bootcamp_cohorts(id),
    title VARCHAR(200) NOT NULL,
    evaluation_type VARCHAR(30) NOT NULL CHECK (evaluation_type IN ('rubric', 'presentation', 'completion_gate')),
    criteria_summary TEXT NOT NULL,
    max_score INTEGER NOT NULL CHECK (max_score > 0),
    weight_percent NUMERIC(5,2) NOT NULL CHECK (weight_percent >= 0 AND weight_percent <= 100),
    scheduled_on DATE NOT NULL
);

CREATE TABLE evaluation_results (
    id BIGINT PRIMARY KEY,
    evaluation_id BIGINT NOT NULL REFERENCES evaluations(id),
    student_id BIGINT NOT NULL REFERENCES students(id),
    attendance_rate NUMERIC(5,2) NOT NULL CHECK (attendance_rate >= 0 AND attendance_rate <= 100),
    score INTEGER NOT NULL CHECK (score >= 0 AND score <= 100),
    result_grade VARCHAR(5) NOT NULL,
    passed BOOLEAN NOT NULL,
    evaluated_at DATE NOT NULL,
    reviewer_notes TEXT,
    UNIQUE (evaluation_id, student_id)
);

CREATE TABLE feedback (
    id BIGINT PRIMARY KEY,
    cohort_id BIGINT NOT NULL REFERENCES bootcamp_cohorts(id),
    student_id BIGINT NOT NULL REFERENCES students(id),
    mentor_id BIGINT NOT NULL REFERENCES employees(id),
    feedback_type VARCHAR(20) NOT NULL CHECK (feedback_type IN ('progress', 'career')),
    feedback_text TEXT NOT NULL,
    created_at DATE NOT NULL
);

CREATE TABLE certificates (
    id BIGINT PRIMARY KEY,
    student_id BIGINT NOT NULL REFERENCES students(id),
    enrollment_id BIGINT NOT NULL UNIQUE REFERENCES enrollments(id),
    course_id BIGINT NOT NULL REFERENCES courses(id),
    issued_at DATE NOT NULL,
    certificate_number VARCHAR(50) NOT NULL UNIQUE,
    status VARCHAR(20) NOT NULL CHECK (status IN ('issued', 'revoked'))
);
