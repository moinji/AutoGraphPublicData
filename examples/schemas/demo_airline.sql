-- Airline Database Schema: flights, pilots, aircraft, bookings
-- Source: Johns Hopkins University CS415 lecture materials
-- (ugrad.cs.jhu.edu/~cs415/oracle-lecture/make-airline-DB.sql)
-- Converted to PostgreSQL dialect

CREATE TABLE person (
    name VARCHAR(15) NOT NULL PRIMARY KEY,
    address VARCHAR(15) NOT NULL,
    phone VARCHAR(12)
);

CREATE TABLE employee (
    name VARCHAR(15) NOT NULL PRIMARY KEY REFERENCES person(name),
    salary DECIMAL(10,2),
    emp_no SMALLINT UNIQUE NOT NULL
);

CREATE TABLE pilot (
    emp_no SMALLINT UNIQUE REFERENCES employee(emp_no)
);

CREATE TABLE plane (
    maker VARCHAR(15) NOT NULL,
    model_no VARCHAR(15) NOT NULL PRIMARY KEY
);

CREATE TABLE aircraft (
    serial_no SMALLINT NOT NULL,
    model_no VARCHAR(15) NOT NULL REFERENCES plane(model_no),
    PRIMARY KEY (serial_no, model_no)
);

CREATE TABLE flight (
    num SMALLINT NOT NULL PRIMARY KEY,
    origin VARCHAR(3),
    dest VARCHAR(3),
    dep_time VARCHAR(5),
    arr_time VARCHAR(5)
);

CREATE TABLE departure (
    dep_date VARCHAR(6) NOT NULL,
    num SMALLINT NOT NULL REFERENCES flight(num),
    PRIMARY KEY (dep_date, num)
);

CREATE TABLE booked_on (
    name VARCHAR(15) NOT NULL,
    dep_date VARCHAR(6) NOT NULL,
    num SMALLINT NOT NULL,
    PRIMARY KEY (name, dep_date, num),
    FOREIGN KEY (dep_date, num) REFERENCES departure(dep_date, num)
);

CREATE TABLE assigned_to (
    emp_no SMALLINT NOT NULL REFERENCES employee(emp_no),
    dep_date VARCHAR(6) NOT NULL,
    num SMALLINT NOT NULL,
    PRIMARY KEY (emp_no, dep_date, num),
    FOREIGN KEY (dep_date, num) REFERENCES departure(dep_date, num)
);

CREATE TABLE can_fly (
    emp_no SMALLINT NOT NULL REFERENCES employee(emp_no),
    model_no VARCHAR(15) NOT NULL REFERENCES plane(model_no),
    PRIMARY KEY (emp_no, model_no)
);

CREATE TABLE equipment (
    dep_date VARCHAR(6) NOT NULL,
    num SMALLINT NOT NULL,
    serial_no SMALLINT NOT NULL,
    model_no VARCHAR(15) NOT NULL,
    PRIMARY KEY (dep_date, num, serial_no, model_no),
    FOREIGN KEY (dep_date, num) REFERENCES departure(dep_date, num),
    FOREIGN KEY (serial_no, model_no) REFERENCES aircraft(serial_no, model_no)
);
