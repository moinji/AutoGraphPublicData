-- Classic Models Database Schema: scale model car dealership
-- Source: NTU Singapore (Nanyang Technological University)
-- (www3.ntu.edu.sg/home/ehchua/programming/sql/SampleDatabases.html)
-- Converted to PostgreSQL dialect

CREATE TABLE offices (
    office_code VARCHAR(10) NOT NULL PRIMARY KEY,
    city VARCHAR(50) NOT NULL,
    phone VARCHAR(50) NOT NULL,
    address_line1 VARCHAR(50) NOT NULL,
    address_line2 VARCHAR(50),
    state VARCHAR(50),
    country VARCHAR(50) NOT NULL,
    postal_code VARCHAR(15) NOT NULL,
    territory VARCHAR(10) NOT NULL
);

CREATE TABLE employees (
    employee_number SERIAL PRIMARY KEY,
    last_name VARCHAR(50) NOT NULL,
    first_name VARCHAR(50) NOT NULL,
    extension VARCHAR(10) NOT NULL,
    email VARCHAR(100) NOT NULL,
    office_code VARCHAR(10) NOT NULL REFERENCES offices(office_code),
    reports_to INTEGER REFERENCES employees(employee_number),
    job_title VARCHAR(50) NOT NULL
);

CREATE TABLE customers (
    customer_number SERIAL PRIMARY KEY,
    customer_name VARCHAR(50) NOT NULL,
    contact_last_name VARCHAR(50) NOT NULL,
    contact_first_name VARCHAR(50) NOT NULL,
    phone VARCHAR(50) NOT NULL,
    address_line1 VARCHAR(50) NOT NULL,
    address_line2 VARCHAR(50),
    city VARCHAR(50) NOT NULL,
    state VARCHAR(50),
    postal_code VARCHAR(15),
    country VARCHAR(50) NOT NULL,
    sales_rep_employee_number INTEGER REFERENCES employees(employee_number),
    credit_limit INTEGER
);

CREATE TABLE product_lines (
    product_line VARCHAR(50) NOT NULL PRIMARY KEY,
    text_description VARCHAR(4000),
    html_description TEXT,
    image BYTEA
);

CREATE TABLE products (
    product_code VARCHAR(15) NOT NULL PRIMARY KEY,
    product_name VARCHAR(70) NOT NULL,
    product_line VARCHAR(50) NOT NULL REFERENCES product_lines(product_line),
    product_scale VARCHAR(10) NOT NULL,
    product_vendor VARCHAR(50) NOT NULL,
    product_description TEXT NOT NULL,
    quantity_in_stock SMALLINT NOT NULL,
    buy_price DECIMAL(8,2) NOT NULL,
    msrp DECIMAL(8,2) NOT NULL
);

CREATE TABLE orders (
    order_number SERIAL PRIMARY KEY,
    order_date DATE NOT NULL,
    required_date DATE NOT NULL,
    shipped_date DATE,
    status VARCHAR(15) NOT NULL,
    comments TEXT,
    customer_number INTEGER NOT NULL REFERENCES customers(customer_number)
);

CREATE TABLE order_details (
    order_number INTEGER NOT NULL REFERENCES orders(order_number),
    product_code VARCHAR(15) NOT NULL REFERENCES products(product_code),
    quantity_ordered SMALLINT NOT NULL,
    price_each DECIMAL(7,2) NOT NULL,
    order_line_number SMALLINT NOT NULL,
    PRIMARY KEY (order_number, product_code)
);

CREATE TABLE payments (
    customer_number INTEGER NOT NULL REFERENCES customers(customer_number),
    check_number VARCHAR(50) NOT NULL,
    payment_date DATE NOT NULL,
    amount DECIMAL(8,2) NOT NULL,
    PRIMARY KEY (customer_number, check_number)
);
