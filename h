CREATE TABLE live_trainer (
    sr_no INT AUTO_INCREMENT PRIMARY KEY,
    faculty_name VARCHAR(255),
    ticket_no VARCHAR(50),
    mail_id VARCHAR(255),
    mobile_number VARCHAR(20),
    area VARCHAR(100),
    dept VARCHAR(100),
    factory VARCHAR(100),
    reporting_manager_name VARCHAR(255),
    reporting_manager_mail_id VARCHAR(255),
    expertise_area VARCHAR(255),
    expertise_category VARCHAR(100),
    hr_coordinator_name VARCHAR(255),
    remark TEXT
);

CREATE TABLE fst (
    sr_no INT AUTO_INCREMENT PRIMARY KEY,
    ticket_no VARCHAR(50),
    name VARCHAR(100),
    gender VARCHAR(20),
    employee_category VARCHAR(50),
    plant_location VARCHAR(100),
    joined_year VARCHAR(10),
    date_from DATE,
    date_to DATE,
    shift VARCHAR(50),
    learning_hours VARCHAR(20),
    training_name VARCHAR(100),
    batch_number VARCHAR(50),
    training_venue_name VARCHAR(100),
    faculty_name VARCHAR(100),
    fst_cell_name VARCHAR(100),
    remark TEXT
);

CREATE TABLE induction (
    sr_no INT AUTO_INCREMENT PRIMARY KEY,
    ticket_no VARCHAR(50),
    name VARCHAR(100),
    gender VARCHAR(10),
    employee_category VARCHAR(50),
    plant_location VARCHAR(100),
    joined_year INT,
    date_from DATE,
    date_to DATE,
    shift VARCHAR(50),
    learning_hours DECIMAL(5,2),
    training_name VARCHAR(200),
    batch_number VARCHAR(50),
    training_venue_name VARCHAR(200),
    faculty_name VARCHAR(100),
    subject_name VARCHAR(100),
    remark TEXT
);

CREATE TABLE jta (
    sr_no INT AUTO_INCREMENT PRIMARY KEY,
    ticket_no VARCHAR(50),
    name VARCHAR(100),
    gender VARCHAR(10),
    joining_year INT,
    date_of_joining DATE,
    jta_batch_number VARCHAR(50),
    date_of_separation DATE,
    trade VARCHAR(100),
    final_result VARCHAR(50),
    training_name VARCHAR(200),
    employee_category VARCHAR(255),
    status VARCHAR(255)
);


CREATE TABLE kaushalya (
    sr_no INT AUTO_INCREMENT PRIMARY KEY,
    ticket_no VARCHAR(50),
    name VARCHAR(50),
    gender VARCHAR(50),
    joining_year YEAR,
    date_of_joining DATE,
    kaushalya_batch_no VARCHAR(50),
    trade VARCHAR(50),
    dei_batch VARCHAR(50),
    sem_1_pass_fail VARCHAR(50),
    sem_2_pass_fail VARCHAR(50),
    sem_3_pass_fail VARCHAR(50),
    sem_4_pass_fail VARCHAR(50),
    sem_5_pass_fail VARCHAR(50),
    sem_6_pass_fail VARCHAR(50),
    final_result VARCHAR(50),
    placement_drive VARCHAR(50),
    training_name VARCHAR(50),
    remark TEXT,
    created_at TIMESTAMP,
    updated_at TIMESTAMP
);

CREATE TABLE lakshya (
    sr_no INT AUTO_INCREMENT PRIMARY KEY,
    ticket_no VARCHAR(50),
    name VARCHAR(50),
    gender VARCHAR(50),
    course_joining_year YEAR,
    date_of_joining DATE,
    date_of_separation DATE,
    lakshya_batch_no VARCHAR(50),
    diploma_name VARCHAR(50),
    diploma_trainee_inplant_shop VARCHAR(50),
    semester_1_pass_fail VARCHAR(50),
    semester_2_pass_fail VARCHAR(50),
    second_year_pass_fail VARCHAR(50),
    third_year_pass_fail VARCHAR(50),
    fourth_year_pass_fail VARCHAR(50),
    final_result VARCHAR(50),
    training_name VARCHAR(50),
    remark TEXT,
    created_at TIMESTAMP,
    updated_at TIMESTAMP
);


CREATE TABLE pragati (
    sr_no INT AUTO_INCREMENT PRIMARY KEY,
    ticket_no VARCHAR(50),
    name VARCHAR(100),
    gender VARCHAR(20),
    employee_category VARCHAR(50),
    factory VARCHAR(100),
    course_joining_year VARCHAR(10),
    date_of_joining DATE,
    pragati_batch_number VARCHAR(50),
    diploma_name VARCHAR(100),
    first_year_result VARCHAR(50),
    second_year_result VARCHAR(50),
    final_result VARCHAR(50),
    training_name VARCHAR(100),
    remark TEXT
);


CREATE TABLE ta (
    sr_no INT AUTO_INCREMENT PRIMARY KEY,
    ticket_no VARCHAR(50),
    name VARCHAR(100),
    gender VARCHAR(10),
    joining_year INT,
    date_of_joining DATE,
    ta_batch_number VARCHAR(50),
    date_of_separation DATE,
    trade VARCHAR(100),
    final_result VARCHAR(50),
    training_name VARCHAR(200)
);


