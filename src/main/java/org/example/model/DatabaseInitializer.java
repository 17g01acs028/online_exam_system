package org.example.model;

import java.sql.*;

public class DatabaseInitializer {
    static void createTablesIfNotExist(Connection connection,String databaseType) throws SQLException {
        boolean isMySQL = "Mysql".equalsIgnoreCase(databaseType);
        boolean isPostgreSQL = "postgresql".equalsIgnoreCase(databaseType);
        boolean isMSSQL = "mssql".equalsIgnoreCase(databaseType);


        try (Statement statement = connection.createStatement()) {


            // Create Tables
            statement.executeUpdate(getStringTeacher(isMySQL, isPostgreSQL, isMSSQL));
            statement.executeUpdate(getStringSubject(isMySQL, isPostgreSQL, isMSSQL));
            statement.executeUpdate(getStringClass(isMySQL, isPostgreSQL, isMSSQL));
            statement.executeUpdate(getStringExam(isMySQL, isPostgreSQL, isMSSQL));
            statement.executeUpdate(getStringExamSchedule(isMySQL, isPostgreSQL, isMSSQL));
            statement.executeUpdate(getStringExamQuestions(isMySQL, isPostgreSQL, isMSSQL));
            statement.executeUpdate(getStringStudents(isMySQL, isPostgreSQL, isMSSQL));
            statement.executeUpdate(getStringExamQuestionOptions(isMySQL, isPostgreSQL, isMSSQL));
            statement.executeUpdate(getStringExamQuestionResponses(isMySQL, isPostgreSQL, isMSSQL));



        if(isMySQL) {
            // Generate indexes
            checkAndCreateIndex(connection, "teacher", "idx_teacher_tsc_number", "tsc_number");
            checkAndCreateIndex(connection, "teacher", "idx_teacher_username", "username");
            checkAndCreateIndex(connection, "teacher", "idx_teacher_id_number", "id_number");
            checkAndCreateIndex(connection, "teacher", "idx_teacher_phone", "phone");
            checkAndCreateIndex(connection, "teacher", "idx_teacher_email", "email");
            checkAndCreateIndex(connection, "teacher", "idx_teacher_date_created", "date_created");
            checkAndCreateIndex(connection, "teacher", "idx_teacher_date_modified", "date_modified");

            // Create class table
            checkAndCreateIndex(connection, "class", "idx_class_class_name", "class_name");
            checkAndCreateIndex(connection, "class", "idx_class_date_created", "date_created");
            checkAndCreateIndex(connection, "class", "idx_class_date_modified", "date_modified");

            // Create exam table
            checkAndCreateIndex(connection, "exam", "idx_exam_exam_name", "exam_name");
            checkAndCreateIndex(connection, "exam", "idx_exam_date_created", "date_created");
            checkAndCreateIndex(connection, "exam", "idx_exam_date_modified", "date_modified");

            // Create exam_schedule table
            checkAndCreateIndex(connection, "exam_schedule", "idx_exam_schedule_exam_date", "exam_date");
            checkAndCreateIndex(connection, "exam_schedule", "idx_exam_schedule_month_year", "month_year");
            checkAndCreateIndex(connection, "exam_schedule", "idx_exam_schedule_date_created", "date_created");
            checkAndCreateIndex(connection, "exam_schedule", "idx_exam_schedule_date_modified", "date_modified");

            // Create questions table
            checkAndCreateIndex(connection, "questions", "idx_questions_date_created", "date_created");
            checkAndCreateIndex(connection, "questions", "idx_questions_date_modified", "date_modified");

            // Create student table
            checkAndCreateIndex(connection, "student", "idx_student_firstname", "firstname");
            checkAndCreateIndex(connection, "student", "idx_student_lastname", "lastname");
            checkAndCreateIndex(connection, "student", "idx_student_username", "username");
            checkAndCreateIndex(connection, "student", "idx_student_gender", "gender");
            checkAndCreateIndex(connection, "student", "idx_student_date_of_birth", "date_of_birth");
            checkAndCreateIndex(connection, "student", "idx_student_guardian_name", "guardian_name");
            checkAndCreateIndex(connection, "student", "idx_student_guardian_phone", "guardian_phone");
            checkAndCreateIndex(connection, "student", "idx_student_class_id", "class_id");
            checkAndCreateIndex(connection, "student", "idx_student_home_address", "home_address");
            checkAndCreateIndex(connection, "student", "idx_student_city", "city");
            checkAndCreateIndex(connection, "student", "idx_student_postal_address", "postal_address");
            checkAndCreateIndex(connection, "student", "idx_student_student_reg_number", "student_reg_number");
            checkAndCreateIndex(connection, "student", "idx_student_date_created", "date_created");
            checkAndCreateIndex(connection, "student", "idx_student_date_modified", "date_modified");

            // Create options table
            checkAndCreateIndex(connection, "options", "idx_options_date_created", "date_created");
            checkAndCreateIndex(connection, "options", "idx_options_date_modified", "date_modified");

            // Create responses table
            checkAndCreateIndex(connection, "responses", "idx_responses_date_created", "date_created");
            checkAndCreateIndex(connection, "responses", "idx_responses_date_modified", "date_modified");
        }

        if(isPostgreSQL){
            // Generate indexes
            checkAndCreateIndexPostgres(connection, "teacher", "idx_teacher_tsc_number", "tsc_number");
            checkAndCreateIndexPostgres(connection, "teacher", "idx_teacher_username", "username");
            checkAndCreateIndexPostgres(connection, "teacher", "idx_teacher_id_number", "id_number");
            checkAndCreateIndexPostgres(connection, "teacher", "idx_teacher_phone", "phone");
            checkAndCreateIndexPostgres(connection, "teacher", "idx_teacher_email", "email");
            checkAndCreateIndexPostgres(connection, "teacher", "idx_teacher_date_created", "date_created");
            checkAndCreateIndexPostgres(connection, "teacher", "idx_teacher_date_modified", "date_modified");

            // Create class table
            checkAndCreateIndexPostgres(connection, "class", "idx_class_class_name", "class_name");
            checkAndCreateIndexPostgres(connection, "class", "idx_class_date_created", "date_created");
            checkAndCreateIndexPostgres(connection, "class", "idx_class_date_modified", "date_modified");

            // Create exam table
            checkAndCreateIndexPostgres(connection, "exam", "idx_exam_exam_name", "exam_name");
            checkAndCreateIndexPostgres(connection, "exam", "idx_exam_date_created", "date_created");
            checkAndCreateIndexPostgres(connection, "exam", "idx_exam_date_modified", "date_modified");

            // Create exam_schedule table
            checkAndCreateIndexPostgres(connection, "exam_schedule", "idx_exam_schedule_exam_date", "exam_date");
            checkAndCreateIndexPostgres(connection, "exam_schedule", "idx_exam_schedule_month_year", "month_year");
            checkAndCreateIndexPostgres(connection, "exam_schedule", "idx_exam_schedule_date_created", "date_created");
            checkAndCreateIndexPostgres(connection, "exam_schedule", "idx_exam_schedule_date_modified", "date_modified");

            // Create questions table
            checkAndCreateIndexPostgres(connection, "questions", "idx_questions_date_created", "date_created");
            checkAndCreateIndexPostgres(connection, "questions", "idx_questions_date_modified", "date_modified");

            // Create student table
            checkAndCreateIndexPostgres(connection, "student", "idx_student_firstname", "firstname");
            checkAndCreateIndexPostgres(connection, "student", "idx_student_lastname", "lastname");
            checkAndCreateIndexPostgres(connection, "student", "idx_student_username", "username");
            checkAndCreateIndexPostgres(connection, "student", "idx_student_gender", "gender");
            checkAndCreateIndexPostgres(connection, "student", "idx_student_date_of_birth", "date_of_birth");
            checkAndCreateIndexPostgres(connection, "student", "idx_student_guardian_name", "guardian_name");
            checkAndCreateIndexPostgres(connection, "student", "idx_student_guardian_phone", "guardian_phone");
            checkAndCreateIndexPostgres(connection, "student", "idx_student_class_id", "class_id");
            checkAndCreateIndexPostgres(connection, "student", "idx_student_home_address", "home_address");
            checkAndCreateIndexPostgres(connection, "student", "idx_student_city", "city");
           // checkAndCreateIndexPostgres(connection, "student", "idx_student_postal_address", "postal_address");
            checkAndCreateIndexPostgres(connection, "student", "idx_student_student_reg_number", "student_reg_number");
            checkAndCreateIndexPostgres(connection, "student", "idx_student_date_created", "date_created");
            checkAndCreateIndexPostgres(connection, "student", "idx_student_date_modified", "date_modified");

            // Create options table
            checkAndCreateIndexPostgres(connection, "options", "idx_options_date_created", "date_created");
            checkAndCreateIndexPostgres(connection, "options", "idx_options_date_modified", "date_modified");

            // Create responses table
            checkAndCreateIndexPostgres(connection, "responses", "idx_responses_date_created", "date_created");
            checkAndCreateIndexPostgres(connection, "responses", "idx_responses_date_modified", "date_modified");

        }
        if(isMSSQL){
            // Generate indexes
            checkAndCreateIndexMSSQL(connection, "teacher", "idx_teacher_tsc_number", "tsc_number");
            checkAndCreateIndexMSSQL(connection, "teacher", "idx_teacher_username", "username");
            checkAndCreateIndexMSSQL(connection, "teacher", "idx_teacher_id_number", "id_number");
            checkAndCreateIndexMSSQL(connection, "teacher", "idx_teacher_phone", "phone");
            checkAndCreateIndexMSSQL(connection, "teacher", "idx_teacher_email", "email");
            checkAndCreateIndexMSSQL(connection, "teacher", "idx_teacher_date_created", "date_created");
            checkAndCreateIndexMSSQL(connection, "teacher", "idx_teacher_date_modified", "date_modified");

            // Create class table
            checkAndCreateIndexMSSQL(connection, "class", "idx_class_class_name", "class_name");
            checkAndCreateIndexMSSQL(connection, "class", "idx_class_date_created", "date_created");
            checkAndCreateIndexMSSQL(connection, "class", "idx_class_date_modified", "date_modified");

            // Create exam table
            checkAndCreateIndexMSSQL(connection, "exam", "idx_exam_exam_name", "exam_name");
            checkAndCreateIndexMSSQL(connection, "exam", "idx_exam_date_created", "date_created");
            checkAndCreateIndexMSSQL(connection, "exam", "idx_exam_date_modified", "date_modified");

            // Create exam_schedule table
            checkAndCreateIndexMSSQL(connection, "exam_schedule", "idx_exam_schedule_exam_date", "exam_date");
            checkAndCreateIndexMSSQL(connection, "exam_schedule", "idx_exam_schedule_month_year", "month_year");
            checkAndCreateIndexMSSQL(connection, "exam_schedule", "idx_exam_schedule_date_created", "date_created");
            checkAndCreateIndexMSSQL(connection, "exam_schedule", "idx_exam_schedule_date_modified", "date_modified");

            // Create questions table
            checkAndCreateIndexMSSQL(connection, "questions", "idx_questions_date_created", "date_created");
            checkAndCreateIndexMSSQL(connection, "questions", "idx_questions_date_modified", "date_modified");

            // Create student table
            checkAndCreateIndexMSSQL(connection, "student", "idx_student_firstname", "firstname");
            checkAndCreateIndexMSSQL(connection, "student", "idx_student_lastname", "lastname");
            checkAndCreateIndexMSSQL(connection, "student", "idx_student_username", "username");
            checkAndCreateIndexMSSQL(connection, "student", "idx_student_gender", "gender");
            checkAndCreateIndexMSSQL(connection, "student", "idx_student_date_of_birth", "date_of_birth");
            checkAndCreateIndexMSSQL(connection, "student", "idx_student_guardian_name", "guardian_name");
            checkAndCreateIndexMSSQL(connection, "student", "idx_student_guardian_phone", "guardian_phone");
            checkAndCreateIndexMSSQL(connection, "student", "idx_student_class_id", "class_id");
            checkAndCreateIndexMSSQL(connection, "student", "idx_student_home_address", "home_address");
            checkAndCreateIndexMSSQL(connection, "student", "idx_student_city", "city");
            checkAndCreateIndexMSSQL(connection, "student", "idx_student_postal_address", "postal_address");
            checkAndCreateIndexMSSQL(connection, "student", "idx_student_student_reg_number", "student_reg_number");
            checkAndCreateIndexMSSQL(connection, "student", "idx_student_date_created", "date_created");
            checkAndCreateIndexMSSQL(connection, "student", "idx_student_date_modified", "date_modified");

            // Create options table
            checkAndCreateIndexMSSQL(connection, "options", "idx_options_date_created", "date_created");
            checkAndCreateIndexMSSQL(connection, "options", "idx_options_date_modified", "date_modified");

            // Create responses table
            checkAndCreateIndexMSSQL(connection, "responses", "idx_responses_date_created", "date_created");
            checkAndCreateIndexMSSQL(connection, "responses", "idx_responses_date_modified", "date_modified");

        }
        }
    }
    private static String getStringClass(boolean isMySQL, boolean isPostgreSQL, boolean isMSSQL) {

        String baseQuery = "";
       if(isMSSQL){
           baseQuery = "IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[class]') AND type in (N'U'))  BEGIN CREATE TABLE dbo.class (";
       }else{
        baseQuery = "CREATE TABLE IF NOT EXISTS class (";
       }
        String primaryKeyPart = "";
        String columnDefinitions = "class_name VARCHAR(50) NOT NULL UNIQUE, "
                + "date_created ";
        String dateTimePart = "";
        String endPart = ", PRIMARY KEY (class_id))";

        if(isMSSQL){
            endPart+=" END";
        }

        if (isMySQL) {
            primaryKeyPart = "class_id BIGINT NOT NULL AUTO_INCREMENT, ";
            dateTimePart = "DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP, "
                    + "date_modified DATETIME DEFAULT NULL";
        } else if (isPostgreSQL) {
            primaryKeyPart = "class_id BIGSERIAL, ";
            dateTimePart = "TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP, "
                    + "date_modified TIMESTAMP DEFAULT NULL";
        } else if (isMSSQL) {
            primaryKeyPart = "class_id BIGINT IDENTITY(1,1), ";
            dateTimePart = "DATETIME NOT NULL DEFAULT GETDATE(), "
                    + "date_modified DATETIME NULL";
        }
       // System.out.println(baseQuery + primaryKeyPart + columnDefinitions + dateTimePart + endPart);
        return baseQuery + primaryKeyPart + columnDefinitions + dateTimePart + endPart;
    }
    private static String getStringTeacher(boolean isMySQL, boolean isPostgreSQL, boolean isMSSQL) {
        String baseQuery = "";
        if(isMSSQL){
            baseQuery = "IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[teacher]') AND type in (N'U'))  BEGIN CREATE TABLE dbo.teacher (";
        }else{
            baseQuery = "CREATE TABLE IF NOT EXISTS teacher (";
        }


        baseQuery += "tsc_number VARCHAR(25) NOT NULL UNIQUE, "
                + "firstname VARCHAR(50) NOT NULL, "
                + "lastname VARCHAR(50) NOT NULL, "
                + "username VARCHAR(50) NOT NULL UNIQUE, "
                + "id_number INT NOT NULL UNIQUE, "
                + "phone VARCHAR(25) NOT NULL UNIQUE, "
                + "email VARCHAR(50) NOT NULL UNIQUE, "
                + "title VARCHAR(25), "
                + "address VARCHAR(50), "
                + "city VARCHAR(50), "
                + "postal_code INT, "
                + "national_insurance_number INT NOT NULL UNIQUE, "
                + "date_of_birth DATE, "
                + "password VARCHAR(255) NOT NULL, ";

        String primaryKeyPart = "";
        String dateTimePart = "";
        String endPart = ")";
        if(isMSSQL){
            endPart+=" END;";
        }else{
            endPart+=";";
        }
        if (isMySQL) {
            primaryKeyPart = "teacher_id BIGINT NOT NULL AUTO_INCREMENT, PRIMARY KEY (teacher_id), ";
            dateTimePart = "date_created DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP, "
                    + "date_modified DATETIME DEFAULT NULL";
        } else if (isPostgreSQL) {
            primaryKeyPart = "teacher_id BIGSERIAL PRIMARY KEY, ";
            dateTimePart = "date_created TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP, "
                    + "date_modified TIMESTAMP DEFAULT NULL";
        } else if (isMSSQL) {
            primaryKeyPart = "teacher_id BIGINT IDENTITY(1,1) PRIMARY KEY, ";
            dateTimePart = "date_created DATETIME NOT NULL DEFAULT GETDATE(), "
                    + "date_modified DATETIME NULL";
        }
        String fullQuery = baseQuery + primaryKeyPart + dateTimePart + endPart;
        return fullQuery;
    }
    private static String getStringSubject(boolean isMySQL, boolean isPostgreSQL, boolean isMSSQL) {

        String baseQuery = "";
        if(isMSSQL){
            baseQuery = "IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[subject]') AND type in (N'U'))  BEGIN CREATE TABLE dbo.subject (";
        }else{
            baseQuery = "CREATE TABLE IF NOT EXISTS subject (";
        }

        String primaryKeyPart = "";
        String columnDefinitions = "subject_name VARCHAR(50) NOT NULL UNIQUE, "
                + "date_created ";
        String dateTimePart = "";
        String endPart = ", PRIMARY KEY (subject_id))";
        if(isMSSQL){
            endPart+=" END;";
        }else{
            endPart+=";";
        }
        if (isMySQL) {
            primaryKeyPart = "subject_id BIGINT NOT NULL AUTO_INCREMENT, ";
            dateTimePart = "DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP, "
                    + "date_modified DATETIME DEFAULT NULL";
        } else if (isPostgreSQL) {
            primaryKeyPart = "subject_id BIGSERIAL, ";
            dateTimePart = "TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP, "
                    + "date_modified TIMESTAMP DEFAULT NULL";
        } else if (isMSSQL) {
            primaryKeyPart = "subject_id BIGINT IDENTITY(1,1), ";
            dateTimePart = "DATETIME NOT NULL DEFAULT GETDATE(), "
                    + "date_modified DATETIME NULL";
        }

        return baseQuery + primaryKeyPart + columnDefinitions + dateTimePart + endPart;
    }



    private static String getStringExam(boolean isMySQL, boolean isPostgreSQL, boolean isMSSQL) {

        String baseQuery = "";
        if(isMSSQL){
            baseQuery = "IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[exam]') AND type in (N'U'))  BEGIN CREATE TABLE dbo.exam (";
        }else{
            baseQuery = "CREATE TABLE IF NOT EXISTS exam (";
        }
        String primaryKeyPart = "";
        String columnDefinitions = "exam_name VARCHAR(50) NOT NULL UNIQUE, "
                + "date_created ";
        String dateTimePart = "";
        String endPart = ", PRIMARY KEY (exam_id))";
        if(isMSSQL){
            endPart+=" END;";
        }else{
            endPart+=";";
        }
        if (isMySQL) {
            primaryKeyPart = "exam_id BIGINT NOT NULL AUTO_INCREMENT, ";
            dateTimePart = "DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP, "
                    + "date_modified DATETIME DEFAULT NULL";
        } else if (isPostgreSQL) {
            primaryKeyPart = "exam_id BIGSERIAL, ";
            dateTimePart = "TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP, "
                    + "date_modified TIMESTAMP DEFAULT NULL";
        } else if (isMSSQL) {
            primaryKeyPart = "exam_id BIGINT IDENTITY(1,1), ";
            dateTimePart = "DATETIME NOT NULL DEFAULT GETDATE(), "
                    + "date_modified DATETIME NULL";
        }

        return baseQuery + primaryKeyPart + columnDefinitions + dateTimePart + endPart;
    }
    private static String getStringExamSchedule(boolean isMySQL, boolean isPostgreSQL, boolean isMSSQL) {
        String baseQuery = "";
        if(isMSSQL){
            baseQuery = "IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[exam_schedule]') AND type in (N'U'))  BEGIN CREATE TABLE dbo.exam_schedule (";
        }else{
            baseQuery = "CREATE TABLE IF NOT EXISTS exam_schedule (";
        }

        String primaryKeyPart = "";
        String columnDefinitions = "";
        String dateTimePart = "";
        String foreignKeyPart = " ,FOREIGN KEY (exam_id) REFERENCES exam(exam_id), "
                + "FOREIGN KEY (subject_id) REFERENCES subject(subject_id), "
                + "FOREIGN KEY (teacher_id) REFERENCES teacher(teacher_id), "
                + "FOREIGN KEY (class_id) REFERENCES class(class_id), "
                + "UNIQUE (exam_id, subject_id, class_id, month_year)";
        String endPart = ", PRIMARY KEY (exam_schedule_id))";
        if(isMSSQL){
            endPart+=" END;";
        }else{
            endPart+=";";
        }
        if (isMySQL) {
            primaryKeyPart = "exam_schedule_id BIGINT NOT NULL AUTO_INCREMENT, ";
            columnDefinitions = "exam_id BIGINT NOT NULL, "
                    + "subject_id BIGINT NOT NULL, "
                    + "exam_date DATETIME, "
                    + "exam_duration INT, "
                    + "teacher_id BIGINT NOT NULL, "
                    + "class_id BIGINT NOT NULL, "
                    + "month_year VARCHAR(7), " // Format: YYYY-MM
                    + "date_created DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP, "
                    + "date_modified DATETIME DEFAULT NULL";
        } else if (isPostgreSQL) {
            primaryKeyPart = "exam_schedule_id BIGSERIAL PRIMARY KEY, "; // Use BIGSERIAL as primary key
            columnDefinitions = "exam_id BIGINT NOT NULL, "
                    + "subject_id BIGINT NOT NULL, "
                    + "exam_date TIMESTAMP, " // Use TIMESTAMP for date and time
                    + "exam_duration INT, "
                    + "teacher_id BIGINT NOT NULL, "
                    + "class_id BIGINT NOT NULL, "
                    + "month_year VARCHAR(7), " // Format: YYYY-MM
                    + "date_created TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP, "
                    + "date_modified TIMESTAMP DEFAULT NULL";
        } else if (isMSSQL) {
            primaryKeyPart = "exam_schedule_id BIGINT IDENTITY(1,1) PRIMARY KEY, "; // Use IDENTITY as primary key
            columnDefinitions = "exam_id BIGINT NOT NULL, "
                    + "subject_id BIGINT NOT NULL, "
                    + "exam_date DATETIME, "
                    + "exam_duration INT, "
                    + "teacher_id BIGINT NOT NULL, "
                    + "class_id BIGINT NOT NULL, "
                    + "month_year VARCHAR(7), " // Format: YYYY-MM
                    + "date_created DATETIME NOT NULL DEFAULT GETDATE(), "
                    + "date_modified DATETIME NULL";
        }

        return baseQuery + primaryKeyPart + columnDefinitions + dateTimePart + foreignKeyPart + endPart;
    }


    private static String getStringExamQuestions(boolean isMySQL, boolean isPostgreSQL, boolean isMSSQL) {
        String baseQuery = "";
        if(isMSSQL){
            baseQuery = "IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[questions]') AND type in (N'U'))  BEGIN CREATE TABLE dbo.questions (";
        }else{
            baseQuery = "CREATE TABLE IF NOT EXISTS questions (";
        }

        String primaryKeyPart = "";
        String columnDefinitions = "exam_schedule_id BIGINT NOT NULL, "
                + "text TEXT NOT NULL, "
                + "marks INT NOT NULL, "
                + "date_created ";
        String dateTimePart = "";
        String foreignKeyPart = ", FOREIGN KEY (exam_schedule_id) REFERENCES exam_schedule(exam_schedule_id), PRIMARY KEY (question_id)";
        String endPart = ")";
        if(isMSSQL){
            endPart+=" END;";
        }else{
            endPart+=";";
        }
        if (isMySQL) {
            primaryKeyPart = "question_id BIGINT NOT NULL AUTO_INCREMENT, ";
            dateTimePart = "DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP, "
                    + "date_modified DATETIME DEFAULT NULL";
        } else if (isPostgreSQL) {
            primaryKeyPart = "question_id BIGSERIAL, ";
            dateTimePart = "TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP, "
                    + "date_modified TIMESTAMP DEFAULT NULL";
        } else if (isMSSQL) {
            primaryKeyPart = "question_id BIGINT IDENTITY(1,1), ";
            dateTimePart = "DATETIME NOT NULL DEFAULT GETDATE(), "
                    + "date_modified DATETIME NULL";
        }

        return baseQuery + primaryKeyPart + columnDefinitions + dateTimePart + foreignKeyPart + endPart;
    }
    private static String getStringExamQuestionOptions(boolean isMySQL, boolean isPostgreSQL, boolean isMSSQL) {
        String baseQuery = "";
        if(isMSSQL){
            baseQuery = "IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[options]') AND type in (N'U'))  BEGIN CREATE TABLE dbo.options (";
        }else{
            baseQuery = "CREATE TABLE IF NOT EXISTS options (";
        }

        String primaryKeyPart = "";
        String columnDefinitions = "question_id BIGINT NOT NULL, "
                + "option_label VARCHAR(2) NOT NULL, "
                + "option_value TEXT NOT NULL, "
                + "correct INT NOT NULL DEFAULT 0, "
                + "date_created ";
        String dateTimePart = "";
        String foreignKeyPart = ", FOREIGN KEY (question_id) REFERENCES questions(question_id), UNIQUE (question_id, option_label), PRIMARY KEY (option_id)";
        String endPart = ")";
        if(isMSSQL){
            endPart+=" END;";
        }else{
            endPart+=";";
        }
        if (isMySQL) {
            primaryKeyPart = "option_id BIGINT NOT NULL AUTO_INCREMENT, ";
            dateTimePart = "DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP, "
                    + "date_modified DATETIME DEFAULT NULL";
        } else if (isPostgreSQL) {
            primaryKeyPart = "option_id BIGSERIAL, ";
            dateTimePart = "TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP, "
                    + "date_modified TIMESTAMP DEFAULT NULL";
        } else if (isMSSQL) {
            primaryKeyPart = "option_id BIGINT IDENTITY(1,1), ";
            dateTimePart = "DATETIME NOT NULL DEFAULT GETDATE(), "
                    + "date_modified DATETIME NULL";
        }

        return baseQuery + primaryKeyPart + columnDefinitions + dateTimePart + foreignKeyPart + endPart;
    }

    private static String getStringExamQuestionResponses(boolean isMySQL, boolean isPostgreSQL, boolean isMSSQL) {
        String baseQuery = "";
        if(isMSSQL){
            baseQuery = "IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[responses]') AND type in (N'U'))  BEGIN CREATE TABLE dbo.responses (";
        }else{
            baseQuery = "CREATE TABLE IF NOT EXISTS responses (";;
        }

        String primaryKeyPart = "";
        String columnDefinitions = "";
        String dateTimePart = "";
        String foreignKeyPart = ", FOREIGN KEY (question_id) REFERENCES questions(question_id), "
                + "FOREIGN KEY (student_id) REFERENCES student(student_id), "
                + "FOREIGN KEY (option_id) REFERENCES options(option_id)";
        String endPart = ", PRIMARY KEY (response_id))";
        if(isMSSQL){
            endPart+=" END;";
        }else{
            endPart+=";";
        }
        if (isMySQL) {
            primaryKeyPart = "response_id BIGINT NOT NULL AUTO_INCREMENT, ";
            columnDefinitions = "question_id BIGINT NOT NULL, "
                    + "student_id BIGINT NOT NULL, "
                    + "option_id BIGINT NOT NULL, "
                    + "date_created DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP, "
                    + "date_modified DATETIME DEFAULT NULL";
        } else if (isPostgreSQL) {
            primaryKeyPart = "response_id BIGSERIAL, ";
            columnDefinitions = "question_id BIGINT NOT NULL, "
                    + "student_id BIGINT NOT NULL, "
                    + "option_id BIGINT NOT NULL, "
                    + "date_created TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP, "
                    + "date_modified TIMESTAMP DEFAULT NULL";
        } else if (isMSSQL) {
            primaryKeyPart = "response_id BIGINT IDENTITY(1,1), ";
            columnDefinitions = "question_id BIGINT NOT NULL, "
                    + "student_id BIGINT NOT NULL, "
                    + "option_id BIGINT NOT NULL, "
                    + "date_created DATETIME NOT NULL DEFAULT GETDATE(), "
                    + "date_modified DATETIME NULL";
        }

        return baseQuery + primaryKeyPart + columnDefinitions + foreignKeyPart + endPart;
    }

    private static String getStringStudents(boolean isMySQL, boolean isPostgreSQL, boolean isMSSQL) {
        String baseQuery = "";
        if(isMSSQL){
            baseQuery = "IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[student]') AND type in (N'U'))  BEGIN CREATE TABLE dbo.student (";
        }else{
            baseQuery = "CREATE TABLE IF NOT EXISTS student (";
        }

        String primaryKeyPart = "";
        String columnDefinitions = "firstname VARCHAR(50) NOT NULL, "
                + "lastname VARCHAR(50) NOT NULL, "
                + "username VARCHAR(50) NOT NULL UNIQUE, "
                + "gender VARCHAR(25) NOT NULL, "
                + "date_of_birth DATE, "
                + "guardian_name VARCHAR(50) NOT NULL, "
                + "password VARCHAR(255) NOT NULL, "
                + "guardian_phone VARCHAR(25) NOT NULL, "
                + "class_id BIGINT NOT NULL, "
                + "home_address VARCHAR(50), "
                + "city VARCHAR(50), "
                + "postal_address INT, "
                + "student_reg_number VARCHAR(50) NOT NULL UNIQUE, "
                + "date_created ";
        String dateTimePart = "";
        String foreignKeyPart = ", FOREIGN KEY (class_id) REFERENCES class(class_id), PRIMARY KEY (student_id)";
        String endPart = ")";
        if(isMSSQL){
            endPart+=" END;";
        }else{
            endPart+=";";
        }
        if (isMySQL) {
            primaryKeyPart = "student_id BIGINT NOT NULL AUTO_INCREMENT, ";
            dateTimePart = "DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP, "
                    + "date_modified DATETIME DEFAULT NULL";
        } else if (isPostgreSQL) {
            primaryKeyPart = "student_id BIGSERIAL, ";
            dateTimePart = "TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP, "
                    + "date_modified TIMESTAMP DEFAULT NULL";
        } else if (isMSSQL) {
            primaryKeyPart = "student_id BIGINT IDENTITY(1,1), ";
            dateTimePart = "DATETIME NOT NULL DEFAULT GETDATE(), "
                    + "date_modified DATETIME NULL";
        }

        return baseQuery + primaryKeyPart + columnDefinitions + dateTimePart + foreignKeyPart + endPart;
    }

    private static void checkAndCreateIndex(Connection connection, String tableName, String indexName, String columnName) throws SQLException {
            // Check if the index exists
            ResultSet resultSet = connection.getMetaData().getIndexInfo(null, null, tableName, false, false);
            boolean indexExists = false;
            while (resultSet.next()) {
                String existingIndexName = resultSet.getString("INDEX_NAME");
                if (indexName.equalsIgnoreCase(existingIndexName)) {
                    indexExists = true;
                    break;
                }
            }
            resultSet.close();

            // If the index doesn't exist, create it
            if (!indexExists) {
                try (Statement statement = connection.createStatement()) {
                    statement.executeUpdate("CREATE INDEX " + indexName + " ON " + tableName + "(" + columnName + ")");
                    System.out.println("Index '" + indexName + "' created for table '" + tableName + "'.");
                }
            }
        }


        private static void checkAndCreateIndexPostgres(Connection connection, String tableName, String indexName, String columnName) throws SQLException {
        // Check if the index exists in PostgreSQL
        String checkQuery = "SELECT * FROM pg_index JOIN pg_class ON pg_class.oid=pg_index.indexrelid "
                + "JOIN pg_namespace ON pg_namespace.oid=pg_class.relnamespace "
                + "WHERE pg_class.relname = ? AND pg_namespace.nspname = 'public';";

        boolean indexExists = false;
        try (PreparedStatement pstmt = connection.prepareStatement(checkQuery)) {
            pstmt.setString(1, indexName);
            ResultSet resultSet = pstmt.executeQuery();
            indexExists = resultSet.next();
            resultSet.close();
        }

        // If the index doesn't exist, create it
        if (!indexExists) {
            try (Statement statement = connection.createStatement()) {
                statement.executeUpdate("CREATE INDEX " + indexName + " ON public." + tableName + "(" + columnName + ")");
                System.out.println("Index '" + indexName + "' created for table '" + tableName + "'.");
            }
        }
    }

    private static void checkAndCreateIndexMSSQL(Connection connection, String tableName, String indexName, String columnName) throws SQLException {
        // Check if the index exists in MSSQL
        String checkQuery = "SELECT * FROM sys.indexes WHERE name = ? AND object_id = OBJECT_ID(?);";

        boolean indexExists = false;
        try (PreparedStatement pstmt = connection.prepareStatement(checkQuery)) {
            pstmt.setString(1, indexName);
            pstmt.setString(2, tableName);
            ResultSet resultSet = pstmt.executeQuery();
            indexExists = resultSet.next();
            resultSet.close();
        }

        // If the index doesn't exist, create it
        if (!indexExists) {
            try (Statement statement = connection.createStatement()) {
                statement.executeUpdate("CREATE INDEX " + indexName + " ON " + tableName + "(" + columnName + ")");
                System.out.println("Index '" + indexName + "' created for table '" + tableName + "'.");
            }
        }
    }

}
