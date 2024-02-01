package org.example.reports;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;

import org.example.libs.Response;
import org.example.model.DatabaseConnection;

import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import static org.example.libs.ConfigFileChecker.configFileChecker;

public class StudentReport {
    public static String fileName;
   public static String type ;
    public StudentReport(Connection conn, String typez) {
        Response checker = configFileChecker("config", fileName);
        DatabaseConnection connection3 = new DatabaseConnection("config/" + checker.getMessage());
        type = typez;
        try (Connection connection = connection3.getConnection()) {
            long examId = 1; // Replace with the desired exam ID
            String query = generateDynamicSubjectMarksQuery(examId);

            try (PreparedStatement statement = connection.prepareStatement(query)) {
                statement.setLong(1, examId); // Bind the exam ID as a parameter
                Map<Integer, Map<Integer, List<StudentMerit>>> classStudentMap = new HashMap<>();

                try (ResultSet resultSet = statement.executeQuery()) {
                    while (resultSet.next()) {
                        int studentId = resultSet.getInt("student_id");
                        String name = resultSet.getString("name");
                        int subjectId = resultSet.getInt("subject_id");
                        String subject = resultSet.getString("Subject");
                        int classId = resultSet.getInt("class_id");
                        String className = resultSet.getString("Class");
                        int marks = resultSet.getInt("Marks");

                        // Check if the class ID exists in the outer HashMap
                        if (!classStudentMap.containsKey(classId)) {
                            classStudentMap.put(classId, new HashMap<>());
                        }

                        // Get the inner HashMap for this class ID
                        Map<Integer, List<StudentMerit>> studentMap = classStudentMap.get(classId);

                        // Check if the student ID exists in the inner HashMap
                        if (!studentMap.containsKey(studentId)) {
                            studentMap.put(studentId, new ArrayList<>());
                        }

                        // Get the list of records for this student ID
                        List<StudentMerit> studentRecords = studentMap.get(studentId);

                        // Create a StudentMerit object and add it to the list
                        StudentMerit studentMerit = new StudentMerit(studentId, name, subjectId, subject, classId, className, marks);
                        studentRecords.add(studentMerit);
                        System.out.println("Student " + studentId + " Name: " + name +
                                " Subject: " + subject + " Class Id " + classId +
                                " Class: " + className + " Marks: " + marks +
                                "Subject Id " + subjectId);
                    }
                    System.out.print(classStudentMap);
                    generateExcelSheets(classStudentMap,connection);
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static void generateExcelSheets(Map<Integer, Map<Integer, List<StudentMerit>>> classStudentMap, Connection conn) {
        Workbook workbook = new XSSFWorkbook();

        // Fetch all subjects from the database and store them in a data structure
        Map<Integer, String> subjectMap = fetchSubjectsFromDatabase(conn);

        // Iterate over classStudentMap
        for (Map.Entry<Integer, Map<Integer, List<StudentMerit>>> classEntry : classStudentMap.entrySet()) {
            int classId = classEntry.getKey();
            Map<Integer, List<StudentMerit>> studentMap = classEntry.getValue();

            Sheet sheet = workbook.createSheet("Class " + classId + " Merit");

            // Create headers for the sheet
            Row headerRow = sheet.createRow(0);
            headerRow.createCell(0).setCellValue("Student ID");
            headerRow.createCell(1).setCellValue("Student Name");

            // Add subject columns
            int columnIndex = 2;
            for (Map.Entry<Integer, String> subjectEntry : subjectMap.entrySet()) {
                headerRow.createCell(columnIndex).setCellValue( subjectEntry.getValue());
                columnIndex++;
            }

            headerRow.createCell(columnIndex).setCellValue("Total");
            headerRow.createCell(columnIndex + 1).setCellValue("Position");

            int rowNum = 1; // Start from row 1 (after headers)



            // Iterate over students in this class
            int no =1;
            for (Map.Entry<Integer, List<StudentMerit>> studentEntry : studentMap.entrySet()) {
                List<StudentMerit> studentRecords = studentEntry.getValue();

                // Create a row for the student
                Row dataRow = sheet.createRow(rowNum++);
                dataRow.createCell(0).setCellValue(no);

                // Set the student's name
                dataRow.createCell(1).setCellValue(studentRecords.get(0).getName()); // Set the student's name

                // Calculate total marks for the student
                int totalMarks = 0;

                // Create a map to store marks for each subject
                Map<Integer, Integer> subjectMarksMap = new HashMap<>();

                // Iterate over student records (for different subjects)
                for (StudentMerit studentRecord : studentRecords) {

                    int subjectId = studentRecord.getSubjectId();
                    int marks = studentRecord.getMarks();
                    System.out.println("subject_id "+subjectId + " Marks "+marks);
                    subjectMarksMap.put(subjectId, marks);
                    totalMarks += marks;
                }

                // Add subject marks to the corresponding columns
                columnIndex = 2;
                for (Integer subjectId : subjectMap.keySet()) {
                    if (subjectMarksMap.containsKey(subjectId)) {
                        dataRow.createCell(columnIndex).setCellValue(subjectMarksMap.get(subjectId));
                    } else {
                        dataRow.createCell(columnIndex).setCellValue("");
                    }
                    columnIndex++;
                }

                // Add total marks and position
                dataRow.createCell(columnIndex).setCellValue(totalMarks);

                // Increment the position for each student
                int position = rowNum - 1;
                dataRow.createCell(columnIndex + 1).setCellValue(position);
                no++;
            }
        }

        try (FileOutputStream outputStream = new FileOutputStream("MeritReport.xlsx")) {
            workbook.write(outputStream);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static String generateDynamicSubjectMarksQuery(long examId) {

        Response checker = configFileChecker("config",fileName);

        StringBuilder queryBuilder = new StringBuilder();
        queryBuilder.append("SELECT ");
        queryBuilder.append("    s.student_id, ");
        queryBuilder.append("    CONCAT(s.firstname, ' ', s.lastname) AS name, ");
        queryBuilder.append("    es.exam_id, ");
        queryBuilder.append("    sub.subject_id, ");
        queryBuilder.append("    sub.subject_name AS Subject, ");
        queryBuilder.append("    c.class_id, ");
        queryBuilder.append("    c.class_name AS Class, ");
        queryBuilder.append("    SUM(q.marks) AS Marks ");
        queryBuilder.append("FROM ");
        queryBuilder.append("    student s ");
        queryBuilder.append("JOIN ");
        queryBuilder.append("    responses r ON s.student_id = r.student_id ");
        queryBuilder.append("JOIN ");
        queryBuilder.append("    questions q ON r.question_id = q.question_id ");
        queryBuilder.append("JOIN ");
        queryBuilder.append("    options o ON r.option_id = o.option_id ");
        queryBuilder.append("JOIN ");
        queryBuilder.append("    exam_schedule es ON q.exam_schedule_id = es.exam_schedule_id ");
        queryBuilder.append("JOIN ");
        queryBuilder.append("    subject sub ON es.subject_id = sub.subject_id ");
        queryBuilder.append("JOIN ");
        queryBuilder.append("    class c ON s.class_id = c.class_id ");
        queryBuilder.append("WHERE ");
        queryBuilder.append("    es.exam_id = ? and o.correct = 1 ");
        queryBuilder.append("GROUP BY ");
        if("mysql".equalsIgnoreCase(type)) {
            queryBuilder.append("    s.student_id, sub.subject_id, c.class_id ");
        } else if ("mssql".equalsIgnoreCase(type)) {
            queryBuilder.append("     s.student_id, s.firstname, s.lastname, sub.subject_id, c.class_id, es.exam_id, c.class_name,sub.subject_name ");
        } else if ("postgresql".equalsIgnoreCase(type)) {
            queryBuilder.append("    s.student_id, sub.subject_id, c.class_id, es.exam_id ");
        }
        queryBuilder.append("ORDER BY ");
        queryBuilder.append("    s.student_id, sub.subject_id; ");

        System.out.println(queryBuilder.toString());
        return queryBuilder.toString();
    }

    // Fetch subjects from the database and return them as a map (subjectId -> subjectName)
    private static Map<Integer, String> fetchSubjectsFromDatabase(Connection connection) {
        Map<Integer, String> subjectMap = new HashMap<>();
        String query = "SELECT subject_id, subject_name FROM subject";

        try (PreparedStatement statement = connection.prepareStatement(query)) {
            try (ResultSet resultSet = statement.executeQuery()) {
                while (resultSet.next()) {
                    int subjectId = resultSet.getInt("subject_id");
                    String subjectName = resultSet.getString("subject_name");
                    subjectMap.put(subjectId, subjectName);
                }
            }
        } catch (SQLException e) {
            e.printStackTrace();
        }

        return subjectMap;
    }
}
