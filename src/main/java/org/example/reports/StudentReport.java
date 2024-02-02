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
import java.util.*;

import static org.example.libs.ConfigFileChecker.configFileChecker;

public class StudentReport {
    public static String fileName;
   public static String type ;

    public StudentReport(Connection conn, String database_type, long ExamId) {
        Response checker = configFileChecker("config", fileName);
        type = database_type;
        try (Connection connection = conn) {
            long examId =  ExamId; // Replace with the desired exam ID
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
                    generateExcelSheets(classStudentMap,connection,examId);
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static Map<Long, Integer> getSubjectMarks(long examId) throws Exception {
        String sql = generateOverallSubjectMarks(examId);
        Response checker = configFileChecker("config",fileName);
        DatabaseConnection connection;
        if(checker.getStatus()){
            connection = new DatabaseConnection("config/"+checker.getMessage());

            Map<Long, Integer> subjectMarks = new HashMap<>();
            try (Connection conn = connection.getConnection();
                 PreparedStatement stmt = conn.prepareStatement(sql)) {
                stmt.setLong(1, examId);
                ResultSet rs = stmt.executeQuery();
                while (rs.next()) {
                    long subjectId = rs.getLong("subject_id");
                    int marks = rs.getInt("Marks");
                    subjectMarks.put(subjectId, marks);
                }
            } catch (SQLException e) {
                e.printStackTrace();
            }
            System.out.println(subjectMarks);
            return subjectMarks;
        }else{
            System.out.println(checker.getMessage());
        }
        return null;
    }


    private static void generateExcelSheets(Map<Integer, Map<Integer, List<StudentMerit>>> classStudentMap, Connection conn, long ExamId) throws IOException {
        Workbook workbook = new XSSFWorkbook();
        Map<Integer, String> subjectMap = fetchSubjectsFromDatabase(conn); // Assumed to be implemented

        for (Map.Entry<Integer, Map<Integer, List<StudentMerit>>> classEntry : classStudentMap.entrySet()) {
            Sheet sheet = workbook.createSheet("Class " + classEntry.getKey() + " Merit");
            createHeaderRow(sheet, subjectMap);

            // Aggregating total marks per student and preparing data for subject marks
            Map<Integer, Integer> totalMarksPerStudent = new HashMap<>();
            Map<Integer, Map<Integer, Integer>> subjectMarksPerStudent = new HashMap<>(); // New structure to hold subject marks per student

            classEntry.getValue().forEach((studentId, studentMerits) -> {
                int totalMarks = 0; //studentMerits.stream().mapToInt(StudentMerit::getMarks).sum();


                // Prepare subject marks map for each student
                Map<Integer, Integer> marks = new HashMap<>();
                for (StudentMerit merit : studentMerits) {
                    Map<Long, Integer> subjectMarks = null;
                    try {
                        subjectMarks = getSubjectMarks(ExamId);
                    } catch (Exception e) {
                        throw new RuntimeException(e);
                    }

                    assert subjectMarks != null;
                    System.out.println(merit.getMarks());
                    double total =  (((double) merit.getMarks() /subjectMarks.get((long) merit.getSubjectId())))*100;
                    int  marksz = (int) total;
                    totalMarks+=marksz;
                    marks.put(merit.getSubjectId(), marksz);
                }
                totalMarksPerStudent.put(studentId, totalMarks);
                subjectMarksPerStudent.put(studentId, marks);
            });

            // Sorting students by total marks in descending order
            List<Map.Entry<Integer, Integer>> sortedEntries = new ArrayList<>(totalMarksPerStudent.entrySet());
            sortedEntries.sort(Map.Entry.<Integer, Integer>comparingByValue().reversed());

            // Creating rows for each student based on sorted total marks
            int rowNum = 1;
            for (Map.Entry<Integer, Integer> entry : sortedEntries) {
                Integer studentId = entry.getKey();
                Row row = sheet.createRow(rowNum++);

                row.createCell(0).setCellValue(rowNum-1);
                row.createCell(1).setCellValue(classEntry.getValue().get(studentId).get(0).getName());
                int cellIndex = 2;

                // Fill in subject marks
                Map<Integer, Integer> studentMarks = subjectMarksPerStudent.get(studentId);
                for (Integer subjectId : subjectMap.keySet()) {
                    if (studentMarks.containsKey(subjectId)) {
                        row.createCell(cellIndex).setCellValue(studentMarks.get(subjectId));
                    } else {
                        row.createCell(cellIndex).setCellValue(0); // or leave blank if preferred
                    }
                    cellIndex++;
                }

                row.createCell(cellIndex).setCellValue(entry.getValue()); // Total marks
                row.createCell(cellIndex + 1).setCellValue(rowNum - 1); // Position
            }
        }

        // Write the workbook to a file
        try (FileOutputStream outputStream = new FileOutputStream("StudentMeritReport.xlsx")) {
            workbook.write(outputStream);
        }
    }

    private static void createHeaderRow(Sheet sheet, Map<Integer, String> subjectMap) {
        Row headerRow = sheet.createRow(0);
        headerRow.createCell(0).setCellValue("NO:");
        headerRow.createCell(1).setCellValue("Student Name");
        int columnIndex = 2;
        for (String subject : subjectMap.values()) {
            headerRow.createCell(columnIndex++).setCellValue(subject);
        }
        headerRow.createCell(columnIndex).setCellValue("Total Marks");
        headerRow.createCell(columnIndex + 1).setCellValue("Position");
    }

    private static String generateDynamicSubjectMarksQuery(long examId) {
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
        return queryBuilder.toString();
    }

    public static String generateOverallSubjectMarks(long examId){
        StringBuilder queryBuilder = new StringBuilder();
        queryBuilder.append("SELECT ");
        queryBuilder.append("    sub.subject_id, ");
        queryBuilder.append("    SUM(q.marks) AS Marks ");
        queryBuilder.append("FROM ");
        queryBuilder.append("    questions q ");
        queryBuilder.append("JOIN ");
        queryBuilder.append("    exam_schedule es ON q.exam_schedule_id = es.exam_schedule_id ");
        queryBuilder.append("JOIN ");
        queryBuilder.append("    subject sub ON es.subject_id = sub.subject_id ");
        queryBuilder.append("JOIN ");
        queryBuilder.append("    class c ON es.class_id = c.class_id ");
        queryBuilder.append("WHERE ");
        queryBuilder.append("    es.exam_id = ? ");
        queryBuilder.append("GROUP BY ");
        queryBuilder.append("    sub.subject_id, c.class_id, es.exam_id, c.class_name, sub.subject_name ");
        queryBuilder.append("ORDER BY ");
        queryBuilder.append(" sub.subject_id; ");
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
