package org.example.reports;

import org.example.QueryBuilder.QueryBuilder;
import org.example.QueryBuilder.QueryBuilderNew;
import org.example.libs.Response;
import java.sql.Connection;
import java.sql.SQLException;

public class ExamReport {
//    public  static Response teacherExam(Connection conn, int id){
//        return QueryBuilder.querySelect(conn,"SELECT e.exam_id, t.username, e.exam_name\n" +
//                "FROM exam e\n" +
//                "INNER JOIN teacher t ON e.teacher_id = t.teacher_id\n" +
//                "WHERE t.teacher_id = "+id);
//    }
//    public static Response studentScore(Connection conn, int exam_id, int student_id){
//        String query = "SELECT st.student_id, st.firstname, st.lastname, ex.exam_id, ex.exam_name, subj.subject_name, q.question_id, q.text AS question_text, opt.option_label, opt.option_value AS student_answer, opt.correct, (SUM(CASE WHEN resp.student_id = "+student_id+" THEN opt.correct * q.marks ELSE 0 END) / SUM(q.marks)) * 100 AS percentage_score FROM student st JOIN responses resp ON st.student_id = resp.student_id JOIN options opt ON resp.option_id = opt.option_id JOIN questions q ON resp.question_id = q.question_id JOIN exam ex ON q.exam_id = ex.exam_id JOIN subject subj ON ex.subject_id = subj.subject_id WHERE st.student_id = "+student_id+" AND ex.exam_id = "+exam_id+" GROUP BY st.student_id, st.firstname, st.lastname, ex.exam_id, ex.exam_name, subj.subject_name, q.question_id, q.text, opt.option_label, opt.option_value, opt.correct";
//        return  QueryBuilder.querySelect(conn,query);
//    }
public static Response teacherExam(Connection conn, int id) {
    String query = "SELECT DISTINCT e.exam_id, c.class_name, s.subject_name, t.username, e.exam_name " +
            "FROM exam e " +
            "JOIN exam_schedule es ON e.exam_id = es.exam_id " +
            "JOIN teacher t ON es.teacher_id = t.teacher_id " +
            "JOIN subject s ON es.subject_id = s.subject_id " +
            "JOIN class c ON es.class_id = c.class_id " +
            "WHERE t.teacher_id = " + id;
    return QueryBuilder.querySelect(conn, query);
}
    public static Response studentScore(Connection conn, int exam_id, int student_id) {
        String query = "SELECT st.student_id, st.firstname, st.lastname, ex.exam_id, ex.exam_name, subj.subject_name, q.question_id, q.text AS question_text, opt.option_label, opt.option_value AS student_answer, opt.correct, (SUM(CASE WHEN resp.student_id = "+student_id+" THEN opt.correct * q.marks ELSE 0 END) / SUM(q.marks)) * 100 AS percentage_score " +
                "FROM student st " +
                "JOIN responses resp ON st.student_id = resp.student_id " +
                "JOIN options opt ON resp.option_id = opt.option_id " +
                "JOIN questions q ON resp.question_id = q.question_id " +
                "JOIN exam_subject es ON q.exam_subject_id = es.exam_subject_id " +
                "JOIN exam ex ON es.exam_id = ex.exam_id " +
                "JOIN subject subj ON es.subject_id = subj.subject_id " +
                "WHERE st.student_id = "+student_id+" AND ex.exam_id = "+exam_id+" " +
                "GROUP BY st.student_id, st.firstname, st.lastname, ex.exam_id, ex.exam_name, subj.subject_name, q.question_id, q.text, opt.option_label, opt.option_value, opt.correct";
        return QueryBuilder.querySelect(conn, query);
    }

    public static Response studentMerit(Connection conn, int limit, int exam_id){
        QueryBuilderNew qb = new QueryBuilderNew(conn);
        qb.setParameter("exam_id",exam_id);
        String query = "SELECT s.firstname, s.lastname, e.exam_name, SUM(CASE WHEN o.correct = 1 THEN q.marks ELSE 0 END) AS total_marks,\n" +
                "       ROW_NUMBER() OVER (ORDER BY SUM(CASE WHEN o.correct = 1 THEN q.marks ELSE 0 END) DESC) AS ranking\n" +
                "FROM student s\n" +
                "JOIN responses r ON s.student_id = r.student_id\n" +
                "JOIN questions q ON r.question_id = q.question_id\n" +
                "JOIN options o ON r.option_id = o.option_id\n" +
                "JOIN exam e ON q.exam_id = e.exam_id\n" +
                "WHERE e.exam_id = :exam_id\n" +
                "GROUP BY s.student_id\n" +
                "ORDER BY total_marks DESC\n" +
                "LIMIT "+limit;
        try {
            return qb.executeSelectQuery(query);
        } catch (SQLException e) {
            return  new Response(500, "Error: "+e.getMessage());
        }
    }
}
