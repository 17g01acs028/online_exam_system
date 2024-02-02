package org.example;

import org.example.controllers.Exam;
import org.example.libs.Response;
import org.example.model.DatabaseConnection;
import org.example.reports.ExamReport;
import org.example.reports.StudentReport;

import java.sql.Connection;

import static org.example.libs.ConfigFileChecker.configFileChecker;


public class Main{
    public static  String fileName;
    public static void main(String[] args){
        try {
            Response checker = configFileChecker("config",fileName);
            if(checker.getStatus()){
                DatabaseConnection connection = new DatabaseConnection("config/"+checker.getMessage());
                Connection conn = connection.getConnection();
                //System.out.println(Exam.Find(conn));
              //  System.out.println(Exam.FindById(conn,1));
              //  System.out.println(ExamReport.teacherExam(conn,21));

                new StudentReport(conn,connection.databaseType,1);

            }else{
                System.out.println(checker.getMessage());
            }
        } catch (Exception e) {
            System.out.println(e.getMessage());
        }
    }


}

