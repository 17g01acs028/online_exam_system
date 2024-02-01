package org.example.controllers;

import org.example.QueryBuilder.Insert;
import org.example.QueryBuilder.Select;
import org.example.libs.Response;
import org.example.QueryBuilder.Update;

import java.sql.Connection;
import java.sql.SQLException;
import java.util.Map;

public class Subject {
    public static Response addSubject(Connection conn, Map<String, String> values){
        Response result = null;
        try {
            result = Insert.insertSingleRow(conn,"subject", values);
        } catch (SQLException e) {
            return new Response(500,e.getMessage());
        }
        return result;
    }

    public static Response updateSubject(Connection conn, Map<String, String> values,String id){
        Response result = null;
        try {
            result = Update.updateSingleRow(conn,"subject",values,id);
        } catch (SQLException e) {
            return new Response(500,e.getMessage());
        }
        return result;
    }
    public static Response Find(Connection conn){
        Response result = null;
        try {
            result = Select.select(conn, "subject");
        } catch (SQLException e) {
            throw new RuntimeException(e);
        }
        return result;
    }

    public static Response FindById(Connection conn,int id){
        Response result = null;
        try {
            result = Select.select(conn, "subject","subject_id ="+id);
        } catch (SQLException e) {
            throw new RuntimeException(e);
        }
        return result;
    }
}
