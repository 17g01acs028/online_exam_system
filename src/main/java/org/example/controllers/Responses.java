package org.example.controllers;

import org.example.QueryBuilder.Insert;
import org.example.QueryBuilder.Select;
import org.example.libs.Response;
import org.example.QueryBuilder.Update;

import java.sql.Connection;
import java.sql.SQLException;
import java.util.Map;

public class Responses {
    public static Response addResponses(Connection conn, Map<String, String> values){
        Response result = null;
        try {
            result = Insert.insertSingleRow(conn,"responses", values);
        } catch (SQLException e) {
            return new Response(500,e.getMessage());
        }
        return result;
    }

    public static Response updateResponses(Connection conn, Map<String, String> values,String id){
        Response result = null;
        try {
            result = Update.updateSingleRow(conn,"responses",values,id);
        } catch (SQLException e) {
            return new Response(500,e.getMessage());
        }
        return result;
    }
    public static Response Find(Connection conn){
        Response result = null;
        try {
            result = Select.select(conn, "responses");
        } catch (SQLException e) {
            throw new RuntimeException(e);
        }
        return result;
    }

    public static Response FindById(Connection conn,int id){
        Response result = null;
        try {
            result = Select.select(conn, "responses","response_id ="+id);
        } catch (SQLException e) {
            throw new RuntimeException(e);
        }
        return result;
    }
}
