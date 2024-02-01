package org.example.QueryBuilder;

import org.example.libs.Response;

import java.sql.*;
import java.util.List;
import java.util.Map;
import java.util.StringJoiner;

import static org.example.QueryBuilder.QueryBuilder.*;

public class Insert {
    public static Response insertSingleRow(Connection conn, String table, Map<String, String> values) throws SQLException {
        StringJoiner columnNames = new StringJoiner(", ");
        StringJoiner valuePlaceholders = new StringJoiner(", ");
        for (String column : values.keySet()) {
            columnNames.add(column);
            valuePlaceholders.add("?");
        }
        String query = "INSERT INTO " + table + " (" + columnNames + ") VALUES (" + valuePlaceholders + ")";
        return insertData(conn, query, values);
    }



    public static Response insertBatchRows(Connection conn, String table, List<Map<String, String>> valuesList) throws SQLException {
        if (valuesList.isEmpty()) {
            return new Response(false,"Failed Please try again");
        }

        StringJoiner columnNames = new StringJoiner(", ");
        StringJoiner valuePlaceholders = new StringJoiner(", ");
        for (String column : valuesList.get(0).keySet()) {
            columnNames.add(column);
            valuePlaceholders.add("?");
        }

        String query = "INSERT INTO " + table + " (" + columnNames + ") VALUES (" + valuePlaceholders + ")";
        return insertDataBatch(conn,query,valuesList);

    }

    public static Response insertIfNotExists(Connection conn, String table, Map<String, String> values, String condition) throws SQLException {
        Response existingData = querySelect(conn, "SELECT * FROM " + table + " WHERE " + condition);
        if (existingData.getResponse().isEmpty()) {
            return insertSingleRow(conn, table, values);
        }
        return existingData;
    }

}

