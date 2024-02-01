package org.example.QueryBuilder;

import org.example.libs.Response;
import org.json.JSONArray;
import org.json.JSONObject;

import java.sql.*;
import java.util.HashMap;
import java.util.Map;

public class QueryBuilderNew {
    private Connection connection;
    private Map<String, Object> parameters;

    public QueryBuilderNew(Connection connection) {
        this.connection = connection;
        this.parameters = new HashMap<>();
    }

    public void setParameter(String name, Object value) {
        parameters.put(name, value);
    }

    public Response executeSelectQuery(String query) throws SQLException {
        JSONArray jsonArray = new JSONArray();

        try {
            // Replace named parameters with ? placeholders
            for (Map.Entry<String, Object> entry : parameters.entrySet()) {
                query = query.replace(":" + entry.getKey(), "?");
            }

            // Preparing the statement
            PreparedStatement stmt = connection.prepareStatement(query);

            // Binding parameters using ? placeholders
            int paramIndex = 1;
            for (Object value : parameters.values()) {
                stmt.setObject(paramIndex, value);
                paramIndex++;
            }

            // Executing the query
            try (ResultSet rs = stmt.executeQuery()) {
                ResultSetMetaData metaData = rs.getMetaData();
                int columnCount = metaData.getColumnCount();

                while (rs.next()) {
                    JSONObject obj = new JSONObject();
                    for (int i = 1; i <= columnCount; i++) {
                        String columnName = metaData.getColumnName(i);
                        String value = rs.getString(i);
                        obj.put(columnName, value);
                    }
                    jsonArray.put(obj);
                }
            }

            stmt.close();

            return new Response(200, jsonArray);
        } catch (Exception e) {
            return new Response(500, e.getMessage());
        } finally {
            try {
                connection.close();
            } catch (Exception e) {
            }
        }
    }
}
