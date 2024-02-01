package org.example.QueryBuilder;

import org.example.libs.Response;
import org.json.JSONArray;
import org.json.JSONObject;

import java.sql.*;
import java.util.List;
import java.util.Map;

public class QueryBuilder {
    public static Response querySelect(Connection conn, String query) {
        JSONArray jsonArray = new JSONArray();

        try (Statement stmt = conn.createStatement();
             ResultSet rs = stmt.executeQuery(query)) {
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
        }catch(Exception e){
            return new Response(500,e.getMessage());
        }finally {
            try{
            conn.close();}
            catch (Exception e){
                return new Response(500,e.getMessage());
            }
        }
        return new Response(200,jsonArray);
    }


    public static Response insertData(Connection conn, String insertQuery, Map<String, String> values) {
        try ( PreparedStatement pstmt = conn.prepareStatement(insertQuery, Statement.RETURN_GENERATED_KEYS)) {
            int index = 1;
            for (String value : values.values()) {
                try {
                    pstmt.setString(index++, value);
                } catch (SQLException e) {
                    throw new RuntimeException(e);
                }
            }
            pstmt.executeUpdate();
            try (ResultSet generatedKeys = pstmt.getGeneratedKeys()) {
                if (generatedKeys.next()) {
                    long generatedId = generatedKeys.getLong(1);
                    Response table = getTableNameFromInsertQuery(insertQuery);
                    Response primary =  findPrimaryKeyColumnName(conn,null,table.getMessage());
                    return querySelect(conn, "SELECT * FROM " + table.getMessage() + " WHERE "+primary.getMessage()+" = " + generatedId);
                }
            }catch(Exception e){
                return  new Response(500, "Error: "+e.getMessage());
            }
        } catch (SQLException e) {
            return  new Response(500, "Error: "+e.getMessage());

        }
        return new Response(false,"Failed Please try again");
    }
    public static Response updateData(Connection conn, String insertQuery, Map<String, String> values, String id){
        try ( PreparedStatement pstmt = conn.prepareStatement(insertQuery, Statement.RETURN_GENERATED_KEYS)) {

            int index = 1;
            for (String value : values.values()) {
                pstmt.setString(index++, value);
            }
            pstmt.setString(index, id);
            pstmt.executeUpdate();
            Response table = getTableNameFromUpdateQuery(insertQuery);
            Response primary =  findPrimaryKeyColumnName(conn,null,table.getMessage());
            return querySelect(conn, "SELECT * FROM " + table.getMessage() + " WHERE "+primary.getMessage()+" = " + id);
        }catch (Exception e){
            try {
                conn.close();
            } catch (SQLException z) {
                return new Response(500, z.getMessage());
            }
                return new Response(404, e.getMessage());

        }
    }
    public static Response insertDataBatch(Connection conn, String insertQuery, List<Map<String, String>> valuesList) throws SQLException {
        try ( PreparedStatement pstmt = conn.prepareStatement(insertQuery, Statement.RETURN_GENERATED_KEYS)) {

            for (Map<String, String> values : valuesList) {
                int index = 1;
                for (String value : values.values()) {
                    pstmt.setString(index++, value);
                }
                pstmt.addBatch();
            }
            pstmt.executeBatch();
            try (ResultSet generatedKeys = pstmt.getGeneratedKeys()) {
                if (generatedKeys.next()) {
                    long generatedId = generatedKeys.getLong(1);
                    Response table = getTableNameFromInsertQuery(insertQuery);
                    Response primary =  findPrimaryKeyColumnName(conn,null,table.getMessage());
                    return querySelect(conn, "SELECT * FROM " + table.getMessage() + " WHERE "+primary.getMessage()+" = " + generatedId);
                }
            }
        } try {
            conn.close();
        } catch (SQLException e) {
            return new Response(500,e.getMessage());
        }
        return new Response(false,"Failed Please try again");
    }
    public static Response updateDataBatch(Connection conn, String insertQuery, List<Map<String, String>> valuesList){
        try ( PreparedStatement pstmt = conn.prepareStatement(insertQuery, Statement.RETURN_GENERATED_KEYS)) {

            for (Map<String, String> values : valuesList) {
                int index = 1;
                for (String value : values.values()) {
                    pstmt.setString(index++, value);
                }
                pstmt.addBatch();
            }
            pstmt.executeBatch();
            try (ResultSet generatedKeys = pstmt.getGeneratedKeys()) {
                if (generatedKeys.next()) {
                    long generatedId = generatedKeys.getLong(1);
                    Response table = getTableNameFromUpdateQuery(insertQuery);
                    Response primary =  findPrimaryKeyColumnName(conn,null,table.getMessage());
                    return querySelect(conn, "SELECT * FROM " + table.getMessage() + " WHERE "+primary.getMessage()+" = " + generatedId);
                }
            }catch (Exception e){
               return new Response(false,e.getMessage());
            }
            conn.close();
        } catch (SQLException e) {
            return new Response(false,e.getMessage());
        }
        return new Response(false,"Failed Please try again");
    }
    public static Response findPrimaryKeyColumnName(Connection conn, String schema, String table) throws SQLException {
        String primaryKeyColumnName = "id"; // default primary key name
        DatabaseMetaData dbMetaData = conn.getMetaData();
        try (ResultSet rs = dbMetaData.getPrimaryKeys(null, schema, table)) {
            if (rs.next()) {
                primaryKeyColumnName = rs.getString("COLUMN_NAME");
            } else {
                System.out.println("No primary key found for table: " + table);
            }
        } catch (SQLException e) {
            try {
                conn.close();
            } catch (SQLException z) {
                return new Response(500, z.getMessage());
            }
                return new Response(500,e.getMessage());

            }
        return new Response(200,primaryKeyColumnName);
    }
    public static Response getTableNameFromInsertQuery(String query) {
        String lowerCaseQuery = query.toLowerCase();
        if (lowerCaseQuery.startsWith("insert into")) {
            int intoIndex = lowerCaseQuery.indexOf("into ");
            int openParenIndex = lowerCaseQuery.indexOf('(', intoIndex + 5);

            if (intoIndex != -1) {
                String tableNamePart = openParenIndex != -1
                        ? lowerCaseQuery.substring(intoIndex + 5, openParenIndex)
                        : lowerCaseQuery.substring(intoIndex + 5);
                return new Response(200,tableNamePart.trim().split("\\s+")[0]);
            }
        }
        return new Response(404,"");
    }

    public static Response getTableNameFromUpdateQuery(String query) {
        String lowerCaseQuery = query.toLowerCase();
        if (lowerCaseQuery.startsWith("update ")) {
            int updateIndex = lowerCaseQuery.indexOf("update ");
            int setIndex = lowerCaseQuery.indexOf(" set", updateIndex + 7);

            if (updateIndex != -1) {
                String tableNamePart = setIndex != -1
                        ? lowerCaseQuery.substring(updateIndex + 7, setIndex)
                        : lowerCaseQuery.substring(updateIndex + 7);
                return new Response(200,tableNamePart.trim().split("\\s+")[0]);
            }
        }
        return new Response(404,"");
    }
}
