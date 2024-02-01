package org.example.QueryBuilder;

import org.example.libs.Response;

import java.sql.*;
import java.util.*;
import static org.example.QueryBuilder.QueryBuilder.querySelect;

public class Select {

    // Select all columns from a table
    public static Response select(Connection conn, String table) throws SQLException {
        return querySelect(conn, "SELECT * FROM " + table);
    }

    // Select specific columns from a table
    public static Response select(Connection conn, String table, String[] columns) throws SQLException {
        String columnList = String.join(", ", columns);
        return querySelect(conn, "SELECT " + columnList + " FROM " + table);
    }

    // Select with WHERE clause
    public static Response select(Connection conn, String table, String where) throws SQLException {
        return querySelect(conn, "SELECT * FROM " + table + " WHERE " + where);
    }

    // Select with WHERE clause and specific columns
    public static Response select(Connection conn, String table, String[] columns, String where) throws SQLException {
        String columnList = String.join(", ", columns);
        return querySelect(conn, "SELECT " + columnList + " FROM " + table + " WHERE " + where);
    }

    // Select with WHERE clause and GROUP BY
    public static Response select(Connection conn, String table, String where, String groupBy) throws SQLException {
        return querySelect(conn, "SELECT * FROM " + table + " WHERE " + where + " GROUP BY " + groupBy);
    }

    // Select with WHERE clause, GROUP BY, and specific columns
    public static Response select(Connection conn, String table, String[] columns, String where, String groupBy) throws SQLException {
        String columnList = String.join(", ", columns);
        return querySelect(conn, "SELECT " + columnList + " FROM " + table + " WHERE " + where + " GROUP BY " + groupBy);
    }


    // Select with LIMIT
    public static Response select(Connection conn, String table, int limit) throws SQLException {
        return querySelect(conn, "SELECT * FROM " + table + " LIMIT " + limit);
    }

    // Combination of WHERE, GROUP BY, ORDER BY, specific columns, and LIMIT
    public static Response select(Connection conn, String table, String[] columns, String where, String groupBy, String orderBy, Integer limit) throws SQLException {
        String columnList = String.join(", ", columns);
        StringJoiner query = new StringJoiner(" ");
        query.add("SELECT").add(columnList).add("FROM").add(table);

        if (where != null && !where.isEmpty()) {
            query.add("WHERE").add(where);
        }

        if (groupBy != null && !groupBy.isEmpty()) {
            query.add("GROUP BY").add(groupBy);
        }

        if (orderBy != null && !orderBy.isEmpty()) {
            query.add("ORDER BY").add(orderBy);
        }

        if (limit != null) {
            query.add("LIMIT").add(limit.toString());
        }

        return querySelect(conn, query.toString());
    }
    public static Response selectJoin(Connection conn, String table, String[] columns, String joinType, String joinTable, String joinCondition) throws SQLException {
        String columnList = String.join(", ", columns);
        String query = "SELECT " + columnList + " FROM " + table + " " + joinType + " JOIN " + joinTable + " ON " + joinCondition;
        return querySelect(conn, query);
    }

    public static Response selectJoin(Connection conn, String table, String joinType, String joinTable, String joinCondition) throws SQLException {
        String query = "SELECT * FROM " + table + " " + joinType + " JOIN " + joinTable + " ON " + joinCondition;
        return querySelect(conn, query);
    }

    public static Response selectAdvanced(Connection conn, String[] columns, String[] tables, String[][] joinClauses, String where, String groupBy, String having, String orderBy, Integer limit) throws SQLException {
        StringJoiner query = new StringJoiner(" ");

        // SELECT clause
        if (columns != null && columns.length > 0) {
            String columnList = String.join(", ", columns);
            query.add("SELECT " + columnList);
        } else {
            query.add("SELECT *");
        }

        // FROM clause
        if (tables != null && tables.length > 0) {
            String tableList = String.join(", ", tables);
            query.add("FROM " + tableList);
        } else {
            // Handle case where no tables are provided
            throw new IllegalArgumentException("At least one table must be specified.");
        }

        // JOIN clauses
        if (joinClauses != null) {
            for (String[] joinClause : joinClauses) {
                if (joinClause != null && joinClause.length >= 3) {
                    String joinType = joinClause[0];
                    String joinTable = joinClause[1];
                    String joinCondition = joinClause[2];
                    if (joinType != null && !joinType.isEmpty() && joinTable != null && !joinTable.isEmpty() && joinCondition != null && !joinCondition.isEmpty()) {
                        query.add(joinType + " JOIN " + joinTable + " ON " + joinCondition);
                    }
                }
            }
        }

        // WHERE clause
        if (where != null && !where.isEmpty()) {
            query.add("WHERE " + where);
        }

        // GROUP BY clause
        if (groupBy != null && !groupBy.isEmpty()) {
            query.add("GROUP BY " + groupBy);
        }

        // HAVING clause
        if (having != null && !having.isEmpty()) {
            query.add("HAVING " + having);
        }

        // ORDER BY clause
        if (orderBy != null && !orderBy.isEmpty()) {
            query.add("ORDER BY " + orderBy);
        }

        // LIMIT clause
        if (limit != null) {
            query.add("LIMIT " + limit.toString());
        }

        return querySelect(conn, query.toString());
    }
}
