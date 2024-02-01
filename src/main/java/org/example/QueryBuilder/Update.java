package org.example.QueryBuilder;

import org.example.libs.Response;

import java.sql.Connection;
import java.sql.SQLException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Map;
import java.util.StringJoiner;

import static org.example.QueryBuilder.QueryBuilder.*;

public class Update {
    public static Response updateSingleRow(Connection conn, String table, Map<String, String> values, String id) throws SQLException {
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        String currentTime = sdf.format(new Date());
        StringJoiner setClauses = new StringJoiner(", ");
        Response primary =  findPrimaryKeyColumnName(conn,null,table);
        for (Map.Entry<String, String> entry : values.entrySet()) {
            if (!entry.getKey().equalsIgnoreCase(primary.getMessage())) {
                setClauses.add(entry.getKey() + " = ?");
            }
        }
        setClauses.add("date_modified = '" + currentTime + "'");

        String query = "UPDATE " + table + " SET " + setClauses + " WHERE " + primary.getMessage() + " = ?";
        return updateData(conn, query, values, id);
    }

}
