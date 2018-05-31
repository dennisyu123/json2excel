package utils;

import com.google.gson.JsonArray;
import com.google.gson.JsonObject;

import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.sql.Types;


public class ResultSetToJson {

    public static JsonObject ResultSetToJsonObject(ResultSet rs) {

        JsonArray ja = new JsonArray();
        JsonObject jo = new JsonObject();
        ResultSetMetaData rsmd = null;
        String columnName;


        try {
            rsmd = rs.getMetaData();
            JsonArray cols = new JsonArray();
            for (int i = 0; i < rsmd.getColumnCount(); i++) {
                cols.add(rsmd.getColumnName(i + 1));
            }

            while (rs.next()) {

                JsonArray elements = new JsonArray();

                for (int i = 0; i < rsmd.getColumnCount(); i++) {
                    columnName = rsmd.getColumnName(i + 1);

                    switch (rsmd.getColumnType(i + 1)){
                        case Types.DOUBLE :{
                            elements.add(rs.getDouble(columnName));
                            break;
                        }
                        case Types.INTEGER :{
                            elements.add(rs.getInt(columnName));
                            break;
                        }
                        case Types.NVARCHAR:{
                            elements.add(rs.getNString(columnName));
                            break;
                        }
                        default:{
                            elements.add(rs.getString(columnName));
                        }
                    }
                }

                ja.add(elements);
            }

            jo.add("columns", cols);
            jo.add("data", ja);

        } catch (SQLException e) {
            e.printStackTrace();
        }
        return jo;
    }

    public static String ResultSetToJsonString(ResultSet rs) {
        return ResultSetToJsonObject(rs).toString();
    }
}
