package config;

import model.DBTable;

import java.io.IOException;
import java.sql.*;
import java.util.ArrayList;
import java.util.LinkedList;
import java.util.List;

public class ConnectionUtils {
    private ConnectionUtils(){}
    public static String SCHEMA_NAME;
    public static String HOST_NAME;
    public static String DB_NAME;
    public static String USERNAME;
    public static String PASSWORD;

    static {
        try {
            SCHEMA_NAME = ConfigProperty.getProperty("schema_name");
            HOST_NAME = ConfigProperty.getProperty("host_name");
            DB_NAME = ConfigProperty.getProperty("db_name");
            USERNAME = ConfigProperty.getProperty("username");
            PASSWORD = ConfigProperty.getProperty("password");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }


    public static Connection getMySQLConnection() throws SQLException {
        return getMySQLConnection(HOST_NAME, DB_NAME, USERNAME, PASSWORD);
    }

    public static Connection getMySQLConnection(String hostName, String dbName, String userName, String password) throws SQLException {
        String connectionURL = "jdbc:mysql://" + hostName + ":3306/" + dbName;
        return DriverManager.getConnection(connectionURL, userName, password);
    }

    public static List<DBTable> getTable() throws SQLException {
        Connection connection = getMySQLConnection();
        Statement statement = connection.createStatement();

        String sql = "SELECT TABLE_NAME, COLUMN_NAME, DATA_TYPE, CHARACTER_MAXIMUM_LENGTH, IS_NULLABLE, COLUMN_KEY, COLUMN_DEFAULT, EXTRA FROM information_schema.columns " +
                "WHERE table_schema = '" + SCHEMA_NAME + "' " +
                "ORDER BY table_name,ordinal_position";

        ResultSet rs = statement.executeQuery(sql);

        //Fetch data
        List<DBTable> dbTables = new ArrayList<>();
        while (rs.next()) {
            int i = 1;
            DBTable dbTable = new DBTable();
            dbTable.setTableName(rs.getString(i++));
            dbTable.setColumnName(rs.getString(i++));
            dbTable.setDataType(rs.getString(i++));
            dbTable.setCharacterMaximumLength(rs.getString(i++));
            dbTable.setIsNullable(rs.getString(i++));
            dbTable.setColumnKey(rs.getString(i++));
            dbTable.setColumnDefault(rs.getString(i++));
            dbTable.setExtra(rs.getString(i++));
            dbTables.add(dbTable);
        }
        connection.close();
        return dbTables;
    }
}