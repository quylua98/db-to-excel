package main;

import config.ConnectionUtils;
import model.DBTable;
import org.apache.poi.ss.usermodel.*;
import util.DBTableExcelUtils;

import java.io.*;
import java.util.*;
import java.util.stream.Collectors;

public class Application {
    public static void main(String... args) throws Exception {
        InputStream input = Application.class.getClassLoader().getResourceAsStream("template.xlsx");
        Workbook workbook = WorkbookFactory.create(input);
        List<DBTable> dbTables = ConnectionUtils.getTable();
        Map<String, List<DBTable>> tables = dbTables.stream().collect(Collectors.groupingBy(DBTable::getTableName, LinkedHashMap::new, Collectors.toList()));

        //Excel
        //Table
        DBTableExcelUtils.createDBTableSheet(workbook, tables);

        //Overview
        DBTableExcelUtils.createOverviewSheet(workbook, tables);

        //Save file
        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        workbook.write(bos);
        workbook.close();
        try (OutputStream outputStream = new FileOutputStream( ConnectionUtils.SCHEMA_NAME + ".xlsx")) {
            bos.writeTo(outputStream);
        }
    }


}
