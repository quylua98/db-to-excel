package util;

import model.DBTable;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;

import java.util.List;
import java.util.Map;
import java.util.Set;

public class DBTableExcelUtils {
    private DBTableExcelUtils(){}
    public static final String[] TYPE_COLUMNS = new String[]{"Field", "Type", "Null", "Key", "Default", "Extra", "Description"};
    public static final int[] COLUMNS_WIDTH = new int[]{35, 25, 6, 6, 30, 35, 50};

    public static final int FIRST_COL = 1;
    public static final int LAST_COL = 7;

    public static final int TITLE_ROW_NUM = 2;
    public static final int TYPE_ROW_NUM = 3;
    public static final int DATA_ROW_NUM = 4;

    public static void createOverviewSheet(Workbook workbook, Map<String, List<DBTable>> tables) {
        int startRow = 3;
        int count = 1;
        Set<String> tableNameSet = tables.keySet();

        CellStyle cellStyle = workbook.createCellStyle();
        setBorder(cellStyle);
        setAlignHorizontalCenter(cellStyle);

        CellStyle fontStyle = workbook.createCellStyle();
        fontStyle.cloneStyleFrom(cellStyle);
        Font font = workbook.createFont();
        font.setUnderline(Font.U_DOUBLE);
        font.setColor(HSSFColor.HSSFColorPredefined.BLUE.getIndex());
        fontStyle.setFont(font);

        for (String tableName : tableNameSet) {
            Sheet sheet = workbook.getSheetAt(0);
            Row row = sheet.createRow(startRow++);
            Cell indexCell = row.createCell(1);
            indexCell.setCellValue(count++);
            indexCell.setCellStyle(cellStyle);

            Cell tableNameCell = row.createCell(2);
            CreationHelper createHelper = workbook.getCreationHelper();
            Hyperlink link2 = createHelper.createHyperlink(HyperlinkType.DOCUMENT);
            link2.setAddress("'" + tableName + "'!A1");
            tableNameCell.setCellValue(tableName);
            tableNameCell.setHyperlink(link2);
            tableNameCell.setCellStyle(fontStyle);

            Cell descriptionCell = row.createCell(3);
            descriptionCell.setCellStyle(cellStyle);
        }
    }

    public static void createDBTableSheet(Workbook workbook, Map<String, List<DBTable>> tables) {
        for (Map.Entry<String, List<DBTable>> table : tables.entrySet()) {
            Sheet sheet = workbook.createSheet(table.getKey());

            //Back link
            createBackCell(workbook, sheet);

            for (int i = FIRST_COL; i <= LAST_COL; i++) {
                sheet.setColumnWidth(i, COLUMNS_WIDTH[i - 1] * 256);
            }

            List<DBTable> dbTableList = table.getValue();
            DBTableExcelUtils.createTitleRow(workbook, sheet, table.getKey());
            DBTableExcelUtils.createTypeRow(workbook, sheet);

            int dataRow = DBTableExcelUtils.DATA_ROW_NUM;
            for (DBTable dbTable : dbTableList) {
                DBTableExcelUtils.createDataRow(workbook, sheet, dataRow++, dbTable);
            }
        }
    }

    public static void createTitleRow(Workbook workbook, Sheet sheet, String value) {
        Row titleRow = sheet.createRow(TITLE_ROW_NUM);
        titleRow.createCell(1).setCellValue(value);

        CellRangeAddress titleRange = new CellRangeAddress(TITLE_ROW_NUM, TITLE_ROW_NUM, FIRST_COL, LAST_COL);

        CellStyle cellStyle = workbook.createCellStyle();

        Font titleFont = workbook.createFont();
        titleFont.setFontHeightInPoints((short) 14);
        titleFont.setItalic(false);
        cellStyle.setFont(titleFont);
        setBorder(cellStyle);

        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        cellStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
        CellStyle cellWordWrap = workbook.createCellStyle();
        cellWordWrap.cloneStyleFrom(cellStyle);

        RegionUtil.setBorderTop(BorderStyle.THIN, titleRange, sheet);
        RegionUtil.setBorderLeft(BorderStyle.THIN, titleRange, sheet);
        RegionUtil.setBorderRight(BorderStyle.THIN, titleRange, sheet);
        RegionUtil.setBorderBottom(BorderStyle.THIN, titleRange, sheet);

        sheet.addMergedRegion(titleRange);
        setStyleTitleRangeCell(sheet, titleRange, cellStyle);
    }

    public static void createTypeRow(Workbook workbook, Sheet sheet) {
        Row typeRow = sheet.createRow(TYPE_ROW_NUM);

        CellStyle cellStyle = workbook.createCellStyle();
        setBorder(cellStyle);

        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        cellStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());

        CellStyle cellWordWrap = workbook.createCellStyle();
        cellWordWrap.cloneStyleFrom(cellStyle);

        for (int i = FIRST_COL; i <= TYPE_COLUMNS.length; i++) {
            Cell cell = typeRow.createCell(i);
            cell.setCellValue(TYPE_COLUMNS[i - 1]);
            cell.setCellStyle(cellWordWrap);
        }
    }

    public static void createDataRow(Workbook workbook, Sheet sheet, int rowNum, DBTable dbTable) {
        Row dataRow = sheet.createRow(rowNum);

        CellStyle cellStyle = workbook.createCellStyle();
        setBorder(cellStyle);

        CellStyle cellWordWrap = workbook.createCellStyle();
        cellWordWrap.cloneStyleFrom(cellStyle);

        int startColumn = 1;
        Cell nameCell = dataRow.createCell(startColumn++);
        nameCell.setCellValue(dbTable.getColumnName());
        nameCell.setCellStyle(cellWordWrap);


        Cell dataTypeCell = dataRow.createCell(startColumn++);
        if (StringUtils.isEmpty(dbTable.getCharacterMaximumLength())) {
            dataTypeCell.setCellValue(dbTable.getDataType());

        } else {
            dataTypeCell.setCellValue(dbTable.getDataType() + "(" + dbTable.getCharacterMaximumLength() + ")");
        }
        dataTypeCell.setCellStyle(cellWordWrap);

        Cell isNullAbleCell = dataRow.createCell(startColumn++);
        isNullAbleCell.setCellValue(dbTable.getIsNullable());
        isNullAbleCell.setCellStyle(cellWordWrap);

        Cell keyCell = dataRow.createCell(startColumn++);
        keyCell.setCellValue(dbTable.getColumnKey());
        keyCell.setCellStyle(cellWordWrap);

        Cell defaultValueCell = dataRow.createCell(startColumn++);
        defaultValueCell.setCellValue(dbTable.getColumnDefault());
        defaultValueCell.setCellStyle(cellWordWrap);

        Cell extraCell = dataRow.createCell(startColumn++);
        extraCell.setCellValue(dbTable.getExtra());
        extraCell.setCellStyle(cellWordWrap);

        Cell descriptionCell = dataRow.createCell(startColumn++);
        descriptionCell.setCellValue("");
        descriptionCell.setCellStyle(cellWordWrap);
    }

    public static void setStyleTitleRangeCell(Sheet sheet, CellRangeAddress region, CellStyle cellStyle) {
        Row row = sheet.getRow(TITLE_ROW_NUM);
        for (int j = region.getFirstColumn(); j < region.getLastColumn(); j++) {
            Cell cell = row.getCell(j);
            cell.setCellStyle(cellStyle);
        }
    }

    public static void setBorder(CellStyle cellStyle) {
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBorderTop(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);
        cellStyle.setBorderLeft(BorderStyle.THIN);
    }

    public static void setAlignHorizontalCenter(CellStyle cellStyle) {
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
    }

    public static void createBackCell(Workbook workbook, Sheet sheet) {
        //Back cell
        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        cellStyle.setFillForegroundColor(IndexedColors.LIGHT_BLUE.index);
        Font font = workbook.createFont();
        font.setColor(HSSFColor.HSSFColorPredefined.WHITE.getIndex());
        font.setBold(true);
        font.setFontHeight((short) (12*20));
        cellStyle.setFont(font);

        Cell backCell = sheet.createRow(0).createCell(1);
        CreationHelper createHelper = workbook.getCreationHelper();
        Hyperlink backLink = createHelper.createHyperlink(HyperlinkType.DOCUMENT);
        backLink.setAddress("'Overview'!A1");
        backCell.setCellValue("Back to overview");
        backCell.setHyperlink(backLink);
        backCell.setCellStyle(cellStyle);
    }
}
