package utils;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFCellUtil;

import java.math.BigDecimal;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.util.Calendar;
import java.util.Date;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.regex.Pattern;

/**
 * Created by dennis.yu on 24/07/2017.
 */
public class ResultSetToExcel {
    private LinkedHashMap<String, String> mapSqlTypeExcelFormat;

    public HSSFWorkbook dump(ResultSet resultSet) throws SQLException {
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFFont boldFont = workbook.createFont();
        boldFont.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
        HSSFSheet sheet = workbook.createSheet("sheet");
        HSSFRow titleRow = sheet.createRow(0);
        ResultSetMetaData metaData = resultSet.getMetaData();
        int columnCount = metaData.getColumnCount();
        HSSFDataFormat dataFormat = workbook.createDataFormat();
        HSSFCellStyle[] dataStyles = new HSSFCellStyle[columnCount + 1];
        for (int colIndex = 0; colIndex < columnCount; colIndex++) {
            String title = metaData.getColumnLabel(colIndex + 1);
            HSSFCell cell = HSSFCellUtil.createCell(titleRow, colIndex, title);
            HSSFCellStyle style = workbook.createCellStyle();
            style.setFont(boldFont);
            cell.setCellStyle(style);
            HSSFCellStyle dataStyle = getDataStyle(workbook, metaData, colIndex, dataFormat);
            dataStyles[colIndex] = dataStyle;
        }
        dumpData(resultSet, sheet, columnCount, dataStyles);
        return workbook;
    }

    private HSSFCellStyle getDataStyle(HSSFWorkbook workbook, ResultSetMetaData metaData, int colIndex, HSSFDataFormat dataFormat) throws SQLException {
        HSSFCellStyle dataStyle = workbook.createCellStyle();
        String columnType = metaData.getColumnTypeName(colIndex + 1);
        columnType += "(" + metaData.getPrecision(colIndex + 1);
        columnType += "," + metaData.getScale(colIndex + 1) + ")";
        String excelFormat = getExcelFormat(columnType);
        final short format = dataFormat.getFormat(excelFormat);
        dataStyle.setDataFormat(format);
        return dataStyle;
    }

    private String getExcelFormat(String columnType) {
        for (Map.Entry<String, String> entry : mapSqlTypeExcelFormat.entrySet()) {
            if (Pattern.matches(entry.getKey(), columnType)) {
                return entry.getValue();
            }
        }
        return "text";
    }

    private void dumpData(ResultSet resultSet, HSSFSheet sheet, int columnCount, HSSFCellStyle[] dataStyles) throws SQLException {
        int currentRow = 1;
        // The result set can have been already browsed
        //resultSet.beforeFirst();
        while (resultSet.next()) {
            HSSFRow row = sheet.createRow(currentRow++);
            for (int colIndex = 0; colIndex < columnCount; colIndex++) {
                Object value = resultSet.getObject(colIndex + 1);
                final HSSFCell cell = row.createCell(colIndex);
                if (value == null) {
                    cell.setCellValue("");
                } else {
                    cell.setCellStyle(dataStyles[colIndex]);
                    if (value instanceof Calendar) {
                        cell.setCellValue((Calendar) value);
                    } else if (value instanceof Date) {
                        cell.setCellValue((Date) value);
                    } else if (value instanceof String) {
                        cell.setCellValue((String) value);
                    } else if (value instanceof Boolean) {
                        cell.setCellValue((Boolean) value);
                    } else if (value instanceof Double) {
                        cell.setCellValue((Double) value);
                    } else if (value instanceof BigDecimal) {
                        cell.setCellValue(((BigDecimal) value).doubleValue());
                    }
                }
            }
        }
        for (int i = 0; i < columnCount; i++) {
            sheet.autoSizeColumn(i);
        }
    }

    public void setMapSqlTypeExcelFormat(LinkedHashMap<String, String> mapSqlTypeExcelFormat) {
        this.mapSqlTypeExcelFormat = mapSqlTypeExcelFormat;
    }
}
