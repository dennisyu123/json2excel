package utils;

import com.mashape.unirest.http.HttpResponse;
import com.mashape.unirest.http.JsonNode;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFCellUtil;
import org.json.JSONArray;

import java.math.BigDecimal;
import java.util.Calendar;
import java.util.Date;

public class JsonObjectToExcel {

    public HSSFWorkbook dump(HttpResponse<JsonNode> jsonNode){

        // get column array
        JSONArray col = (JSONArray) jsonNode.getBody().getObject().get("columns");
        JSONArray data = (JSONArray) jsonNode.getBody().getObject().get("data");

        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFFont boldFont = workbook.createFont();
        boldFont.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
        HSSFSheet sheet = workbook.createSheet("sheet");
        HSSFRow titleRow = sheet.createRow(0);

        int columnCount = col.length();

        HSSFCellStyle[] dataStyles = new HSSFCellStyle[columnCount + 1];
        for (int colIndex = 0; colIndex < columnCount; colIndex++) {
            String title = col.getString(colIndex);
            HSSFCell cell = HSSFCellUtil.createCell(titleRow, colIndex, title);
            HSSFCellStyle style = workbook.createCellStyle();
            style.setFont(boldFont);
            cell.setCellStyle(style);
            HSSFCellStyle dataStyle = workbook.createCellStyle();

            dataStyles[colIndex] = dataStyle;
        }
        dumpData(data, sheet, columnCount, dataStyles);

        return workbook;
    }

    private void dumpData(JSONArray data, HSSFSheet sheet, int columnCount, HSSFCellStyle[] dataStyles){
        int rowCount = data.length();
        int currentRow = 1;

        for(int i = 0; i < rowCount; i++){
            HSSFRow row = sheet.createRow(currentRow++);

            JSONArray rowData = (JSONArray) data.get(i);

            for (int colIndex = 0; colIndex < columnCount; colIndex++) {
                Object value = rowData.get(colIndex);

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
}
