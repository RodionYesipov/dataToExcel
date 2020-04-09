import jxl.CellView;
import jxl.Workbook;
import jxl.format.Alignment;
import jxl.format.Border;
import jxl.format.BorderLineStyle;
import jxl.format.Colour;
import jxl.write.*;
import jxl.write.Number;

import java.io.File;
import java.io.IOException;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.util.HashMap;

public class ExcelWrite {
    /////////////
    public static WritableWorkbook createWritableWorkbook(String EXCEL_FILE_LOCATION) throws IOException {
        return Workbook.createWorkbook(new File(EXCEL_FILE_LOCATION));
    }

    public static WritableSheet createWritableSheet(WritableWorkbook writableWorkbook, String sheetName) {
        return writableWorkbook.createSheet(sheetName, 0);
    }

    private static void writeHeadersToSheet(WritableSheet ws, ResultSetMetaData rs, CellView cellView) throws SQLException, WriteException {
        for (int i = 0; i < rs.getColumnCount(); i++) {
            WritableFont cellFont = new WritableFont(WritableFont.COURIER,14);
            cellFont.setBoldStyle(WritableFont.BOLD);
            WritableCellFormat cellFormat = new WritableCellFormat(cellFont);
            cellFormat.setBackground(Colour.YELLOW);
            cellFormat.setAlignment(Alignment.CENTRE);
            cellFormat.setBorder(Border.ALL, BorderLineStyle.THIN);
            ws.setColumnView(i, cellView);

            Label header = new Label(i, 0, rs.getColumnName(i + 1),cellFormat);
            try {
                ws.addCell(header);
            } catch (WriteException ex) {
                ex.printStackTrace();
            }
        }
    }

    private static HashMap dataTypesMap(ResultSetMetaData resultSetMetaData) throws SQLException {
        HashMap typesMap = new HashMap();
        for (int i = 0; i < resultSetMetaData.getColumnCount(); i++) {
            typesMap.put(i, resultSetMetaData.getColumnTypeName(i+1));
        }
        return typesMap;
    }

    public static boolean resultSetToFile(ResultSet resultSet, String filePath) throws SQLException, IOException {
        ResultSetMetaData resultSetMetaData = resultSet.getMetaData();
        int columnCount = resultSetMetaData.getColumnCount();
        WritableWorkbook workbook = createWritableWorkbook(filePath);
        WritableSheet sheet = createWritableSheet(workbook, "Sheet 1");
        //
        CellView cellView = new CellView();
        cellView.setAutosize(true);

        HashMap hashMap = dataTypesMap(resultSetMetaData);
        try {
            //headers
            writeHeadersToSheet(sheet, resultSetMetaData, cellView);
            //data
            WritableCellFormat cellFormat = new WritableCellFormat();
            cellFormat.setBorder(Border.ALL,BorderLineStyle.THIN);
            while (resultSet.next()) {
                for (int i = 0; i < columnCount; i++) {
                    if (hashMap.get(i) == "DOUBLE" || hashMap.get(i) == "NUMERIC"
                            || hashMap.get(i) == "INT" || hashMap.get(i) == "BIGINT") {
                        sheet.addCell(new Number(i, resultSet.getRow(), resultSet.getDouble(i + 1), cellFormat));
                    } else {
                        sheet.addCell(new Label(i, resultSet.getRow(), resultSet.getString(i + 1), cellFormat));
                    }
                }
            }
            workbook.write();
        } catch (WriteException e){
            e.printStackTrace();
        } finally {
            if (workbook != null) {
                try {
                    workbook.close();
                } catch (IOException e) {
                    e.printStackTrace();
                } catch (WriteException e) {
                    e.printStackTrace();
                }
            }
        }
        return true;
    }
}
