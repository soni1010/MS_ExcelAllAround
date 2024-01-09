package org.harender;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

public class ReUsableExcelMethods {

    public static void readXLSX_File() {

        try (FileInputStream fileInputStream = new FileInputStream("src/main/resources/FinanceSppecial.xlsx") ) {
            Workbook workbook = new XSSFWorkbook(fileInputStream);
            Sheet sheet = workbook.getSheetAt(0); // Assuming you're working with the first sheet

            for ( Row row : sheet ) {
                for ( Cell cell : row ) {
                    System.out.print(cell.toString() + "\t");
                }
                System.out.println();
            }

        } catch (Exception e ) {
            e.printStackTrace();
        }
    }

    public static List<String> getColumnData(String filePath, String sheetName, String columnHeader) {
        List<String> columnData = new ArrayList<>();

        try (FileInputStream fileInputStream = new FileInputStream(filePath)) {
            Workbook workbook = new XSSFWorkbook(fileInputStream);
            Sheet sheet = workbook.getSheet(sheetName);

            // Find the column index based on the header
            int columnIndex = -1;
            Row headerRow = sheet.getRow(0);
            Iterator<Cell> cellIterator = headerRow.cellIterator();
            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();
                System.out.println("Column "+Integer.sum(cell.getColumnIndex(),1) +" - "+cell.getStringCellValue());
                if (cell.getStringCellValue().equals(columnHeader)) {
                    columnIndex = cell.getColumnIndex();
                    break;
                }
            }

            // If the column header is not found, return an empty list
            if (columnIndex == -1) {
                System.out.println("Column header not found: " + columnHeader);
                return columnData;
            }

            // Read data from the specified column
            Iterator<Row> rowIterator = sheet.iterator();
            rowIterator.next(); // Skip the header row
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                Cell cell = row.getCell(columnIndex, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                columnData.add(cell.toString());
            }

        } catch (Exception e) {
            e.printStackTrace();
        }

        return columnData;
    }


}
