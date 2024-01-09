package org.harender;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.FileInputStream;

public class ReUsableExcelMethods {

    public static void readXLSX_File() {

        try (FileInputStream fileInputStream = new FileInputStream("path/to/your/file.xlsx") ) {
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

}
