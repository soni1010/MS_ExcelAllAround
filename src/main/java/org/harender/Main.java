package org.harender;

public class Main {
    public static void main(String[] args) {
        System.out.println("Reading File");

        String filePath = "src/main/resources/FinanceSppecial.xlsx";
        String sheetName = "Sheet1";
        String columnHeader = "YourColumnName";

        /*
        Column 1 - Test Case ID
        Column 2 - Environment Type
        Column 3 - Version
        Column 4 - TC_Type
        Column 5 - TC_Header
        Column 6 - TC_Steps
        Column 7 - Expected
        Column 8 - Execution Output
        Column 9 - Status
        Column 10 - Defect ID
        */

        //ReUsableExcelMethods.readXLSX_File();
        ReUsableExcelMethods.getColumnData(filePath,sheetName,columnHeader);
    }
}