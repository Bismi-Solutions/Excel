package solutions.bismi.excel;

import java.io.File;
import java.util.Date;

public class ExcelLibraryDemo {

    public static void main(String[] args) {
        ExcelApplication xlApp = new ExcelApplication();
        String outputFileName = "demo_workbook.xlsx";
        String outputFilePath = System.getProperty("user.dir") + File.separator + outputFileName;

        System.out.println("Attempting to create demo workbook at: " + outputFilePath);

        ExcelWorkBook xlBook = xlApp.createWorkBook(outputFilePath);

        if (xlBook == null || xlBook.getWb() == null) {
            System.err.println("Failed to create the demo workbook. Please check permissions and path.");
            return;
        }

        // Sheet 1: Data Showcase
        try {
            ExcelWorkSheet sheet1 = xlBook.addSheet("Data Showcase");
            if (sheet1 == null) {
                System.err.println("Failed to create sheet 'Data Showcase'");
            } else {
                sheet1.activate();

                // Title
                sheet1.cell(1, 1).setText("Data Showcase");
                sheet1.cell(1, 1).setFontStyle(true, false, false).setFontColor("DARK_BLUE"); // Bold, Dark Blue
                sheet1.cell(1,1).getCell().getCellStyle().getFont().setFontHeightInPoints((short)16); // Example direct POI for font size

                // Text examples
                sheet1.cell(3, 1).setText("Simple Text:");
                sheet1.cell(3, 2).setText("Hello, Excel World!");
                sheet1.cell(3, 2).setFontColor("GREEN");

                sheet1.cell(4, 1).setText("Styled Text:");
                sheet1.cell(4, 2).setText("Bold & Italic");
                sheet1.cell(4, 2).setFontStyle(true, true, false).setFontColor("RED");

                // Number examples
                sheet1.cell(6, 1).setText("Integer:");
                sheet1.cell(6, 2).setNumericValue(12345);
                sheet1.cell(6, 2).setFillColor("LIGHT_YELLOW");

                sheet1.cell(7, 1).setText("Decimal:");
                sheet1.cell(7, 2).setNumericValue(678.90);
                sheet1.cell(7, 2).setFillColor("LIGHT_GREEN");

                // Date example
                sheet1.cell(9, 1).setText("Current Date:");
                sheet1.cell(9, 2).setDateValue(new Date());
                sheet1.cell(9, 2).setNumberFormat("dd-mmm-yyyy hh:mm AM/PM"); // Date format

                // Alignment and Wrapping
                sheet1.cell(11, 1).setText("Alignment & Wrap Test:");
                ExcelCell cellForWrap = sheet1.cell(11, 2);
                cellForWrap.setText("This is a long piece of text that should wrap within the cell to demonstrate text wrapping capabilities of the library.");
                // Text wrapping needs to be set on CellStyle directly via POI for now
                org.apache.poi.ss.usermodel.CellStyle wrapStyle = cellForWrap.getCell().getCellStyle();
                if (wrapStyle == null) wrapStyle = xlBook.getWb().createCellStyle();
                wrapStyle.setWrapText(true);
                cellForWrap.getCell().setCellStyle(wrapStyle);
                // Adjust row height manually for wrapped text to be visible
                sheet1.getSheet().getRow(10).setHeightInPoints((short)(3 * sheet1.getSheet().getDefaultRowHeightInPoints()));


                sheet1.cell(12, 1).setText("Top-Left");
                sheet1.cell(12, 1).setHorizontalAlignment("LEFT").setVerticalAlignment("TOP").setFillColor("LIGHT_CORNFLOWER_BLUE");
                sheet1.cell(12, 2).setText("Center-Center");
                sheet1.cell(12, 2).setHorizontalAlignment("CENTER").setVerticalAlignment("CENTER").setFillColor("LIGHT_CORNFLOWER_BLUE");
                sheet1.cell(12, 3).setText("Bottom-Right");
                sheet1.cell(12, 3).setHorizontalAlignment("RIGHT").setVerticalAlignment("BOTTOM").setFillColor("LIGHT_CORNFLOWER_BLUE");

                // Column width for wrapped text
                sheet1.getSheet().setColumnWidth(1, 25 * 256); // Approx 25 characters width for column B (index 1)

                System.out.println("'Data Showcase' sheet created.");
            }
        } catch (Exception e) {
            System.err.println("Error creating 'Data Showcase' sheet: " + e.getMessage());
            e.printStackTrace();
        }

        // Sheet 2: Formulas and Structures
        try {
            ExcelWorkSheet sheet2 = xlBook.addSheet("Formulas and Structures");
            if (sheet2 == null) {
                System.err.println("Failed to create sheet 'Formulas and Structures'");
            } else {
                sheet2.activate();

                // Title
                sheet2.cell(1, 1).setText("Sales Report Q1");
                sheet2.cell(1, 1).setFontStyle(true, false, false).setFontColor("DARK_RED");
                sheet2.cell(1,1).getCell().getCellStyle().getFont().setFontHeightInPoints((short)14);
                sheet2.mergeCells(1, 1, 1, 4); // Merge R1C1 to R1C4
                sheet2.cell(1,1).setHorizontalAlignment("CENTER");


                // Data table
                sheet2.cell(3, 1).setText("Month").setFontStyle(true, false, false);
                sheet2.cell(3, 2).setText("Product A Sales").setFontStyle(true, false, false);
                sheet2.cell(3, 3).setText("Product B Sales").setFontStyle(true, false, false);
                sheet2.cell(3, 4).setText("Monthly Total").setFontStyle(true, false, false);

                sheet2.cell(4, 1).setText("January");
                sheet2.cell(4, 2).setNumericValue(1500);
                sheet2.cell(4, 3).setNumericValue(1200);
                sheet2.cell(4, 4).setFormula("B4+C4");

                sheet2.cell(5, 1).setText("February");
                sheet2.cell(5, 2).setNumericValue(1800);
                sheet2.cell(5, 3).setNumericValue(1100);
                sheet2.cell(5, 4).setFormula("B5+C5");

                sheet2.cell(6, 1).setText("March");
                sheet2.cell(6, 2).setNumericValue(2100);
                sheet2.cell(6, 3).setNumericValue(1350);
                sheet2.cell(6, 4).setFormula("B6+C6");

                // Totals and Average
                sheet2.cell(8, 1).setText("Total Sales").setFontStyle(true, false, false);
                sheet2.cell(8, 2).setFormula("SUM(B4:B6)");
                sheet2.cell(8, 3).setFormula("SUM(C4:C6)");
                sheet2.cell(8, 4).setFormula("SUM(D4:D6)");

                sheet2.cell(9, 1).setText("Average Sales").setFontStyle(true, false, false);
                sheet2.cell(9, 2).setFormula("AVERAGE(B4:B6)");
                sheet2.cell(9, 3).setFormula("AVERAGE(C4:C6)");

                // Apply borders
                for(int r=3; r<=9; r++) {
                    for(int c=1; c<=4; c++) {
                        if (sheet2.cell(r,c).getTextValue() != null && !sheet2.cell(r,c).getTextValue().isEmpty()) {
                             sheet2.cell(r,c).setFullBorder("GREY_50_PERCENT");
                        } else if (sheet2.cell(r,c).getValue() != null && !sheet2.cell(r,c).getValue().isEmpty()){
                             sheet2.cell(r,c).setFullBorder("GREY_50_PERCENT");
                        }
                    }
                }
                 System.out.println("'Formulas and Structures' sheet created.");
            }
        } catch (Exception e) {
            System.err.println("Error creating 'Formulas and Structures' sheet: " + e.getMessage());
            e.printStackTrace();
        }

        // Sheet 3: Advanced Formatting
        try {
            ExcelWorkSheet sheet3 = xlBook.addSheet("Advanced Formatting");
            if (sheet3 == null) {
                System.err.println("Failed to create sheet 'Advanced Formatting'");
            } else {
                sheet3.activate();

                sheet3.cell(1,1).setText("Formatting Showcase").setFontStyle(true, false, false);

                sheet3.cell(3, 1).setText("Currency (USD):");
                sheet3.cell(3, 2).setNumericValue(12345.67);
                sheet3.cell(3, 2).setNumberFormat("$#,##0.00");

                sheet3.cell(4, 1).setText("Currency (EUR):");
                sheet3.cell(4, 2).setNumericValue(12345.67);
                sheet3.cell(4, 2).setNumberFormat("€#,##0.00");


                sheet3.cell(6, 1).setText("Percentage:");
                sheet3.cell(6, 2).setNumericValue(0.8567);
                sheet3.cell(6, 2).setNumberFormat("0.00%");

                sheet3.cell(8, 1).setText("Scientific Notation:");
                sheet3.cell(8, 2).setNumericValue(1234567890.123);
                sheet3.cell(8, 2).setNumberFormat("0.00E+00");

                sheet3.cell(10, 1).setText("Custom Date:");
                sheet3.cell(10, 2).setDateValue(new Date());
                sheet3.cell(10, 2).setNumberFormat("ddd, mmm d, yyyy");

                // Column Widths
                sheet3.cell(12, 1).setText("Wide Column (A)");
                sheet3.getSheet().setColumnWidth(0, 30 * 256); // Column A (index 0) width 30 chars

                sheet3.cell(12, 2).setText("Narrow Column (B)");
                 sheet3.getSheet().setColumnWidth(1, 10 * 256); // Column B (index 1) width 10 chars

                // Text Rotation - requires direct POI manipulation if not in ExcelCell API
                ExcelCell rotationCell = sheet3.cell(14, 1);
                rotationCell.setText("Rotated Text");
                org.apache.poi.ss.usermodel.CellStyle rotationStyle = rotationCell.getCell().getCellStyle();
                if (rotationStyle == null) rotationStyle = xlBook.getWb().createCellStyle();
                rotationStyle.setRotation((short)45); // 45 degrees
                rotationCell.getCell().setCellStyle(rotationStyle);

                System.out.println("'Advanced Formatting' sheet created.");
            }
        } catch (Exception e) {
            System.err.println("Error creating 'Advanced Formatting' sheet: " + e.getMessage());
            e.printStackTrace();
        }

        // Save and Close
        if (xlBook.saveWorkbook()) {
            System.out.println("Demo workbook saved successfully to: " + outputFilePath);
        } else {
            System.err.println("Failed to save demo workbook to: " + outputFilePath);
        }
        xlApp.closeAllWorkBooks();
        System.out.println("All workbooks closed.");
    }
}
