package solutions.bismi.excel;

import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.MethodOrderer;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.TestMethodOrder;


import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.UUID;

@TestMethodOrder(MethodOrderer.MethodName.class)
class ExcelRowTest {

    private static final String TEST_DATA_DIR = "./resources/testdata";

    // Helper to get a unique file path and ensure directory exists
    private String getTestFilePath(String baseName, String extension) throws IOException {
        Files.createDirectories(Path.of(TEST_DATA_DIR));
        String uniqueId = UUID.randomUUID().toString().substring(0, 8);
        String fileName = baseName + "_" + uniqueId + "." + extension;
        return Path.of(TEST_DATA_DIR, fileName).toString();
    }

    // Helper for cleanup
    private void cleanupTestFile(String filePath) {
        try {
            Files.deleteIfExists(Path.of(filePath));
        } catch (IOException e) {
            System.err.println("Could not delete test file: " + filePath + " : " + e.getMessage());
        }
    }

    @Test
    void aSetfontcolor() throws IOException {
        String xlsxFile = getTestFilePath("rowFontColor", "xlsx");
        setColor2(xlsxFile);
        cleanupTestFile(xlsxFile);

        String xlsFile = getTestFilePath("rowFontColor", "xls");
        setColor2(xlsFile);
        cleanupTestFile(xlsFile);
    }

    private void setColor2(String strCompleteFileName) {
        ExcelApplication xlApp = new ExcelApplication();
        ExcelWorkBook xlbook = xlApp.createWorkBook(strCompleteFileName);
        int cnt = 0;
        cnt = xlbook.getSheetCount();
        Assertions.assertEquals(1, cnt);
        ExcelWorkSheet sh1 = xlbook.addSheet("Bismi1");
        sh1.activate();

        String[] arrRow = {"A", "B", "C", "D", "E", "0", "1.1", "0.0000", ".000A", ".0000"};
        sh1.row(11).setRowValues(arrRow);
        sh1.row(11).setFontColor("Red", 1, 3);
        sh1.row(11).setFontColor("green", 3, 7);
        sh1.row(2).setRowValues(arrRow);
        sh1.row(2).setFontColor("White");
        sh1.row(2).setFillColor("Green");
        sh1.row(2).setFullBorder("Red");
        sh1.row(11).setFullBorder("blue");
        sh1.row(5).setFullBorder("blue", 1, 10);
        sh1.saveWorkBook();
        xlApp.closeAllWorkBooks();

        // Reopen and Assert
        ExcelApplication assertApp = new ExcelApplication();
        ExcelWorkBook assertBook = assertApp.openWorkbook(strCompleteFileName);
        Assertions.assertNotNull(assertBook, "Failed to reopen workbook for row formatting assertions.");
        ExcelWorkSheet assertSheet = assertBook.getExcelSheet("Bismi1");
        Assertions.assertNotNull(assertSheet, "Failed to get sheet Bismi1 for row formatting assertions.");

        // Assertions for row 11
        // Cell (11,1) - Font Red
        org.apache.poi.ss.usermodel.Cell poiCell11_1 = assertSheet.cell(11,1).getCell();
        org.apache.poi.ss.usermodel.CellStyle style11_1 = poiCell11_1.getCellStyle();
        org.apache.poi.ss.usermodel.Font font11_1 = assertBook.getWb().getFontAt(style11_1.getFontIndexAsInt());
        Assertions.assertEquals(org.apache.poi.ss.usermodel.IndexedColors.RED.getIndex(), font11_1.getColor(), "Font color for (11,1) should be RED");

        // Cell (11,3) - Font Green (overwrote Red)
        org.apache.poi.ss.usermodel.Cell poiCell11_3 = assertSheet.cell(11,3).getCell();
        org.apache.poi.ss.usermodel.CellStyle style11_3 = poiCell11_3.getCellStyle();
        org.apache.poi.ss.usermodel.Font font11_3 = assertBook.getWb().getFontAt(style11_3.getFontIndexAsInt());
        Assertions.assertEquals(org.apache.poi.ss.usermodel.IndexedColors.GREEN.getIndex(), font11_3.getColor(), "Font color for (11,3) should be GREEN");

        // Cell (11,7) - Font Green
        org.apache.poi.ss.usermodel.Cell poiCell11_7 = assertSheet.cell(11,7).getCell();
        org.apache.poi.ss.usermodel.CellStyle style11_7 = poiCell11_7.getCellStyle();
        org.apache.poi.ss.usermodel.Font font11_7 = assertBook.getWb().getFontAt(style11_7.getFontIndexAsInt());
        Assertions.assertEquals(org.apache.poi.ss.usermodel.IndexedColors.GREEN.getIndex(), font11_7.getColor(), "Font color for (11,7) should be GREEN");


        // Assertions for row 2
        // Cell (2,1) - Font White, Fill Green
        org.apache.poi.ss.usermodel.Cell poiCell2_1 = assertSheet.cell(2,1).getCell();
        org.apache.poi.ss.usermodel.CellStyle style2_1 = poiCell2_1.getCellStyle();
        org.apache.poi.ss.usermodel.Font font2_1 = assertBook.getWb().getFontAt(style2_1.getFontIndexAsInt());
        Assertions.assertEquals(org.apache.poi.ss.usermodel.IndexedColors.WHITE.getIndex(), font2_1.getColor(), "Font color for (2,1) should be WHITE");
        Assertions.assertEquals(org.apache.poi.ss.usermodel.IndexedColors.GREEN.getIndex(), style2_1.getFillForegroundColor(), "Fill color for (2,1) should be GREEN");
        Assertions.assertEquals(org.apache.poi.ss.usermodel.FillPatternType.SOLID_FOREGROUND, style2_1.getFillPattern());
        // Border check (e.g., bottom border)
        Assertions.assertEquals(org.apache.poi.ss.usermodel.IndexedColors.RED.getIndex(), style2_1.getBottomBorderColor(), "Bottom border for (2,1) should be RED");
        Assertions.assertEquals(org.apache.poi.ss.usermodel.BorderStyle.MEDIUM, style2_1.getBorderBottom(), "Bottom border style for (2,1) should be MEDIUM");


        assertApp.closeAllWorkBooks();
    }


    @Test
    void bTestSetRowValuesVariants() throws IOException {
        String xlsxFile = getTestFilePath("rowValuesVariants", "xlsx");
        testSetRowValuesVariants(xlsxFile);
        cleanupTestFile(xlsxFile);

        String xlsFile = getTestFilePath("rowValuesVariants", "xls");
        testSetRowValuesVariants(xlsFile);
        cleanupTestFile(xlsFile);
    }

    private void testSetRowValuesVariants(String strCompleteFileName) {
        ExcelApplication xlApp = new ExcelApplication();
        ExcelWorkBook xlbook = xlApp.createWorkBook(strCompleteFileName);
        ExcelWorkSheet sheet = xlbook.addSheet("RowTest");

        // Test setRowValues(String[])
        String[] values1 = {"Value1", "Value2", "Value3"};
        sheet.row(1).setRowValues(values1);

        // Test setRowValues(String[], int)
        String[] values2 = {"Value4", "Value5", "Value6"};
        sheet.row(2).setRowValues(values2, 2);

        // Test setRowValues(String[], boolean)
        String[] values3 = {"Value7", "Value8", "Value9"};
        sheet.row(3).setRowValues(values3, true);

        // Test setRowValues(String[], int, boolean)
        String[] values4 = {"Value10", "Value11", "Value12"};
        sheet.row(4).setRowValues(values4, 3, true);
        // Verify row 1
        Assertions.assertEquals("Value1", sheet.cell(1, 1).getTextValue());

        // Verify row 2 (with offset)
        Assertions.assertEquals("Value4", sheet.cell(2, 3).getTextValue());

        // Verify row 3
        Assertions.assertEquals("Value7", sheet.cell(3, 1).getTextValue());

        // Verify row 4 (with offset)
        Assertions.assertEquals("Value10", sheet.cell(4, 4).getTextValue());

        xlApp.closeAllWorkBooks();
    }

    /**
     * Test setFillColor with start and end column
     */
    @Test
    void cTestSetFillColorWithRange() throws IOException {
        String xlsxFile = getTestFilePath("fillColorRange", "xlsx");
        testSetFillColorWithRange(xlsxFile);
        cleanupTestFile(xlsxFile);

        String xlsFile = getTestFilePath("fillColorRange", "xls");
        testSetFillColorWithRange(xlsFile);
        cleanupTestFile(xlsFile);
    }

    private void testSetFillColorWithRange(String strCompleteFileName) {
        ExcelApplication xlApp = new ExcelApplication();
        ExcelWorkBook xlbook = xlApp.createWorkBook(strCompleteFileName);
        ExcelWorkSheet sheet = xlbook.addSheet("FillTest");

        // Add some data
        String[] values = {"Col1", "Col2", "Col3", "Col4", "Col5"};
        sheet.row(1).setRowValues(values);

        // Set fill color for specific columns
        sheet.row(1).setFillColor("Yellow", 2, 4);

        sheet.saveWorkBook();
        xlApp.closeAllWorkBooks();
    }

    /**
     * Test error handling and edge cases
     */
    @Test
    void dTestErrorHandling() throws IOException {
        String testFile = getTestFilePath("rowErrorHandling", "xlsx");
        ExcelApplication xlApp = new ExcelApplication();
        ExcelWorkBook xlbook = xlApp.createWorkBook(testFile);
        ExcelWorkSheet sheet = xlbook.addSheet("ErrorTest");

        // Test with empty array
        String[] emptyArray = {};
        sheet.row(1).setRowValues(emptyArray);

        // Test with null array (should not throw exception)
        try {
            sheet.row(2).setRowValues(null);
        } catch (Exception e) {
            Assertions.fail("Should not throw exception with null array");
        }

        // Test with negative row index
        // Note: This might throw an exception depending on implementation
        // Just making sure it doesn't crash the application
        try {
            sheet.row(-1).setFontColor("Red");
        } catch (Exception e) {
            // It's acceptable for this to throw an exception
            // Just making sure the application continues to function
        }

        // Test with invalid color name (should not throw exception)
        try {
            sheet.row(3).setFontColor("InvalidColor");
        } catch (Exception e) {
            Assertions.fail("Should not throw exception with invalid color name");
        }

        // Test with invalid column range
        try {
            sheet.row(4).setFillColor("Blue", 5, 2); // end < start
        } catch (Exception e) {
            Assertions.fail("Should not throw exception with invalid column range");
        }

        sheet.saveWorkBook(); // Ensure sheet is saved before closing
        xlApp.closeAllWorkBooks();
        cleanupTestFile(testFile);
    }

    /**
     * Test with large data set to verify performance
     */
    @Test
    void eTestLargeDataSet() throws java.io.IOException {
        // Using a single unique file for this test as it's more about data volume.
        String filePath = getTestFilePath("largeDataSet", "xlsx");
        // No need to deleteIfExists here as getTestFilePath ensures uniqueness,
        // and cleanupTestFile will handle it at the end.

        ExcelApplication xlApp = new ExcelApplication();
        ExcelWorkBook xlbook = xlApp.createWorkBook(filePath);
        ExcelWorkSheet sheet = xlbook.addSheet("LargeData");
        Assertions.assertNotNull(sheet, "Sheet 'LargeData' should be created.");
    
        int numRows = 100; // Increased from 20, can be set higher for true perf testing
        int numCols = 3;

        for (int i = 1; i <= numRows; i++) {
            String[] rowData = new String[numCols];
            for (int j = 0; j < numCols; j++) {
                rowData[j] = "Row" + i + "Col" + (j + 1);
            }
            sheet.row(i).setRowValues(rowData);
    
            if (i % 2 == 0) {
                sheet.row(i).setFillColor("LIGHT_GREY"); // LIGHT_GREY is a valid color name
            } else {
                sheet.row(i).setFontColor("BLUE");
            }
        }

        // Verify data integrity before saving
        Assertions.assertEquals("Row1Col1", sheet.cell(1, 1).getTextValue());
        Assertions.assertEquals("Row" + (numRows / 2) + "Col2", sheet.cell(numRows / 2, 2).getTextValue());
        Assertions.assertEquals("Row" + numRows + "Col3", sheet.cell(numRows, numCols).getTextValue());

        boolean saveResult = sheet.saveWorkBook();
        Assertions.assertTrue(saveResult, "Saving large dataset workbook should succeed.");
        xlApp.closeAllWorkBooks();

        // Reopen and Assert
        ExcelApplication assertApp = new ExcelApplication();
        ExcelWorkBook assertBook = assertApp.openWorkbook(filePath);
        Assertions.assertNotNull(assertBook, "Failed to reopen workbook for large dataset assertions.");
        ExcelWorkSheet assertSheet = assertBook.getExcelSheet("LargeData");
        Assertions.assertNotNull(assertSheet, "Failed to get sheet LargeData for assertions.");

        // Verify data for a few key cells after reopening
        Assertions.assertEquals("Row1Col1", assertSheet.cell(1, 1).getTextValue(), "Value at (1,1) after reopen");

        org.apache.poi.ss.usermodel.Cell poiCell1_1 = assertSheet.cell(1,1).getCell();
        org.apache.poi.ss.usermodel.CellStyle style1_1 = poiCell1_1.getCellStyle();
        org.apache.poi.ss.usermodel.Font font1_1 = assertBook.getWb().getFontAt(style1_1.getFontIndexAsInt());
        Assertions.assertEquals(org.apache.poi.ss.usermodel.IndexedColors.BLUE.getIndex(), font1_1.getColor(), "Font color for (1,1) should be BLUE");

        Assertions.assertEquals("Row" + (numRows/2) + "Col2", assertSheet.cell(numRows/2, 2).getTextValue(), "Value at (numRows/2,2) after reopen");
        
        org.apache.poi.ss.usermodel.Cell poiCellEvenRow = assertSheet.cell(2,1).getCell(); // An even row
        org.apache.poi.ss.usermodel.CellStyle styleEvenRow = poiCellEvenRow.getCellStyle();
        Assertions.assertEquals(org.apache.poi.ss.usermodel.IndexedColors.GREY_25_PERCENT.getIndex(), styleEvenRow.getFillForegroundColor(), "Fill color for an even row (2,1) should be LIGHT_GREY (GREY_25_PERCENT)");
        Assertions.assertEquals(org.apache.poi.ss.usermodel.FillPatternType.SOLID_FOREGROUND, styleEvenRow.getFillPattern());

        Assertions.assertEquals("Row" + numRows + "Col3", assertSheet.cell(numRows, numCols).getTextValue(), "Value at (numRows,numCols) after reopen");

        // Verify data integrity for all cells after reopening
        for (int i = 1; i <= numRows; i++) {
            for (int j = 1; j <= numCols; j++) {
                String expectedValue = "Row" + i + "Col" + j;
                String actualValue = assertSheet.cell(i, j).getTextValue();
                Assertions.assertEquals(expectedValue, actualValue, 
                        "Cell value at row " + i + ", column " + j + " should match after reopen");
            }
        }
        
        assertApp.closeAllWorkBooks();
        // Clean up the large file
        try {
            java.nio.file.Files.deleteIfExists(java.nio.file.Paths.get(filePath));
        } catch (java.io.IOException e) {
            System.err.println("Could not delete test file: " + filePath + " : " + e.getMessage());
        }
        cleanupTestFile(filePath); // Ensure cleanup even if assertions fail mid-way (though JUnit structure is better for this)
    }

    /**
     * Test edge cases and additional error handling
     */
    @Test
    void fTestEdgeCases() throws IOException {
        String testFile = getTestFilePath("rowEdgeCases", "xlsx");
        ExcelApplication xlApp = new ExcelApplication();
        ExcelWorkBook xlbook = xlApp.createWorkBook(testFile);
        ExcelWorkSheet sheet = xlbook.addSheet("EdgeCases");

        // Test with very large column index
        try {
            String[] values = {"Value1"};
            sheet.row(1).setRowValues(values, 1000);
        } catch (Exception e) {
            Assertions.fail("Should not throw exception with large column index");
        }

        // Test with very large row index
        try {
            ExcelRow largeRow = sheet.row(10000);
            largeRow.setFillColor("Yellow");
        } catch (Exception e) {
            Assertions.fail("Should not throw exception with large row index");
        }

        // Test with empty string array
        try {
            String[] emptyArray = {""};
            sheet.row(2).setRowValues(emptyArray);
        } catch (Exception e) {
            Assertions.fail("Should not throw exception with empty string array");
        }

        // Test with null values in array
        try {
            String[] nullArray = {null, "Value", null};
            sheet.row(3).setRowValues(nullArray);
        } catch (Exception e) {
            Assertions.fail("Should not throw exception with null values in array");
        }

        // Test with invalid color
        try {
            sheet.row(4).setFillColor("NonExistentColor");
        } catch (Exception e) {
            Assertions.fail("Should not throw exception with invalid color");
        }

        sheet.saveWorkBook(); // Ensure sheet is saved before closing
        xlApp.closeAllWorkBooks();
        cleanupTestFile(testFile);
    }

    @Test
    void gTestRowFormattingCombinations() throws IOException {
        String xlsxFile = getTestFilePath("rowFormattingCombinations", "xlsx");
        testRowFormattingCombinations(xlsxFile);
        cleanupTestFile(xlsxFile);

        String xlsFile = getTestFilePath("rowFormattingCombinations", "xls");
        testRowFormattingCombinations(xlsFile);
        cleanupTestFile(xlsFile);
    }

    private void testRowFormattingCombinations(String strCompleteFileName) {
        ExcelApplication xlApp = new ExcelApplication();
        ExcelWorkBook xlbook = xlApp.createWorkBook(strCompleteFileName);
        ExcelWorkSheet sheet = xlbook.addSheet("FormattingCombinations");

        // Test various formatting combinations
        String[] values = {"Value1", "Value2", "Value3", "Value4", "Value5"};
        
        // Test font color combinations
        ExcelRow row1 = sheet.row(1);
        row1.setRowValues(values);
        row1.setFontColor("RED", 1, 3);
        row1.setFontColor("BLUE", 3, 6);
        row1.setFontColor("GREEN", 2, 4);

        // Test fill color combinations
        ExcelRow row2 = sheet.row(2);
        row2.setRowValues(values);
        row2.setFillColor("YELLOW", 1, 3);
        row2.setFillColor("LIGHT_BLUE", 3, 6);
        row2.setFillColor("LIGHT_GREEN", 2, 4);

        // Test border combinations
        ExcelRow row3 = sheet.row(3);
        row3.setRowValues(values);
        row3.setFullBorder("BLACK", 1, 3);
        row3.setFullBorder("RED", 3, 6);
        row3.setFullBorder("BLUE", 2, 4);

        // Test mixed formatting
        ExcelRow row4 = sheet.row(4);
        row4.setRowValues(values);
        row4.setFontColor("RED", 1, 3);
        row4.setFillColor("YELLOW", 1, 3);
        row4.setFullBorder("BLACK", 1, 3);
        row4.setFontColor("BLUE", 3, 6);
        row4.setFillColor("LIGHT_BLUE", 3, 6);
        row4.setFullBorder("RED", 3, 6);

        sheet.saveWorkBook();
        xlApp.closeAllWorkBooks();
    }

    @Test
    void hTestRowDataTypes() throws IOException {
        String xlsxFile = getTestFilePath("rowDataTypes", "xlsx");
        testRowDataTypes(xlsxFile);
        cleanupTestFile(xlsxFile);

        String xlsFile = getTestFilePath("rowDataTypes", "xls");
        testRowDataTypes(xlsFile);
        cleanupTestFile(xlsFile);
    }

    private void testRowDataTypes(String strCompleteFileName) {
        ExcelApplication xlApp = new ExcelApplication();
        ExcelWorkBook xlbook = xlApp.createWorkBook(strCompleteFileName);
        ExcelWorkSheet sheet = xlbook.addSheet("DataTypes");

        // Test different data types in a row
        String[] textValues = {"Text1", "Text2", "Text3"};
        String[] numericValues = {"123", "456.789", "0.001"};
        String[] dateValues = {"2024-01-01", "2024-02-29", "2024-12-31"};
        String[] mixedValues = {"Text", "123", "2024-01-01", "TRUE", "FALSE"};

        // Test text values
        ExcelRow textRow = sheet.row(1);
        textRow.setRowValues(textValues);

        // Test numeric values
        ExcelRow numericRow = sheet.row(2);
        numericRow.setRowValues(numericValues);

        // Test date values
        ExcelRow dateRow = sheet.row(3);
        dateRow.setRowValues(dateValues);

        // Test mixed values
        ExcelRow mixedRow = sheet.row(4);
        mixedRow.setRowValues(mixedValues);

        // Test special characters
        String[] specialChars = {"Line\nBreak", "Tab\tHere", "Quote\"Test", "Ampersand&Test"};
        ExcelRow specialRow = sheet.row(5);
        specialRow.setRowValues(specialChars);

        // Test empty values
        String[] emptyValues = {"", " ", "  ", "\t", "\n"};
        ExcelRow emptyRow = sheet.row(6);
        emptyRow.setRowValues(emptyValues);

        sheet.saveWorkBook();
        xlApp.closeAllWorkBooks();
    }

}
