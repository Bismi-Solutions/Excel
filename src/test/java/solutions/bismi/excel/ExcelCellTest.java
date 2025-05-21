package solutions.bismi.excel;

import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.MethodOrderer;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.TestMethodOrder;


@TestMethodOrder(MethodOrderer.MethodName.class)
class ExcelCellTest {


    @Test
    void aSetFontColor() {
        setColor("./resources/testdata/cellFormatCheck1.XLSX");
        setColor("./resources/testdata/cellFormatCheck1.XLS");
    }

    private void setColor(String strCompleteFileName) {
        ExcelApplication xlApp = new ExcelApplication();
        ExcelWorkBook xlbook = xlApp.createWorkBook(strCompleteFileName);
        int cnt = 0;
        cnt = xlbook.getSheetCount();
        Assertions.assertEquals(1, cnt);
        ExcelWorkSheet sh1 = xlbook.addSheet("Bismi1");
        sh1.activate();
        ExcelCell c1 = sh1.cell(3, 4);
        c1.setText("Alpha lion");
        c1.setFontColor("blue");
        ExcelCell c2 = sh1.cell(5, 7);
        c2.setText("Beta");
        c2.setFontColor("RED");

        // Test with hex color codes
        ExcelCell c3 = sh1.cell(7, 4);
        c3.setText("Hex Color Test");
        c3.setFontColor("#FF0000"); // Red in hex

        ExcelCell c4 = sh1.cell(8, 4);
        c4.setText("Another Hex Color");
        c4.setFontColor("#0000FF"); // Blue in hex

        sh1.saveWorkBook();
        xlApp.closeAllWorkBooks();
    }

    @Test
    void bSetFillcolor() {
        setColor2("./resources/testdata/cellFormatCheck2.XLSX");
        setColor2("./resources/testdata/cellFormatCheck2.XLS");
    }

    private void setColor2(String strCompleteFileName) {
        ExcelApplication xlApp = new ExcelApplication();
        ExcelWorkBook xlbook = xlApp.createWorkBook(strCompleteFileName);
        int cnt = 0;
        cnt = xlbook.getSheetCount();
        Assertions.assertEquals(1, cnt);
        ExcelWorkSheet sh1 = xlbook.addSheet("Bismi1");
        sh1.activate();

        // Test with named colors
        sh1.cell(10, 10).setText("TestColor");
        sh1.cell(10, 10).setFontColor("blue");
        sh1.cell(10, 10).setFillColor("yellow");
        sh1.cell(1, 1).setFillColor("GREEN");
        sh1.cell(3, 17).setFullBorder("Red");

        // Test with hex color codes
        sh1.cell(5, 5).setText("Hex Fill Color");
        sh1.cell(5, 5).setFillColor("#00FF00"); // Green in hex

        sh1.cell(6, 6).setText("Hex Border Color");
        sh1.cell(6, 6).setFullBorder("#0000FF"); // Blue in hex

        sh1.cell(7, 7).setText("Hex Font and Fill");
        sh1.cell(7, 7).setFontColor("#FF0000"); // Red in hex
        sh1.cell(7, 7).setFillColor("#FFFF00"); // Yellow in hex

        // Other cell operations
        sh1.cell(12, 13).setCellValue("123");
        sh1.cell(12, 13).setCellValue("0");
        sh1.cell(13, 13).setText("0");
        sh1.cell(12, 22).setCellValue("0.123", true);

        sh1.saveWorkBook();
        xlApp.closeAllWorkBooks();
    }

    @Test
    void cTestFormulas() {
        testFormulas("./resources/testdata/formulaTest.XLSX");
        testFormulas("./resources/testdata/formulaTest.XLS");
    }

    private void testFormulas(String strCompleteFileName) {
        ExcelApplication xlApp = new ExcelApplication();
        ExcelWorkBook xlbook = xlApp.createWorkBook(strCompleteFileName);
        ExcelWorkSheet sh1 = xlbook.addSheet("FormulaTest");
        sh1.activate();

        // Set up some test data
        sh1.cell(1, 1).setCellValue("Item");
        sh1.cell(1, 2).setCellValue("Quantity");
        sh1.cell(1, 3).setCellValue("Price");
        sh1.cell(1, 4).setCellValue("Total");

        // Format the header row
        for (int i = 1; i <= 4; i++) {
            sh1.cell(1, i).setFontStyle(true, false, false);
            sh1.cell(1, i).setFillColor("GREY_25_PERCENT");
        }

        // Add some data rows
        sh1.cell(2, 1).setCellValue("Apples");
        sh1.cell(2, 2).setNumericValue(10);
        sh1.cell(2, 3).setNumericValue(1.5);

        sh1.cell(3, 1).setCellValue("Oranges");
        sh1.cell(3, 2).setNumericValue(5);
        sh1.cell(3, 3).setNumericValue(2.0);

        sh1.cell(4, 1).setCellValue("Bananas");
        sh1.cell(4, 2).setNumericValue(8);
        sh1.cell(4, 3).setNumericValue(1.2);

        // Add formulas for calculating totals
        sh1.cell(2, 4).setFormula("B2*C2");
        sh1.cell(3, 4).setFormula("B3*C3");
        sh1.cell(4, 4).setFormula("B4*C4");

        // Add a sum formula
        sh1.cell(5, 1).setCellValue("Total");
        sh1.cell(5, 1).setFontStyle(true, false, false);
        sh1.cell(5, 4).setFormula("SUM(D2:D4)");

        // Format the total cell
        sh1.cell(5, 4).setFontStyle(true, false, false);
        sh1.cell(5, 4).setFillColor("LIGHT_YELLOW");

        // Format the number cells
        for (int row = 2; row <= 5; row++) {
            for (int col = 2; col <= 4; col++) {
                if (col == 3 || col == 4) {
                    sh1.cell(row, col).setNumberFormat("$#,##0.00");
                }
            }
        }

        // Set column alignment
        for (int row = 2; row <= 5; row++) {
            sh1.cell(row, 1).setHorizontalAlignment("LEFT");
            sh1.cell(row, 2).setHorizontalAlignment("CENTER");
            sh1.cell(row, 3).setHorizontalAlignment("RIGHT");
            sh1.cell(row, 4).setHorizontalAlignment("RIGHT");
        }

        sh1.saveWorkBook();
        xlApp.closeAllWorkBooks();
    }

    @Test
    void dTestCellMerging() {
        testCellMerging("./resources/testdata/mergeTest.XLSX");
        testCellMerging("./resources/testdata/mergeTest.XLS");
    }

    private void testCellMerging(String strCompleteFileName) {
        ExcelApplication xlApp = new ExcelApplication();
        ExcelWorkBook xlbook = xlApp.createWorkBook(strCompleteFileName);
        ExcelWorkSheet sh1 = xlbook.addSheet("MergeTest");
        sh1.activate();

        // Create a title and merge cells
        sh1.cell(1, 1).setCellValue("Sales Report");
        sh1.cell(1, 1).setFontStyle(true, false, false);
        sh1.cell(1, 1).setHorizontalAlignment("CENTER");

            // Merge cells for the title (row1, col1, row2, col2)
            sh1.mergeCells(1, 1, 1, 5);

            // Add a subtitle
            sh1.cell(2, 1).setCellValue("Q1 2023");
            sh1.cell(2, 1).setFontStyle(false, true, false);
            sh1.cell(2, 1).setHorizontalAlignment("CENTER");

            // Merge cells for the subtitle (row1, col1, row2, col2)
            sh1.mergeCells(2, 1, 2, 5);

            // Verify that cells are merged
            boolean isMerged = sh1.isCellMerged(1, 3);
            Assertions.assertTrue(isMerged, "Cell (1,3) should be part of a merged region");

            // Test unmerging (row1, col1, row2, col2)
            sh1.unmergeCells(2, 1, 2, 5);
        isMerged = sh1.isCellMerged(2, 3);
        Assertions.assertFalse(isMerged, "Cell (2,3) should not be merged after unmerging");

        sh1.saveWorkBook();
        xlApp.closeAllWorkBooks();
    }


    @Test
    void eTestDateAndVerticalAlignment() {
        testDateAndVerticalAlignment("./resources/testdata/dateAlignTest.XLSX");
        testDateAndVerticalAlignment("./resources/testdata/dateAlignTest.XLS");
    }

    private void testDateAndVerticalAlignment(String strCompleteFileName) {
        ExcelApplication xlApp = new ExcelApplication();
        ExcelWorkBook xlbook = xlApp.createWorkBook(strCompleteFileName);
        ExcelWorkSheet sh1 = xlbook.addSheet("DateAlignTest");
        sh1.activate();

        // Test date value
        java.util.Date today = new java.util.Date();
        sh1.cell(1, 1).setCellValue("Date:");
        sh1.cell(1, 2).setDateValue(today);
        sh1.cell(1, 2).setNumberFormat("dd/mm/yyyy");

        // Test date with auto-size
        sh1.cell(2, 1).setCellValue("Date with auto-size:");
        sh1.cell(2, 2).setDateValue(today, true);
        sh1.cell(2, 2).setNumberFormat("yyyy-mm-dd");

        // Test vertical alignment
        sh1.cell(4, 1).setCellValue("Top aligned");
        sh1.cell(4, 1).setVerticalAlignment("TOP");

        sh1.cell(5, 1).setCellValue("Center aligned");
        sh1.cell(5, 1).setVerticalAlignment("CENTER");

        sh1.cell(6, 1).setCellValue("Bottom aligned");
        sh1.cell(6, 1).setVerticalAlignment("BOTTOM");

        sh1.cell(7, 1).setCellValue("Justified");
        sh1.cell(7, 1).setVerticalAlignment("JUSTIFY");

        sh1.cell(8, 1).setCellValue("Distributed");
        sh1.cell(8, 1).setVerticalAlignment("DISTRIBUTED");

        sh1.cell(9, 1).setCellValue("Default alignment");
        sh1.cell(9, 1).setVerticalAlignment("INVALID"); // Should default to CENTER

        sh1.saveWorkBook();
        xlApp.closeAllWorkBooks();
    }

    @Test
    void fTestErrorHandling() {
        ExcelApplication xlApp = new ExcelApplication();
        ExcelWorkBook xlbook = xlApp.createWorkBook("./resources/testdata/errorHandlingTest.XLSX");
        ExcelWorkSheet sh1 = xlbook.addSheet("ErrorTest");
        sh1.activate();

        // Test with invalid formula (should not throw exception)
        try {
            sh1.cell(1, 1).setFormula("INVALID_FORMULA(A1)");
        } catch (Exception e) {
            Assertions.fail("Should not throw exception with invalid formula");
        }

        // Test with invalid color name (should not throw exception)
        try {
            sh1.cell(2, 1).setFontColor("NONEXISTENT_COLOR");
        } catch (Exception e) {
            Assertions.fail("Should not throw exception with invalid color name");
        }

        // Test getValue on empty cell
        String value = sh1.cell(3, 3).getValue();
        Assertions.assertEquals("", value, "Empty cell should return empty string");

        sh1.saveWorkBook();
        xlApp.closeAllWorkBooks();
    }

    @Test
    void fTestEdgeCasesAndErrorConditions() {
        testEdgeCases("./resources/testdata/edgeCasesTest.XLSX");
        testEdgeCases("./resources/testdata/edgeCasesTest.XLS");
    }

    private void testEdgeCases(String strCompleteFileName) {
        ExcelApplication xlApp = new ExcelApplication();
        ExcelWorkBook xlbook = xlApp.createWorkBook(strCompleteFileName);
        ExcelWorkSheet sh1 = xlbook.addSheet("EdgeCases");
        sh1.activate();

        // Test null values
        ExcelCell nullCell = sh1.cell(1, 1);
        Assertions.assertTrue(nullCell.setText(null), "Should handle null text value");
        Assertions.assertEquals("", nullCell.getTextValue(), "Null text should be converted to empty string");
        sh1.saveWorkBook();
        // Test empty string
        ExcelCell emptyCell = sh1.cell(1, 2);
        Assertions.assertTrue(emptyCell.setText(""), "Should handle empty string");
        Assertions.assertEquals("", emptyCell.getTextValue(), "Empty string should be preserved");
        sh1.saveWorkBook();
        // Test very long text
        String longText = "x".repeat(32767); // Maximum cell text length
        ExcelCell longTextCell = sh1.cell(1, 3);
        Assertions.assertTrue(longTextCell.setText(longText), "Should handle maximum length text");
        Assertions.assertEquals(longText, longTextCell.getTextValue(), "Long text should be preserved");
        sh1.saveWorkBook();
        // Test invalid numeric values
        ExcelCell invalidNumericCell = sh1.cell(1, 4);
        Assertions.assertTrue(invalidNumericCell.setNumericValue(Double.NaN), "Should handle NaN");
        Assertions.assertTrue(invalidNumericCell.setNumericValue(Double.POSITIVE_INFINITY), "Should handle infinity");
        Assertions.assertTrue(invalidNumericCell.setNumericValue(Double.NEGATIVE_INFINITY), "Should handle negative infinity");
        sh1.saveWorkBook();
        // Test invalid date values
        ExcelCell invalidDateCell = sh1.cell(1, 5);
        Assertions.assertTrue(invalidDateCell.setDateValue(null), "Should handle null date");
        sh1.saveWorkBook();
        // Test invalid formula
        ExcelCell invalidFormulaCell = sh1.cell(1, 6);
        Assertions.assertTrue(invalidFormulaCell.setFormula("=INVALID_FORMULA()"), "Should handle invalid formula");
        sh1.saveWorkBook();
        // Test invalid color values
        ExcelCell invalidColorCell = sh1.cell(1, 7);
        invalidColorCell.setFontColor("INVALID_COLOR");
        invalidColorCell.setFillColor("INVALID_COLOR");
        sh1.saveWorkBook();
        // Should not throw exception

        // Test invalid alignment values
        ExcelCell invalidAlignCell = sh1.cell(1, 8);
        Assertions.assertTrue(invalidAlignCell.setHorizontalAlignment("INVALID_ALIGNMENT"), "Should handle invalid horizontal alignment");
        Assertions.assertTrue(invalidAlignCell.setVerticalAlignment("INVALID_ALIGNMENT"), "Should handle invalid vertical alignment");
        sh1.saveWorkBook();
        // Test invalid number format
        ExcelCell invalidFormatCell = sh1.cell(1, 9);
        Assertions.assertTrue(invalidFormatCell.setNumberFormat("INVALID_FORMAT"), "Should handle invalid number format");
        sh1.saveWorkBook();
        sh1.saveWorkBook();
        xlApp.closeAllWorkBooks();
    }

    @Test
    void gTestCellFormattingCombinations() {
        testFormattingCombinations("./resources/testdata/formattingCombinations.XLSX");
        testFormattingCombinations("./resources/testdata/formattingCombinations.XLS");
    }

    private void testFormattingCombinations(String strCompleteFileName) {
        ExcelApplication xlApp = new ExcelApplication();
        ExcelWorkBook xlbook = xlApp.createWorkBook(strCompleteFileName);
        ExcelWorkSheet sh1 = xlbook.addSheet("FormattingCombinations");
        sh1.activate();

        // Test various formatting combinations
        ExcelCell cell1 = sh1.cell(1, 1);
        cell1.setText("Bold Italic Underline");
        cell1.setFontStyle(true, true, true);
        cell1.setFontColor("RED");
        cell1.setFillColor("YELLOW");
        cell1.setHorizontalAlignment("CENTER");
        cell1.setVerticalAlignment("CENTER");
        cell1.setFullBorder("BLUE");

        // Test number formatting with different locales
        ExcelCell cell2 = sh1.cell(2, 1);
        cell2.setNumericValue(1234.56);
        cell2.setNumberFormat("#,##0.00");
        cell2.setFontStyle(true, false, false);
        cell2.setHorizontalAlignment("RIGHT");

        // Test date formatting with different patterns
        ExcelCell cell3 = sh1.cell(3, 1);
        cell3.setDateValue(new java.util.Date());
        cell3.setNumberFormat("yyyy-mm-dd");
        cell3.setFontStyle(false, true, false);
        cell3.setHorizontalAlignment("CENTER");

        // Test currency formatting
        ExcelCell cell4 = sh1.cell(4, 1);
        cell4.setNumericValue(1234.56);
        cell4.setNumberFormat("$#,##0.00");
        cell4.setFontStyle(true, true, false);
        cell4.setHorizontalAlignment("RIGHT");

        // Test percentage formatting
        ExcelCell cell5 = sh1.cell(5, 1);
        cell5.setNumericValue(0.1234);
        cell5.setNumberFormat("0.00%");
        cell5.setFontStyle(false, false, true);
        cell5.setHorizontalAlignment("RIGHT");

        // Test scientific notation
        ExcelCell cell6 = sh1.cell(6, 1);
        cell6.setNumericValue(1234567.89);
        cell6.setNumberFormat("0.00E+00");
        cell6.setFontStyle(true, true, true);
        cell6.setHorizontalAlignment("RIGHT");

        sh1.saveWorkBook();
        xlApp.closeAllWorkBooks();
    }

    @Test
    void hTestCellValueRetrieval() {
        testValueRetrieval("./resources/testdata/valueRetrieval.XLSX");
        testValueRetrieval("./resources/testdata/valueRetrieval.XLS");
    }

    private void testValueRetrieval(String strCompleteFileName) {
        ExcelApplication xlApp = new ExcelApplication();
        ExcelWorkBook xlbook = xlApp.createWorkBook(strCompleteFileName);
        ExcelWorkSheet sh1 = xlbook.addSheet("ValueRetrieval");
        sh1.activate();

        // Test text value retrieval
        ExcelCell textCell = sh1.cell(1, 1);
        textCell.setText("Test Text");
        Assertions.assertEquals("Test Text", textCell.getTextValue(), "Text value should be retrieved correctly");

        // Test numeric value retrieval
        ExcelCell numericCell = sh1.cell(2, 1);
        numericCell.setNumericValue(123.45);
        Assertions.assertEquals("123.45", numericCell.getValue(), "Numeric value should be retrieved correctly");

        // Test date value retrieval
        ExcelCell dateCell = sh1.cell(3, 1);
        java.util.Date testDate = new java.util.Date();
        dateCell.setDateValue(testDate);
        Assertions.assertNotNull(dateCell.getValue(), "Date value should be retrieved correctly");

        // Test formula value retrieval
        ExcelCell formulaCell = sh1.cell(4, 1);
        formulaCell.setFormula("SUM(A1:A3)");
        Assertions.assertEquals("SUM(A1:A3)", formulaCell.getValue(), "Formula value should be retrieved correctly");

        // Test empty cell value retrieval
        ExcelCell emptyCell = sh1.cell(5, 1);
        Assertions.assertEquals("", emptyCell.getValue(), "Empty cell should return empty string");

        // Test cell with special characters
        ExcelCell specialCharCell = sh1.cell(6, 1);
        specialCharCell.setText("Special\nChars\tHere");
        Assertions.assertEquals("Special\nChars\tHere", specialCharCell.getTextValue(), "Special characters should be preserved");

        sh1.saveWorkBook();
        xlApp.closeAllWorkBooks();
    }

    @Test
    void iTestRowOperations() {
        testRowOperations("./resources/testdata/rowOperations.XLSX");
        testRowOperations("./resources/testdata/rowOperations.XLS");
    }

    private void testRowOperations(String strCompleteFileName) {
        ExcelApplication xlApp = new ExcelApplication();
        ExcelWorkBook xlbook = xlApp.createWorkBook(strCompleteFileName);
        ExcelWorkSheet sh1 = xlbook.addSheet("RowOperations");
        sh1.activate();

        // Test setting multiple values in a row
        String[] values = {"Value1", "Value2", "Value3", "Value4"};
        ExcelRow row1 = sh1.row(1);
        row1.setRowValues(values);
        row1.setRowValues(values, 2); // Start from column 2
        row1.setRowValues(values, true); // With auto-size
        row1.setRowValues(values, 2, true); // Start from column 2 with auto-size

        // Test row formatting
        ExcelRow row2 = sh1.row(2);
        row2.setRowValues(values);
        row2.setFontColor("RED");
        row2.setFillColor("YELLOW");
        row2.setFullBorder("BLUE");

        // Test partial row formatting
        ExcelRow row3 = sh1.row(3);
        row3.setRowValues(values);
        row3.setFontColor("GREEN", 2, 4); // Format columns 2-3
        row3.setFillColor("LIGHT_BLUE", 1, 3); // Format columns 1-2
        row3.setFullBorder("BLACK", 3, 5); // Format columns 3-4

        // Test empty row operations
        ExcelRow emptyRow = sh1.row(4);
        emptyRow.setRowValues(new String[0]);
        emptyRow.setFontColor("RED");
        emptyRow.setFillColor("YELLOW");
        emptyRow.setFullBorder("BLUE");

        // Test null values in row
        ExcelRow nullRow = sh1.row(5);
        String[] nullValues = {null, "Value", null, "Value"};
        nullRow.setRowValues(nullValues);

        sh1.saveWorkBook();
        xlApp.closeAllWorkBooks();
    }

    @Test
    void jTestSheetOperations() {
        testSheetOperations("./resources/testdata/sheetOperations.XLSX");
        testSheetOperations("./resources/testdata/sheetOperations.XLS");
    }

    private void testSheetOperations(String strCompleteFileName) {
        ExcelApplication xlApp = new ExcelApplication();
        ExcelWorkBook xlbook = xlApp.createWorkBook(strCompleteFileName);
        
        // Test adding multiple sheets
        String[] sheetNames = {"Sheet1", "Sheet2", "Sheet3"};
        ExcelWorkSheet sh1 = xlbook.addSheet("MainSheet");
        sh1.activate();
        xlbook.addSheets(sheetNames);

        // Test sheet activation
        ExcelWorkSheet sh2 = xlbook.addSheet("TestSheet");
        sh2.activate();
        Assertions.assertEquals("TestSheet", sh2.getSheetName(), "Sheet name should match");

        // Test sheet content operations
        sh2.cell(1, 1).setText("Test Content");
        Assertions.assertEquals(1, sh2.getRowNumber(), "Should have 1 row");
        Assertions.assertEquals(1, sh2.getColNumber(), "Should have 1 column");

        // Test clearing sheet contents
        sh2.cell(2, 1).setText("Content to be cleared");
        sh2.cell(2, 2).setText("More content");
        Assertions.assertTrue(sh2.clearContents(), "Should clear contents successfully");
        Assertions.assertEquals(0, sh2.getRowNumber(), "Should have 0 rows after clearing");

        // Test cell operations after clearing
        sh2.cell(1, 1).setText("New Content");
        Assertions.assertEquals(1, sh2.getRowNumber(), "Should have 1 row after adding new content");

        // Test row operations
        ExcelRow row = sh2.row(2);
        row.setRowValues(new String[]{"Row", "Content", "Test"});
        Assertions.assertEquals(2, sh2.getRowNumber(), "Should have 2 rows");

        sh2.saveWorkBook();
        xlApp.closeAllWorkBooks();
    }

    @Test
    void kTestAdvancedFormatting() {
        testAdvancedFormatting("./resources/testdata/advancedFormatting.XLSX");
        testAdvancedFormatting("./resources/testdata/advancedFormatting.XLS");
    }

    private void testAdvancedFormatting(String strCompleteFileName) {
        ExcelApplication xlApp = new ExcelApplication();
        ExcelWorkBook xlbook = xlApp.createWorkBook(strCompleteFileName);
        ExcelWorkSheet sh1 = xlbook.addSheet("AdvancedFormatting");
        sh1.activate();

        // Test complex number formats
        ExcelCell cell1 = sh1.cell(1, 1);
        cell1.setNumericValue(1234.5678);
        cell1.setNumberFormat("#,##0.00_);[Red](#,##0.00)"); // Positive in black, negative in red

        // Test date formats with time
        ExcelCell cell2 = sh1.cell(2, 1);
        cell2.setDateValue(new java.util.Date());
        cell2.setNumberFormat("yyyy-mm-dd hh:mm:ss");

        // Test currency with different symbols
        ExcelCell cell3 = sh1.cell(3, 1);
        cell3.setNumericValue(1234.56);
        cell3.setNumberFormat("€#,##0.00");

        // Test percentage with different decimal places
        ExcelCell cell4 = sh1.cell(4, 1);
        cell4.setNumericValue(0.1234);
        cell4.setNumberFormat("0.000%");

        // Test scientific notation with different formats
        ExcelCell cell5 = sh1.cell(5, 1);
        cell5.setNumericValue(1234567.89);
        cell5.setNumberFormat("0.00E+00");

        // Test custom text format
        ExcelCell cell6 = sh1.cell(6, 1);
        cell6.setText("Test");
        cell6.setNumberFormat("@"); // Text format

        // Test mixed format
        ExcelCell cell7 = sh1.cell(7, 1);
        cell7.setNumericValue(1234.56);
        cell7.setNumberFormat("#,##0.00_);[Red](#,##0.00);\"Zero\""); // Positive, negative, and zero formats

        sh1.saveWorkBook();
        xlApp.closeAllWorkBooks();
    }

    @Test
    void lTestErrorRecovery() {
        testErrorRecovery("./resources/testdata/errorRecovery.XLSX");
        testErrorRecovery("./resources/testdata/errorRecovery.XLS");
    }

    private void testErrorRecovery(String strCompleteFileName) {
        ExcelApplication xlApp = new ExcelApplication();
        ExcelWorkBook xlbook = xlApp.createWorkBook(strCompleteFileName);
        ExcelWorkSheet sh1 = xlbook.addSheet("ErrorRecovery");
        sh1.activate();

        // Test recovery from invalid cell operations
        ExcelCell cell1 = sh1.cell(1, 1);
        cell1.setText("Original Text");
        
        // Test invalid operations that should not crash
        cell1.setFontColor("INVALID_COLOR");
        cell1.setFillColor("INVALID_COLOR");
        cell1.setHorizontalAlignment("INVALID_ALIGNMENT");
        cell1.setVerticalAlignment("INVALID_ALIGNMENT");
        cell1.setNumberFormat("INVALID_FORMAT");
        
        // Verify cell still has original content
        Assertions.assertEquals("Original Text", cell1.getTextValue(), "Cell should retain original content after invalid operations");

        // Test recovery from invalid row operations
        ExcelRow row1 = sh1.row(2);
        row1.setRowValues(new String[]{"Row", "Content"});
        
        // Test invalid row formatting
        row1.setFontColor("INVALID_COLOR");
        row1.setFillColor("INVALID_COLOR");
        row1.setFullBorder("INVALID_COLOR");
        
        // Verify row still has content
        Assertions.assertEquals("Row", sh1.cell(2, 1).getTextValue(), "Row should retain content after invalid operations");

        // Test recovery from invalid sheet operations
        sh1.cell(3, 1).setText("Sheet Content");
        
        // Test invalid sheet operations
        sh1.mergeCells(3, 1, 3, 2); // Should not crash
        sh1.unmergeCells(3, 1, 3, 2); // Should not crash
        
        // Verify sheet still has content
        Assertions.assertEquals("Sheet Content", sh1.cell(3, 1).getTextValue(), "Sheet should retain content after invalid operations");

        sh1.saveWorkBook();
        xlApp.closeAllWorkBooks();
    }
}
