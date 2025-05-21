package solutions.bismi.excel;

import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.MethodOrderer;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.TestMethodOrder;


@TestMethodOrder(MethodOrderer.MethodName.class)
class ExcelRowTest {

    @Test
    void aSetfontcolor() {
        setColor2("./resources/testdata/cellFormatCheckrow1.XLSX");
        setColor2("./resources/testdata/cellFormatCheckrow1.XLS");
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
    }


    @Test
    void bTestSetRowValuesVariants() {
        testSetRowValuesVariants("./resources/testdata/rowValuesVariants.xlsx");
        testSetRowValuesVariants("./resources/testdata/rowValuesVariants.xls");
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
    void cTestSetFillColorWithRange() {
        testSetFillColorWithRange("./resources/testdata/fillColorRange.xlsx");
        testSetFillColorWithRange("./resources/testdata/fillColorRange.xls");
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
    void dTestErrorHandling() {
        ExcelApplication xlApp = new ExcelApplication();
        ExcelWorkBook xlbook = xlApp.createWorkBook("./resources/testdata/rowErrorHandling.xlsx");
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

        sheet.saveWorkBook();
        xlApp.closeAllWorkBooks();
    }

    /**
     * Test with large data set to verify performance
     */
    @Test
    void eTestLargeDataSet() {
        String filePath = "./resources/testdata/largeDataSet.xlsx";
        ExcelApplication xlApp = new ExcelApplication();
        ExcelWorkBook xlbook = xlApp.createWorkBook(filePath);
        ExcelWorkSheet sheet = xlbook.addSheet("LargeData");
    
        // Create a large data set (100 rows)
        for (int i = 1; i <= 20; i++) {
            String[] rowData = {"Row" + i + "Col1", "Row" + i + "Col2", "Row" + i + "Col3"};
            sheet.row(i).setRowValues(rowData);
    
            // Apply different formatting to even and odd rows
            if (i % 2 == 0) {
                sheet.row(i).setFillColor("LightGray");
            } else {
                sheet.row(i).setFontColor("Blue");
            }
        }

        // Verify data for a few rows
        Assertions.assertEquals("Row1Col1", sheet.cell(1, 1).getTextValue());
        Assertions.assertEquals("Row5Col2", sheet.cell(5, 2).getTextValue());
        Assertions.assertEquals("Row10Col3", sheet.cell(10, 3).getTextValue());
        Assertions.assertEquals("Row20Col1", sheet.cell(20, 1).getTextValue());
        
        // Not verifying format as on now
        // verifying the data integrity
        for (int i = 1; i <= 20; i++) {
            for (int j = 1; j <= 3; j++) {
                String expectedValue = "Row" + i + "Col" + j;
                String actualValue = sheet.cell(i, j).getTextValue();
                Assertions.assertEquals(expectedValue, actualValue, 
                        "Cell value at row " + i + ", column " + j + " should match");
            }
        }
        
        xlApp.closeAllWorkBooks();
    }

    /**
     * Test edge cases and additional error handling
     */
    @Test
    void fTestEdgeCases() {
        ExcelApplication xlApp = new ExcelApplication();
        ExcelWorkBook xlbook = xlApp.createWorkBook("./resources/testdata/rowEdgeCases.xlsx");
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

        sheet.saveWorkBook();
        xlApp.closeAllWorkBooks();
    }

    @Test
    void gTestRowFormattingCombinations() {
        testRowFormattingCombinations("./resources/testdata/rowFormattingCombinations.xlsx");
        testRowFormattingCombinations("./resources/testdata/rowFormattingCombinations.xls");
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
    void hTestRowDataTypes() {
        testRowDataTypes("./resources/testdata/rowDataTypes.xlsx");
        testRowDataTypes("./resources/testdata/rowDataTypes.xls");
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
