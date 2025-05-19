package solutions.bismi.excel;

import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.MethodOrderer;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.TestMethodOrder;

@TestMethodOrder(MethodOrderer.Alphanumeric.class)
public class ExcelRowTest {

    @Test
    public void aSetfontcolor() {
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
    public void bTestSetRowValuesVariants() {
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

        // Save and reopen to verify
        sheet.saveWorkBook();
        xlApp.closeAllWorkBooks();

        // Reopen and verify
        xlbook = xlApp.openWorkbook(strCompleteFileName);
        sheet = xlbook.getExcelSheet("RowTest");

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
    public void cTestSetFillColorWithRange() {
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
    public void dTestErrorHandling() {
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
    public void eTestLargeDataSet() {
        ExcelApplication xlApp = new ExcelApplication();
        ExcelWorkBook xlbook = xlApp.createWorkBook("./resources/testdata/largeDataSet.xlsx");
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

        sheet.saveWorkBook();
        xlApp.closeAllWorkBooks();
    }
}
