package solutions.bismi.excel;

import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.MethodOrderer;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.TestMethodOrder;

@TestMethodOrder(MethodOrderer.MethodName.class)
class ExcelWorkSheetTest {

    @Test
    void aVerifySheetActivationXLSX() {

        verifySheetActivation("./resources/testdata/sheetactivation.xlsx");
        verifySheetActivation("./resources/testdata/sheetactivation.xls");
    }

    private void verifySheetActivation(String strCompleteFileName) {
        ExcelApplication xlApp = new ExcelApplication();
        ExcelWorkBook xlbook = xlApp.createWorkBook(strCompleteFileName);
        int cnt = xlbook.getSheetCount();
        Assertions.assertEquals(1, cnt);
        xlbook.addSheet("Bismi1");
        cnt = xlbook.getSheetCount();
        xlbook.getExcelSheet("Bismi1").activate();
        xlApp.closeAllWorkBooks();
        Assertions.assertEquals(2, cnt);
        xlbook = xlApp.openWorkbook(strCompleteFileName);
        cnt = xlApp.getOpenWorkbookCount();
        Assertions.assertEquals(1, cnt);
        cnt = xlbook.getSheetCount();
        Assertions.assertEquals(2, cnt);
        xlbook.getExcelSheet("Bismi1").activate();
        String sheetName = xlbook.getActiveSheetName();
        Assertions.assertEquals("Bismi1", sheetName);
        xlApp.closeAllWorkBooks();
    }


    @Test
    void cAddRowsXLSX() {

        addRowsXL("./resources/testdata/rowCreation.xlsx");
        addRowsXL("./resources/testdata/rowCreation.xls");
    }


    private void addRowsXL(String strCompleteFileName) {

        ExcelApplication xlApp = new ExcelApplication();
        ExcelWorkBook xlbook = xlApp.createWorkBook(strCompleteFileName);
        int cnt = xlbook.getSheetCount();
        Assertions.assertEquals(1, cnt);
        xlbook.addSheet("Bismi1");
        cnt = xlbook.getSheetCount();
        ExcelWorkSheet cSheet = xlbook.getExcelSheet("Bismi1");
        cSheet.activate();
        int cRownum = cSheet.getRowNumber();
        Assertions.assertEquals(0, cRownum);
        String[] arrRow = {"A", "B", "C", "D"};
        cSheet.row(11).setRowValues(arrRow);
        cRownum = cSheet.getRowNumber();
        Assertions.assertEquals(11, cRownum);

        int cColNumber = cSheet.getColNumber();
        Assertions.assertEquals(4, cColNumber);

        cSheet.row(16).setRowValues(arrRow, 5);
        cRownum = cSheet.getRowNumber();
        Assertions.assertEquals(16, cRownum);

        cColNumber = cSheet.getColNumber();
        Assertions.assertEquals(9, cColNumber);

        String[] arrRow1 = {"A", "10.5", "Cwwwwwwwwwwwwwwwwwwwwww", "D", "02/04/2019"};
        cSheet.row(34).setRowValues(arrRow1, true);
        cRownum = cSheet.getRowNumber();
        Assertions.assertEquals(34, cRownum);
        cColNumber = cSheet.getColNumber();
        Assertions.assertEquals(9, cColNumber);
        cSheet.saveWorkBook();
        xlApp.closeAllWorkBooks();
    }


    @Test
    void dAddRowsInMultipleSheet() {
        addRowsInMultipleSheet("./resources/testdata/multiRowMultiSheet.xlsx");
        addRowsInMultipleSheet("./resources/testdata/multiRowMultiSheet.xls");
    }

    private void addRowsInMultipleSheet(String strMultipleFileName) {
        ExcelApplication xlApp = new ExcelApplication();
        ExcelWorkBook xlbook = xlApp.createWorkBook(strMultipleFileName);
        int cnt = xlbook.getSheetCount();
        Assertions.assertTrue(cnt > 0, "Sheet count should be greater than zero");
        Assertions.assertEquals(1, cnt);
        xlbook.addSheet("Bismi1");
        xlbook.addSheet("Bismi2");
        cnt = xlbook.getSheetCount();
        ExcelWorkSheet cSheet = xlbook.getExcelSheet("Bismi1");

        ExcelWorkSheet cSheet2 = xlbook.getExcelSheet("Bismi2");
        cSheet2.activate();

        String[] arrRow1 = {"A", "10.5", "Cwwwwwwwwwwwwwwwwwwwwww", "D", "02/04/2019"};
        cSheet.row(34).setRowValues(arrRow1, true);
        int cRownum = cSheet.getRowNumber();
        Assertions.assertEquals(34, cRownum);
        int cColNumber = cSheet.getColNumber();
        Assertions.assertEquals(5, cColNumber);

        String[] arrRow2 = {"A", "10.5", "Cwwwwwwwwwwwwwwwwwwwwww", "D", "02/04/2019"};
        cSheet2.row(4).setRowValues(arrRow2, true);
        cRownum = cSheet2.getRowNumber();
        Assertions.assertEquals(4, cRownum);
        cColNumber = cSheet2.getColNumber();
        Assertions.assertEquals(5, cColNumber);
        cSheet2.saveWorkBook();
        xlApp.closeAllWorkBooks();
        xlbook = xlApp.openWorkbook(strMultipleFileName);
        ExcelWorkSheet gSheet = xlbook.getExcelSheet("Bismi2");
        String val = gSheet.cell(4, 2).getTextValue();
        Assertions.assertEquals("10.5", val);

        xlApp.closeAllWorkBooks();
    }


    @Test
    void eSetCellValues() {
        setCellValues("./resources/testdata/cellvalcheck.xlsx");
        setCellValues("./resources/testdata/cellvalcheck.xls");
    }

    private void setCellValues(String strCompleteFilePath) {
        ExcelApplication xlApp = new ExcelApplication();
        ExcelWorkBook xlbook = xlApp.createWorkBook(strCompleteFilePath);
        int cnt = xlbook.getSheetCount();
        Assertions.assertEquals(1, cnt);
        xlbook.addSheet("Bismi1");
        xlbook.addSheet("Bismi2");
        cnt = xlbook.getSheetCount();
        ExcelWorkSheet cSheet = xlbook.getExcelSheet("Bismi1");

        ExcelWorkSheet cSheet2 = xlbook.getExcelSheet("Bismi2");
        cSheet2.activate();
        cSheet.cell(5, 4).setText("Alpha lion");
        cSheet2.cell(3, 4).setText("Alpha lion");
        cSheet2.saveWorkBook();


        String val = cSheet2.cell(3, 4).getTextValue();
        Assertions.assertEquals("Alpha lion", val);

        xlApp.closeAllWorkBooks();
    }

    @Test
    void fTestCellMerging() {
        testCellMerging("./resources/testdata/cellmerge.xlsx");
        testCellMerging("./resources/testdata/cellmerge.xls");
    }

    private void testCellMerging(String strCompleteFilePath) {
        ExcelApplication xlApp = new ExcelApplication();
        ExcelWorkBook xlbook = xlApp.createWorkBook(strCompleteFilePath);

        // Add a sheet
        xlbook.addSheet("MergeTest");
        ExcelWorkSheet sheet = xlbook.getExcelSheet("MergeTest");

        // Add some data
        sheet.cell(1, 1).setText("Merged Cell Content");
        sheet.cell(1, 2).setText("Cell 1,2");
        sheet.cell(2, 1).setText("Cell 2,1");
        sheet.cell(2, 2).setText("Cell 2,2");

            // Test merge cells (row1, col1, row2, col2)
            boolean mergeResult = sheet.mergeCells(1, 1, 1, 2);
            Assertions.assertTrue(mergeResult, "Cell merge should succeed");

            // Test if cell is merged
            boolean isMerged = sheet.isCellMerged(1, 1);
            Assertions.assertTrue(isMerged, "Cell (1,1) should be merged");

            isMerged = sheet.isCellMerged(1, 2);
            Assertions.assertTrue(isMerged, "Cell (1,2) should be merged");

            isMerged = sheet.isCellMerged(2, 1);
            Assertions.assertFalse(isMerged, "Cell (2,1) should not be merged");

            // Test unmerge cells (row1, col1, row2, col2)
        boolean unmergeResult = sheet.unmergeCells(1, 1, 1, 2);
        Assertions.assertTrue(unmergeResult, "Cell unmerge should succeed");

        // Verify cells are no longer merged
        isMerged = sheet.isCellMerged(1, 1);
        Assertions.assertFalse(isMerged, "Cell (1,1) should no longer be merged");

        // Save and close
        sheet.saveWorkBook();
        xlApp.closeAllWorkBooks();
    }

    @Test
    void gTestClearContents() {
        testClearContents("./resources/testdata/clearcontent.xlsx");
        testClearContents("./resources/testdata/clearcontent.xls");
    }

    private void testClearContents(String strCompleteFilePath) {
        ExcelApplication xlApp = new ExcelApplication();
        ExcelWorkBook xlbook = xlApp.createWorkBook(strCompleteFilePath);

        // Add a sheet
        xlbook.addSheet("ClearTest");
        ExcelWorkSheet sheet = xlbook.getExcelSheet("ClearTest");

        // Add some data
        sheet.cell(1, 1).setText("Test Data 1");
        sheet.cell(2, 1).setText("Test Data 2");
        sheet.cell(3, 1).setText("Test Data 3");

        // Verify row count before clearing
        int rowCount = sheet.getRowNumber();
        Assertions.assertEquals(3, rowCount, "Sheet should have 3 rows before clearing");

        // Clear contents
        boolean clearResult = sheet.clearContents();
        Assertions.assertTrue(clearResult, "Clear contents should succeed");

        // Verify row count after clearing
        // Note: In some Excel implementations, clearing doesn't reset the row count until reopening
        // So reopen the workbook to check
        xlApp.closeAllWorkBooks();

        xlbook = xlApp.openWorkbook(strCompleteFilePath);
        sheet = xlbook.getExcelSheet("ClearTest");
        rowCount = sheet.getRowNumber();
        Assertions.assertEquals(0, rowCount, "Sheet should have 0 rows after clearing");

        xlApp.closeAllWorkBooks();
    }

    /**
     * Test for error handling in cell operations
     */
    @Test
    void hTestErrorHandling() {
        ExcelApplication xlApp = new ExcelApplication();
        ExcelWorkBook xlbook = xlApp.createWorkBook("./resources/testdata/errorhandling.xlsx");

        // Add a sheet
        xlbook.addSheet("ErrorTest");
        ExcelWorkSheet sheet = xlbook.getExcelSheet("ErrorTest");

        // Test unmerging cells that aren't merged
        boolean unmergeResult = sheet.unmergeCells(1, 1, 1, 2);
        Assertions.assertFalse(unmergeResult, "Unmerging non-merged cells should return false");

        // Test merging with invalid indices
        boolean mergeResult = sheet.mergeCells(-1, 1, 1, 2);
        // This should not throw an exception, but may return false
        Assertions.assertFalse(mergeResult, "Merging cells with invalid indices should return false");

        // Test isCellMerged with invalid indices
        boolean isMerged = sheet.isCellMerged(-1, -1);
        Assertions.assertFalse(isMerged, "Invalid cell indices should not be reported as merged");

        xlApp.closeAllWorkBooks();
    }
}
