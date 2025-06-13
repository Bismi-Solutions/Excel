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
class ExcelWorkSheetTest {

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
    void aVerifySheetActivationXLSX() throws IOException {
        String xlsxFile = getTestFilePath("sheetActivation", "xlsx");
        verifySheetActivation(xlsxFile);
        cleanupTestFile(xlsxFile);

        String xlsFile = getTestFilePath("sheetActivation", "xls");
        verifySheetActivation(xlsFile);
        cleanupTestFile(xlsFile);
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
    void cAddRowsXLSX() throws IOException {
        String xlsxFile = getTestFilePath("rowCreation", "xlsx");
        addRowsXL(xlsxFile);
        cleanupTestFile(xlsxFile);

        String xlsFile = getTestFilePath("rowCreation", "xls");
        addRowsXL(xlsFile);
        cleanupTestFile(xlsFile);
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
    void dAddRowsInMultipleSheet() throws IOException {
        String xlsxFile = getTestFilePath("multiRowMultiSheet", "xlsx");
        addRowsInMultipleSheet(xlsxFile);
        cleanupTestFile(xlsxFile);

        String xlsFile = getTestFilePath("multiRowMultiSheet", "xls");
        addRowsInMultipleSheet(xlsFile);
        cleanupTestFile(xlsFile);
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


    }


    @Test
    void eSetCellValues() throws IOException {
        String xlsxFile = getTestFilePath("cellValCheck", "xlsx");
        setCellValues(xlsxFile);
        cleanupTestFile(xlsxFile);

        String xlsFile = getTestFilePath("cellValCheck", "xls");
        setCellValues(xlsFile);
        cleanupTestFile(xlsFile);
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
        Assertions.assertEquals("Alpha lion", val, "Value should be set correctly before save-close-reopen");

        xlApp.closeAllWorkBooks();

        // Reopen and Assert
        ExcelApplication assertApp = new ExcelApplication();
        ExcelWorkBook assertBook = assertApp.openWorkbook(strCompleteFilePath);
        Assertions.assertNotNull(assertBook, "Failed to reopen workbook for setCellValues assertions.");

        ExcelWorkSheet assertCSheet = assertBook.getExcelSheet("Bismi1");
        Assertions.assertNotNull(assertCSheet, "Failed to get sheet Bismi1 for assertions.");
        Assertions.assertEquals("Alpha lion", assertCSheet.cell(5, 4).getTextValue(), "Value in Bismi1 after reopen should match.");

        ExcelWorkSheet assertCSheet2 = assertBook.getExcelSheet("Bismi2");
        Assertions.assertNotNull(assertCSheet2, "Failed to get sheet Bismi2 for assertions.");
        Assertions.assertEquals("Alpha lion", assertCSheet2.cell(3, 4).getTextValue(), "Value in Bismi2 after reopen should match.");

        assertApp.closeAllWorkBooks();
    }

    @Test
    void fTestCellMerging() throws IOException {
        String xlsxFile = getTestFilePath("cellMerge", "xlsx");
        testCellMerging(xlsxFile);
        cleanupTestFile(xlsxFile);

        String xlsFile = getTestFilePath("cellMerge", "xls");
        testCellMerging(xlsFile);
        cleanupTestFile(xlsFile);
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
    void gTestClearContents() throws IOException {
        String xlsxFile = getTestFilePath("clearContent", "xlsx");
        testClearContents(xlsxFile);
        cleanupTestFile(xlsxFile);

        String xlsFile = getTestFilePath("clearContent", "xls");
        testClearContents(xlsFile);
        cleanupTestFile(xlsFile);
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
    void hTestErrorHandling() throws IOException {
        String testFile = getTestFilePath("worksheetErrorHandling", "xlsx");
        ExcelApplication xlApp = new ExcelApplication();
        ExcelWorkBook xlbook = xlApp.createWorkBook(testFile);

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
        cleanupTestFile(testFile);
    }

    @Test
    void zCoverageGettersSettersAndConstructors() {
        ExcelWorkSheet ws = new ExcelWorkSheet();
        // No setters/getters, but test all constructors
        // The following constructors are package-private or protected and might not be directly testable
        // without a concrete Sheet/Workbook object or being in the same package.
        // new ExcelWorkSheet(null, null, null, null);
        // For the protected constructor:
        // Need mock objects or actual POI objects to satisfy the constructor arguments.
        // For now, focusing on behavioral tests. If a Sheet object is null, it's caught by constructor.
        // Example: Assertions.assertThrows(IllegalArgumentException.class, () -> new ExcelWorkSheet(null, "sheet", null, null, null, "file.xlsx"));
    }

    @Test
    void iTestInvalidStateOperations() throws IOException {
        ExcelApplication xlApp = new ExcelApplication();
        // Create a workbook in memory, but don't save it initially, so its path might be unset/default
        // However, ExcelWorkBook constructor tries to set a name.
        // To truly test ExcelWorkSheet with a null sCompleteFileName, we'd need to construct it manually.
        // The public API to get an ExcelWorkSheet is via ExcelWorkBook.getExcelSheet(), which sets sCompleteFileName.

        // Let's test the sCompleteFileName check in saveWorkBook() by temporarily nullifying it (not ideal, but for specific test)
        // This requires a bit of setup to get a valid sheet object first.
        String tempFilePath = getTestFilePath("invalidStateTest", "xlsx");
        ExcelWorkBook tempBook = xlApp.createWorkBook(tempFilePath);
        Assertions.assertNotNull(tempBook, "Tempbook for invalidStateTest should be created.");
        ExcelWorkSheet sheet = tempBook.getExcelSheet("Sheet1");
        Assertions.assertNotNull(sheet, "Sheet should be created for invalidStateTest.");

        // Scenario 1: Test saveWorkBook() with null sCompleteFileName
        String originalFileName = sheet.sCompleteFileName; // store original
        sheet.sCompleteFileName = null;
        boolean saveResult = sheet.saveWorkBook();
        Assertions.assertFalse(saveResult, "saveWorkBook should return false with null sCompleteFileName");
        sheet.sCompleteFileName = ""; // Test with empty string
        saveResult = sheet.saveWorkBook();
        Assertions.assertFalse(saveResult, "saveWorkBook should return false with empty sCompleteFileName");
        sheet.sCompleteFileName = originalFileName; // restore

        // Scenario 2: Test cell() and row() with null sCompleteFileName
        // These methods in ExcelWorkSheet already have checks for null/empty sCompleteFileName.
        // We can test this by directly manipulating sCompleteFileName on a valid sheet object.
        if (sheet != null) { // sheet could be null if tempBook creation failed.
            sheet.sCompleteFileName = null;
            Assertions.assertNull(sheet.cell(1,1), "cell() should return null if sCompleteFileName is null");
            Assertions.assertNull(sheet.row(1), "row() should return null if sCompleteFileName is null");
            sheet.sCompleteFileName = originalFileName; // restore
        }


        xlApp.closeAllWorkBooks();
        cleanupTestFile(tempFilePath);
    }

    @Test
    void jTestGetColNumberScenarios() throws IOException {
        String filePath = getTestFilePath("colNumberTest", "xlsx");
        ExcelApplication xlApp = new ExcelApplication();
        ExcelWorkBook xlBook = xlApp.createWorkBook(filePath);
        Assertions.assertNotNull(xlBook, "Workbook for colNumberTest should be created.");
        ExcelWorkSheet sheet = xlBook.getExcelSheet("Sheet1"); // Assuming default sheet
        Assertions.assertNotNull(sheet, "Sheet1 for colNumberTest should exist.");

        // 1. Test with an empty sheet
        Assertions.assertEquals(0, sheet.getColNumber(), "Column number should be 0 for a new, empty sheet.");

        // 2. Test with a sheet that has data in A1 and C5 (sparse)
        sheet.cell(1, 1).setText("A1"); // Column A (index 0)
        sheet.cell(5, 3).setText("C5"); // Column C (index 2) -> max col number should be 3
        Assertions.assertEquals(3, sheet.getColNumber(), "Max column should be 3 for data in A1 and C5.");

        // 3. Test with rows that have varying numbers of cells
        sheet.cell(2, 5).setText("E2"); // Row 2, Col E (index 4) -> max col 5
        sheet.cell(3, 2).setText("B3"); // Row 3, Col B (index 1)
        Assertions.assertEquals(5, sheet.getColNumber(), "Max column should be 5 after adding E2.");

        // Add a cell further out
        sheet.cell(1, 10).setText("J1"); // Col J (index 9) -> max col 10
        Assertions.assertEquals(10, sheet.getColNumber(), "Max column should be 10 after adding J1.");


        // 4. Test after clearContents()
        boolean clearResult = sheet.clearContents(); // This also saves the workbook
        Assertions.assertTrue(clearResult, "Clear contents should succeed.");

        // Re-open to ensure state is read from file
        xlApp.closeAllWorkBooks();
        xlBook = xlApp.openWorkbook(filePath);
        sheet = xlBook.getExcelSheet("Sheet1");
        Assertions.assertNotNull(sheet);

        Assertions.assertEquals(0, sheet.getColNumber(), "Column number should be 0 after clearContents and reopen.");

        xlApp.closeAllWorkBooks();
        cleanupTestFile(filePath);
    }
}
