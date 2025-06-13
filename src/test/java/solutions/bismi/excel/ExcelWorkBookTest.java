package solutions.bismi.excel;

import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.MethodOrderer;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.TestMethodOrder;

import java.util.List;


import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.UUID;

@TestMethodOrder(MethodOrderer.MethodName.class)
class ExcelWorkBookTest {

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
    void aAddSheetXLSXWorkSheet() throws IOException {
        String xlsxFile = getTestFilePath("addSheet", "xlsx");
        addSheetXLWorkSheet(xlsxFile);
        cleanupTestFile(xlsxFile);

        String xlsFile = getTestFilePath("addSheet", "xls");
        addSheetXLWorkSheet(xlsFile);
        cleanupTestFile(xlsFile);
    }

    private void addSheetXLWorkSheet(String strCompleteFileName) {
        ExcelApplication xlApp = new ExcelApplication();
        ExcelWorkBook xlbook = xlApp.createWorkBook(strCompleteFileName);
        int cnt = 0;
        cnt = xlbook.getSheetCount();
        Assertions.assertEquals(1, cnt);
        xlbook.addSheet("Bismi1");
        cnt = xlbook.getSheetCount();
        xlApp.closeAllWorkBooks();
        Assertions.assertEquals(2, cnt);
    }

    @Test
    void bAddSheetXLSXWorkSheets() throws IOException {
        String xlsxFile = getTestFilePath("addSheets", "xlsx");
        dAddSheetXLWorkSheets(xlsxFile);
        cleanupTestFile(xlsxFile);

        String xlsFile = getTestFilePath("addSheets", "xls");
        dAddSheetXLWorkSheets(xlsFile);
        cleanupTestFile(xlsFile);
    }

    private void dAddSheetXLWorkSheets(String strCompleteFileName) {
        ExcelApplication xlApp = new ExcelApplication();
        ExcelWorkBook xlbook = xlApp.createWorkBook(strCompleteFileName);
        int cnt = 0;
        cnt = xlbook.getSheetCount();
        Assertions.assertEquals(1, cnt);

        String[] shArray = {"Bismi", "Solutions", "Sulfikar", "Ali", "Nazar"};
        xlbook.addSheets(shArray);
        cnt = xlbook.getSheetCount();
        xlApp.closeAllWorkBooks();
        Assertions.assertEquals(6, cnt);
    }

    @Test
    void eGetXLSXWorkSheet() throws IOException {
        String xlsFile = getTestFilePath("getSheet", "xls");
        getXLSWorkSheet(xlsFile);
        cleanupTestFile(xlsFile);

        String xlsxFile = getTestFilePath("getSheet", "xlsx");
        getXLSWorkSheet(xlsxFile);
        cleanupTestFile(xlsxFile);
    }

    private void getXLSWorkSheet(String strCompleteFileName) {
        ExcelApplication xlApp = new ExcelApplication();
        ExcelWorkBook xlbook = xlApp.createWorkBook(strCompleteFileName);
        int cnt = 0;
        cnt = xlbook.getSheetCount();
        Assertions.assertEquals(1, cnt);

        ExcelWorkSheet xlsht = xlbook.addSheet("Bismi");
        String shName = xlsht.getSheetName();
        String shName1 = xlbook.getExcelSheet(0).getSheetName();
        xlApp.closeAllWorkBooks();
        Assertions.assertEquals("Sheet1", shName1);
        Assertions.assertEquals("Bismi", shName);
    }

    private void gGetXLSXWorkSheets(String strCompleteFileName) {
        ExcelApplication xlApp = new ExcelApplication();
        ExcelWorkBook xlbook = xlApp.createWorkBook(strCompleteFileName);
        int cnt = 0;
        cnt = xlbook.getSheetCount();
        Assertions.assertEquals(1, cnt);

        String[] shArray = {"Bismi", "Solutions", "Sulfikar", "Ali", "Nazar"};
        xlbook.addSheets(shArray);
        List<ExcelWorkSheet> sheetList = xlbook.getExcelSheets();
        ExcelWorkSheet sht3 = sheetList.get(2);
        String shName = sht3.getSheetName();
        xlApp.closeAllWorkBooks();
        Assertions.assertEquals("Solutions", shName);
    }

    @Test
    void hgetExcelWokSheets() throws IOException {
        String xlsxFile = getTestFilePath("getSheets", "xlsx");
        gGetXLSXWorkSheets(xlsxFile);
        cleanupTestFile(xlsxFile);

        String xlsFile = getTestFilePath("getSheets", "xls");
        gGetXLSXWorkSheets(xlsFile);
        cleanupTestFile(xlsFile);
    }

    @Test
    void iTestWorkbookOperations() throws IOException {
        String xlsxFile = getTestFilePath("workbookOperations", "xlsx");
        testWorkbookOperations(xlsxFile);
        cleanupTestFile(xlsxFile);

        String xlsFile = getTestFilePath("workbookOperations", "xls");
        testWorkbookOperations(xlsFile);
        cleanupTestFile(xlsFile);
    }

    private void testWorkbookOperations(String strCompleteFileName) {
        ExcelApplication xlApp = new ExcelApplication();
        ExcelWorkBook xlbook = xlApp.createWorkBook(strCompleteFileName);

        // Test workbook name operations
        String originalName = xlbook.getExcelBookName();
        Assertions.assertNotNull(originalName, "Workbook should have a name");
        Assertions.assertTrue(originalName.endsWith(".xlsx") || originalName.endsWith(".xls"), "Workbook name should have correct extension");

        // Test sheet operations
        xlbook.addSheet("TestSheet1");
        xlbook.addSheet("TestSheet2");

        // Test sheet retrieval by index
        ExcelWorkSheet retrievedSheet = xlbook.getExcelSheet(0);
        Assertions.assertEquals("Sheet1", retrievedSheet.getSheetName(), "Should retrieve correct sheet by index");

        // Test sheet retrieval by name
        ExcelWorkSheet namedSheet = xlbook.getExcelSheet("TestSheet1");
        Assertions.assertNotNull(namedSheet, "Should retrieve sheet by name");
        Assertions.assertEquals("TestSheet1", namedSheet.getSheetName(), "Should retrieve correct sheet by name");

        // Test sheet count
        int sheetCount = xlbook.getSheetCount();
        Assertions.assertEquals(3, sheetCount, "Should have correct number of sheets");

        // Test sheet list retrieval
        List<ExcelWorkSheet> sheets = xlbook.getExcelSheets();
        Assertions.assertEquals(3, sheets.size(), "Should retrieve correct number of sheets");
        Assertions.assertEquals("Sheet1", sheets.get(0).getSheetName(), "First sheet should be Sheet1");
        Assertions.assertEquals("TestSheet1", sheets.get(1).getSheetName(), "Second sheet should be TestSheet1");
        Assertions.assertEquals("TestSheet2", sheets.get(2).getSheetName(), "Third sheet should be TestSheet2");

        xlApp.closeAllWorkBooks();
    }

    @Test
    void jTestWorkbookErrorHandling() throws IOException {
        String xlsxFile = getTestFilePath("workbookErrorHandling", "xlsx");
        testWorkbookErrorHandling(xlsxFile);
        cleanupTestFile(xlsxFile);

        String xlsFile = getTestFilePath("workbookErrorHandling", "xls");
        testWorkbookErrorHandling(xlsFile);
        cleanupTestFile(xlsFile);
    }

    private void testWorkbookErrorHandling(String strCompleteFileName) {
        ExcelApplication xlApp = new ExcelApplication();
        ExcelWorkBook xlbook = xlApp.createWorkBook(strCompleteFileName);

        // Test invalid sheet index
        ExcelWorkSheet invalidSheet = xlbook.getExcelSheet(-1);
        Assertions.assertNull(invalidSheet, "Should return null for negative sheet index");

        invalidSheet = xlbook.getExcelSheet(999);
        Assertions.assertNull(invalidSheet, "Should return null for out-of-bounds sheet index");

        // Test invalid sheet name
        invalidSheet = xlbook.getExcelSheet("NonExistentSheet");
        Assertions.assertNull(invalidSheet, "Should return null for non-existent sheet name");

        // Test adding sheet with null name
        ExcelWorkSheet nullSheet = xlbook.addSheet(null);
        Assertions.assertNull(nullSheet, "Should return null when adding sheet with null name");

        // Test adding sheet with empty name
        ExcelWorkSheet emptySheet = xlbook.addSheet("");
        Assertions.assertNull(emptySheet, "Should return null when adding sheet with empty name");

        // Test adding multiple sheets with invalid names
        String[] invalidSheetNames = {null, "", "ValidSheet", null, ""};
        xlbook.addSheets(invalidSheetNames);
        xlbook.getExcelSheet(0).saveWorkBook(); // Save to ensure changes are persisted
        int sheetCount = xlbook.getSheetCount();
        Assertions.assertEquals(2, sheetCount, "Should only add valid sheet names");

        xlApp.closeAllWorkBooks();
    }

    @Test
    void kTestWorkbookContentOperations() throws IOException {
        String xlsxFile = getTestFilePath("workbookContent", "xlsx");
        testWorkbookContentOperations(xlsxFile);
        cleanupTestFile(xlsxFile);

        String xlsFile = getTestFilePath("workbookContent", "xls");
        testWorkbookContentOperations(xlsFile);
        cleanupTestFile(xlsFile);
    }

    private void testWorkbookContentOperations(String strCompleteFileName) {
        ExcelApplication xlApp = new ExcelApplication();
        ExcelWorkBook xlbook = xlApp.createWorkBook(strCompleteFileName);

        // Create test sheets with content
        ExcelWorkSheet sheet1 = xlbook.addSheet("ContentSheet1");
        ExcelWorkSheet sheet2 = xlbook.addSheet("ContentSheet2");

        // Add content to first sheet
        sheet1.cell(1, 1).setText("Sheet1 Content");
        sheet1.cell(1, 2).setNumericValue(123.45);
        sheet1.cell(2, 1).setDateValue(new java.util.Date());

        // Add content to second sheet
        sheet2.cell(1, 1).setText("Sheet2 Content");
        sheet2.cell(1, 2).setFormula("SUM(A1:A10)");

        // Save to ensure content is persisted
        sheet1.saveWorkBook();

        // Verify content in first sheet
        ExcelWorkSheet retrievedSheet1 = xlbook.getExcelSheet("ContentSheet1");
        Assertions.assertEquals("Sheet1 Content", retrievedSheet1.cell(1, 1).getTextValue());
        Assertions.assertEquals("123.45", retrievedSheet1.cell(1, 2).getValue());
        Assertions.assertNotNull(retrievedSheet1.cell(2, 1).getValue());

        // Verify content in second sheet
        ExcelWorkSheet retrievedSheet2 = xlbook.getExcelSheet("ContentSheet2");
        Assertions.assertEquals("Sheet2 Content", retrievedSheet2.cell(1, 1).getTextValue());
        Assertions.assertEquals("SUM(A1:A10)", retrievedSheet2.cell(1, 2).getValue());

        // Test sheet activation
        sheet1.activate();
        Assertions.assertEquals("ContentSheet1", xlbook.getExcelSheet(1).getSheetName());

        sheet2.activate();
        Assertions.assertEquals("ContentSheet2", xlbook.getExcelSheet(2).getSheetName());

        retrievedSheet1 = xlbook.getExcelSheet("ContentSheet1");
        Assertions.assertEquals("Sheet1 Content", retrievedSheet1.cell(1, 1).getTextValue());

        xlApp.closeAllWorkBooks();
    }

    @Test
    void lTestWorkbookFileOperations() throws IOException {
        String xlsxFile = getTestFilePath("workbookFile", "xlsx");
        testWorkbookFileOperations(xlsxFile);
        cleanupTestFile(xlsxFile);

        String xlsFile = getTestFilePath("workbookFile", "xls");
        testWorkbookFileOperations(xlsFile);
        cleanupTestFile(xlsFile);
    }

    private void testWorkbookFileOperations(String strCompleteFileName) {
        ExcelApplication xlApp = new ExcelApplication();
        ExcelWorkBook xlbook = xlApp.createWorkBook(strCompleteFileName);

        // Test file extension handling
        String fileExtension = strCompleteFileName.substring(strCompleteFileName.lastIndexOf("."));
        Assertions.assertTrue(fileExtension.equalsIgnoreCase(".xlsx") || fileExtension.equalsIgnoreCase(".xls"), "File should have correct extension");

        // Test workbook creation with content
        ExcelWorkSheet sheet = xlbook.addSheet("FileTest");
        sheet.cell(1, 1).setText("Test Content");
        sheet.saveWorkBook();
        xlApp.closeAllWorkBooks();
    }

    @Test
    void lTestWorkbookOpen() throws IOException {
        String xlsxFile = getTestFilePath("openWorkbookTest", "xlsx");
        setupTestFileForOpen(xlsxFile); // Create a file with known content
        testWorkbookOpen(xlsxFile);
        cleanupTestFile(xlsxFile);

        String xlsFile = getTestFilePath("openWorkbookTest", "xls");
        setupTestFileForOpen(xlsFile); // Create a file with known content
        testWorkbookOpen(xlsFile);
        cleanupTestFile(xlsFile);
    }

    private void setupTestFileForOpen(String filePath) {
        ExcelApplication app = new ExcelApplication();
        ExcelWorkBook book = app.createWorkBook(filePath);
        Assertions.assertNotNull(book, "Setup: Book should be created.");
        ExcelWorkSheet sheet = book.getExcelSheet("Sheet1"); // Default sheet
        Assertions.assertNotNull(sheet, "Setup: Sheet1 should exist.");
        sheet.cell(1,1).setText("Test");
        boolean saved = book.saveWorkbook();
        Assertions.assertTrue(saved, "Setup: Workbook should be saved.");
        app.closeAllWorkBooks();
    }

    private void testWorkbookOpen(String strCompleteFileName) {
        ExcelApplication xlApp = new ExcelApplication();
        ExcelWorkBook xlbook = xlApp.openWorkbook(strCompleteFileName);
        Assertions.assertNotNull(xlbook, "Should be able to open workbook: " + strCompleteFileName);
        ExcelWorkSheet sheet = xlbook.getExcelSheet(0);
        Assertions.assertNotNull(sheet, "Should be able to get sheet 0 after opening.");
        Assertions.assertEquals("Sheet1", sheet.getSheetName(), "Should retrieve correct sheet by index");
        Assertions.assertEquals("Test",sheet.cell(1,1).getTextValue(),"Should retrieve correct cell value");

        xlApp.closeAllWorkBooks();
    }

    @Test
    void zCoverageGettersSettersAndConstructors() {
        ExcelWorkBook wb = new ExcelWorkBook();
        wb.setExcelSheets(null);
        wb.setWb(null);
        wb.setExcelBookName("testbook");
        wb.setExcelBookPath("/tmp");
        wb.setInputStream(null);
        wb.setOutputStream(null);
        wb.setXlFile(null);
        wb.setXlFileExtension(".xlsx");
        wb.setLog(org.apache.logging.log4j.LogManager.getLogger(ExcelWorkBook.class));

        // Call all getters
        wb.getExcelSheets();
        wb.getWb();
        wb.getExcelBookName();
        wb.getExcelBookPath();
        wb.getInputStream();
        wb.getOutputStream();
        wb.getXlFile();
        wb.getXlFileExtension();
        wb.getLog();

        // Test constructors
        new ExcelWorkBook(null, "testbook", "/tmp");
    }

    @Test
    void mTestGetFileNameScenarios() {
        ExcelWorkBook testWb = new ExcelWorkBook(); // Instance to call the private method indirectly

        // Test case 1: Filename with multiple dots
        String path1 = "./resources/testdata/test.archive.v1.xlsx";
        // ExcelWorkBook.getFileName is private. We test its effect on excelBookName and xlFileExtension
        // by calling a method that uses it, like createWorkBook (even if file creation itself isn't the focus).
        // createWorkBook will set the internal fields based on getFileName.
        // For this test, we don't need xlApp or to actually create the file fully if getFileName is robust.
        // We'll use a dummy ExcelWorkBook instance to inspect the results of getFileName.

        // To test getFileName, we'd ideally call it directly. Since it's private,
        // we observe its side-effects on excelBookName and xlFileExtension
        // when createWorkBook or openWorkbook is called.
        // The createWorkBook method in ExcelWorkBook calls getFileName.

        ExcelApplication xlApp = new ExcelApplication(); // Needed for createWorkBook in this context
        ExcelWorkBook wb1 = xlApp.createWorkBook(path1); // This will call getFileName
        Assertions.assertNotNull(wb1, "Workbook creation should not fail for path1");
        Assertions.assertEquals("test.archive.v1.xlsx", wb1.getExcelBookName(), "Test Case 1 Filename");
        Assertions.assertEquals(".xlsx", wb1.getXlFileExtension(), "Test Case 1 Extension");
        xlApp.closeAllWorkBooks(); // Close to release resources, file might have been created

        // Test case 2: Filename with no extension
        String path2 = "./resources/testdata/myworkbook";
        ExcelWorkBook wb2 = xlApp.createWorkBook(path2);
        Assertions.assertNotNull(wb2);
        Assertions.assertEquals("myworkbook", wb2.getExcelBookName(), "Test Case 2 Filename");
        Assertions.assertEquals("", wb2.getXlFileExtension(), "Test Case 2 Extension");
        xlApp.closeAllWorkBooks();

        // Test case 3: Filename starting with a dot (hidden file, no extension)
        String path3 = "./resources/testdata/.configfile";
        ExcelWorkBook wb3 = xlApp.createWorkBook(path3);
        Assertions.assertNotNull(wb3);
        Assertions.assertEquals(".configfile", wb3.getExcelBookName(), "Test Case 3 Filename");
        Assertions.assertEquals("", wb3.getXlFileExtension(), "Test Case 3 Extension");
        xlApp.closeAllWorkBooks();

        // Test case 4: Filename starting with a dot and having an extension
        String path4 = "./resources/testdata/.myhidden.config";
        ExcelWorkBook wb4 = xlApp.createWorkBook(path4);
        Assertions.assertNotNull(wb4);
        // Based on current getFileName logic: fileName.lastIndexOf(".")
        // For ".myhidden.config", lastIndexOf('.') is at index 9 (before "config")
        // lastDotIndex > 0 is true (9 > 0). So, extension becomes ".config"
        Assertions.assertEquals(".myhidden.config", wb4.getExcelBookName(), "Test Case 4 Filename");
        Assertions.assertEquals(".config", wb4.getXlFileExtension(), "Test Case 4 Extension");
        xlApp.closeAllWorkBooks();

        // Test case 5: Null input for sCompleteFilePath
        // ExcelWorkBook.createWorkBook has a null check at the beginning and returns null.
        // So, getFileName won't be called with null directly from there.
        // The private getFileName method itself has a null check.
        // This is implicitly tested by ensuring createWorkBook(null) doesn't throw NPE from getFileName.
        ExcelWorkBook wb5 = xlApp.createWorkBook(null);
        Assertions.assertNull(wb5, "Test Case 5: createWorkBook with null path should return null");
        xlApp.closeAllWorkBooks();


        // Test case 6: Empty string input for sCompleteFilePath
        ExcelWorkBook wb6 = xlApp.createWorkBook("");
        Assertions.assertNull(wb6, "Test Case 6: createWorkBook with empty path should return null");
        xlApp.closeAllWorkBooks();

        // Clean up created files if any (optional, as they are in testdata and might be overwritten)
        // For robustness, good to delete test files.
        // Test specific cleanup now handled in each test method for clarity.
        // If createWorkBook in mTestGetFileNameScenarios actually creates files that are not cleaned up by test specific cleanup,
        // then general cleanup might be needed, but it's better if each test manages its own files.
        // The current structure of mTestGetFileNameScenarios calls createWorkBook which might create files.
        // Let's ensure its own cleanup is robust.
        try {
            if (path1 != null) java.nio.file.Files.deleteIfExists(java.nio.file.Paths.get(path1));
            if (path2 != null) java.nio.file.Files.deleteIfExists(java.nio.file.Paths.get(path2)); // This path might be problematic if it became a directory.
            if (path3 != null) java.nio.file.Files.deleteIfExists(java.nio.file.Paths.get(path3));
            if (path4 != null) java.nio.file.Files.deleteIfExists(java.nio.file.Paths.get(path4));
        } catch (java.io.IOException e) {
            // Use the class-level logger if available, or System.err
             System.err.println("Warning: Could not clean up all files in mTestGetFileNameScenarios: " + e.getMessage());
        }
    }

    @Test
    void nTestCreateWorkbookWithRelativePathAndSubdirectory() throws IOException {
        String uniqueSubDirName = "temp-test-output_" + UUID.randomUUID().toString().substring(0, 8);
        String relativePathString = "./" + uniqueSubDirName + "/relative_workbook_test.xlsx";
        java.nio.file.Path testFilePath = java.nio.file.Paths.get(relativePathString);
        java.nio.file.Path parentDirPath = testFilePath.getParent();

        // Ensure directory does not exist before test
        if (parentDirPath != null && Files.exists(parentDirPath)) {
            // Clean up if it exists from a previous failed run
            Files.walk(parentDirPath)
                 .sorted(java.util.Comparator.reverseOrder())
                 .map(Path::toFile)
                 .forEach(java.io.File::delete);
        }
        Assertions.assertFalse(parentDirPath != null && Files.exists(parentDirPath), "Pre-condition: Subdirectory should not exist.");

        ExcelApplication xlApp = new ExcelApplication();
        ExcelWorkBook xlBook = xlApp.createWorkBook(relativePathString);

        Assertions.assertNotNull(xlBook, "Workbook object should not be null when creating with relative path.");
        Assertions.assertNotNull(xlBook.getWb(), "Internal POI workbook should not be null.");
        Assertions.assertTrue(xlBook.getExcelBookPath().contains(uniqueSubDirName), "ExcelBookPath should reflect the new subdirectory.");

        ExcelWorkSheet sheet = xlBook.addSheet("TestSheet");
        Assertions.assertNotNull(sheet, "Sheet should be added to the workbook.");
        sheet.cell(1, 1).setText("Relative Path Test Successful");

        boolean saveResult = xlBook.saveWorkbook();
        Assertions.assertTrue(saveResult, "Workbook should be saved successfully.");

        java.io.File testFile = testFilePath.toFile();
        Assertions.assertTrue(testFile.exists(), "Excel file should exist at relative path: " + relativePathString);
        Assertions.assertTrue(testFile.length() > 0, "Excel file should not be empty.");

        xlApp.closeAllWorkBooks();

        // Cleanup
        try {
            if (Files.exists(testFilePath)) {
                Files.delete(testFilePath);
            }
            if (parentDirPath != null && Files.exists(parentDirPath)) {
                 // Attempt to delete the directory. It should be empty now.
                Files.deleteIfExists(parentDirPath);
            }
        } catch (java.io.IOException e) {
            System.err.println("Warning: Could not clean up test file/directory: " + e.getMessage());
        }
    }
}
