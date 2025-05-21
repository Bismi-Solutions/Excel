package solutions.bismi.excel;

import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.UUID;

class ExcelApplicationTest {

    private static final String TEST_DATA_DIR = "./resources/testdata";

    private String createTestFile(String testName) throws IOException {
        // Create test data directory if it doesn't exist
        Files.createDirectories(Path.of(TEST_DATA_DIR));
        
        String uniqueId = UUID.randomUUID().toString().substring(0, 8);
        String fileName = testName + uniqueId + ".XLSX";
        String filePath = Path.of(TEST_DATA_DIR, fileName).toString();
        
        // Create a new Excel workbook
        ExcelApplication xlApp = new ExcelApplication();
        xlApp.createWorkBook(filePath);
        xlApp.closeAllWorkBooks();
        
        return filePath;
    }

    @Test
    void testCreateExcelWorkBook() throws IOException {
        String file1 = createTestFile("createWorkbook");
        String file2 = createTestFile("createWorkbook");
        testCreateXLSWorkBook(file1);
        testCreateXLSWorkBook(file2);
    }

    private void testCreateXLSWorkBook(String strCompleteFileName) {
        ExcelApplication xlApp = new ExcelApplication();
        xlApp.createWorkBook(strCompleteFileName);
        int cnt = xlApp.getOpenWorkbookCount();
        xlApp.closeAllWorkBooks();
        //Verify only one work book is created
        Assertions.assertEquals(1, cnt);
    }

    @Test
    void testOpenExcelWorkbook() throws IOException {
        String file1 = createTestFile("openWorkbook");
        String file2 = createTestFile("openWorkbook");
        testOpenXLSWorkbook(file1);
        testOpenXLSWorkbook(file2);
    }

    private void testOpenXLSWorkbook(String strCompleteFileName) {
        ExcelApplication xlApp = new ExcelApplication();
        xlApp.openWorkbook(strCompleteFileName);
        int cnt = xlApp.getOpenWorkbookCount();
        xlApp.closeAllWorkBooks();
        //Verify only one work book is created
        Assertions.assertEquals(1, cnt);
    }

    @Test
    void testCloseWorkBookByIndex() throws IOException {
        String file1 = createTestFile("closeByIndex");
        String file2 = createTestFile("closeByIndex");
        testCloseBookByIndex(file1);
        testCloseBookByIndex(file2);
    }

    private void testCloseBookByIndex(String strCompleteFileName) {
        ExcelApplication xlApp = new ExcelApplication();
        xlApp.openWorkbook(strCompleteFileName);
        int cnt = xlApp.getOpenWorkbookCount();
        Assertions.assertEquals(1, cnt);
        xlApp.closeWorkBook(0);
        cnt = xlApp.getOpenWorkbookCount();
        Assertions.assertEquals(0, cnt);
    }

    @Test
    void testCloseWorkBookByName() throws IOException {
        String file1 = createTestFile("closeByName");
        String file2 = createTestFile("closeByName");
        testCloseBookByName(file1);
        testCloseBookByName(file2);
    }

    private void testCloseBookByName(String strCompleteFileName) {
        ExcelApplication xlApp = new ExcelApplication();
        ExcelWorkBook xlBook = xlApp.openWorkbook(strCompleteFileName);
        Assertions.assertNotNull(xlBook, "Workbook should be opened successfully");
        int cnt = xlApp.getOpenWorkbookCount();
        Assertions.assertEquals(1, cnt, "Should have one workbook open");
        xlApp.closeWorkBook(xlBook.getExcelBookName());
        cnt = xlApp.getOpenWorkbookCount();
        Assertions.assertEquals(0, cnt, "Workbook should be closed");
    }

    @Test
    void testCloseAllWorkBooks() throws IOException {
        ExcelApplication xlApp = new ExcelApplication();
        String file1 = createTestFile("closeAll1");
        String file2 = createTestFile("closeAll2");
        
        xlApp.openWorkbook(file1);
        int cnt = xlApp.getOpenWorkbookCount();
        Assertions.assertEquals(1, cnt);

        //Verifying the close functionality of xlsx
        xlApp.openWorkbook(file2);
        cnt = xlApp.getOpenWorkbookCount();
        Assertions.assertEquals(2, cnt);
        xlApp.closeAllWorkBooks();
        cnt = xlApp.getOpenWorkbookCount();
        Assertions.assertEquals(0, cnt);
    }

    @Test
    void testGetWorkbooks() throws IOException {
        ExcelApplication xlApp = new ExcelApplication();
        String file1 = createTestFile("getWorkbooks1");
        String file2 = createTestFile("getWorkbooks2");

        // Initially should be empty
        List<ExcelWorkBook> workbooks = xlApp.getWorkbooks();
        Assertions.assertEquals(0, workbooks.size());

        // Add a workbook and check
        xlApp.openWorkbook(file1);
        workbooks = xlApp.getWorkbooks();
        Assertions.assertEquals(1, workbooks.size());

        // Add another workbook and check
        xlApp.openWorkbook(file2);
        workbooks = xlApp.getWorkbooks();
        Assertions.assertEquals(2, workbooks.size());

        xlApp.closeAllWorkBooks();
    }

    @Test
    void testGetWorkbookByIndex() throws IOException {
        ExcelApplication xlApp = new ExcelApplication();
        String file1 = createTestFile("getByIndex1");
        String file2 = createTestFile("getByIndex2");

        // Test with valid index
        ExcelWorkBook xlBook1 = xlApp.openWorkbook(file1);
        ExcelWorkBook xlBook2 = xlApp.openWorkbook(file2);

        ExcelWorkBook retrievedBook = xlApp.getWorkbook(0);
        Assertions.assertNotNull(retrievedBook);
        Assertions.assertEquals(xlBook1.getExcelBookName(), retrievedBook.getExcelBookName());

        retrievedBook = xlApp.getWorkbook(1);
        Assertions.assertNotNull(retrievedBook);
        Assertions.assertEquals(xlBook2.getExcelBookName(), retrievedBook.getExcelBookName());

        // Test with invalid index
        retrievedBook = xlApp.getWorkbook(2);
        Assertions.assertNull(retrievedBook);

        retrievedBook = xlApp.getWorkbook(-1);
        Assertions.assertNull(retrievedBook);

        xlApp.closeAllWorkBooks();
    }

    @Test
    void testGetWorkbookByName() throws IOException {
        ExcelApplication xlApp = new ExcelApplication();
        String file1 = createTestFile("getByName");

        // Test with valid name
        ExcelWorkBook xlBook1 = xlApp.openWorkbook(file1);
        String bookName = xlBook1.getExcelBookName();

        ExcelWorkBook retrievedBook = xlApp.getWorkbook(bookName);
        Assertions.assertNotNull(retrievedBook);
        Assertions.assertEquals(bookName, retrievedBook.getExcelBookName());

        // Test with invalid name
        retrievedBook = xlApp.getWorkbook("NonExistentBook");
        Assertions.assertNull(retrievedBook);

        // Test with null and empty name
        retrievedBook = xlApp.getWorkbook(null);
        Assertions.assertNull(retrievedBook);

        retrievedBook = xlApp.getWorkbook("");
        Assertions.assertNull(retrievedBook);

        xlApp.closeAllWorkBooks();
    }

    @Test
    void testCloseWorkBookErrorCases() throws IOException {
        ExcelApplication xlApp = new ExcelApplication();
        String file1 = createTestFile("errorCases");

        // Test closing by invalid index
        xlApp.closeWorkBook(0); // Should not throw exception
        Assertions.assertEquals(0, xlApp.getOpenWorkbookCount(), "Workbook count should remain zero after closing invalid index");

        // Test closing by null or empty name
        xlApp.closeWorkBook(null); // Should not throw exception
        Assertions.assertEquals(0, xlApp.getOpenWorkbookCount(), "Workbook count should remain zero after closing null name");

        xlApp.closeWorkBook(""); // Should not throw exception
        Assertions.assertEquals(0, xlApp.getOpenWorkbookCount(), "Workbook count should remain zero after closing empty name");

        // Test closing non-existent workbook
        xlApp.closeWorkBook("NonExistentBook"); // Should not throw exception
        Assertions.assertEquals(0, xlApp.getOpenWorkbookCount(), "Workbook count should remain zero after closing non-existent workbook");
    }
}
