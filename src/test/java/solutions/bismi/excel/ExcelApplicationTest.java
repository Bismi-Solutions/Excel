package solutions.bismi.excel;

import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.MethodOrderer;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.TestMethodOrder;

import java.util.List;


@TestMethodOrder(MethodOrderer.MethodName.class)
class ExcelApplicationTest {

    @Test
    void aCreateExcelWorkBook() {
        bCreateXLSWorkBook("./resources/testdata/basicWorkbook.xlsx");
        bCreateXLSWorkBook("./resources/testdata/basicWorkbook.xls");

    }

    private void bCreateXLSWorkBook(String strCompleteFileName) {
        ExcelApplication xlApp = new ExcelApplication();
        xlApp.createWorkBook(strCompleteFileName);
        int cnt = xlApp.getOpenWorkbookCount();
        xlApp.closeAllWorkBooks();
        //Verify only one work book is created
        Assertions.assertEquals(1, cnt);

    }

    @Test
    void cOpenExcelXWorkbook() {
        dOpenXLSWorkbook("./resources/testdata/basicWorkbook.xlsx");
        dOpenXLSWorkbook("./resources/testdata/basicWorkbook.xls");
    }

    private void dOpenXLSWorkbook(String strCompleteFileName) {
        ExcelApplication xlApp = new ExcelApplication();
        xlApp.openWorkbook(strCompleteFileName);
        int cnt = xlApp.getOpenWorkbookCount();
        xlApp.closeAllWorkBooks();
        //Verify only one work book is created
        Assertions.assertEquals(1, cnt);
    }

    @Test
    void eCloseWorkBookByIndex() {
        closeBookByindex("./resources/testdata/basicWorkbook.xls");
        closeBookByindex("./resources/testdata/basicWorkbook.xlsx");

    }

    private void closeBookByindex(String strCompleteFileName) {
        ExcelApplication xlApp = new ExcelApplication();
        xlApp.openWorkbook(strCompleteFileName);
        int cnt = xlApp.getOpenWorkbookCount();
        Assertions.assertEquals(1, cnt);
        xlApp.closeWorkBook(0);
        cnt = xlApp.getOpenWorkbookCount();
        Assertions.assertEquals(0, cnt);
        Assertions.assertEquals(0, cnt);
    }

    @Test
    void fCloseWorkBookByName() {
        closeBookByName("./resources/testdata/basicWorkbook.xls");
        closeBookByName("./resources/testdata/basicWorkbook.xlsx");
    }

    private void closeBookByName(String strCompleteFileName) {
        ExcelApplication xlApp = new ExcelApplication();
        ExcelWorkBook xlBook = xlApp.openWorkbook(strCompleteFileName);
        int cnt = xlApp.getOpenWorkbookCount();
        Assertions.assertEquals(1, cnt);
        xlApp.closeWorkBook(xlBook.getExcelBookName());
        cnt = xlApp.getOpenWorkbookCount();
        Assertions.assertEquals(0, cnt);
    }

    @Test
    void gCloseAllWorkBooks() {
        ExcelApplication xlApp = new ExcelApplication();
        xlApp.openWorkbook("./resources/testdata/basicWorkbook.xls");
        int cnt = xlApp.getOpenWorkbookCount();
        Assertions.assertEquals(1, cnt);

        //Verifying the close functionality of xlsx
        xlApp.openWorkbook("./resources/testdata/basicWorkbook.xlsx");
        cnt = xlApp.getOpenWorkbookCount();
        Assertions.assertEquals(2, cnt);
        xlApp.closeAllWorkBooks();
        cnt = xlApp.getOpenWorkbookCount();
        Assertions.assertEquals(0, cnt);


    }

    @Test
    void hGetWorkbooks() {
        ExcelApplication xlApp = new ExcelApplication();

        // Initially should be empty
        List<ExcelWorkBook> workbooks = xlApp.getWorkbooks();
        Assertions.assertEquals(0, workbooks.size());

        // Add a workbook and check
        xlApp.openWorkbook("./resources/testdata/basicWorkbook.xlsx");
        workbooks = xlApp.getWorkbooks();
        Assertions.assertEquals(1, workbooks.size());

        // Add another workbook and check
        xlApp.openWorkbook("./resources/testdata/basicWorkbook.xls");
        workbooks = xlApp.getWorkbooks();
        Assertions.assertEquals(2, workbooks.size());

        xlApp.closeAllWorkBooks();
    }

    @Test
    void iGetWorkbookByIndex() {
        ExcelApplication xlApp = new ExcelApplication();

        // Test with valid index
        ExcelWorkBook xlBook1 = xlApp.openWorkbook("./resources/testdata/basicWorkbook.xlsx");
        ExcelWorkBook xlBook2 = xlApp.openWorkbook("./resources/testdata/basicWorkbook.xls");

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
    void jGetWorkbookByName() {
        ExcelApplication xlApp = new ExcelApplication();

        // Test with valid name
        ExcelWorkBook xlBook1 = xlApp.openWorkbook("./resources/testdata/basicWorkbook.xlsx");
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
    void kCloseWorkBookErrorCases() {
        ExcelApplication xlApp = new ExcelApplication();

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
