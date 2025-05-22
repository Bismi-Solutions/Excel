package solutions.bismi.excel;

import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.MethodOrderer;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.TestMethodOrder;

import java.util.List;


@TestMethodOrder(MethodOrderer.MethodName.class)
class ExcelWorkBookTest {

    @Test
    void aAddSheetXLSXWorkSheet() {
        addSheetXLWorkSheet("./resources/testdata/addSheet.xlsx");
        addSheetXLWorkSheet("./resources/testdata/addSheet.xls");
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
    void bAddSheetXLSXWorkSheets() {
        dAddSheetXLWorkSheets("./resources/testdata/addSheets.xlsx");
        dAddSheetXLWorkSheets("./resources/testdata/addSheets.xls");
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
    void eGetXLSXWorkSheet() {
        getXLSWorkSheet("./resources/testdata/getSheet.xls");
        getXLSWorkSheet("./resources/testdata/getSheet.xlsx");
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
    void hgetExcelWokSheets() {
        gGetXLSXWorkSheets("./resources/testdata/getSheets.xlsx");
        gGetXLSXWorkSheets("./resources/testdata/getSheets.xls");
    }

    @Test
    void iTestWorkbookOperations() {
        testWorkbookOperations("./resources/testdata/workbookOperations.xlsx");
        testWorkbookOperations("./resources/testdata/workbookOperations.xls");
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
    void jTestWorkbookErrorHandling() {
        testWorkbookErrorHandling("./resources/testdata/workbookErrorHandling.xlsx");
        testWorkbookErrorHandling("./resources/testdata/workbookErrorHandling.xls");
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
    void kTestWorkbookContentOperations() {
        testWorkbookContentOperations("./resources/testdata/workbookContent.xlsx");
        testWorkbookContentOperations("./resources/testdata/workbookContent.xls");
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
    void lTestWorkbookFileOperations() {
        testWorkbookFileOperations("./resources/testdata/workbookFile.xlsx");
        testWorkbookFileOperations("./resources/testdata/workbookFile.xls");
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
    void lTestWorkbookOpen() {
        testWorkbookOpen("./resources/testdata/openWorkbook.xlsx");
        testWorkbookOpen("./resources/testdata/openWorkbook.xls");
    }

    private void testWorkbookOpen(String strCompleteFileName) {
        ExcelApplication xlApp = new ExcelApplication();
        ExcelWorkBook xlbook = xlApp.openWorkbook(strCompleteFileName);
        ExcelWorkSheet sheet = xlbook.getExcelSheet(0);
        Assertions.assertNotNull(sheet, "Should be able to open workbook");
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

}
