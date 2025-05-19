package solutions.bismi.excel;

import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.MethodOrderer;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.TestMethodOrder;

import java.util.List;


@TestMethodOrder(MethodOrderer.Alphanumeric.class)
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


}
