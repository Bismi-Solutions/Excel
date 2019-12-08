package solutions.bismi.excel;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;

import org.junit.FixMethodOrder;
import org.junit.Test;
import org.junit.runners.MethodSorters;

/**
 * Unit test for simple App.
 */
@FixMethodOrder(MethodSorters.NAME_ASCENDING)
public class ExcelApplicationTest
{
    /**
     * @author - Sulfikar Ali Nazar
     * Create xlsx work book
     */
    @Test
    public void aCreateExcelWorkBook()
    {
        bCreateXLSWorkBook("./resources/testdata/er.xlsx");
        bCreateXLSWorkBook("./resources/testdata/er.xls");
        //bCreateXLSWorkBook("./resources/testdata/er.xlsm");
    }

    private void bCreateXLSWorkBook(String strCompleteFileName)
    {
        ExcelApplication xlApp =new ExcelApplication();
        ExcelWorkBook  xlbook=xlApp.createWorkBook(strCompleteFileName);
        int cnt=xlApp.getOpenWorkbookCount();
        xlApp.closeAllWorkBooks();
        //Verify only one work book is created
        assertEquals(1,cnt);

    }

    /**
     * @author - Sulfikar Ali Nazar
     */
    @Test
    public void cOpenExcelXWorkbook(){
        dOpenXLSWorkbook("./resources/testdata/er.xlsx");
        dOpenXLSWorkbook("./resources/testdata/er.xls");
    }

    /**
     * @author - Sulfikar Ali Nazar
     */

    private void dOpenXLSWorkbook(String strCompleteFileName){
        ExcelApplication xlApp =new ExcelApplication();
        xlApp.openWorkbook(strCompleteFileName);
        int cnt=xlApp.getOpenWorkbookCount();
        xlApp.closeAllWorkBooks();
        //Verify only one work book is created
        assertEquals(1,cnt);
    }

    /**
     * @author - Sulfikar Ali Nazar
     */
    @Test
    public void eCloseWorkBookByIndex(){
        closeBookByindex("./resources/testdata/er.xls");
        closeBookByindex("./resources/testdata/er.xlsx");

    }

    private void closeBookByindex(String strCompleteFileName){
        ExcelApplication xlApp =new ExcelApplication();
        xlApp.openWorkbook(strCompleteFileName);
        int cnt=xlApp.getOpenWorkbookCount();
        assertEquals(1,cnt);
        xlApp.closeWorkBook(0);
        cnt=xlApp.getOpenWorkbookCount();
        assertEquals(0,cnt);
        assertEquals(0,cnt);
    }

    /**
     * @author - Sulfikar Ali Nazar
     */
    @Test
    public void fCloseWorkBookByName(){


        closeBookByName("./resources/testdata/er.xls");
        closeBookByName("./resources/testdata/er.xlsx");

    }

    private void closeBookByName(String strCompleteFileName ){
        ExcelApplication xlApp =new ExcelApplication();
        ExcelWorkBook xlBook = xlApp.openWorkbook(strCompleteFileName);
        int cnt=xlApp.getOpenWorkbookCount();
        assertEquals(1,cnt);
        xlApp.closeWorkBook(xlBook.getExcelBookName());
        cnt=xlApp.getOpenWorkbookCount();
        assertEquals(0,cnt);
    }

    /**
     * @author - Sulfikar Ali Nazar
     */
    @Test
    public void gCloseAllWorkBooks(){
        ExcelApplication xlApp =new ExcelApplication();
        xlApp.openWorkbook("./er.xls");
        int cnt=xlApp.getOpenWorkbookCount();
        assertEquals(1,cnt);

        //Verifying the close functionality of xlsx
        xlApp.openWorkbook("./resources/testdata/er.xlsx");
        cnt=xlApp.getOpenWorkbookCount();
        assertEquals(2,cnt);
        xlApp.closeAllWorkBooks();
        cnt=xlApp.getOpenWorkbookCount();
        assertEquals(0,cnt);


    }






}
