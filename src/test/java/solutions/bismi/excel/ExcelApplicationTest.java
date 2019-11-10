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
    public void aCreateXLSXWorkBook()
    {
        ExcelApplication xlApp =new ExcelApplication();
        ExcelWorkBook  xlbook=xlApp.createWorkBook("./resources/testdata/er.xlsx");
        int cnt=xlApp.getOpenWorkbookCount();
        xlApp.closeAllWorkBooks();
        //Verify only one work book is created
        assertEquals(1,cnt);
    }
    /**
     * @author - Sulfikar Ali Nazar
     * Create xls work book
     */
    @Test
    public void bCreateXLSWorkBook()
    {
        ExcelApplication xlApp =new ExcelApplication();
        ExcelWorkBook  xlbook=xlApp.createWorkBook("./resources/testdata/er.xls");
        int cnt=xlApp.getOpenWorkbookCount();
        xlApp.closeAllWorkBooks();
        //Verify only one work book is created
        assertEquals(1,cnt);

    }

    /**
     * @author - Sulfikar Ali Nazar
     */
    @Test
    public void cOpenXLSXWorkbook(){
        ExcelApplication xlApp =new ExcelApplication();
        xlApp.openWorkbook("./resources/testdata/er.xlsx");
        int cnt=xlApp.getOpenWorkbookCount();
        xlApp.closeAllWorkBooks();
        //Verify only one work book is created
        assertEquals(1,cnt);


    }

    /**
     * @author - Sulfikar Ali Nazar
     */
    @Test
    public void dOpenXLSWorkbook(){
        ExcelApplication xlApp =new ExcelApplication();
        xlApp.openWorkbook("./resources/testdata/er.xls");
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
        ExcelApplication xlApp =new ExcelApplication();
        xlApp.openWorkbook("./resources/testdata/er.xls");
        int cnt=xlApp.getOpenWorkbookCount();
        assertEquals(1,cnt);
        xlApp.closeWorkBook(0);
        cnt=xlApp.getOpenWorkbookCount();
        assertEquals(0,cnt);

        //Verifying the close functionality of xlsx
        xlApp.openWorkbook("./resources/testdata/er.xlsx");
        cnt=xlApp.getOpenWorkbookCount();
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
        ExcelApplication xlApp =new ExcelApplication();
        xlApp.openWorkbook("./resources/testdata/er.xls");
        int cnt=xlApp.getOpenWorkbookCount();
        assertEquals(1,cnt);
        xlApp.closeWorkBook("er.xls");
        cnt=xlApp.getOpenWorkbookCount();
        assertEquals(0,cnt);

        //Verifying the close functionality of xlsx
        xlApp.openWorkbook("./resources/testdata/er.xlsx");
        cnt=xlApp.getOpenWorkbookCount();
        assertEquals(1,cnt);
        xlApp.closeWorkBook("er.xlsx");
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
