/*
 * Copyright (c) 2019. Bismi Solutions
 *
 * https://bismi.solutions/
 * support@bismi.solutions
 * sulfikar.ali.nazar@gmail.com
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in all
 * copies or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
 * SOFTWARE.
 */

package solutions.bismi.excel;

import org.junit.jupiter.api.MethodOrderer;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.TestMethodOrder;


/**
 * @author Sulfikar Ali Nazar
 */
@TestMethodOrder(MethodOrderer.Alphanumeric.class)
public class ExcelWorkSheetTest {


    /**
     * @author - Sulfikar Ali Nazar
     *
     * Function to verify sheet acticvation. Which verifies by closing and reopening the work book.
     */
    @Test
    public void aVerifySheetActivationXLSX()
    {

        verifySheetActivation("./resources/testdata/sheetactivation.xlsx");
        verifySheetActivation("./resources/testdata/sheetactivation.xls");
    }
    /**
     * @author - Sulfikar Ali Nazar
     *
     * Function to verify sheet acticvation. Which verifies by closing and reopening the work book.
     */

    private void verifySheetActivation(String strCompleteFileName)
    {
        ExcelApplication xlApp =new ExcelApplication();
        ExcelWorkBook  xlbook=xlApp.createWorkBook(strCompleteFileName);
        int cnt=0;
        cnt=xlbook.getSheetCount();
        Assertions.assertEquals(1,cnt);
        xlbook.addSheet("Bismi1");
        cnt=xlbook.getSheetCount();
        xlbook.getExcelSheet("Bismi1").activate();
        xlApp.closeAllWorkBooks();
        Assertions.assertEquals(2,cnt);
        xlbook=xlApp.openWorkbook(strCompleteFileName);
        cnt=xlApp.getOpenWorkbookCount();
        Assertions.assertEquals(1,cnt);
        cnt=xlbook.getSheetCount();
        Assertions.assertEquals(2,cnt);
        xlbook.getExcelSheet("Bismi1").activate();
        String sheetName=xlbook.getActiveSheetName();
        Assertions.assertEquals("Bismi1",sheetName);
        xlApp.closeAllWorkBooks();
    }
    /**
     * @author - Sulfikar Ali Nazar
     * Functions to verify the row functionality in excel
     */
    @Test
    public void cAddRowsXLSX(){

        addRowsXL("./resources/testdata/rowCreation.xlsx");
        addRowsXL("./resources/testdata/rowCreation.xls");
    }

    /**
     * @author - Sulfikar Ali Nazar
     * Functions to verify the row functionality in excel
     */

    private void addRowsXL(String strCompleteFileName){

        ExcelApplication xlApp =new ExcelApplication();
        ExcelWorkBook  xlbook=xlApp.createWorkBook(strCompleteFileName);
        int cnt=0;
        cnt=xlbook.getSheetCount();
        Assertions.assertEquals(1,cnt);
        xlbook.addSheet("Bismi1");
        cnt=xlbook.getSheetCount();
        ExcelWorkSheet cSheet = xlbook.getExcelSheet("Bismi1");
        cSheet.activate();
        int cRownum = cSheet.getRowNumber();
        Assertions.assertEquals(0,cRownum);
        String[] arrRow= {"A","B","C","D"};
        cSheet.row(11).setRowValues(arrRow);
        cRownum = cSheet.getRowNumber();
        Assertions.assertEquals(11,cRownum);

        int cColNumber = cSheet.getColNumber();
        Assertions.assertEquals(4,cColNumber);

        cSheet.row(16).setRowValues(arrRow,5);
        cRownum = cSheet.getRowNumber();
        Assertions.assertEquals(16,cRownum);

        cColNumber = cSheet.getColNumber();
        Assertions.assertEquals(9,cColNumber);

        String[] arrRow1= {"A","10.5","Cwwwwwwwwwwwwwwwwwwwwww","D","02/04/2019"};
        cSheet.row(34).setRowValues(arrRow1,true);
        cRownum = cSheet.getRowNumber();
        Assertions.assertEquals(34,cRownum);
        cColNumber = cSheet.getColNumber();
        Assertions.assertEquals(9,cColNumber);
        cSheet.saveWorkBook();
        xlApp.closeAllWorkBooks();
    }

    /**
     * @author - Sulfikar Ali Nazar
     */
    @Test
    public void dAddRowsInMultipleSheet(){
        addRowsInMultipleSheet("./resources/testdata/multiRowMultiSheet.xlsx");
        addRowsInMultipleSheet("./resources/testdata/multiRowMultiSheet.xls");
    }

    private  void addRowsInMultipleSheet(String strMultipleFileName){
        ExcelApplication xlApp =new ExcelApplication();
        ExcelWorkBook  xlbook=xlApp.createWorkBook(strMultipleFileName);
        int cnt=0;
        cnt=xlbook.getSheetCount();
        Assertions.assertEquals(1,cnt);
        xlbook.addSheet("Bismi1");
        xlbook.addSheet("Bismi2");
        cnt=xlbook.getSheetCount();
        ExcelWorkSheet cSheet = xlbook.getExcelSheet("Bismi1");

        ExcelWorkSheet cSheet2 = xlbook.getExcelSheet("Bismi2");
        cSheet2.activate();

        String[] arrRow1= {"A","10.5","Cwwwwwwwwwwwwwwwwwwwwww","D","02/04/2019"};
        cSheet.row(34).setRowValues(arrRow1,true);
        int cRownum = cSheet.getRowNumber();
        Assertions.assertEquals(34,cRownum);
        int cColNumber = cSheet.getColNumber();
        Assertions.assertEquals(5,cColNumber);

        String[] arrRow2= {"A","10.5","Cwwwwwwwwwwwwwwwwwwwwww","D","02/04/2019"};
        cSheet2.row(4).setRowValues(arrRow2,true);
        cRownum = cSheet2.getRowNumber();
        Assertions.assertEquals(4,cRownum);
        cColNumber = cSheet2.getColNumber();
        Assertions.assertEquals(5,cColNumber);
        cSheet2.saveWorkBook();
        xlApp.closeAllWorkBooks();
        xlbook=xlApp.openWorkbook(strMultipleFileName);
        ExcelWorkSheet gSheet = xlbook.getExcelSheet("Bismi2");
        String val = gSheet.cell(4, 2).getTextValue();
        Assertions.assertEquals("10.5",val);

        xlApp.closeAllWorkBooks();
    }


    @Test
    public void eSetCellValues(){
        setCellValues("./resources/testdata/cellvalcheck.xlsx");
        setCellValues("./resources/testdata/cellvalcheck.xls");
    }

    private void setCellValues(String strCompleteFilePath){
        ExcelApplication xlApp =new ExcelApplication();
        ExcelWorkBook  xlbook=xlApp.createWorkBook(strCompleteFilePath);
        int cnt=0;
        cnt=xlbook.getSheetCount();
        Assertions.assertEquals(1,cnt);
        xlbook.addSheet("Bismi1");
        xlbook.addSheet("Bismi2");
        cnt=xlbook.getSheetCount();
        ExcelWorkSheet cSheet = xlbook.getExcelSheet("Bismi1");

        ExcelWorkSheet cSheet2 = xlbook.getExcelSheet("Bismi2");
        cSheet2.activate();
        cSheet.cell(5,4).setText("Alpha lion");
        cSheet2.cell(3,4).setText("Alpha lion");
        cSheet2.saveWorkBook();

        xlApp.closeAllWorkBooks();

        xlbook=xlApp.openWorkbook(strCompleteFilePath);
        ExcelWorkSheet gSheet = xlbook.getExcelSheet("Bismi2");
        String val = gSheet.cell(3, 4).getTextValue();
        Assertions.assertEquals("Alpha lion",val);

        xlApp.closeAllWorkBooks();

    }








}
