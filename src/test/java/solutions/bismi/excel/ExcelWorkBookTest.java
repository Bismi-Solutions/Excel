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

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.Assertions;

import org.junit.jupiter.api.MethodOrderer;
import org.junit.jupiter.api.TestMethodOrder;

import java.util.List;



/**
 * @author Sulfikar Ali Nazar
 */
@TestMethodOrder(MethodOrderer.Alphanumeric.class)
public class ExcelWorkBookTest {
    /**
     * @author - Sulfikar Ali Nazar
     */
    @Test
    public void aAddSheetXLSXWorkSheet()
    {
        addSheetXLWorkSheet("./resources/testdata/addSheet.xlsx");
        addSheetXLWorkSheet("./resources/testdata/addSheet.xls");
    }
    /**
     * @author - Sulfikar Ali Nazar
     */

    private void addSheetXLWorkSheet(String strCompleteFileName)
    {
        ExcelApplication xlApp =new ExcelApplication();
        ExcelWorkBook  xlbook=xlApp.createWorkBook(strCompleteFileName);
        int cnt=0;
        cnt=xlbook.getSheetCount();
        Assertions.assertEquals(1,cnt);
        xlbook.addSheet("Bismi1");
        cnt=xlbook.getSheetCount();
        xlApp.closeAllWorkBooks();
        Assertions.assertEquals(2,cnt);

    }

    /**
     * @author - Sulfikar Ali Nazar
     */
    @Test
    public void bAddSheetXLSXWorkSheets()
    {
        dAddSheetXLWorkSheets("./resources/testdata/addSheets.xlsx");
        dAddSheetXLWorkSheets("./resources/testdata/addSheets.xls");
    }


    /**
     * @author - Sulfikar Ali Nazar
     */

    private void dAddSheetXLWorkSheets(String strCompleteFileName)
    {
        ExcelApplication xlApp =new ExcelApplication();
        ExcelWorkBook  xlbook=xlApp.createWorkBook(strCompleteFileName);
        int cnt=0;
        cnt=xlbook.getSheetCount();
        Assertions.assertEquals(1,cnt);

        String[] shArray={"Bismi","Solutions","Sulfikar","Ali","Nazar"};
        xlbook.addSheets(shArray);
        cnt=xlbook.getSheetCount();
        xlApp.closeAllWorkBooks();
        Assertions.assertEquals(6,cnt);
    }


    /**
     * @author - Sulfikar Ali Nazar
     */
    @Test
    public void eGetXLSXWorkSheet()
    {

        getXLSWorkSheet("./resources/testdata/getSheet.xls");
        getXLSWorkSheet("./resources/testdata/getSheet.xlsx");
    }

    /**
     * @author - Sulfikar Ali Nazar
     */

    private void getXLSWorkSheet(String strCompleteFileName)
    {
        ExcelApplication xlApp =new ExcelApplication();
        ExcelWorkBook  xlbook=xlApp.createWorkBook(strCompleteFileName);
        int cnt=0;
        cnt=xlbook.getSheetCount();
        Assertions.assertEquals(1,cnt);

        ExcelWorkSheet xlsht=xlbook.addSheet("Bismi");
        String shName=xlsht.getSheetName();
        String shName1=xlbook.getExcelSheet(0).getSheetName();
        xlApp.closeAllWorkBooks();
        Assertions.assertEquals("Sheet1",shName1);
        Assertions.assertEquals("Bismi",shName);
    }

    /**
     * @author - Sulfikar Ali Nazar
     */

    private void gGetXLSXWorkSheets(String strCompleteFileName)
    {
        ExcelApplication xlApp =new ExcelApplication();
        ExcelWorkBook  xlbook=xlApp.createWorkBook(strCompleteFileName);
        int cnt=0;
        cnt=xlbook.getSheetCount();
        Assertions.assertEquals(1,cnt);

        String[] shArray={"Bismi","Solutions","Sulfikar","Ali","Nazar"};
        xlbook.addSheets(shArray);
        List<ExcelWorkSheet> sheetList = xlbook.getExcelSheets();
        ExcelWorkSheet sht3 = sheetList.get(2);
        String shName=sht3.getSheetName();
        xlApp.closeAllWorkBooks();
        Assertions.assertEquals("Solutions",shName);
    }

    /**
     * @author - Sulfikar Ali Nazar
     */
    @Test
    public  void hgetExcelWokSheets(){
        gGetXLSXWorkSheets("./resources/testdata/getSheets.xlsx");
        gGetXLSXWorkSheets("./resources/testdata/getSheets.xls");
    }









}
