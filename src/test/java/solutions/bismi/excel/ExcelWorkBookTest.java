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

import org.junit.FixMethodOrder;
import org.junit.Test;
import org.junit.runners.MethodSorters;

import java.util.List;

import static org.junit.Assert.assertEquals;

/**
 * @author Sulfikar Ali Nazar
 */
@FixMethodOrder(MethodSorters.NAME_ASCENDING)
public class ExcelWorkBookTest {
    /**
     * @author - Sulfikar Ali Nazar
     */
    @Test
    public void aAddSheetXLSXWorkSheet()
    {
        ExcelApplication xlApp =new ExcelApplication();
        ExcelWorkBook  xlbook=xlApp.createWorkBook("./resources/testdata/addSheet.xlsx");
        int cnt=0;
        cnt=xlbook.getSheetCount();
        assertEquals(1,cnt);
        xlbook.addSheet("Bismi1");
        cnt=xlbook.getSheetCount();
        xlApp.closeAllWorkBooks();
        assertEquals(2,cnt);

    }
    /**
     * @author - Sulfikar Ali Nazar
     */
    @Test
    public void bAddSheetXLSWorkSheet()
    {
        ExcelApplication xlApp =new ExcelApplication();
        ExcelWorkBook  xlbook=xlApp.createWorkBook("./resources/testdata/addSheet.xls");
        int cnt=0;
        cnt=xlbook.getSheetCount();
        assertEquals(1,cnt);
        xlbook.addSheet("Bismi1");
        cnt=xlbook.getSheetCount();
        xlApp.closeAllWorkBooks();
        assertEquals(2,cnt);

    }

    /**
     * @author - Sulfikar Ali Nazar
     */
    @Test
    public void cAddSheetXLSXWorkSheets()
    {
        ExcelApplication xlApp =new ExcelApplication();
        ExcelWorkBook  xlbook=xlApp.createWorkBook("./resources/testdata/addSheets.xlsx");
        int cnt=0;
        cnt=xlbook.getSheetCount();
        assertEquals(1,cnt);

        String[] shArray={"Bismi","Solutions","Sulfikar","Ali","Nazar"};
        xlbook.addSheets(shArray);
        cnt=xlbook.getSheetCount();
        xlApp.closeAllWorkBooks();
        assertEquals(6,cnt);
    }


    /**
     * @author - Sulfikar Ali Nazar
     */
    @Test
    public void dAddSheetXLSWorkSheets()
    {
        ExcelApplication xlApp =new ExcelApplication();
        ExcelWorkBook  xlbook=xlApp.createWorkBook("./resources/testdata/addSheets.xls");
        int cnt=0;
        cnt=xlbook.getSheetCount();
        assertEquals(1,cnt);

        String[] shArray={"Bismi","Solutions","Sulfikar","Ali","Nazar"};
        xlbook.addSheets(shArray);
        cnt=xlbook.getSheetCount();
        xlApp.closeAllWorkBooks();
        assertEquals(6,cnt);
    }


    /**
     * @author - Sulfikar Ali Nazar
     */
    @Test
    public void eGetXLSXWorkSheet()
    {
        ExcelApplication xlApp =new ExcelApplication();
        ExcelWorkBook  xlbook=xlApp.createWorkBook("./resources/testdata/getSheet.xlsx");
        int cnt=0;
        cnt=xlbook.getSheetCount();
        assertEquals(1,cnt);

        ExcelWorkSheet xlsht=xlbook.addSheet("Bismi");
        String shName=xlsht.getSheetName();


        String shName1=xlbook.getExcelSheet(0).getSheetName();
        xlApp.closeAllWorkBooks();
        assertEquals("Bismi",shName);
        assertEquals("Sheet1",shName1);
    }

    /**
     * @author - Sulfikar Ali Nazar
     */
    @Test
    public void fGetXLSWorkSheet()
    {
        ExcelApplication xlApp =new ExcelApplication();
        ExcelWorkBook  xlbook=xlApp.createWorkBook("./resources/testdata/getSheet.xls");
        int cnt=0;
        cnt=xlbook.getSheetCount();
        assertEquals(1,cnt);

        ExcelWorkSheet xlsht=xlbook.addSheet("Bismi");
        String shName=xlsht.getSheetName();
        String shName1=xlbook.getExcelSheet(0).getSheetName();
        xlApp.closeAllWorkBooks();
        assertEquals("Sheet1",shName1);
        assertEquals("Bismi",shName);
    }

    /**
     * @author - Sulfikar Ali Nazar
     */
    @Test
    public void gGetXLSXWorkSheets()
    {
        ExcelApplication xlApp =new ExcelApplication();
        ExcelWorkBook  xlbook=xlApp.createWorkBook("./resources/testdata/getSheets.xlsx");
        int cnt=0;
        cnt=xlbook.getSheetCount();
        assertEquals(1,cnt);

        String[] shArray={"Bismi","Solutions","Sulfikar","Ali","Nazar"};
        xlbook.addSheets(shArray);
        List<ExcelWorkSheet> sheetList = xlbook.getExcelSheets();
        ExcelWorkSheet sht3 = sheetList.get(2);
        String shName=sht3.getSheetName();
        xlApp.closeAllWorkBooks();
        assertEquals("Solutions",shName);
    }










}
