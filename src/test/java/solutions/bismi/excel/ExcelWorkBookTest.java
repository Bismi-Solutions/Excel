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
    public void aAddSheetXLSXWorkBook()
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
    public void bAddSheetXLSWorkBook()
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


}
