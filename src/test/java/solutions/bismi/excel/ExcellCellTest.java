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

import org.junit.Test;

import static org.junit.Assert.assertEquals;

/**
 * @author Sulfikar Ali Nazar
 */
public class ExcellCellTest {


    @Test
    public void aSetFontColor(){
        setColor("./resources/testdata/cellFormatCheck1.XLSX");
        setColor("./resources/testdata/cellFormatCheck1.XLS");
    }

    private void setColor(String strCompleteFileName){
        ExcelApplication xlApp =new ExcelApplication();
        ExcelWorkBook  xlbook=xlApp.createWorkBook(strCompleteFileName);
        int cnt=0;
        cnt=xlbook.getSheetCount();
        assertEquals(1,cnt);
        ExcelWorkSheet sh1 = xlbook.addSheet("Bismi1");
        sh1.activate();
        ExcelCell c1 = sh1.cell(3, 4);
        c1.setText("Alpha lion");
        c1.setFontColor("green");
        ExcelCell c2 = sh1.cell(5, 7);
        c2.setText("Beta");
        c2.setFontColor("RED");
        sh1.saveWorkBook();
        xlApp.closeAllWorkBooks();
    }




}
