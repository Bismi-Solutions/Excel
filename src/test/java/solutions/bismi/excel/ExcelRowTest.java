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

public class ExcelRowTest {

    @Test
    public void aSetfontcolor(){
        setColor2("./resources/testdata/cellFormatCheckrow1.XLSX");
        setColor2("./resources/testdata/cellFormatCheckrow1.XLS");
    }

    private void setColor2(String strCompleteFileName){
        ExcelApplication xlApp =new ExcelApplication();
        ExcelWorkBook  xlbook=xlApp.createWorkBook(strCompleteFileName);
        int cnt=0;
        cnt=xlbook.getSheetCount();
        assertEquals(1,cnt);
        ExcelWorkSheet sh1 = xlbook.addSheet("Bismi1");
        sh1.activate();

        String[] arrRow= {"A","B","C","D","E"};
        sh1.row(11).setRowValues(arrRow);
        sh1.row(11).setFontColor("Red",1,3);
        sh1.row(11).setFontColor("green",3,7);
        sh1.row(2).setRowValues(arrRow);
        sh1.row(2).setFontColor("White");
        sh1.row(2).setFillColor("Green");
        sh1.row(2).setFullBorder("Red");
        sh1.row(11).setFullBorder("blue");
        sh1.row(5).setFullBorder("blue",1,10);
        sh1.saveWorkBook();
        xlApp.closeAllWorkBooks();
    }





}
