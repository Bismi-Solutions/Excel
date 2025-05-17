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

/**
 * @author Sulfikar Ali Nazar
 */
@TestMethodOrder(MethodOrderer.Alphanumeric.class)
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
        Assertions.assertEquals(1,cnt);
        ExcelWorkSheet sh1 = xlbook.addSheet("Bismi1");
        sh1.activate();
        ExcelCell c1 = sh1.cell(3, 4);
        c1.setText("Alpha lion");
        c1.setFontColor("blue");
        ExcelCell c2 = sh1.cell(5, 7);
        c2.setText("Beta");
        c2.setFontColor("RED");
        sh1.saveWorkBook();
        xlApp.closeAllWorkBooks();
    }
    @Test
    public void bSetFillcolor(){
        setColor2("./resources/testdata/cellFormatCheck2.XLSX");
        setColor2("./resources/testdata/cellFormatCheck2.XLS");
    }

    private void setColor2(String strCompleteFileName){
        ExcelApplication xlApp =new ExcelApplication();
        ExcelWorkBook  xlbook=xlApp.createWorkBook(strCompleteFileName);
        int cnt=0;
        cnt=xlbook.getSheetCount();
        Assertions.assertEquals(1,cnt);
        ExcelWorkSheet sh1 = xlbook.addSheet("Bismi1");
        sh1.activate();
        sh1.cell(10,10).setText("TestColor");
        sh1.cell(10,10).setFontColor("blue");
        sh1.cell(10,10).setFillColor("yellow");
        sh1.cell(1,1).setFillColor("GREEN");
        sh1.cell(3,17).setFullBorder("Red");
        sh1.cell(12,13).setCellValue("123");
        sh1.cell(12,13).setCellValue("0");
        sh1.cell(13,13).setText("0");
        sh1.cell(12,22).setCellValue("0.123",true);
        sh1.saveWorkBook();
        xlApp.closeAllWorkBooks();
    }

    @Test
    public void cTestFormulas() {
        testFormulas("./resources/testdata/formulaTest.XLSX");
        testFormulas("./resources/testdata/formulaTest.XLS");
    }

    private void testFormulas(String strCompleteFileName) {
        ExcelApplication xlApp = new ExcelApplication();
        ExcelWorkBook xlbook = xlApp.createWorkBook(strCompleteFileName);
        ExcelWorkSheet sh1 = xlbook.addSheet("FormulaTest");
        sh1.activate();

        // Set up some test data
        sh1.cell(1, 1).setCellValue("Item");
        sh1.cell(1, 2).setCellValue("Quantity");
        sh1.cell(1, 3).setCellValue("Price");
        sh1.cell(1, 4).setCellValue("Total");

        // Format the header row
        for (int i = 1; i <= 4; i++) {
            sh1.cell(1, i).setFontStyle(true, false, false);
            sh1.cell(1, i).setFillColor("GREY_25_PERCENT");
        }

        // Add some data rows
        sh1.cell(2, 1).setCellValue("Apples");
        sh1.cell(2, 2).setNumericValue(10);
        sh1.cell(2, 3).setNumericValue(1.5);

        sh1.cell(3, 1).setCellValue("Oranges");
        sh1.cell(3, 2).setNumericValue(5);
        sh1.cell(3, 3).setNumericValue(2.0);

        sh1.cell(4, 1).setCellValue("Bananas");
        sh1.cell(4, 2).setNumericValue(8);
        sh1.cell(4, 3).setNumericValue(1.2);

        // Add formulas for calculating totals
        sh1.cell(2, 4).setFormula("B2*C2");
        sh1.cell(3, 4).setFormula("B3*C3");
        sh1.cell(4, 4).setFormula("B4*C4");

        // Add a sum formula
        sh1.cell(5, 1).setCellValue("Total");
        sh1.cell(5, 1).setFontStyle(true, false, false);
        sh1.cell(5, 4).setFormula("SUM(D2:D4)");

        // Format the total cell
        sh1.cell(5, 4).setFontStyle(true, false, false);
        sh1.cell(5, 4).setFillColor("LIGHT_YELLOW");

        // Format the number cells
        for (int row = 2; row <= 5; row++) {
            for (int col = 2; col <= 4; col++) {
                if (col == 3 || col == 4) {
                    sh1.cell(row, col).setNumberFormat("$#,##0.00");
                }
            }
        }

        // Set column alignment
        for (int row = 2; row <= 5; row++) {
            sh1.cell(row, 1).setHorizontalAlignment("LEFT");
            sh1.cell(row, 2).setHorizontalAlignment("CENTER");
            sh1.cell(row, 3).setHorizontalAlignment("RIGHT");
            sh1.cell(row, 4).setHorizontalAlignment("RIGHT");
        }

        sh1.saveWorkBook();
        xlApp.closeAllWorkBooks();
    }

    @Test
    public void dTestCellMerging() {
        testCellMerging("./resources/testdata/mergeTest.XLSX");
        testCellMerging("./resources/testdata/mergeTest.XLS");
    }

    private void testCellMerging(String strCompleteFileName) {
        ExcelApplication xlApp = new ExcelApplication();
        ExcelWorkBook xlbook = xlApp.createWorkBook(strCompleteFileName);
        ExcelWorkSheet sh1 = xlbook.addSheet("MergeTest");
        sh1.activate();

        // Create a title and merge cells
        sh1.cell(1, 1).setCellValue("Sales Report");
        sh1.cell(1, 1).setFontStyle(true, false, false);
        sh1.cell(1, 1).setHorizontalAlignment("CENTER");

        // Merge cells for the title
        sh1.mergeCells(1, 1, 1, 5);

        // Add a subtitle
        sh1.cell(2, 1).setCellValue("Q1 2023");
        sh1.cell(2, 1).setFontStyle(false, true, false);
        sh1.cell(2, 1).setHorizontalAlignment("CENTER");

        // Merge cells for the subtitle
        sh1.mergeCells(2, 2, 1, 5);

        // Verify that cells are merged
        boolean isMerged = sh1.isCellMerged(1, 3);
        Assertions.assertTrue(isMerged, "Cell (1,3) should be part of a merged region");

        // Test unmerging
        sh1.unmergeCells(2, 2, 1, 5);
        isMerged = sh1.isCellMerged(2, 3);
        Assertions.assertFalse(isMerged, "Cell (2,3) should not be merged after unmerging");

        sh1.saveWorkBook();
        xlApp.closeAllWorkBooks();
    }







}
