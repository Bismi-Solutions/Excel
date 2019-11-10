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

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import java.util.ArrayList;
import java.util.List;

/**
 * @author - Sulfikar Ali Nazar
 * Excel Application base class to hold all application level operations.This is the rrot class.
 *
 */
public class ExcelApplication
{
    private Logger log=LogManager.getLogger(ExcelApplication.class);
    List<ExcelWorkBook> exlWorkBooks = new ArrayList<>();

    /**
     * @author - Sulfikar Ali Nazar
     */
    public ExcelApplication() {


    }

    /**
     * @author - Sulfikar Ali Nazar
     * @return
     */
    public List<ExcelWorkBook> getWorkbooks(){

        return exlWorkBooks;

    }

    /**
     * @author - Sulfikar Ali Nazar
     * @return
     */

    public int getOpenWorkbookCount() {

        return exlWorkBooks.size();
    }

    /**
     * @author - Sulfikar Ali Nazar
     *
     * @param sCompleteFileName
     * @return
     */
    public ExcelWorkBook createWorkBook(String sCompleteFileName) {
        log.info("Creating excel workbook");
        ExcelWorkBook exlWorkBook = new ExcelWorkBook();
        exlWorkBook.createWorkBook(sCompleteFileName);
        exlWorkBooks.add(exlWorkBook);//adding excel book to main application list
        return exlWorkBook;

    }

    /**
     * @author - Sulfikar Ali Nazar
     * @param sCompleteFileName
     * @return
     */
    public ExcelWorkBook openWorkbook(String sCompleteFileName) {
        ExcelWorkBook exlWorkBook = new ExcelWorkBook();
        exlWorkBook.openWorkbook(sCompleteFileName);
        exlWorkBooks.add(exlWorkBook);
        return exlWorkBook;
    }

    /**
     * @author - Sulfikar Ali Nazar
     * @param iIndex
     * @return
     */
    public ExcelWorkBook getWorkbook(int iIndex) {
        ExcelWorkBook tempBook = null;

        try {
            tempBook = this.exlWorkBooks.get(iIndex);
        } catch (Exception e) {
            log.info("Error during work book retrieval " + " - - " + e.toString());
        }

        return tempBook;
    }

    /**
     * @author - Sulfikar Ali Nazar
     * @param sWorkBookName
     * @return
     */
    public ExcelWorkBook getWorkbook(String sWorkBookName) {
        ExcelWorkBook tempBook = null;
        boolean fFound = false;
        for (ExcelWorkBook excelWorkBook : this.exlWorkBooks) {

            try {

                if (excelWorkBook.getExcelBookName().equalsIgnoreCase(sWorkBookName)) {
                    fFound = true;
                    tempBook = excelWorkBook;
                    break;
                }

            } catch (Exception e) {
                log.info("Error during work book retrieval " + sWorkBookName + " - - " + e.toString());
            }

        }

        if (!fFound)
            log.info("Work book not opened. Please open work book -- " + sWorkBookName);

        return tempBook;
    }

    /**
     * @author - Sulfikar Ali Nazar
     * @param iIndex
     */
    public void closeWorkBook(int iIndex) {
        ExcelWorkBook tempBook = null;
        try {
            tempBook = this.exlWorkBooks.get(iIndex);
            tempBook.closeWorkBook();
            this.exlWorkBooks.remove(iIndex);
        } catch (Exception e) {
            log.info("Error during work book retrieval " + " - - " + e.toString());
        }

    }

    /**
     * @author - Sulfikar Ali Nazar
     * @param sWorkBookName
     */
    public void closeWorkBook(String sWorkBookName) {
        int iIndex = -1;
        boolean fFound = false;
        for (ExcelWorkBook excelWorkBook : this.exlWorkBooks) {
            ++iIndex;
            try {

                if (excelWorkBook.getExcelBookName().equalsIgnoreCase(sWorkBookName)) {
                    excelWorkBook.closeWorkBook();

                    fFound = true;
                    break;
                }

            } catch (Exception e) {

            }

        }

        if (iIndex >= 0 && fFound == true) {
            this.exlWorkBooks.remove(iIndex);

        } else {
            log.info("Work book " + sWorkBookName + " not available for close operation");
        }

    }

    /**
     * @author - Sulfikar Ali Nazar
     */
    public void closeAllWorkBooks() {

        for (ExcelWorkBook excelWorkBook : this.exlWorkBooks) {

            try {
                excelWorkBook.closeWorkBook();

            } catch (Exception e) {
                log.info("Work Book " + excelWorkBook.excelBookName + " is already closed");
            }

        }

        this.exlWorkBooks.clear();

    }

}
