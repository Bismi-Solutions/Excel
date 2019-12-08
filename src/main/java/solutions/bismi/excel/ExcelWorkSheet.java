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
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
/**
 * @author Sulfikar Ali Nazar
 */
public class ExcelWorkSheet {
    int colNumber;
    String xlFileExtension=null;

    FileInputStream inputStream=null;

    private Logger log = LogManager.getLogger(ExcelWorkSheet.class);
    FileOutputStream outputStream=null;
    private int rowNumber;
    String sCompleteFileName=null;
    Sheet sheet;
    String sheetName = null;
    Workbook wb;

    /**
     *
     */
    ExcelWorkSheet() {

    }

    /**
     * @author - Sulfikar Ali Nazar
     * @param sheet
     * @param log
     * @param sheetName
     * @param wb
     */
    ExcelWorkSheet(Sheet sheet, Logger log, String sheetName, Workbook wb) {

        this.sheet = sheet;
        this.log = log;
        this.sheetName = sheetName;
        this.wb = wb;

    }

    /**
     * @author - Sulfikar Ali Nazar
     * @param sheet
     * @param sheetName
     * @param wb
     * @param inputStream
     * @param outputStream
     * @param sCompleteFileName
     */
    protected ExcelWorkSheet(Sheet sheet, String sheetName, Workbook wb,FileInputStream inputStream,FileOutputStream outputStream, String sCompleteFileName) {

        this.sheet = sheet;

        this.sheetName = sheetName;
        this.wb = wb;
        this.inputStream=inputStream;
        this.outputStream=outputStream;
        this.sCompleteFileName=sCompleteFileName;

    }

    /**
     * @author - Sulfikar Ali Nazar
     * Function to activate excel sheet
     */
    public void activate() {

        this.wb.setActiveSheet(this.wb.getSheetIndex(this.sheet.getSheetName()));
        this.wb.setSelectedTab(this.wb.getSheetIndex(this.sheet.getSheetName()));

        saveWorkBook( inputStream, outputStream, sCompleteFileName);

    }

    /**
     * @author - Sulfikar Ali Nazar
     * @param sSheetName
     * @return
     */
    protected ExcelWorkSheet addSheet(String sSheetName) {
        ExcelWorkSheet shExcel = null;
        try {
            sheet = this.wb.createSheet(sSheetName);

            this.sheetName = sSheetName;
            log.info("Created sheet " + sSheetName);
            shExcel = new ExcelWorkSheet(this.sheet, this.log, this.sheetName, this.wb);

        } catch (Exception e) {
            log.info("Error in creating sheet " + sSheetName + e.toString());

        }

        return shExcel;
    }

    /**
     * @author - Sulfikar Ali Nazar
     * @param inputStream
     * @param outputStream
     * @param sCompleteFileName
     */
    private void saveWorkBook(FileInputStream inputStream,FileOutputStream outputStream, String sCompleteFileName ) {
        try {
            if (inputStream != null)
                inputStream.close();
            OutputStream fileOut = new FileOutputStream(sCompleteFileName);
            wb.write(fileOut);
            //wb.close();
            fileOut.close();


        } catch (Exception e) {
            log.info("Error in saving record " + e.toString());
        }


    }

    public void saveWorkBook( ) {
        saveWorkBook(this.inputStream,this.outputStream, this.sCompleteFileName );

    }





    /**
     * @author - Sulfikar Ali Nazar
     * @param strArrSheets
     * @return
     */
    protected void addSheets(String[] strArrSheets) {

        String eShName=null;
        ExcelWorkSheet shExcel = null;
        try {

            for (String sSheetName : strArrSheets) {
                sheet = this.wb.createSheet(sSheetName);
                this.sheetName = sSheetName;
                log.info("Created sheet " + sSheetName);
                shExcel = new ExcelWorkSheet(this.sheet, this.log, this.sheetName, this.wb);

            }


        } catch (Exception e) {
            log.info("Error in creating sheet " + eShName + e.toString());

        }


    }

    /**
     * @author - Sulfikar Ali Nazar
     * @return
     */
    public boolean clearContents() {

        try {

            for (int i = this.sheet.getLastRowNum(); i >= 0; i--) {

                try {
                    sheet.removeRow(sheet.getRow(i));

                } catch (Exception e) {
                    log.info("Empty row removal error " + i);
                }

            }

            SaveWorkBook( inputStream, outputStream, sCompleteFileName);
            log.info("Cleared sheet contents successfully " );
            return true;

        } catch (Exception e) {
            log.info("Error in clearing sheet contents " + e.toString());
            return false;
        }



    }

    public int getColNumber() {
        int rn=this.sheet.getLastRowNum();
        int maxcol = 0;
        for(int i=0;i<=rn;i++){
            try{
                if(this.sheet.getRow(i).getLastCellNum()>maxcol){
                    maxcol=this.sheet.getRow(i).getLastCellNum();
                }
            }catch(Exception e){

            }

        }

        return maxcol;
    }


    public int getRowNumber() {
        return this.sheet.getLastRowNum()+1;
    }

    public String getSheetName() {
        return sheetName;
    }

    /**
     * @author - Sulfikar Ali Nazar
     * @param inputStream
     * @param outputStream
     * @param sCompleteFileName
     */
    private void SaveWorkBook(FileInputStream inputStream,FileOutputStream outputStream, String sCompleteFileName ) {
        try {

            if (inputStream != null)
                inputStream.close();
            OutputStream fileOut = new FileOutputStream(sCompleteFileName);
            wb.write(fileOut);
            //wb.close();
            fileOut.close();


        } catch (Exception e) {
            log.info("Error in saving record " + e.toString());
        }


    }


    /**
     * @author - Sulfikar Ali Nazar
     * @param row
     * @param col
     * @return
     */
    public ExcelCell cell(int row,int col) {
        ExcelCell cells =null;
        try {
            cells=new ExcelCell(this.sheet,this.wb,row,col,this.inputStream,this.outputStream,this.sCompleteFileName);

        } catch (Exception e) {
            log.info("Error Excel cells operation " + e.toString());
        }

        return cells;

    }

    /**
     * @author - Sulfikar Ali Nazar
     * @param row
     * @return
     */
    public ExcelRow row(int row) {
        ExcelRow rows =null;
        try {
            rows=new ExcelRow(this.sheet,this.wb,row,this.inputStream,this.outputStream,this.sCompleteFileName);

        } catch (Exception e) {
            log.info("Error Excel rows operation " + e.toString());
        }

        return rows;

    }

}
