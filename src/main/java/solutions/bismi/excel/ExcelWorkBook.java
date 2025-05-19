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
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.InvalidPathException;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;

import org.apache.logging.log4j.Logger;
import org.apache.logging.log4j.LogManager;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelWorkBook {

    Workbook wb;
    String excelBookName = "";
    String excelBookPath = "";
    FileInputStream inputStream = null;
    FileOutputStream outputStream = null;
    File xlFile = null;
    //List to keep currently available sheets in excel
    protected List<ExcelWorkSheet> excelSheets = new ArrayList<>();
    /** The log. */
    private Logger log = LogManager.getLogger(ExcelWorkBook.class);
    String xlFileExtension=null;


    /**
     * @author - Sulfikar Ali Nazar
     */
    ExcelWorkBook() {

    }

    /**
     * @author - Sulfikar Ali Nazar
     * @param wb
     * @param excelBookName
     * @param excelBookPath
     */
    ExcelWorkBook(Workbook wb, String excelBookName, String excelBookPath) {
        this.wb = wb;
        this.excelBookName = excelBookName;
        this.excelBookPath = excelBookPath;

    }

    /**
     * @author - Sulfikar Ali Nazar
     * @return
     */
    protected boolean closeWorkBook() {
        try {

            this.wb.close();
            this.wb = null;
            log.info("Closed work book " + this.excelBookName);
            return true;
        } catch (IOException e) {
            log.info("Error in closing work book " + e.toString());
            return false;
        }

    }

    /**
     * @author - Sulfikar Ali Nazar
     * @param sCompleteFileName
     * @return
     */
    protected ExcelWorkBook createWorkBook(String sCompleteFileName) {
        ExcelWorkBook xlWbook = new ExcelWorkBook();// WorkbookFactory
        try {
            if (isValidPath(sCompleteFileName)) {
                this.xlFile = new File(sCompleteFileName);
                this.excelBookName = getFileName(sCompleteFileName);
                String[] fileExtn = this.excelBookName.split("[.]");
                if (fileExtn[1].equalsIgnoreCase("xlsx") || fileExtn[1].equalsIgnoreCase("xlsm")) {
                    wb = new XSSFWorkbook();

                    OutputStream fileOut = new FileOutputStream(sCompleteFileName);
                    Sheet sheet1 = wb.createSheet("Sheet1");
                    wb.write(fileOut);
                    updateSheetList();
                  //  wb.close();
                    fileOut.close();
                    log.info("Created file " + this.excelBookName);

                } else if (fileExtn[1].equalsIgnoreCase("xls")) {
                    wb = new HSSFWorkbook();
                    OutputStream fileOut = new FileOutputStream(sCompleteFileName);
                    Sheet sheet1 = wb.createSheet("Sheet1");

                    wb.write(fileOut);
                    updateSheetList();
                  //  wb.close();
                    fileOut.close();

                    log.info("Created file " + this.excelBookName);
                }
            }

        } catch (Exception e) {
            log.info("Exception occured while creating work book " + e.toString());
        }
        xlWbook = new ExcelWorkBook(this.wb, this.excelBookName, this.excelBookPath);

        return xlWbook;

    }

    /**
     * @author - Sulfikar Ali Nazar
     * @param sCompleteFilePath
     * @return
     */
    private String getFileName(String sCompleteFilePath) {
        try {
            if (isValidPath(sCompleteFilePath)) {
                File file = new File(sCompleteFilePath);
                log.info("GetFileName: -> Returning file name " + file.getName());

                this.excelBookName = file.getName();
                this.xlFileExtension = this.excelBookName.substring(this.excelBookName.lastIndexOf("."));
                return file.getName();

            }
        } catch (Exception e) {
            log.info("Error in getting file name " + e.toString());

        }

        return "";// return null on error
    }

    /**
     * @author - Sulfikar Ali Nazar
     * @param path
     * @return
     */

    private boolean isValidPath(String path) {
        try {
            Paths.get(path);
            log.info("Verified path " + path + " is valid");
            File file = new File(path);

            if (!file.isDirectory()) {
                file = file.getParentFile();
                new File(file.getPath()).mkdirs();
                log.info("Created missing directories");
                this.excelBookPath = file.getPath();
            } else {
                new File(file.getPath()).mkdirs();
                this.excelBookPath = file.getPath();
            }
        } catch (InvalidPathException | NullPointerException ex) {
            log.info("Returned error message during -- " + ex.toString());
            return false;
        }
        return true;
    }

    /**
     * @author - Sulfikar Ali Nazar
     * @return
     */
    public String getExcelBookName() {
        return excelBookName;
    }

    /**
     * @author - Sulfikar Ali Nazar
     * @return
     */
    public String getExcelBookPath() {
        return excelBookPath;
    }

    protected void updateSheetList(){
        //excelSheets
        excelSheets.clear();
        int sheetcount=this.wb.getNumberOfSheets();
        String pth= excelBookPath+"\\"+excelBookName;
        try {
            for (int i=0;i<sheetcount;i++){
                ExcelWorkSheet shTemp=null;
                String shName=this.wb.getSheetName(i);
                shTemp=new ExcelWorkSheet(this.wb.getSheet(shName), shName, this.wb,this.inputStream,this.outputStream,pth);
                excelSheets.add(shTemp);
            }
            log.info("updateSheetList: -->  Updated sheet count in workbook");
        }catch (Exception e){
            log.error("updateSheetList:  -->  Error in retrieving excel sheets");
        }

    }


    /**
     * @author - Sulfikar Ali Nazar
     * @param sCompleteFileName
     * @return
     */
    protected ExcelWorkBook openWorkbook(String sCompleteFileName) {

        ExcelWorkBook xlWbook = null;
        log.info("Opening Excel work book");

        try {

            if (isValidPath(sCompleteFileName)) {
                this.xlFile = new File(sCompleteFileName);
                this.excelBookName = getFileName(sCompleteFileName);
                inputStream = new FileInputStream(this.xlFile);
                this.wb = WorkbookFactory.create(inputStream);

                // Get the active sheet index before creating the ExcelWorkBook
                int activeSheetIndex = this.wb.getActiveSheetIndex();

                xlWbook = new ExcelWorkBook(this.wb, this.excelBookName, this.excelBookPath);
                updateSheetList();

                // Ensure the active sheet is properly set
                if (activeSheetIndex >= 0 && activeSheetIndex < this.wb.getNumberOfSheets()) {
                    this.wb.setActiveSheet(activeSheetIndex);
                    log.info("Set active sheet to: " + this.wb.getSheetName(activeSheetIndex));
                }
            }

        } catch (Exception e) {
            log.info("Error in opening excel work book" + e.toString());
        }

        return xlWbook;

    }

    /**
     * @author - Sulfikar Ali Nazar
     * @param sSheetName
     * @return
     */
    public ExcelWorkSheet addSheet(String sSheetName) {
        ExcelWorkSheet wSheet = new ExcelWorkSheet();
        wSheet.wb = this.wb;
        ExcelWorkSheet tempSheet = wSheet.addSheet(sSheetName);

        try {
            if (inputStream != null)
                inputStream.close();
            outputStream = new FileOutputStream(this.xlFile);
            this.wb.write(outputStream);
            outputStream.close();

        } catch (Exception e) {
            log.info("Error in creating excel sheet" + e.toString());

        }
        excelSheets.add(tempSheet);

        return getExcelSheet(sSheetName);

    }

    /**
     * @author - Sulfikar Ali Nazar
     * @param strArrSheets
     * @return
     */
    public void addSheets(String[] strArrSheets) {

        ExcelWorkSheet wSheet = new ExcelWorkSheet();
        wSheet.wb = this.wb;
        wSheet.addSheets(strArrSheets);

        try {
            if (inputStream != null)
                inputStream.close();
            outputStream = new FileOutputStream(this.xlFile);
            this.wb.write(outputStream);
            outputStream.close();
            updateSheetList();

        } catch (Exception e) {
            log.error("Error in creating excel sheet" + e.toString());

        }

    }

    protected Workbook getWb() {
        return wb;
    }

    protected void setWb(Workbook wb) {
        this.wb = wb;
    }

    protected FileInputStream getInputStream() {
        return inputStream;
    }

    protected void setInputStream(FileInputStream inputStream) {
        this.inputStream = inputStream;
    }

    protected FileOutputStream getOutputStream() {
        return outputStream;
    }

    protected void setOutputStream(FileOutputStream outputStream) {
        this.outputStream = outputStream;
    }

    protected File getXlFile() {
        return xlFile;
    }

    protected void setXlFile(File xlFile) {
        this.xlFile = xlFile;
    }




    protected Logger getLog() {
        return log;
    }

    protected void setLog(Logger log) {
        this.log = log;
    }

    protected List<ExcelWorkSheet> getExcelSheets() {
        return excelSheets;
    }

    protected void setExcelSheets(List<ExcelWorkSheet> excelSheets) {
        this.excelSheets = excelSheets;
    }

    protected void setExcelBookName(String excelBookName) {
        this.excelBookName = excelBookName;
    }

    protected void setExcelBookPath(String excelBookPath) {
        this.excelBookPath = excelBookPath;
    }

    /**
     * @author - Sulfikar Ali Nazar
     * @param sSheetName
     * @return
     */
    public ExcelWorkSheet getExcelSheet(String sSheetName) {

        Sheet sheet1=null;
        ExcelWorkSheet shTemp=null;
        try {
            sheet1=this.wb.getSheet(sSheetName);

            if(sheet1 != null)   {
                int index = this.wb.getSheetIndex(sheet1);
                // this.wb.removeSheetAt(index);

                String pth= excelBookPath+"\\"+excelBookName;
                shTemp=new ExcelWorkSheet(this.wb.getSheet(sSheetName), this.wb.getSheet(sSheetName).getSheetName(), this.wb,this.inputStream,this.outputStream,pth);
            }

        } catch (Exception e) {
            log.info("Error in getting sheet " + sSheetName);
        }


        return shTemp;
    }

    /**
     * @author - Sulfikar Ali Nazar
     * @param iIndex
     * @return
     */
    public ExcelWorkSheet getExcelSheet(int iIndex) {

        Sheet sheet1=null;
        ExcelWorkSheet shTemp=null;
        try {

            sheet1=this.wb.getSheetAt(iIndex);
            if(sheet1 != null)   {
                int index = this.wb.getSheetIndex(sheet1);
                // this.wb.removeSheetAt(index);
                String pth= excelBookPath+"\\"+excelBookName;
                shTemp=new ExcelWorkSheet(this.wb.getSheetAt(iIndex), this.wb.getSheetAt(iIndex).getSheetName(), this.wb,this.inputStream,this.outputStream,pth);
            }
        } catch (Exception e) {
            log.info("Error in getting sheet index" + iIndex);
        }

        return shTemp;//new ExcelWorkSheet(this.wb.getSheetAt(iIndex), this.wb.getSheetAt(iIndex).getSheetName(), this.wb);
    }

    public ExcelWorkSheet getActiveSheet() {
        int iIndex=getActiveSheetIndex();
        Sheet sheet1=null;
        ExcelWorkSheet shTemp=null;
        try {

            sheet1=this.wb.getSheetAt(iIndex);
            if(sheet1 != null)   {
                int index = this.wb.getSheetIndex(sheet1);
                // this.wb.removeSheetAt(index);
                String pth= excelBookPath+"\\"+excelBookName;
                shTemp=new ExcelWorkSheet(this.wb.getSheetAt(iIndex), this.wb.getSheetAt(iIndex).getSheetName(), this.wb,this.inputStream,this.outputStream,pth);
            }
        } catch (Exception e) {
            log.info("Error in getting sheet index" + iIndex);
        }

        return shTemp;//new ExcelWorkSheet(this.wb.getSheetAt(iIndex), this.wb.getSheetAt(iIndex).getSheetName(), this.wb);
    }




    public int getSheetCount(){

        return excelSheets.size();
    }

    public String getActiveSheetName(){

        return this.wb.getSheetName(this.wb.getActiveSheetIndex());
    }

    /**
     * @author - Sulfikar Ali Nazar
     * @return
     */
    public int getActiveSheetIndex(){
        return this.wb.getActiveSheetIndex();
    }


    /**
     * Saves the workbook to the file
     * 
     * @return true if successful, false otherwise
     */
    public boolean saveWorkbook() {
        try {
            if (inputStream != null)
                inputStream.close();
            outputStream = new FileOutputStream(this.xlFile);
            this.wb.write(outputStream);
            outputStream.close();
            log.info("Workbook saved successfully: " + this.excelBookName);
            return true;
        } catch (Exception e) {
            log.error("Error saving workbook: " + e.getMessage());
            return false;
        }
    }


}
