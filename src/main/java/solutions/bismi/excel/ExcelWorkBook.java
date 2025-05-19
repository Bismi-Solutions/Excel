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

/**
 * Represents an Excel workbook.
 * Provides methods for creating, opening, and manipulating Excel workbooks and their sheets.
 */
public class ExcelWorkBook {

    Workbook wb;
    String excelBookName = "";
    String excelBookPath = "";
    FileInputStream inputStream = null;
    FileOutputStream outputStream = null;
    File xlFile = null;
    // List to keep currently available sheets in excel
    protected List<ExcelWorkSheet> excelSheets = new ArrayList<>();
    private Logger log = LogManager.getLogger(ExcelWorkBook.class);
    String xlFileExtension = null;


    /**
     * Default constructor for ExcelWorkBook.
     */
    ExcelWorkBook() {
        // Default constructor
    }

    /**
     * Creates a new ExcelWorkBook with the specified parameters.
     * 
     * @param wb The Apache POI Workbook object
     * @param excelBookName The name of the Excel workbook
     * @param excelBookPath The path to the Excel workbook
     */
    ExcelWorkBook(Workbook wb, String excelBookName, String excelBookPath) {
        this.wb = wb;
        this.excelBookName = excelBookName;
        this.excelBookPath = excelBookPath;
    }

    /**
     * Closes the workbook and releases resources.
     * 
     * @return true if the workbook was closed successfully, false otherwise
     */
    protected boolean closeWorkBook() {
        try {
            this.wb.close();
            this.wb = null;
            log.info("Closed workbook: " + this.excelBookName);
            return true;
        } catch (IOException e) {
            log.error("Error in closing workbook: " + e.getMessage());
            return false;
        }
    }

    /**
     * Creates a new Excel workbook at the specified file path.
     * 
     * @param sCompleteFileName The complete file path and name for the new workbook
     * @return The newly created ExcelWorkBook object
     */
    protected ExcelWorkBook createWorkBook(String sCompleteFileName) {
        ExcelWorkBook xlWbook = new ExcelWorkBook();
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
                    fileOut.close();
                    log.info("Created XLSX file: " + this.excelBookName);
                } else if (fileExtn[1].equalsIgnoreCase("xls")) {
                    wb = new HSSFWorkbook();
                    OutputStream fileOut = new FileOutputStream(sCompleteFileName);
                    Sheet sheet1 = wb.createSheet("Sheet1");
                    wb.write(fileOut);
                    updateSheetList();
                    fileOut.close();
                    log.info("Created XLS file: " + this.excelBookName);
                }
            }
        } catch (Exception e) {
            log.error("Error creating workbook: " + e.getMessage());
        }

        xlWbook = new ExcelWorkBook(this.wb, this.excelBookName, this.excelBookPath);
        return xlWbook;
    }

    /**
     * Extracts the file name from a complete file path.
     * 
     * @param sCompleteFilePath The complete file path
     * @return The file name, or an empty string if an error occurs
     */
    private String getFileName(String sCompleteFilePath) {
        try {
            if (isValidPath(sCompleteFilePath)) {
                File file = new File(sCompleteFilePath);
                this.excelBookName = file.getName();
                this.xlFileExtension = this.excelBookName.substring(this.excelBookName.lastIndexOf("."));
                log.debug("Extracted file name: " + file.getName());
                return file.getName();
            }
        } catch (Exception e) {
            log.error("Error extracting file name: " + e.getMessage());
        }

        return ""; // Return empty string on error
    }

    /**
     * Validates a file path and creates any missing directories.
     * 
     * @param path The file path to validate
     * @return true if the path is valid, false otherwise
     */
    private boolean isValidPath(String path) {
        try {
            Paths.get(path);
            log.debug("Verified path is valid: " + path);
            File file = new File(path);

            if (!file.isDirectory()) {
                file = file.getParentFile();
                new File(file.getPath()).mkdirs();
                log.debug("Created missing directories");
                this.excelBookPath = file.getPath();
            } else {
                new File(file.getPath()).mkdirs();
                this.excelBookPath = file.getPath();
            }
            return true;
        } catch (InvalidPathException | NullPointerException ex) {
            log.error("Invalid path: " + ex.getMessage());
            return false;
        }
    }

    /**
     * Gets the name of the Excel workbook.
     * 
     * @return The name of the workbook
     */
    public String getExcelBookName() {
        return excelBookName;
    }

    /**
     * Gets the path to the Excel workbook.
     * 
     * @return The path to the workbook
     */
    public String getExcelBookPath() {
        return excelBookPath;
    }

    /**
     * Updates the list of worksheets in the workbook.
     * This method clears the existing list and repopulates it with all sheets from the workbook.
     */
    protected void updateSheetList() {
        excelSheets.clear();
        int sheetCount = this.wb.getNumberOfSheets();
        String path = excelBookPath + "\\" + excelBookName;

        try {
            for (int i = 0; i < sheetCount; i++) {
                String sheetName = this.wb.getSheetName(i);
                ExcelWorkSheet sheetTemp = new ExcelWorkSheet(
                    this.wb.getSheet(sheetName), 
                    sheetName, 
                    this.wb,
                    this.inputStream,
                    this.outputStream,
                    path
                );
                excelSheets.add(sheetTemp);
            }
            log.debug("Updated sheet list in workbook. Total sheets: " + sheetCount);
        } catch (Exception e) {
            log.error("Error updating sheet list: " + e.getMessage());
        }
    }


    /**
     * Opens an existing Excel workbook from the specified file path.
     * 
     * @param sCompleteFileName The complete file path and name of the workbook to open
     * @return The opened ExcelWorkBook object, or null if an error occurs
     */
    protected ExcelWorkBook openWorkbook(String sCompleteFileName) {
        ExcelWorkBook xlWbook = null;
        log.info("Opening Excel workbook: " + sCompleteFileName);

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
                    log.debug("Set active sheet to: " + this.wb.getSheetName(activeSheetIndex));
                }
            }
        } catch (Exception e) {
            log.error("Error opening workbook: " + e.getMessage());
        }

        return xlWbook;
    }

    /**
     * Adds a new worksheet to the workbook.
     * 
     * @param sSheetName The name of the new worksheet
     * @return The newly created ExcelWorkSheet object
     */
    public ExcelWorkSheet addSheet(String sSheetName) {
        ExcelWorkSheet wSheet = new ExcelWorkSheet();
        wSheet.wb = this.wb;
        ExcelWorkSheet tempSheet = wSheet.addSheet(sSheetName);

        try {
            if (inputStream != null) {
                inputStream.close();
            }
            outputStream = new FileOutputStream(this.xlFile);
            this.wb.write(outputStream);
            outputStream.close();
            log.debug("Added new sheet: " + sSheetName);
        } catch (Exception e) {
            log.error("Error creating worksheet: " + e.getMessage());
        }

        excelSheets.add(tempSheet);
        return getExcelSheet(sSheetName);
    }

    /**
     * Adds multiple worksheets to the workbook.
     * 
     * @param strArrSheets Array of sheet names to add
     */
    public void addSheets(String[] strArrSheets) {
        ExcelWorkSheet wSheet = new ExcelWorkSheet();
        wSheet.wb = this.wb;
        wSheet.addSheets(strArrSheets);

        try {
            if (inputStream != null) {
                inputStream.close();
            }
            outputStream = new FileOutputStream(this.xlFile);
            this.wb.write(outputStream);
            outputStream.close();
            updateSheetList();
            log.debug("Added " + strArrSheets.length + " new sheets");
        } catch (Exception e) {
            log.error("Error creating worksheets: " + e.getMessage());
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
            log.debug("Error in getting sheet " + sSheetName);
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
            log.debug("Error in getting sheet index" + iIndex);
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
            log.debug("Error in getting sheet index" + iIndex);
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
