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
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.HashMap;
import java.util.Map;

import org.apache.commons.lang3.StringUtils;
import org.apache.logging.log4j.Logger;
import org.apache.logging.log4j.LogManager;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellUtil;

/**
 * @author Sulfikar Ali Nazar
 */
public class ExcelCell {

    private Logger log = LogManager.getLogger(ExcelCell.class);
    Sheet sheet;
    Workbook wb;
    int row;
    int col;
    FileInputStream inputStream = null;
    FileOutputStream outputStream = null;
    String sCompleteFileName = null;
    Cell CELL = null;
    String xlFileExtension=null;
    private Cell getCELL() {
        return CELL;
    }

    private void setCELL(Cell cELL) {
        CELL = cELL;
    }

    private int getRow() {
        return row;
    }

    private void setRow(int row) {
        this.row = row;
    }

    private int getCol() {
        return col;
    }

    private void setCol(int col) {
        this.col = col;
    }

    public boolean setText(String sText) {


        return setText(sText, false);

    }
    public boolean setCellValue(String sText) {


        return setCellValue(sText, false);

    }
    /**
     * @author - Sulfikar Ali Nazar
     * @param sText
     * @param autoSizeColoumn
     * @return
     */
    public boolean setCellValue(String sText, boolean autoSizeColoumn) {

        try {

            this.getCELL().setCellValue(sText);
            if (autoSizeColoumn) {
                this.sheet.autoSizeColumn(col - 1);
            }

            //SaveWorkBook(inputStream, outputStream, sCompleteFileName);

        } catch (Exception e) {
            log.info("Error in setting text value " + e.toString());
        }

        return false;

    }
    /**
     * @author - Sulfikar Ali Nazar
     * @param sText
     * @param autoSizeColoumn
     * @return
     */
    public boolean setText(String sText, boolean autoSizeColoumn) {
        Cell cCell=null;
        String val="";
        CellStyle cStyle=null;


        try {
            cCell = this.getCELL();
           //

            Map<String, Object> properties = new HashMap<String, Object>();
            properties.put(CellUtil.DATA_FORMAT, BuiltinFormats.getBuiltinFormat("Text"));
            cCell.setCellValue(sText);
            CellUtil.setCellStyleProperties(cCell, properties);

            if (autoSizeColoumn) {
                this.sheet.autoSizeColumn(col - 1);
            }

            //SaveWorkBook(inputStream, outputStream, sCompleteFileName);

        } catch (Exception e) {
            log.info("Error in setting text value " + e.toString());
        }

        return false;

    }


    public void setFontColor(String fontColor){
        Cell cCell=null;
        String val="";
        CellStyle cStyle=null;
        try{
            cCell = this.getCELL();
            val = cCell.getStringCellValue();

            Font font = null;//this.wb.findFont(false, )

            if (font == null) { // Create new font
                font = cCell.getSheet().getWorkbook().createFont();

                font.setColor(Common.getColorCode(fontColor));

            }
            CellUtil.setFont(cCell, font);


        }catch (Exception e){
            log.info("Error in setting Font color " + e.toString());
        }


    }


    public void setFillColor(String fillColor){
        Cell cCell=null;
        String val="";
        CellStyle cStyle=null;
        try{
            cCell = this.getCELL();
            val = cCell.getStringCellValue();

            Map<String, Object> properties = new HashMap<String, Object>();
            properties.put(CellUtil.FILL_PATTERN, FillPatternType.SOLID_FOREGROUND);
            properties.put(CellUtil.FILL_FOREGROUND_COLOR, Common.getColorCode(fillColor));
            CellUtil.setCellStyleProperties(cCell, properties);

        }catch (Exception e){
            log.info("Error in setting Font color " + e.toString());
        }


    }
    public void setFullBorder(String fillColor){
        Cell cCell=null;
        String val="";
        CellStyle cStyle=null;
        try{
            cCell = this.getCELL();
            val = cCell.getStringCellValue();

            Map<String, Object> properties = new HashMap<String, Object>();
            properties.put(CellUtil.BORDER_LEFT, BorderStyle.MEDIUM);
            properties.put(CellUtil.BORDER_RIGHT, BorderStyle.MEDIUM);
            properties.put(CellUtil.BORDER_TOP, BorderStyle.MEDIUM);
            properties.put(CellUtil.BORDER_BOTTOM, BorderStyle.MEDIUM);
            properties.put(CellUtil.BOTTOM_BORDER_COLOR, Common.getColorCode(fillColor));
            properties.put(CellUtil.LEFT_BORDER_COLOR, Common.getColorCode(fillColor));
            properties.put(CellUtil.TOP_BORDER_COLOR, Common.getColorCode(fillColor));
            properties.put(CellUtil.RIGHT_BORDER_COLOR, Common.getColorCode(fillColor));

            CellUtil.setCellStyleProperties(cCell, properties);

        }catch (Exception e){
            log.info("Error in setting Font color " + e.toString());
        }


    }


    private CellStyle getCellStyle(CellStyle cStyle){
            return cStyle;
    }

    public String getValue(){
        return getTextValue();
    }

    public String getTextValue() {
        String val = null;
        try {
            inputStream = new FileInputStream(this.sCompleteFileName);
            this.wb = WorkbookFactory.create(inputStream);
            String strSheetName = sheet.getSheetName();
            this.sheet = this.wb.getSheet(strSheetName);
            // val = this.getCELL().getStringCellValue();//
            try {
                val = this.sheet.getRow(this.row - 1).getCell(this.col - 1).getStringCellValue();
            } catch (Exception e) {
                val = "";
            }

            inputStream.close();
            this.wb.close();
            // SaveWorkBook(inputStream, outputStream, sCompleteFileName);

        } catch (Exception e) {
            log.info("Error in setting text value " + e.toString());
        }

        return val;

    }

    private ExcelCell() {

    }
    protected ExcelCell(Sheet sheet, Workbook wb, int row, int col, FileInputStream inputStream,
                        FileOutputStream outputStream, String sCompleteFileName) {

        this.sheet = sheet;
        this.wb = wb;
        this.row = row;
        this.col = col;
        this.inputStream = inputStream;
        this.outputStream = outputStream;
        this.sCompleteFileName = sCompleteFileName;
        this.xlFileExtension=getFileExtension(sCompleteFileName);
        try {
            Row cRow = null;
            try {
                cRow = sheet.getRow(this.row - 1);

            } catch (Exception e) {

            }
            Row row1 = null;
            if (checkIfRowIsEmpty(cRow)) {
                row1 = sheet.createRow(this.row - 1);
            } else {
                row1 = cRow;
            }

            Cell cell1=null;
            try{
                cell1=row1.getCell(this.col - 1);
            }catch(Exception e){

            }


            if(checkIfCellIsEmpty(cell1)){
                cell1 = row1.createCell(this.col - 1);
            }

            this.setCELL(cell1);
            // SaveWorkBook(inputStream, outputStream, sCompleteFileName);

        } catch (Exception e) {
            log.info("Error in Cells constructor creation " + e.toString());
        }
    }

    /**
     * Saves the workbook to the specified file.
     *
     * @param inputStream The input stream to close
     * @param outputStream The output stream (not used)
     * @param sCompleteFileName The complete file path and name to save to
     */
    private void saveWorkBook(FileInputStream inputStream, FileOutputStream outputStream, String sCompleteFileName) {
        try {
            if (inputStream != null) {
                inputStream.close();
            }
            OutputStream fileOut = new FileOutputStream(sCompleteFileName);
            wb.write(fileOut);
            // wb.close();
            fileOut.close();
            log.info("Workbook saved successfully to: " + sCompleteFileName);
        } catch (Exception e) {
            log.error("Error in saving workbook: " + e.getMessage());
        }
    }



    private boolean checkIfRowIsEmpty(Row ROW) {
        try {
            if (ROW == null || ROW.getLastCellNum() <= 0) {
                return true;
            }

            for (int cellNum = ROW.getFirstCellNum(); cellNum < ROW.getLastCellNum(); cellNum++) {
                Cell cell = ROW.getCell(cellNum);
                if (cell != null && cell.getCellType() != CellType.BLANK && StringUtils.isNotBlank(cell.toString())) {
                    return false;
                }
            }
        }catch (Exception e){
            return true;
        }

        return true;
    }

    private boolean checkIfCellIsEmpty(Cell cell){
        if (cell != null && cell.getCellType() != CellType.BLANK && StringUtils.isNotBlank(cell.toString())) {
            return false;
        }
        return true;
    }


    private String getFileExtension(String sCompleteFilePath) {
        try {

                File file = new File(sCompleteFilePath);
                return sCompleteFilePath.substring(sCompleteFilePath.lastIndexOf("."));

        } catch (Exception e) {
            log.info("Error in getting file name " + e.toString());

        }

        return "";// return null on error
    }

}
