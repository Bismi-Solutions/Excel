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
public class ExcelRow {
    private Logger log = LogManager.getLogger(ExcelRow.class);
    private Sheet sheet;
    private Workbook wb;
    private int row;

    FileInputStream inputStream = null;
    FileOutputStream outputStream = null;
    String sCompleteFileName = null;
    Row ROW=null;
    String xlFileExtension=null;
    /**
     * @author - Sulfikar Ali Nazar
     * @param sheet
     * @param wb
     * @param row
     * @param inputStream
     * @param outputStream
     * @param sCompleteFileName
     */
    protected ExcelRow(Sheet sheet, Workbook wb, int row, FileInputStream inputStream,
                       FileOutputStream outputStream, String sCompleteFileName) {

        this.sheet = sheet;
        this.wb = wb;
        this.row = row;
        this.inputStream = inputStream;
        this.outputStream = outputStream;
        this.sCompleteFileName = sCompleteFileName;
        try {





            //SaveWorkBook(inputStream, outputStream, sCompleteFileName);
        } catch (Exception e) {
            log.info("Error in Cells constructor creation " + e.toString());
        }



    }

    /**
     * @author - Sulfikar Ali Nazar
     * @param sValues
     * @param startColNumber
     * @param autoSizeColoumns
     */
    public void setRowValues(String[] sValues,int startColNumber,boolean autoSizeColoumns) {

        Row cRow = null;

        try {

            try {
                cRow = sheet.getRow(this.row - 1);

            } catch (Exception e1) {

            }
            if (checkIfRowIsEmpty(cRow)) {
                cRow = sheet.createRow(this.row - 1);
            }
            int i=startColNumber;
            for (String sValu : sValues) {
                Cell cell1=null;
                try{
                    cell1=cRow.getCell(i);
                }catch (Exception e){

                }
                if(checkIfCellIsEmpty(cell1)){
                    cell1 = cRow.createCell(i);
                }
                cell1.setCellValue(sValu);
                if(autoSizeColoumns) {
                    sheet.autoSizeColumn(i);
                }
                ++i;

            }
            //SaveWorkBook(inputStream, outputStream, sCompleteFileName);

        } catch (Exception e) {
            log.info("Error in writing record to excel" + e.toString());
        }

    }

    public void setRowValues(String[] sValues,int startColNumber) {
        setRowValues(sValues, startColNumber, false);
    }
    public void setRowValues(String[] sValues,boolean autoSizeColoumns) {
        setRowValues(sValues, 0, autoSizeColoumns);
    }
    public void setRowValues(String[] sValues) {
        setRowValues(sValues, 0, false);
    }

    /**
     * @author - Sulfikar Ali Nazar
     *
     */
    public void setFontColor(String fontColor,int startColNumber, int endColNumber) {

        Row cRow = null;

        try {

            try {
                cRow = sheet.getRow(this.row - 1);
            } catch (Exception e1) {

            }
            if (checkIfRowIsEmpty(cRow)) {
                cRow = sheet.createRow(this.row - 1);
            }


            for (int iCol=startColNumber;iCol<endColNumber;iCol++) {
                Cell cell1=null;
                try{
                    cell1=cRow.getCell(iCol-1);
                }catch (Exception e){

                }
                if(checkIfCellIsEmpty(cell1)){
                    cell1 = cRow.createCell(iCol-1);
                }

                Font font = null;//this.wb.findFont(false, )

                if (font == null) { // Create new font
                    font = cell1.getSheet().getWorkbook().createFont();

                    font.setColor(Common.getColorCode(fontColor));

                }
                CellUtil.setFont(cell1, font);



            }
            //SaveWorkBook(inputStream, outputStream, sCompleteFileName);

        } catch (Exception e) {
            log.info("Error in setting font color record to excel" + e.toString());
        }

    }


    public void setFontColor(String fontColor) {
        Row cRow = null;



            try {
                cRow = sheet.getRow(this.row - 1);
            } catch (Exception e1) {

            }
            if (checkIfRowIsEmpty(cRow)) {
                cRow = sheet.createRow(this.row - 1);
            }

            int endColNumber=cRow.getLastCellNum();
            setFontColor(fontColor, 1, endColNumber+1);

    }

    /**
     * @author - Sulfikar Ali Nazar
     * @param fillColor
     * @param startColNumber
     * @param endColNumber
     */
    public void setFillColor(String fillColor,int startColNumber, int endColNumber) {

        Row cRow = null;

        try {

            try {
                cRow = sheet.getRow(this.row - 1);
            } catch (Exception e1) {

            }
            if (checkIfRowIsEmpty(cRow)) {
                cRow = sheet.createRow(this.row - 1);
            }


            for (int iCol=startColNumber;iCol<endColNumber;iCol++) {
                Cell cell1=null;
                try{
                    cell1=cRow.getCell(iCol-1);
                }catch (Exception e){

                }
                if(checkIfCellIsEmpty(cell1)){
                    cell1 = cRow.createCell(iCol-1);
                }

                Map<String, Object> properties = new HashMap<String, Object>();
                properties.put(CellUtil.FILL_PATTERN, FillPatternType.SOLID_FOREGROUND);
                properties.put(CellUtil.FILL_FOREGROUND_COLOR, Common.getColorCode(fillColor));
                CellUtil.setCellStyleProperties(cell1, properties);



            }
            //SaveWorkBook(inputStream, outputStream, sCompleteFileName);

        } catch (Exception e) {
            log.info("Error in setting font color record to excel" + e.toString());
        }

    }



    public void setFillColor(String fillColor) {
        Row cRow = null;



        try {
            cRow = sheet.getRow(this.row - 1);
        } catch (Exception e1) {

        }
        if (checkIfRowIsEmpty(cRow)) {
            cRow = sheet.createRow(this.row - 1);
        }

        int endColNumber=cRow.getLastCellNum();
        setFillColor(fillColor, 1, endColNumber+1);

    }

    /**
     * @author Sulfikar Ali Nazar
     * @param fillColor
     * @param startColNumber
     * @param endColNumber
     */

    public void setFullBorder(String fillColor,int startColNumber, int endColNumber) {

        Row cRow = null;

        try {

            try {
                cRow = sheet.getRow(this.row - 1);
            } catch (Exception e1) {

            }
            if (checkIfRowIsEmpty(cRow)) {
                cRow = sheet.createRow(this.row - 1);
            }


            for (int iCol=startColNumber;iCol<endColNumber;iCol++) {
                Cell cell1=null;
                try{
                    cell1=cRow.getCell(iCol-1);
                }catch (Exception e){

                }
                if(checkIfCellIsEmpty(cell1)){
                    cell1 = cRow.createCell(iCol-1);
                }

                Map<String, Object> properties = new HashMap<String, Object>();
                properties.put(CellUtil.BORDER_LEFT, BorderStyle.MEDIUM);
                properties.put(CellUtil.BORDER_RIGHT, BorderStyle.MEDIUM);
                properties.put(CellUtil.BORDER_TOP, BorderStyle.MEDIUM);
                properties.put(CellUtil.BORDER_BOTTOM, BorderStyle.MEDIUM);
                properties.put(CellUtil.BOTTOM_BORDER_COLOR, Common.getColorCode(fillColor));
                properties.put(CellUtil.LEFT_BORDER_COLOR, Common.getColorCode(fillColor));
                properties.put(CellUtil.TOP_BORDER_COLOR, Common.getColorCode(fillColor));
                properties.put(CellUtil.RIGHT_BORDER_COLOR, Common.getColorCode(fillColor));

                CellUtil.setCellStyleProperties(cell1, properties);



            }
            //SaveWorkBook(inputStream, outputStream, sCompleteFileName);

        } catch (Exception e) {
            log.info("Error in setting font color record to excel" + e.toString());
        }

    }

    public void setFullBorder(String fillColor) {
        Row cRow = null;



        try {
            cRow = sheet.getRow(this.row - 1);
        } catch (Exception e1) {

        }
        if (checkIfRowIsEmpty(cRow)) {
            cRow = sheet.createRow(this.row - 1);
        }

        int endColNumber=cRow.getLastCellNum();
        setFullBorder(fillColor, 1, endColNumber+1);

    }


    private void SaveWorkBook(FileInputStream inputStream, FileOutputStream outputStream, String sCompleteFileName) {
        try {

            if (inputStream != null)
                inputStream.close();
            OutputStream fileOut = new FileOutputStream(sCompleteFileName);
            wb.write(fileOut);
            // wb.close();
            fileOut.close();

        } catch (Exception e) {
            log.info("Error in saving record " + e.toString());
        }

    }

    /**
     * @author - Sulfikar Ali Nazar
     *
     * @param ROW
     * @return
     */
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





}
