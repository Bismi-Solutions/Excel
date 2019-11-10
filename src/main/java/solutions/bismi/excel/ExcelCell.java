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

import org.apache.commons.lang3.StringUtils;
import org.apache.logging.log4j.Logger;
import org.apache.logging.log4j.LogManager;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
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

    protected Cell getCELL() {
        return CELL;
    }

    protected void setCELL(Cell cELL) {
        CELL = cELL;
    }

    protected int getRow() {
        return row;
    }

    protected void setRow(int row) {
        this.row = row;
    }

    protected int getCol() {
        return col;
    }

    protected void setCol(int col) {
        this.col = col;
    }

    public boolean SetText(String sText) {

        try {

            this.getCELL().setCellValue(sText);

            SaveWorkBook(inputStream, outputStream, sCompleteFileName);

        } catch (Exception e) {
            log.info("Error in setting text value " + e.toString());
        }

        return false;

    }

    public boolean SetText(String sText, boolean autoSizeColoumn) {

        try {

            this.getCELL().setCellValue(sText);
            if (autoSizeColoumn) {
                this.sheet.autoSizeColumn(col - 1);
            }

            SaveWorkBook(inputStream, outputStream, sCompleteFileName);

        } catch (Exception e) {
            log.info("Error in setting text value " + e.toString());
        }

        return false;

    }

    public String GetTextValue() {
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

            Cell cell1 = row1.createCell(this.col - 1);
            this.setCELL(cell1);
            // SaveWorkBook(inputStream, outputStream, sCompleteFileName);

        } catch (Exception e) {
            log.info("Error in Cells constructor creation " + e.toString());
        }
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


    @SuppressWarnings("deprecation")
    private boolean checkIfRowIsEmpty(Row ROW) {
        if (ROW == null) {
            return true;
        }
        if (ROW.getLastCellNum() <= 0) {
            return true;
        }
        for (int cellNum = ROW.getFirstCellNum(); cellNum < ROW.getLastCellNum(); cellNum++) {
            Cell cell = ROW.getCell(cellNum);
            if (cell != null && cell.getCellTypeEnum() != CellType.BLANK && StringUtils.isNotBlank(cell.toString())) {
                return false;
            }
        }
        return true;
    }



}
