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
 * Represents a row in an Excel worksheet.
 * Provides methods for manipulating row content, formatting, and styling.
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
     * Creates a new ExcelRow with the specified parameters.
     * 
     * @param sheet The worksheet containing the row
     * @param wb The workbook containing the sheet
     * @param row The row number (1-based)
     * @param inputStream The input stream for the workbook file
     * @param outputStream The output stream for the workbook file
     * @param sCompleteFileName The complete file path and name of the workbook
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
            // Initialize row if needed
        } catch (Exception e) {
            log.error("Error in row constructor creation: " + e.getMessage());
        }
    }

    /**
     * Sets values for multiple cells in the row.
     * 
     * @param sValues Array of string values to set in the row
     * @param startColNumber The starting column number (0-based)
     * @param autoSizeColumns Whether to auto-size the columns
     */
    public void setRowValues(String[] sValues, int startColNumber, boolean autoSizeColumns) {
        Row cRow = null;

        try {
            try {
                cRow = sheet.getRow(this.row - 1);
            } catch (Exception e1) {
                // Row doesn't exist
            }

            if (checkIfRowIsEmpty(cRow)) {
                cRow = sheet.createRow(this.row - 1);
            }

            int i = startColNumber;
            for (String value : sValues) {
                Cell cell1 = null;
                try {
                    cell1 = cRow.getCell(i);
                } catch (Exception e) {
                    // Cell doesn't exist
                }

                if (checkIfCellIsEmpty(cell1)) {
                    cell1 = cRow.createCell(i);
                }

                cell1.setCellValue(value);
                if (autoSizeColumns) {
                    sheet.autoSizeColumn(i);
                }
                ++i;
            }
        } catch (Exception e) {
            log.error("Error in writing values to row: " + e.getMessage());
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
     * Sets the font color for a range of cells in the row.
     * 
     * @param fontColor The color name to set for the font
     * @param startColNumber The starting column number (1-based)
     * @param endColNumber The ending column number (exclusive)
     */
    public void setFontColor(String fontColor, int startColNumber, int endColNumber) {
        Row cRow = null;

        try {
            try {
                cRow = sheet.getRow(this.row - 1);
            } catch (Exception e1) {
                // Row doesn't exist
            }

            if (checkIfRowIsEmpty(cRow)) {
                cRow = sheet.createRow(this.row - 1);
            }

            for (int iCol = startColNumber; iCol < endColNumber; iCol++) {
                Cell cell1 = null;
                try {
                    cell1 = cRow.getCell(iCol - 1);
                } catch (Exception e) {
                    // Cell doesn't exist
                }

                if (checkIfCellIsEmpty(cell1)) {
                    cell1 = cRow.createCell(iCol - 1);
                }

                Font font = cell1.getSheet().getWorkbook().createFont();
                font.setColor(Common.getColorCode(fontColor));
                CellUtil.setFont(cell1, font);
            }
        } catch (Exception e) {
            log.error("Error in setting font color: " + e.getMessage());
        }
    }


    /**
     * Sets the font color for all cells in the row.
     * 
     * @param fontColor The color name to set for the font
     */
    public void setFontColor(String fontColor) {
        Row cRow = null;
        try {
            cRow = sheet.getRow(this.row - 1);
        } catch (Exception e1) {
            // Row doesn't exist
        }

        if (checkIfRowIsEmpty(cRow)) {
            cRow = sheet.createRow(this.row - 1);
        }

        int endColNumber = cRow.getLastCellNum();
        setFontColor(fontColor, 1, endColNumber + 1);
    }

    /**
     * Sets the background fill color for a range of cells in the row.
     * 
     * @param fillColor The color name to set for the cell background
     * @param startColNumber The starting column number (1-based)
     * @param endColNumber The ending column number (exclusive)
     */
    public void setFillColor(String fillColor, int startColNumber, int endColNumber) {
        Row cRow = null;

        try {
            try {
                cRow = sheet.getRow(this.row - 1);
            } catch (Exception e1) {
                // Row doesn't exist
            }

            if (checkIfRowIsEmpty(cRow)) {
                cRow = sheet.createRow(this.row - 1);
            }

            for (int iCol = startColNumber; iCol < endColNumber; iCol++) {
                Cell cell1 = null;
                try {
                    cell1 = cRow.getCell(iCol - 1);
                } catch (Exception e) {
                    // Cell doesn't exist
                }

                if (checkIfCellIsEmpty(cell1)) {
                    cell1 = cRow.createCell(iCol - 1);
                }

                Map<String, Object> properties = new HashMap<>();
                properties.put(CellUtil.FILL_PATTERN, FillPatternType.SOLID_FOREGROUND);
                properties.put(CellUtil.FILL_FOREGROUND_COLOR, Common.getColorCode(fillColor));
                CellUtil.setCellStyleProperties(cell1, properties);
            }
        } catch (Exception e) {
            log.error("Error in setting fill color: " + e.getMessage());
        }
    }



    /**
     * Sets the background fill color for all cells in the row.
     * 
     * @param fillColor The color name to set for the cell background
     */
    public void setFillColor(String fillColor) {
        Row cRow = null;
        try {
            cRow = sheet.getRow(this.row - 1);
        } catch (Exception e1) {
            // Row doesn't exist
        }

        if (checkIfRowIsEmpty(cRow)) {
            cRow = sheet.createRow(this.row - 1);
        }

        int endColNumber = cRow.getLastCellNum();
        setFillColor(fillColor, 1, endColNumber + 1);
    }

    /**
     * Sets a full border around a range of cells in the row.
     * 
     * @param borderColor The color name to set for the border
     * @param startColNumber The starting column number (1-based)
     * @param endColNumber The ending column number (exclusive)
     */
    public void setFullBorder(String borderColor, int startColNumber, int endColNumber) {
        Row cRow = null;

        try {
            try {
                cRow = sheet.getRow(this.row - 1);
            } catch (Exception e1) {
                // Row doesn't exist
            }

            if (checkIfRowIsEmpty(cRow)) {
                cRow = sheet.createRow(this.row - 1);
            }

            for (int iCol = startColNumber; iCol < endColNumber; iCol++) {
                Cell cell1 = null;
                try {
                    cell1 = cRow.getCell(iCol - 1);
                } catch (Exception e) {
                    // Cell doesn't exist
                }

                if (checkIfCellIsEmpty(cell1)) {
                    cell1 = cRow.createCell(iCol - 1);
                }

                Map<String, Object> properties = new HashMap<>();
                properties.put(CellUtil.BORDER_LEFT, BorderStyle.MEDIUM);
                properties.put(CellUtil.BORDER_RIGHT, BorderStyle.MEDIUM);
                properties.put(CellUtil.BORDER_TOP, BorderStyle.MEDIUM);
                properties.put(CellUtil.BORDER_BOTTOM, BorderStyle.MEDIUM);
                properties.put(CellUtil.BOTTOM_BORDER_COLOR, Common.getColorCode(borderColor));
                properties.put(CellUtil.LEFT_BORDER_COLOR, Common.getColorCode(borderColor));
                properties.put(CellUtil.TOP_BORDER_COLOR, Common.getColorCode(borderColor));
                properties.put(CellUtil.RIGHT_BORDER_COLOR, Common.getColorCode(borderColor));

                CellUtil.setCellStyleProperties(cell1, properties);
            }
        } catch (Exception e) {
            log.error("Error in setting border: " + e.getMessage());
        }
    }

    /**
     * Sets a full border around all cells in the row.
     * 
     * @param borderColor The color name to set for the border
     */
    public void setFullBorder(String borderColor) {
        Row cRow = null;
        try {
            cRow = sheet.getRow(this.row - 1);
        } catch (Exception e1) {
            // Row doesn't exist
        }

        if (checkIfRowIsEmpty(cRow)) {
            cRow = sheet.createRow(this.row - 1);
        }

        int endColNumber = cRow.getLastCellNum();
        setFullBorder(borderColor, 1, endColNumber + 1);
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
            log.debug("Workbook saved successfully to: " + sCompleteFileName);
        } catch (Exception e) {
            log.error("Error in saving workbook: " + e.getMessage());
        }
    }

    /**
     * Checks if a row is empty or contains only blank cells.
     * 
     * @param row The row to check
     * @return true if the row is empty or null, false otherwise
     */
    private boolean checkIfRowIsEmpty(Row row) {
        try {
            if (row == null || row.getLastCellNum() <= 0) {
                return true;
            }

            for (int cellNum = row.getFirstCellNum(); cellNum < row.getLastCellNum(); cellNum++) {
                Cell cell = row.getCell(cellNum);
                if (cell != null && cell.getCellType() != CellType.BLANK && StringUtils.isNotBlank(cell.toString())) {
                    return false;
                }
            }
        } catch (Exception e) {
            return true;
        }

        return true;
    }

    /**
     * Checks if a cell is empty or blank.
     * 
     * @param cell The cell to check
     * @return true if the cell is empty, null, or blank, false otherwise
     */
    private boolean checkIfCellIsEmpty(Cell cell) {
        if (cell != null && cell.getCellType() != CellType.BLANK && StringUtils.isNotBlank(cell.toString())) {
            return false;
        }
        return true;
    }





}
