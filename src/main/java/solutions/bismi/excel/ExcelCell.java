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
 * Represents a cell in an Excel worksheet.
 * Provides methods for manipulating cell content, formatting, and styling.
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
     * Sets a text value in the cell.
     * 
     * @param sText The text value to set
     * @param autoSizeColoumn Whether to auto-size the column
     * @return true if successful, false otherwise
     */
    public boolean setCellValue(String sText, boolean autoSizeColoumn) {
        try {
            this.getCELL().setCellValue(sText);
            if (autoSizeColoumn) {
                this.sheet.autoSizeColumn(col - 1);
            }
            return true;
        } catch (Exception e) {
            log.error("Error in setting cell value: " + e.getMessage());
            return false;
        }
    }

    /**
     * Sets a numeric value in the cell
     * 
     * @param value The numeric value to set
     * @param autoSizeColumn Whether to auto-size the column
     * @return true if successful, false otherwise
     */
    public boolean setNumericValue(double value, boolean autoSizeColumn) {
        try {
            this.getCELL().setCellValue(value);
            if (autoSizeColumn) {
                this.sheet.autoSizeColumn(col - 1);
            }
            return true;
        } catch (Exception e) {
            log.error("Error in setting numeric value: " + e.getMessage());
            return false;
        }
    }

    /**
     * Sets a numeric value in the cell
     * 
     * @param value The numeric value to set
     * @return true if successful, false otherwise
     */
    public boolean setNumericValue(double value) {
        return setNumericValue(value, false);
    }

    /**
     * Sets a date value in the cell
     * 
     * @param date The date value to set
     * @param autoSizeColumn Whether to auto-size the column
     * @return true if successful, false otherwise
     */
    public boolean setDateValue(java.util.Date date, boolean autoSizeColumn) {
        try {
            this.getCELL().setCellValue(date);
            if (autoSizeColumn) {
                this.sheet.autoSizeColumn(col - 1);
            }
            return true;
        } catch (Exception e) {
            log.error("Error in setting date value: " + e.getMessage());
            return false;
        }
    }

    /**
     * Sets a date value in the cell
     * 
     * @param date The date value to set
     * @return true if successful, false otherwise
     */
    public boolean setDateValue(java.util.Date date) {
        return setDateValue(date, false);
    }

    /**
     * Sets a formula in the cell
     * 
     * @param formula The formula to set (without the leading '=')
     * @param autoSizeColumn Whether to auto-size the column
     * @return true if successful, false otherwise
     */
    public boolean setFormula(String formula, boolean autoSizeColumn) {
        try {
            this.getCELL().setCellFormula(formula);
            if (autoSizeColumn) {
                this.sheet.autoSizeColumn(col - 1);
            }
            return true;
        } catch (Exception e) {
            log.error("Error in setting formula: " + e.getMessage());
            return false;
        }
    }

    /**
     * Sets a formula in the cell
     * 
     * @param formula The formula to set (without the leading '=')
     * @return true if successful, false otherwise
     */
    public boolean setFormula(String formula) {
        return setFormula(formula, false);
    }
    /**
     * Sets a text value in the cell with text formatting.
     * 
     * @param sText The text value to set
     * @param autoSizeColoumn Whether to auto-size the column
     * @return true if successful, false otherwise
     */
    public boolean setText(String sText, boolean autoSizeColoumn) {
        Cell cCell = null;

        try {
            cCell = this.getCELL();

            Map<String, Object> properties = new HashMap<>();
            properties.put(CellUtil.DATA_FORMAT, BuiltinFormats.getBuiltinFormat("Text"));
            cCell.setCellValue(sText);
            CellUtil.setCellStyleProperties(cCell, properties);

            if (autoSizeColoumn) {
                this.sheet.autoSizeColumn(col - 1);
            }

            return true;
        } catch (Exception e) {
            log.error("Error in setting text value: " + e.getMessage());
            return false;
        }
    }


    /**
     * Sets the font color for the cell.
     * 
     * @param fontColor The color name to set for the font
     */
    public void setFontColor(String fontColor) {
        Cell cCell = null;
        try {
            cCell = this.getCELL();

            Font font = cCell.getSheet().getWorkbook().createFont();
            font.setColor(Common.getColorCode(fontColor));

            CellUtil.setFont(cCell, font);
        } catch (Exception e) {
            log.error("Error in setting font color: " + e.getMessage());
        }
    }


    /**
     * Sets the background fill color for the cell.
     * 
     * @param fillColor The color name to set for the cell background
     */
    public void setFillColor(String fillColor) {
        Cell cCell = null;
        try {
            cCell = this.getCELL();

            Map<String, Object> properties = new HashMap<>();
            properties.put(CellUtil.FILL_PATTERN, FillPatternType.SOLID_FOREGROUND);
            properties.put(CellUtil.FILL_FOREGROUND_COLOR, Common.getColorCode(fillColor));
            CellUtil.setCellStyleProperties(cCell, properties);
        } catch (Exception e) {
            log.error("Error in setting fill color: " + e.getMessage());
        }
    }

    /**
     * Sets the horizontal alignment of the cell content
     * 
     * @param alignment The alignment to set (LEFT, CENTER, RIGHT, etc.)
     * @return true if successful, false otherwise
     */
    public boolean setHorizontalAlignment(String alignment) {
        try {
            Cell cCell = this.getCELL();
            Map<String, Object> properties = new HashMap<String, Object>();

            HorizontalAlignment hAlign;
            switch(alignment.toUpperCase()) {
                case "LEFT":
                    hAlign = HorizontalAlignment.LEFT;
                    break;
                case "CENTER":
                    hAlign = HorizontalAlignment.CENTER;
                    break;
                case "RIGHT":
                    hAlign = HorizontalAlignment.RIGHT;
                    break;
                case "JUSTIFY":
                    hAlign = HorizontalAlignment.JUSTIFY;
                    break;
                case "FILL":
                    hAlign = HorizontalAlignment.FILL;
                    break;
                default:
                    hAlign = HorizontalAlignment.GENERAL;
            }

            properties.put(CellUtil.ALIGNMENT, hAlign);
            CellUtil.setCellStyleProperties(cCell, properties);
            return true;
        } catch (Exception e) {
            log.error("Error in setting horizontal alignment: " + e.getMessage());
            return false;
        }
    }

    /**
     * Sets the vertical alignment of the cell content
     * 
     * @param alignment The alignment to set (TOP, CENTER, BOTTOM, etc.)
     * @return true if successful, false otherwise
     */
    public boolean setVerticalAlignment(String alignment) {
        try {
            Cell cCell = this.getCELL();
            Map<String, Object> properties = new HashMap<String, Object>();

            VerticalAlignment vAlign;
            switch(alignment.toUpperCase()) {
                case "TOP":
                    vAlign = VerticalAlignment.TOP;
                    break;
                case "CENTER":
                    vAlign = VerticalAlignment.CENTER;
                    break;
                case "BOTTOM":
                    vAlign = VerticalAlignment.BOTTOM;
                    break;
                case "JUSTIFY":
                    vAlign = VerticalAlignment.JUSTIFY;
                    break;
                case "DISTRIBUTED":
                    vAlign = VerticalAlignment.DISTRIBUTED;
                    break;
                default:
                    vAlign = VerticalAlignment.CENTER;
            }

            properties.put(CellUtil.VERTICAL_ALIGNMENT, vAlign);
            CellUtil.setCellStyleProperties(cCell, properties);
            return true;
        } catch (Exception e) {
            log.error("Error in setting vertical alignment: " + e.getMessage());
            return false;
        }
    }

    /**
     * Sets the number format for the cell
     * 
     * @param formatPattern The format pattern to use (e.g., "#,##0.00", "dd/mm/yyyy")
     * @return true if successful, false otherwise
     */
    public boolean setNumberFormat(String formatPattern) {
        try {
            Cell cCell = this.getCELL();
            DataFormat format = wb.createDataFormat();
            Map<String, Object> properties = new HashMap<String, Object>();
            properties.put(CellUtil.DATA_FORMAT, format.getFormat(formatPattern));
            CellUtil.setCellStyleProperties(cCell, properties);
            return true;
        } catch (Exception e) {
            log.error("Error in setting number format: " + e.getMessage());
            return false;
        }
    }

    /**
     * Sets the font style (bold, italic, etc.)
     * 
     * @param bold true to make the font bold
     * @param italic true to make the font italic
     * @param underline true to underline the text
     * @return true if successful, false otherwise
     */
    public boolean setFontStyle(boolean bold, boolean italic, boolean underline) {
        try {
            Cell cCell = this.getCELL();
            Font font = wb.createFont();
            font.setBold(bold);
            font.setItalic(italic);
            if (underline) {
                font.setUnderline(Font.U_SINGLE);
            }
            CellUtil.setFont(cCell, font);
            return true;
        } catch (Exception e) {
            log.error("Error in setting font style: " + e.getMessage());
            return false;
        }
    }
    /**
     * Sets a full border around the cell with the specified color.
     * 
     * @param borderColor The color name to set for the border
     */
    public void setFullBorder(String borderColor) {
        Cell cCell = null;
        try {
            cCell = this.getCELL();

            Map<String, Object> properties = new HashMap<>();
            properties.put(CellUtil.BORDER_LEFT, BorderStyle.MEDIUM);
            properties.put(CellUtil.BORDER_RIGHT, BorderStyle.MEDIUM);
            properties.put(CellUtil.BORDER_TOP, BorderStyle.MEDIUM);
            properties.put(CellUtil.BORDER_BOTTOM, BorderStyle.MEDIUM);
            properties.put(CellUtil.BOTTOM_BORDER_COLOR, Common.getColorCode(borderColor));
            properties.put(CellUtil.LEFT_BORDER_COLOR, Common.getColorCode(borderColor));
            properties.put(CellUtil.TOP_BORDER_COLOR, Common.getColorCode(borderColor));
            properties.put(CellUtil.RIGHT_BORDER_COLOR, Common.getColorCode(borderColor));

            CellUtil.setCellStyleProperties(cCell, properties);
        } catch (Exception e) {
            log.error("Error in setting border: " + e.getMessage());
        }
    }


    /**
     * Returns the cell style.
     * 
     * @param cStyle The cell style to return
     * @return The provided cell style
     */
    private CellStyle getCellStyle(CellStyle cStyle) {
        return cStyle;
    }

    /**
     * Gets the value of the cell as a string.
     * 
     * @return The cell value as a string
     */
    public String getValue() {
        return getTextValue();
    }

    /**
     * Gets the text value of the cell by reading from the workbook file.
     * 
     * @return The text value of the cell, or an empty string if not found or on error
     */
    public String getTextValue() {
        String val = null;
        try {
            inputStream = new FileInputStream(this.sCompleteFileName);
            this.wb = WorkbookFactory.create(inputStream);
            String strSheetName = sheet.getSheetName();
            this.sheet = this.wb.getSheet(strSheetName);

            try {
                val = this.sheet.getRow(this.row - 1).getCell(this.col - 1).getStringCellValue();
            } catch (Exception e) {
                val = "";
            }

            inputStream.close();
            this.wb.close();
        } catch (Exception e) {
            log.error("Error in getting text value: " + e.getMessage());
        }

        return val;
    }

    /**
     * Private constructor for internal use.
     */
    private ExcelCell() {
        // Default constructor
    }

    /**
     * Creates a new ExcelCell with the specified parameters.
     * This constructor initializes the cell and creates it if it doesn't exist.
     * 
     * @param sheet The worksheet containing the cell
     * @param wb The workbook containing the sheet
     * @param row The row number (1-based)
     * @param col The column number (1-based)
     * @param inputStream The input stream for the workbook file
     * @param outputStream The output stream for the workbook file
     * @param sCompleteFileName The complete file path and name of the workbook
     */
    protected ExcelCell(Sheet sheet, Workbook wb, int row, int col, FileInputStream inputStream,
                        FileOutputStream outputStream, String sCompleteFileName) {
        this.sheet = sheet;
        this.wb = wb;
        this.row = row;
        this.col = col;
        this.inputStream = inputStream;
        this.outputStream = outputStream;
        this.sCompleteFileName = sCompleteFileName;
        this.xlFileExtension = getFileExtension(sCompleteFileName);

        try {
            Row cRow = null;
            try {
                cRow = sheet.getRow(this.row - 1);
            } catch (Exception e) {
                // Row doesn't exist
            }

            Row row1;
            if (checkIfRowIsEmpty(cRow)) {
                row1 = sheet.createRow(this.row - 1);
            } else {
                row1 = cRow;
            }

            Cell cell1 = null;
            try {
                cell1 = row1.getCell(this.col - 1);
            } catch (Exception e) {
                // Cell doesn't exist
            }

            if (checkIfCellIsEmpty(cell1)) {
                cell1 = row1.createCell(this.col - 1);
            }

            this.setCELL(cell1);
        } catch (Exception e) {
            log.error("Error in cell constructor creation: " + e.getMessage());
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

    /**
     * Gets the file extension from a complete file path.
     * 
     * @param sCompleteFilePath The complete file path
     * @return The file extension including the dot, or an empty string on error
     */
    private String getFileExtension(String sCompleteFilePath) {
        try {
            return sCompleteFilePath.substring(sCompleteFilePath.lastIndexOf("."));
        } catch (Exception e) {
            log.error("Error in getting file extension: " + e.getMessage());
        }

        return ""; // Return empty string on error
    }

}
