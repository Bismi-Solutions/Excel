package solutions.bismi.excel;

import lombok.AccessLevel;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;
import lombok.AllArgsConstructor;
import lombok.Builder;
import org.apache.commons.lang3.StringUtils;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellUtil;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.HashMap;
import java.util.Map;

/**
 * Represents a cell in an Excel worksheet.
 * Provides methods for manipulating cell content, formatting, and styling.
 */
@NoArgsConstructor(access = AccessLevel.PRIVATE)
@AllArgsConstructor(access = AccessLevel.PRIVATE)
@Builder
public class ExcelCell {

    @Getter(AccessLevel.PRIVATE) @Setter(AccessLevel.PRIVATE)
    Sheet sheet;

    @Getter(AccessLevel.PRIVATE) @Setter(AccessLevel.PRIVATE)
    Workbook wb;

    @Getter(AccessLevel.PRIVATE) @Setter(AccessLevel.PRIVATE)
    int row;

    @Getter(AccessLevel.PRIVATE) @Setter(AccessLevel.PRIVATE)
    int col;

    @Getter(AccessLevel.PRIVATE) @Setter(AccessLevel.PRIVATE)
    FileInputStream inputStream = null;

    @Getter(AccessLevel.PRIVATE) @Setter(AccessLevel.PRIVATE)
    FileOutputStream outputStream = null;

    @Getter(AccessLevel.PRIVATE) @Setter(AccessLevel.PRIVATE)
    String sCompleteFileName = null;

    @Getter(AccessLevel.PRIVATE) @Setter(AccessLevel.PRIVATE)
    Cell CELL = null;

    @Getter(AccessLevel.PRIVATE) @Setter(AccessLevel.PRIVATE)
    String xlFileExtension = null;

    private final Logger log = LogManager.getLogger(ExcelCell.class);

    /**
     * Creates a new ExcelCell with the specified parameters.
     * This constructor initializes the cell and creates it if it doesn't exist.
     *
     * @param sheet             The worksheet containing the cell
     * @param wb                The workbook containing the sheet
     * @param row               The row number (1-based)
     * @param col               The column number (1-based)
     * @param inputStream       The input stream for the workbook file
     * @param outputStream      The output stream for the workbook file
     * @param sCompleteFileName The complete file path and name of the workbook
     */
    protected ExcelCell(Sheet sheet, Workbook wb, int row, int col, FileInputStream inputStream, FileOutputStream outputStream, String sCompleteFileName) {
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
            log.error("Error in cell constructor creation: {}", e.getMessage());
        }
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
     * @param sText           The text value to set
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
            log.error("Error in setting cell value: {}", e.getMessage());
            return false;
        }
    }

    /**
     * Sets a numeric value in the cell
     *
     * @param value          The numeric value to set
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
            log.error("Error in setting numeric value: {}", e.getMessage());
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
     * @param date           The date value to set
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
            log.error("Error in setting date value: {}", e.getMessage());
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
     * @param formula        The formula to set (without the leading '=')
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
            log.error("Error in setting formula: {}", e.getMessage());
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
     * @param sText           The text value to set
     * @param autoSizeColoumn Whether to auto-size the column
     * @return true if successful, false otherwise
     */
    public boolean setText(String sText, boolean autoSizeColoumn) {
        Cell cCell = null;

        try {
            cCell = this.getCELL();

            // Get the current cell style
            Workbook workbook = cCell.getSheet().getWorkbook();
            CellStyle originalStyle = cCell.getCellStyle();
            CellStyle newStyle = workbook.createCellStyle();
            newStyle.cloneStyleFrom(originalStyle);

            // Set the data format to Text
            DataFormat format = workbook.createDataFormat();
            newStyle.setDataFormat(format.getFormat("@"));  // "@" is the format code for Text

            cCell.setCellValue(sText);
            cCell.setCellStyle(newStyle);

            if (autoSizeColoumn) {
                this.sheet.autoSizeColumn(col - 1);
            }

            return true;
        } catch (Exception e) {
            log.error("Error in setting text value: {}", e.getMessage());
            return false;
        }
    }

    /**
     * Sets the font color for the cell while preserving existing formatting.
     *
     * @param fontColor The color name to set for the font
     * @return this ExcelCell instance for method chaining
     */
    public ExcelCell setFontColor(String fontColor) {
        try {
            Cell cCell = this.getCELL();
            Workbook workbook = cCell.getSheet().getWorkbook();
            CellStyle originalStyle = cCell.getCellStyle();
            CellStyle newStyle = workbook.createCellStyle();
            newStyle.cloneStyleFrom(originalStyle);
            Font existingFont = workbook.getFontAt(originalStyle.getFontIndex());
            Font newFont = workbook.createFont();
            Common.copyFontProperties(existingFont, newFont);
            newFont.setColor(Common.getColorCode(fontColor));
            newStyle.setFont(newFont);
            cCell.setCellStyle(newStyle);

            return this;
        } catch (Exception e) {
            log.error("Error in setting font color: {}", e.getMessage());
            return this;
        }
    }

    /**
     * Sets the background fill color for the cell.
     *
     * @param fillColor The color name to set for the cell background
     */
    public void setFillColor(String fillColor) {
        try {
            Cell cCell = this.getCELL();

            // Get the current cell style
            Workbook workbook = cCell.getSheet().getWorkbook();
            CellStyle originalStyle = cCell.getCellStyle();
            CellStyle newStyle = workbook.createCellStyle();
            newStyle.cloneStyleFrom(originalStyle);

            // Set the fill pattern and color
            newStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            newStyle.setFillForegroundColor(Common.getColorCode(fillColor));

            // Apply the new style to the cell
            cCell.setCellStyle(newStyle);
        } catch (Exception e) {
            log.error("Error in setting fill color: {}", e.getMessage());
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

            // Get the current cell style
            Workbook workbook = cCell.getSheet().getWorkbook();
            CellStyle originalStyle = cCell.getCellStyle();
            CellStyle newStyle = workbook.createCellStyle();
            newStyle.cloneStyleFrom(originalStyle);

            HorizontalAlignment hAlign;
            switch (alignment.toUpperCase()) {
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

            // Set the horizontal alignment
            newStyle.setAlignment(hAlign);

            // Apply the new style to the cell
            cCell.setCellStyle(newStyle);
            return true;
        } catch (Exception e) {
            log.error("Error in setting horizontal alignment: {}", e.getMessage());
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

            // Get the current cell style
            Workbook workbook = cCell.getSheet().getWorkbook();
            CellStyle originalStyle = cCell.getCellStyle();
            CellStyle newStyle = workbook.createCellStyle();
            newStyle.cloneStyleFrom(originalStyle);

            VerticalAlignment vAlign;
            switch (alignment.toUpperCase()) {
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

            // Set the vertical alignment
            newStyle.setVerticalAlignment(vAlign);

            // Apply the new style to the cell
            cCell.setCellStyle(newStyle);
            return true;
        } catch (Exception e) {
            log.error("Error in setting vertical alignment: {}", e.getMessage());
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

            // Get the current cell style
            Workbook workbook = cCell.getSheet().getWorkbook();
            CellStyle originalStyle = cCell.getCellStyle();
            CellStyle newStyle = workbook.createCellStyle();
            newStyle.cloneStyleFrom(originalStyle);

            // Set the number format
            DataFormat format = workbook.createDataFormat();
            newStyle.setDataFormat(format.getFormat(formatPattern));

            // Apply the new style to the cell
            cCell.setCellStyle(newStyle);
            return true;
        } catch (Exception e) {
            log.error("Error in setting number format: {}", e.getMessage());
            return false;
        }
    }

    /**
     * Sets the font style (bold, italic, etc.)
     *
     * @param bold      true to make the font bold
     * @param italic    true to make the font italic
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
                font.setUnderline(org.apache.poi.ss.usermodel.FontUnderline.SINGLE.getByteValue());
            }
            CellUtil.setFont(cCell, font);
            return true;
        } catch (Exception e) {
            log.error("Error in setting font style: {}", e.getMessage());
            return false;
        }
    }

    /**
     * Sets a full border around the cell with the specified color.
     *
     * @param borderColor The color name to set for the border
     */
    public void setFullBorder(String borderColor) {
        try {
            Cell cCell = this.getCELL();

            // Get the current cell style
            Workbook workbook = cCell.getSheet().getWorkbook();
            CellStyle originalStyle = cCell.getCellStyle();
            CellStyle newStyle = workbook.createCellStyle();
            newStyle.cloneStyleFrom(originalStyle);

            // Set border styles directly
            short colorCode = Common.getColorCode(borderColor);

            newStyle.setBorderLeft(BorderStyle.MEDIUM);
            newStyle.setBorderRight(BorderStyle.MEDIUM);
            newStyle.setBorderTop(BorderStyle.MEDIUM);
            newStyle.setBorderBottom(BorderStyle.MEDIUM);

            newStyle.setLeftBorderColor(colorCode);
            newStyle.setRightBorderColor(colorCode);
            newStyle.setTopBorderColor(colorCode);
            newStyle.setBottomBorderColor(colorCode);

            // Apply the new style to the cell
            cCell.setCellStyle(newStyle);
        } catch (Exception e) {
            log.error("Error in setting border: {}", e.getMessage());
        }
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
        String val = "";
        try (FileInputStream fis = new FileInputStream(this.sCompleteFileName);
             Workbook workbook = WorkbookFactory.create(fis)) {

            String strSheetName = sheet.getSheetName();
            Sheet currentSheet = workbook.getSheet(strSheetName);

            try {
                val = currentSheet.getRow(this.row - 1).getCell(this.col - 1).getStringCellValue();
            } catch (Exception e) {
                // If cell value can't be retrieved, return empty string
            }
        } catch (Exception e) {
            log.error("Error in getting text value: {}", e.getMessage());
        }

        return val;
    }


    /**
     * Checks if a row is empty or contains only blank cells.
     * Delegates to Common utility method.
     *
     * @param row The row to check
     * @return true if the row is empty or null, false otherwise
     */
    private boolean checkIfRowIsEmpty(Row row) {
        return Common.checkIfRowIsEmpty(row);
    }

    /**
     * Checks if a cell is empty or blank.
     * Delegates to Common utility method.
     *
     * @param cell The cell to check
     * @return true if the cell is empty, null, or blank, false otherwise
     */
    private boolean checkIfCellIsEmpty(Cell cell) {
        return Common.checkIfCellIsEmpty(cell);
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
            log.error("Error in getting file extension: {}", e.getMessage());
        }

        return ""; // Return empty string on error
    }

}
