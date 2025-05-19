package solutions.bismi.excel;

import lombok.AccessLevel;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;
import java.io.FileOutputStream;


/**
 * Represents a row in an Excel worksheet.
 * Provides methods for manipulating row content, formatting, and styling.
 */
@NoArgsConstructor(access = AccessLevel.PRIVATE)
public class ExcelRow {
    private final Logger log = LogManager.getLogger(ExcelRow.class);
    @Getter
    @Setter
    FileInputStream inputStream = null;
    @Getter
    @Setter
    FileOutputStream outputStream = null;
    @Getter
    @Setter
    String sCompleteFileName = null;
    @Getter
    @Setter
    Row ROW = null;
    @Getter
    @Setter
    String xlFileExtension = null;
    @Getter
    @Setter
    private Sheet sheet;
    @Getter
    @Setter
    private Workbook wb;
    @Getter
    @Setter
    private int rowNumber;

    /**
     * Creates a new ExcelRow with the specified parameters.
     *
     * @param sheet             The worksheet containing the row
     * @param wb                The workbook containing the sheet
     * @param rowNumber         The row number (1-based)
     * @param inputStream       The input stream for the workbook file
     * @param outputStream      The output stream for the workbook file
     * @param sCompleteFileName The complete file path and name of the workbook
     */
    protected ExcelRow(Sheet sheet, Workbook wb, int rowNumber, FileInputStream inputStream, FileOutputStream outputStream, String sCompleteFileName) {
        this.sheet = sheet;
        this.wb = wb;
        this.rowNumber = rowNumber;
        this.inputStream = inputStream;
        this.outputStream = outputStream;
        this.sCompleteFileName = sCompleteFileName;
    }

    /**
     * Sets values for multiple cells in the row.
     *
     * @param sValues         Array of string values to set in the row
     * @param startColNumber  The starting column number (0-based)
     * @param autoSizeColumns Whether to auto-size the columns
     */
    public void setRowValues(String[] sValues, int startColNumber, boolean autoSizeColumns) {
        try {
            if (sValues == null || sValues.length == 0) {
                return;
            }

            Row cRow = getOrCreateRow();

            int i = startColNumber;
            for (String value : sValues) {
                Cell cell = getOrCreateCell(cRow, i);
                cell.setCellValue(value);

                if (autoSizeColumns) {
                    sheet.autoSizeColumn(i);
                }
                ++i;
            }
        } catch (Exception e) {
            log.error("Error in writing values to row: {}", e.getMessage());
        }
    }

    public void setRowValues(String[] sValues, int startColNumber) {
        setRowValues(sValues, startColNumber, false);
    }

    public void setRowValues(String[] sValues, boolean autoSizeColoumns) {
        setRowValues(sValues, 0, autoSizeColoumns);
    }

    public void setRowValues(String[] sValues) {
        setRowValues(sValues, 0, false);
    }

    /**
     * Gets or creates a row in the sheet.
     *
     * @return The row object
     */
    private Row getOrCreateRow() {
        Row cRow = null;
        try {
            cRow = sheet.getRow(this.rowNumber - 1);
        } catch (Exception e1) {
            // Row doesn't exist
        }

        if (checkIfRowIsEmpty(cRow)) {
            cRow = sheet.createRow(this.rowNumber - 1);
        }

        return cRow;
    }

    /**
     * Gets or creates a cell in the given row.
     *
     * @param row      The row containing the cell
     * @param colIndex The column index (0-based)
     * @return The cell object
     */
    private Cell getOrCreateCell(Row row, int colIndex) {
        Cell cell = null;
        try {
            cell = row.getCell(colIndex);
        } catch (Exception e) {
            // Cell doesn't exist
        }

        if (checkIfCellIsEmpty(cell)) {
            cell = row.createCell(colIndex);
        }

        return cell;
    }

    /**
     * Applies a function to a range of cells in the row.
     *
     * @param startColNumber The starting column number (1-based)
     * @param endColNumber   The ending column number (exclusive)
     * @param cellOperation  The operation to apply to each cell
     * @param errorMessage   Error message to log if operation fails
     */
    private void applyCellOperation(int startColNumber, int endColNumber, CellOperation cellOperation, String errorMessage) {
        try {
            Row cRow = getOrCreateRow();

            for (int iCol = startColNumber; iCol < endColNumber; iCol++) {
                Cell cell = getOrCreateCell(cRow, iCol - 1);
                cellOperation.apply(cell);
            }
        } catch (Exception e) {
            log.error("{}: {}", errorMessage, e.getMessage());
        }
    }

    /**
     * Sets the font color for a range of cells in the row.
     * Preserves all existing cell style properties including borders and fill.
     *
     * @param fontColor      The color name to set for the font
     * @param startColNumber The starting column number (1-based)
     * @param endColNumber   The ending column number (exclusive)
     */
    public void setFontColor(String fontColor, int startColNumber, int endColNumber) {
        applyCellOperation(startColNumber, endColNumber, cell -> {
            Workbook workbook = cell.getSheet().getWorkbook();
            CellStyle originalStyle = cell.getCellStyle();
            CellStyle newStyle = workbook.createCellStyle();
            newStyle.cloneStyleFrom(originalStyle);
            Font oldFont = workbook.getFontAt(originalStyle.getFontIndex());
            Font newFont = workbook.createFont();
            Common.copyFontProperties(oldFont, newFont);
            newFont.setColor(Common.getColorCode(fontColor));
            newStyle.setFont(newFont);
            cell.setCellStyle(newStyle);
        }, "Error in setting font color");
    }

    /**
     * Sets the font color for all cells in the row.
     *
     * @param fontColor The color name to set for the font
     */
    public void setFontColor(String fontColor) {
        Row cRow = getOrCreateRow();
        int endColNumber = cRow.getLastCellNum();
        if (endColNumber <= 0) {
            endColNumber = 1; // If row is empty, assume at least one cell
        }
        setFontColor(fontColor, 1, endColNumber + 1);
    }

    /**
     * Sets the background fill color for a range of cells in the row.
     * Uses direct cell style manipulation instead of setCellStyleProperties.
     *
     * @param fillColor      The color name to set for the cell background
     * @param startColNumber The starting column number (1-based)
     * @param endColNumber   The ending column number (exclusive)
     */
    public void setFillColor(String fillColor, int startColNumber, int endColNumber) {
        applyCellOperation(startColNumber, endColNumber, cell -> {
            // Get the current cell style
            CellStyle currentStyle = cell.getCellStyle();

            // Create a new style with the same properties
            Workbook workbook = cell.getSheet().getWorkbook();
            CellStyle newStyle = workbook.createCellStyle();
            newStyle.cloneStyleFrom(currentStyle);

            // Set the fill pattern and color directly
            newStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            newStyle.setFillForegroundColor(Common.getColorCode(fillColor));

            // Apply the new style to the cell
            cell.setCellStyle(newStyle);
        }, "Error in setting fill color");
    }

    /**
     * Sets the background fill color for all cells in the row.
     *
     * @param fillColor The color name to set for the cell background
     */
    public void setFillColor(String fillColor) {
        Row cRow = getOrCreateRow();
        int endColNumber = cRow.getLastCellNum();
        setFillColor(fillColor, 1, endColNumber + 1);
    }

    /**
     * Sets a full border around a range of cells in the row.
     * Uses direct cell style manipulation instead of setCellStyleProperties.
     *
     * @param borderColor    The color name to set for the border
     * @param startColNumber The starting column number (1-based)
     * @param endColNumber   The ending column number (exclusive)
     */
    public void setFullBorder(String borderColor, int startColNumber, int endColNumber) {
        applyCellOperation(startColNumber, endColNumber, cell -> {
            // Get the current cell style
            CellStyle currentStyle = cell.getCellStyle();

            // Create a new style with the same properties
            Workbook workbook = cell.getSheet().getWorkbook();
            CellStyle newStyle = workbook.createCellStyle();
            newStyle.cloneStyleFrom(currentStyle);

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
            cell.setCellStyle(newStyle);
        }, "Error in setting border");
    }

    /**
     * Sets a full border around all cells in the row.
     *
     * @param borderColor The color name to set for the border
     */
    public void setFullBorder(String borderColor) {
        Row cRow = getOrCreateRow();
        int endColNumber = cRow.getLastCellNum();
        setFullBorder(borderColor, 1, endColNumber + 1);
    }

    /**
     * Saves the workbook to the specified file.
     * Delegates to Common utility method.
     *
     * @param filePath The complete file path and name to save to
     */
    private void saveWorkBook(String filePath) {
        Common.saveWorkBook(wb, filePath);
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
     * Functional interface for cell operations.
     */
    @FunctionalInterface
    private interface CellOperation {
        void apply(Cell cell);
    }


}
