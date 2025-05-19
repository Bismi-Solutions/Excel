package solutions.bismi.excel;

import lombok.AccessLevel;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;
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
 * Represents a row in an Excel worksheet.
 * Provides methods for manipulating row content, formatting, and styling.
 */
@NoArgsConstructor(access = AccessLevel.PRIVATE)
public class ExcelRow {
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
    private final Logger log = LogManager.getLogger(ExcelRow.class);
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
        try {
            // Initialize row if needed
        } catch (Exception e) {
            log.error("Error in row constructor creation: {}", e.getMessage());
        }
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
     * @param row The row containing the cell
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
     * @param endColNumber The ending column number (exclusive)
     * @param cellOperation The operation to apply to each cell
     * @param errorMessage Error message to log if operation fails
     */
    private void applyCellOperation(int startColNumber, int endColNumber, 
                                   CellOperation cellOperation, String errorMessage) {
        try {
            Row cRow = getOrCreateRow();

            for (int iCol = startColNumber; iCol < endColNumber; iCol++) {
                Cell cell = getOrCreateCell(cRow, iCol - 1);
                cellOperation.apply(cell);
            }
        } catch (Exception e) {
            log.error(errorMessage + ": " + e.getMessage());
        }
    }

    /**
     * Functional interface for cell operations.
     */
    @FunctionalInterface
    private interface CellOperation {
        void apply(Cell cell);
    }

    /**
     * Sets the font color for a range of cells in the row.
     *
     * @param fontColor      The color name to set for the font
     * @param startColNumber The starting column number (1-based)
     * @param endColNumber   The ending column number (exclusive)
     */
    public void setFontColor(String fontColor, int startColNumber, int endColNumber) {
        applyCellOperation(startColNumber, endColNumber, cell -> {
            Font font = cell.getSheet().getWorkbook().createFont();
            font.setColor(Common.getColorCode(fontColor));
            CellUtil.setFont(cell, font);
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
        setFontColor(fontColor, 1, endColNumber + 1);
    }

    /**
     * Sets the background fill color for a range of cells in the row.
     *
     * @param fillColor      The color name to set for the cell background
     * @param startColNumber The starting column number (1-based)
     * @param endColNumber   The ending column number (exclusive)
     */
    public void setFillColor(String fillColor, int startColNumber, int endColNumber) {
        applyCellOperation(startColNumber, endColNumber, cell -> {
            Map<String, Object> properties = new HashMap<>();
            properties.put(CellUtil.FILL_PATTERN, FillPatternType.SOLID_FOREGROUND);
            properties.put(CellUtil.FILL_FOREGROUND_COLOR, Common.getColorCode(fillColor));
            CellUtil.setCellStyleProperties(cell, properties);
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
     *
     * @param borderColor    The color name to set for the border
     * @param startColNumber The starting column number (1-based)
     * @param endColNumber   The ending column number (exclusive)
     */
    public void setFullBorder(String borderColor, int startColNumber, int endColNumber) {
        applyCellOperation(startColNumber, endColNumber, cell -> {
            Map<String, Object> properties = new HashMap<>();
            properties.put(CellUtil.BORDER_LEFT, BorderStyle.MEDIUM);
            properties.put(CellUtil.BORDER_RIGHT, BorderStyle.MEDIUM);
            properties.put(CellUtil.BORDER_TOP, BorderStyle.MEDIUM);
            properties.put(CellUtil.BORDER_BOTTOM, BorderStyle.MEDIUM);
            properties.put(CellUtil.BOTTOM_BORDER_COLOR, Common.getColorCode(borderColor));
            properties.put(CellUtil.LEFT_BORDER_COLOR, Common.getColorCode(borderColor));
            properties.put(CellUtil.TOP_BORDER_COLOR, Common.getColorCode(borderColor));
            properties.put(CellUtil.RIGHT_BORDER_COLOR, Common.getColorCode(borderColor));

            CellUtil.setCellStyleProperties(cell, properties);
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
     *
     * @param filePath The complete file path and name to save to
     */
    private void saveWorkBook(String filePath) {
        try (OutputStream fileOut = new FileOutputStream(filePath)) {
            wb.write(fileOut);
            log.debug("Workbook saved successfully to: {}", filePath);
        } catch (Exception e) {
            log.error("Error in saving workbook: {}", e.getMessage());
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
        return cell == null || cell.getCellType() == CellType.BLANK || !StringUtils.isNotBlank(cell.toString());
    }


}
