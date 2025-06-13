package solutions.bismi.excel;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;

/**
 */
public class ExcelWorkSheet {
    int colNumber;
    String xlFileExtension = null;

    FileInputStream inputStream = null;
    FileOutputStream outputStream = null;
    String sCompleteFileName = null;
    Sheet sheet;
    String sheetName = null;
    Workbook wb;
    private Logger log = LogManager.getLogger(ExcelWorkSheet.class);


    /**
     *
     */
    ExcelWorkSheet() {

    }

    /**
     * @param sheet
     * @param log
     * @param sheetName
     * @param wb
     */
    ExcelWorkSheet(Sheet sheet, Logger log, String sheetName, Workbook wb) {

        this.sheet = sheet;
        this.log = log;
        this.sheetName = sheetName;
        this.wb = wb;

    }

    /**
     * @param sheet
     * @param sheetName
     * @param wb
     * @param inputStream
     * @param outputStream
     * @param sCompleteFileName
     */
    protected ExcelWorkSheet(Sheet sheet, String sheetName, Workbook wb, FileInputStream inputStream, FileOutputStream outputStream, String sCompleteFileName) {
        if (sheet == null) {
            log.error("Sheet cannot be null in ExcelWorkSheet constructor");
            throw new IllegalArgumentException("Sheet cannot be null");
        }
        if (wb == null) {
            log.error("Workbook cannot be null in ExcelWorkSheet constructor");
            throw new IllegalArgumentException("Workbook cannot be null");
        }
        if (sCompleteFileName == null || sCompleteFileName.trim().isEmpty()) {
            log.error("sCompleteFileName cannot be null or empty in ExcelWorkSheet constructor");
            throw new IllegalArgumentException("sCompleteFileName cannot be null or empty");
        }

        this.sheet = sheet;
        this.sheetName = sheetName; // sheetName can be null if wb.getSheetName(sheetIndex) returns null
        this.wb = wb;
        this.inputStream = inputStream;
        this.outputStream = outputStream;
        this.sCompleteFileName = sCompleteFileName;

    }

    /**
     * Function to activate excel sheet
     */
    public void activate() {
        if (this.wb == null || this.sheet == null) {
            log.error("Workbook or Sheet is null, cannot activate sheet: {}", this.sheetName);
            return;
        }
        try {
            int sheetIndex = this.wb.getSheetIndex(this.sheet.getSheetName());
            if (sheetIndex == -1) { // POI returns -1 if sheet name not found
                log.error("Sheet named '{}' not found in workbook. Cannot activate.", this.sheet.getSheetName());
                return;
            }
            this.wb.setActiveSheet(sheetIndex);
            this.wb.setSelectedTab(sheetIndex); // For older Excel versions compatibility
            if (!saveWorkBook()) {
                log.error("Failed to save workbook after activating sheet: {}", this.sheet.getSheetName());
                // Depending on desired behavior, could throw an exception or set an internal error state.
            }
        } catch (Exception e) {
            log.error("Error activating sheet '{}': " + e.getMessage(), this.sheet.getSheetName(), e);
        }
    }

    /**
     * @param sSheetName
     * @return
     */
    protected ExcelWorkSheet addSheet(String sSheetName) {
        ExcelWorkSheet shExcel = null;
        try {
            sheet = this.wb.createSheet(sSheetName);

            this.sheetName = sSheetName;
            log.debug("Created sheet {}", sSheetName);
            // Ensure sCompleteFileName is passed to retain context, though this constructor doesn't take it.
            // This implies shExcel created here might not be fully functional for saving or creating cells if it relies on sCompleteFileName.
            // For now, sticking to existing constructor signature.
            shExcel = new ExcelWorkSheet(this.sheet, this.log, this.sheetName, this.wb);
            // To make it fully functional, it would need:
            // shExcel = new ExcelWorkSheet(this.sheet, this.sheetName, this.wb, this.inputStream, this.outputStream, this.sCompleteFileName);


        } catch (Exception e) {
            log.error("Error in creating sheet {}: " + e.getMessage(), e);

        }

        return shExcel;
    }

    public boolean saveWorkBook() {
        if (this.sCompleteFileName == null || this.sCompleteFileName.trim().isEmpty()) {
            log.error("Cannot save workbook for sheet '{}', sCompleteFileName is null or empty.", this.sheetName);
            return false;
        }
        if (this.wb == null) {
            log.error("Cannot save workbook for sheet '{}', workbook (wb) is null.", this.sheetName);
            return false;
        }
        try (FileOutputStream fileOut = new FileOutputStream(this.sCompleteFileName)) {
            wb.write(fileOut);
            // fileOut.flush(); // flush() is called by close() in try-with-resources
            log.debug("Workbook saved successfully for sheet: {}", this.sheetName);
            return true;
        } catch (java.io.FileNotFoundException e) {
            log.error("Error saving workbook for sheet '{}' to '{}': File not found or path is a directory. " + e.getMessage(), this.sheetName, this.sCompleteFileName, e);
        } catch (java.io.IOException e) {
            log.error("IO error saving workbook for sheet '{}' to '{}': " + e.getMessage(), this.sheetName, this.sCompleteFileName, e);
        } catch (Exception e) { // Catch any other unexpected exceptions
            log.error("Unexpected error saving workbook for sheet '{}' to '{}': " + e.getMessage(), this.sheetName, this.sCompleteFileName, e);
        }
        return false;
    }


    /**
     * @param strArrSheets
     * @return
     */
    protected void addSheets(String[] strArrSheets) {
        try {
            for (String sSheetName : strArrSheets) {
                if (sSheetName != null && !sSheetName.trim().isEmpty()) {
                    // Ensure wb is not null before creating a sheet
                    if (this.wb == null) {
                        log.error("Workbook (wb) is null, cannot create sheet: {}", sSheetName);
                        continue; // or throw exception
                    }
                    sheet = this.wb.createSheet(sSheetName);
                    // this.sheetName = sSheetName; // This would change the current worksheet's name, which is likely not intended here.
                                                // The 'sheet' variable is overwritten in the loop, but this.sheetName should ideally refer to the primary sheet of this ExcelWorkSheet instance.
                    log.debug("Created sheet {}", sSheetName);
                }
            }
        } catch (Exception e) {
            log.error("Error in creating sheets: " + e.getMessage(), e);
        }
    }

    /**
     * Removes a row at the specified index, handling any exceptions internally
     * 
     * @param rowIndex The index of the row to remove
     */
    private void removeRow(int rowIndex) {
        try {
            if (this.sheet == null) {
                log.warn("Sheet is null, cannot remove row at index {}", rowIndex);
                return;
            }
            org.apache.poi.ss.usermodel.Row row = sheet.getRow(rowIndex);
            if (row != null) { // Only remove if row exists
                sheet.removeRow(row);
            }
        } catch (Exception e) {
            log.debug("Error removing row at index {}: " + e.getMessage(), rowIndex, e);
        }
    }

    /**
     * @return true if contents were cleared successfully, false otherwise
     */
    public boolean clearContents() {
        if (this.sheet == null) {
            log.error("Sheet is null. Cannot clear contents for sheet named: {}", this.sheetName);
            return false;
        }
        try {
            for (int i = this.sheet.getLastRowNum(); i >= 0; i--) {
                removeRow(i);
            }
            if (!saveWorkBook()) { // This might be slow if clearing many rows one by one.
                log.error("Failed to save workbook after clearing contents for sheet: {}", this.sheetName);
                return false; // Indicate that operation failed because save failed
            }
            log.debug("Cleared sheet contents successfully for sheet: {}", this.sheetName);
            return true;
        } catch (Exception e) {
            log.error("Error in clearing sheet contents for sheet '{}': " + e.getMessage(), this.sheetName, e);
            return false;
        }
    }

    public int getColNumber() {
        if (this.sheet == null) {
            log.warn("Sheet is null, cannot get column number. Returning 0.");
            return 0;
        }
        int rn = this.sheet.getLastRowNum();
        int maxcol = 0;
        for (int i = 0; i <= rn; i++) {
            try {
                org.apache.poi.ss.usermodel.Row row = this.sheet.getRow(i);
                if (row != null && row.getLastCellNum() > maxcol) {
                    maxcol = row.getLastCellNum();
                }
            } catch (Exception e) {
                log.debug("Error getting column count for row {}: " + e.getMessage(), i, e);
            }
        }
        return maxcol;
    }


    public int getRowNumber() {
        if (this.sheet == null) {
            log.warn("Sheet is null, cannot get row number. Returning 0.");
            return 0;
        }
        return this.sheet.getLastRowNum() + 1;
    }

    public String getSheetName() {
        if (this.sheet != null) {
            return this.sheet.getSheetName();
        } else if (this.sheetName != null) {
            return this.sheetName; // Fallback to stored name if sheet object is null
        }
        log.warn("Sheet object and sheetName field are both null. Cannot get sheet name.");
        return null;
    }


    /**
     * @param row
     * @param col
     * @return
     */
    public ExcelCell cell(int row, int col) {
        if (this.sCompleteFileName == null || this.sCompleteFileName.trim().isEmpty()) {
            log.error("sCompleteFileName is null or empty in ExcelWorkSheet. Cannot create ExcelCell.");
            // Or throw new IllegalStateException("sCompleteFileName is not initialized, cannot create cell.");
            return null;
        }
        if (this.sheet == null) {
            log.error("Sheet object is null in ExcelWorkSheet. Cannot create ExcelCell for row {}, col {}.", row, col);
            return null;
        }
        if (this.wb == null) {
            log.error("Workbook object (wb) is null in ExcelWorkSheet. Cannot create ExcelCell for row {}, col {}.", row, col);
            return null;
        }
        ExcelCell cells = null;
        try {
            cells = new ExcelCell(this.sheet, this.wb, row, col, this.inputStream, this.outputStream, this.sCompleteFileName);
        } catch (Exception e) {
            log.error("Error in ExcelCell creation for sheet '{}', row {}, col {}: " + e.getMessage(), this.getSheetName(), row, col, e);
        }
        return cells;
    }

    /**
     * @param row
     * @return
     */
    public ExcelRow row(int row) {
        if (this.sCompleteFileName == null || this.sCompleteFileName.trim().isEmpty()) {
            log.error("sCompleteFileName is null or empty in ExcelWorkSheet. Cannot create ExcelRow.");
            // Or throw new IllegalStateException("sCompleteFileName is not initialized, cannot create row.");
            return null;
        }
        if (this.sheet == null) {
            log.error("Sheet object is null in ExcelWorkSheet. Cannot create ExcelRow for row {}.", row);
            return null;
        }
        if (this.wb == null) {
            log.error("Workbook object (wb) is null in ExcelWorkSheet. Cannot create ExcelRow for row {}.", row);
            return null;
        }
        ExcelRow rows = null;
        try {
            rows = new ExcelRow(this.sheet, this.wb, row, this.inputStream, this.outputStream, this.sCompleteFileName);
        } catch (Exception e) {
            log.error("Error in ExcelRow creation for sheet '{}', row {}: " + e.getMessage(), this.getSheetName(), row, e);
        }
        return rows;
    }

    /**
     * Merges cells in the specified range
     *
     * @param row1 First row of the range (1-based)
     * @param col1 First column of the range (1-based)
     * @param row2 Last row of the range (1-based)
     * @param col2 Last column of the range (1-based)
     * @return true if successful, false otherwise
     */
    public boolean mergeCells(int row1, int col1, int row2, int col2) {
        try {
            // Input validation
            if (row1 <= 0 || col1 <= 0 || row2 <= 0 || col2 <= 0) {
                log.error("Invalid merge range: ({},{}) to ({},{})", row1, col1, row2, col2);
                return false;
            }

            // Determine actual first/last row/column
            int firstRow = Math.min(row1, row2);
            int lastRow = Math.max(row1, row2);
            int firstCol = Math.min(col1, col2);
            int lastCol = Math.max(col1, col2);

            // Convert to 0-based indices for POI
            org.apache.poi.ss.util.CellRangeAddress range = new org.apache.poi.ss.util.CellRangeAddress(firstRow - 1, lastRow - 1, firstCol - 1, lastCol - 1);

            // Check for overlapping regions
            if (this.sheet == null) {
                log.error("Sheet is null, cannot merge cells.");
                return false;
            }
            for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
                if (range.intersects(sheet.getMergedRegion(i))) {
                    log.warn("Cannot merge cells on sheet '{}' - overlaps with existing merged region at ({},{}) to ({},{})", this.getSheetName(), row1, col1, row2, col2);
                    return false;
                }
            }

            this.sheet.addMergedRegion(range);
            if (!saveWorkBook()) { // saveWorkBook already checks for sCompleteFileName and wb
                log.error("Failed to save workbook after merging cells on sheet '{}' for range ({},{}) to ({},{})", this.getSheetName(), row1, col1, row2, col2);
                return false; // Indicate failure if save fails
            }
            return true;
        } catch (Exception e) {
            log.error("Error merging cells on sheet '{}' for range ({},{}) to ({},{}): " + e.getMessage(), this.getSheetName(), row1, col1, row2, col2, e);
            return false;
        }
    }

    /**
     * Unmerges the cells in the specified range
     *
     * @param firstRow First row of the range (1-based)
     * @param lastRow  Last row of the range (1-based)
     * @param firstCol First column of the range (1-based)
     * @param lastCol  Last column of the range (1-based)
     * @return true if successful, false otherwise
     */
    public boolean unmergeCells(int firstRow, int lastRow, int firstCol, int lastCol) {
        try {
            // Input validation
            if (firstRow <= 0 || firstCol <= 0 || lastRow <= 0 || lastCol <= 0) {
                log.error("Invalid unmerge range: ({},{}) to ({},{})", firstRow, firstCol, lastRow, lastCol);
                return false;
            }

            // Determine actual first/last row/column
            int actualFirstRow = Math.min(firstRow, lastRow);
            int actualLastRow = Math.max(firstRow, lastRow);
            int actualFirstCol = Math.min(firstCol, lastCol);
            int actualLastCol = Math.max(firstCol, lastCol);

            // Convert to 0-based indices for POI
            org.apache.poi.ss.util.CellRangeAddress rangeToRemove = new org.apache.poi.ss.util.CellRangeAddress(actualFirstRow - 1, actualLastRow - 1, actualFirstCol - 1, actualLastCol - 1);

            // Find and remove any merged region that matches or overlaps with our range
            boolean removed = false;
            if (this.sheet == null) {
                log.error("Sheet is null, cannot unmerge cells.");
                return false;
            }
            for (int i = sheet.getNumMergedRegions() - 1; i >= 0; i--) {
                org.apache.poi.ss.util.CellRangeAddress mergedRegion = sheet.getMergedRegion(i);

                // Check if the regions are equal or if they overlap
                if (mergedRegion.equals(rangeToRemove) || (mergedRegion.getFirstRow() <= rangeToRemove.getLastRow() && mergedRegion.getLastRow() >= rangeToRemove.getFirstRow() && mergedRegion.getFirstColumn() <= rangeToRemove.getLastColumn() && mergedRegion.getLastColumn() >= rangeToRemove.getFirstColumn())) {
                    sheet.removeMergedRegion(i);
                    removed = true;
                }
            }

            if (removed) {
                if (!saveWorkBook()) { // saveWorkBook already checks for sCompleteFileName and wb
                    log.error("Failed to save workbook after unmerging cells on sheet '{}' for range ({},{}) to ({},{})", this.getSheetName(), firstRow, firstCol, lastRow, lastCol);
                    return false; // Indicate failure if save fails
                }
                return true;
            }

            log.warn("No matching or overlapping merged region found to unmerge on sheet '{}' for range ({},{}) to ({},{})", this.getSheetName(), firstRow, firstCol, lastRow, lastCol);
            return false; // No region found is not necessarily a save failure, but the operation didn't change anything to save.
        } catch (Exception e) {
            log.error("Error unmerging cells on sheet '{}': " + e.getMessage(), this.getSheetName(), e);
            return false;
        }
    }

    /**
     * Checks if a cell is part of a merged region
     *
     * @param row Row index (1-based)
     * @param col Column index (1-based)
     * @return true if the cell is part of a merged region, false otherwise
     */
    public boolean isCellMerged(int row, int col) {
        try {
            // Convert to 0-based indices for POI
            int rowIndex = row - 1;
            int colIndex = col - 1;

            if (this.sheet == null) {
                log.error("Sheet is null, cannot check if cell is merged.");
                return false;
            }
            for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
                org.apache.poi.ss.util.CellRangeAddress mergedRegion = sheet.getMergedRegion(i);
                if (mergedRegion.isInRange(rowIndex, colIndex)) {
                    return true;
                }
            }
            return false;
        } catch (Exception e) {
            log.error("Error checking if cell ({},{}) is merged on sheet '{}': " + e.getMessage(), row, col, this.getSheetName(), e);
            return false;
        }
    }


}
