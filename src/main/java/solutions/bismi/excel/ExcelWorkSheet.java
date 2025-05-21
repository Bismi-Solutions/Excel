package solutions.bismi.excel;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.OutputStream;

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

        this.sheet = sheet;

        this.sheetName = sheetName;
        this.wb = wb;
        this.inputStream = inputStream;
        this.outputStream = outputStream;
        this.sCompleteFileName = sCompleteFileName;

    }

    /**
     * Function to activate excel sheet
     */
    public void activate() {
        this.wb.setActiveSheet(this.wb.getSheetIndex(this.sheet.getSheetName()));
        this.wb.setSelectedTab(this.wb.getSheetIndex(this.sheet.getSheetName()));
        saveWorkBook();
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
            shExcel = new ExcelWorkSheet(this.sheet, this.log, this.sheetName, this.wb);

        } catch (Exception e) {
            log.debug("Error in creating sheet {}: {}", sSheetName, e.toString());

        }

        return shExcel;
    }

    /**
     * @param inputStream
     * @param outputStream
     * @param sCompleteFileName
     */
    private void saveWorkBook(FileInputStream inputStream, FileOutputStream outputStream, String sCompleteFileName) {
        try {
            if (inputStream != null) inputStream.close();

            // Use the provided outputStream if it's not null, otherwise create a new one
            OutputStream fileOut;
            if (outputStream != null) {
                fileOut = outputStream;
            } else {
                fileOut = new FileOutputStream(sCompleteFileName);
            }

            wb.write(fileOut);

            // Only close the stream if we created it
            if (outputStream == null) {
                fileOut.close();
            }

        } catch (Exception e) {
            log.debug("Error in saving record: {}", e.toString());
        }
    }

    public void saveWorkBook() {
        try (FileOutputStream fileOut = new FileOutputStream(this.sCompleteFileName)) {
            wb.write(fileOut);
            fileOut.flush();
        } catch (Exception e) {
            log.debug("Error in saving record: {}", e.toString());
        }
    }


    /**
     * @param strArrSheets
     * @return
     */
    protected void addSheets(String[] strArrSheets) {
        try {
            for (String sSheetName : strArrSheets) {
                if (sSheetName != null && !sSheetName.trim().isEmpty()) {
                    sheet = this.wb.createSheet(sSheetName);
                    this.sheetName = sSheetName;
                    log.debug("Created sheet {}", sSheetName);
                }
            }
        } catch (Exception e) {
            log.debug("Error in creating sheet: {}", e.toString());
        }
    }

    /**
     * Removes a row at the specified index, handling any exceptions internally
     * 
     * @param rowIndex The index of the row to remove
     */
    private void removeRow(int rowIndex) {
        try {
            sheet.removeRow(sheet.getRow(rowIndex));
        } catch (Exception e) {
            log.debug("Empty row removal error {}", rowIndex);
        }
    }

    /**
     * @return true if contents were cleared successfully, false otherwise
     */
    public boolean clearContents() {
        try {
            for (int i = this.sheet.getLastRowNum(); i >= 0; i--) {
                removeRow(i);
            }
            saveWorkBook();
            log.debug("Cleared sheet contents successfully");
            return true;

        } catch (Exception e) {
            log.debug("Error in clearing sheet contents: {}", e.toString());
            return false;
        }
    }

    public int getColNumber() {
        int rn = this.sheet.getLastRowNum();
        int maxcol = 0;
        for (int i = 0; i <= rn; i++) {
            try {
                if (this.sheet.getRow(i).getLastCellNum() > maxcol) {
                    maxcol = this.sheet.getRow(i).getLastCellNum();
                }
            } catch (Exception e) {
                log.debug("Error getting column count for row {}: {}", i, e.getMessage());
            }

        }

        return maxcol;
    }


    public int getRowNumber() {
        return this.sheet.getLastRowNum() + 1;
    }

    public String getSheetName() {
        return sheetName;
    }


    /**
     * @param row
     * @param col
     * @return
     */
    public ExcelCell cell(int row, int col) {
        ExcelCell cells = null;
        try {
            cells = new ExcelCell(this.sheet, this.wb, row, col, this.inputStream, this.outputStream, this.sCompleteFileName);

        } catch (Exception e) {
            log.debug("Error Excel cells operation: {}", e.toString());
        }

        return cells;

    }

    /**
     * @param row
     * @return
     */
    public ExcelRow row(int row) {
        ExcelRow rows = null;
        try {
            rows = new ExcelRow(this.sheet, this.wb, row, this.inputStream, this.outputStream, this.sCompleteFileName);
        } catch (Exception e) {
            log.error("Error in Excel rows operation: {}", e.getMessage());
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
            for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
                if (range.intersects(sheet.getMergedRegion(i))) {
                    log.warn("Cannot merge cells - overlaps with existing merged region");
                    return false;
                }
            }

            this.sheet.addMergedRegion(range);
            saveWorkBook();
            return true;
        } catch (Exception e) {
            log.error("Error merging cells: {}", e.getMessage(), e);
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
            for (int i = sheet.getNumMergedRegions() - 1; i >= 0; i--) {
                org.apache.poi.ss.util.CellRangeAddress mergedRegion = sheet.getMergedRegion(i);

                // Check if the regions are equal or if they overlap
                if (mergedRegion.equals(rangeToRemove) || (mergedRegion.getFirstRow() <= rangeToRemove.getLastRow() && mergedRegion.getLastRow() >= rangeToRemove.getFirstRow() && mergedRegion.getFirstColumn() <= rangeToRemove.getLastColumn() && mergedRegion.getLastColumn() >= rangeToRemove.getFirstColumn())) {

                    sheet.removeMergedRegion(i);
                    removed = true;
                }
            }

            if (removed) {
                saveWorkBook();
                return true;
            }

            log.warn("No matching or overlapping merged region found to unmerge");
            return false;
        } catch (Exception e) {
            log.error("Error unmerging cells: {}", e.getMessage());
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

            for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
                org.apache.poi.ss.util.CellRangeAddress mergedRegion = sheet.getMergedRegion(i);
                if (mergedRegion.isInRange(rowIndex, colIndex)) {
                    return true;
                }
            }
            return false;
        } catch (Exception e) {
            log.error("Error checking if cell is merged: {}", e.getMessage());
            return false;
        }
    }


}
