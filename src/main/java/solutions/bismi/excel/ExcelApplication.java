package solutions.bismi.excel;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import java.util.ArrayList;
import java.util.List;

/**
 * Excel Application base class to hold all application level operations. This is the root class.
 */
public class ExcelApplication {
    private final Logger log = LogManager.getLogger(ExcelApplication.class);
    private final List<ExcelWorkBook> exlWorkBooks = new ArrayList<>();

    /**
     * Default constructor for ExcelApplication.
     */
    public ExcelApplication() {
        // Initialize with empty workbook list
    }

    /**
     * Gets all workbooks currently open in the application.
     *
     * @return List of ExcelWorkBook objects
     */
    public List<ExcelWorkBook> getWorkbooks() {
        return exlWorkBooks;
    }

    /**
     * Gets the count of workbooks currently open in the application.
     *
     * @return The number of open workbooks
     */
    public int getOpenWorkbookCount() {
        return exlWorkBooks.size();
    }

    /**
     * Creates a new Excel workbook with the specified file name.
     *
     * @param sCompleteFileName The complete file path and name for the new workbook
     * @return The newly created ExcelWorkBook object
     */
    public ExcelWorkBook createWorkBook(String sCompleteFileName) {
        log.info("Creating excel workbook: {}", sCompleteFileName);
        ExcelWorkBook exlWorkBook = new ExcelWorkBook();
        exlWorkBook.createWorkBook(sCompleteFileName);
        exlWorkBooks.add(exlWorkBook); // adding excel book to main application list
        return exlWorkBook;
    }

    /**
     * Opens an existing Excel workbook with the specified file name.
     *
     * @param sCompleteFileName The complete file path and name of the workbook to open
     * @return The opened ExcelWorkBook object
     */
    public ExcelWorkBook openWorkbook(String sCompleteFileName) {
        log.debug("Opening excel workbook: {}", sCompleteFileName);
        ExcelWorkBook exlWorkBook = new ExcelWorkBook();
        exlWorkBook.openWorkbook(sCompleteFileName);
        exlWorkBooks.add(exlWorkBook);
        return exlWorkBook;
    }

    /**
     * Gets a workbook by its index in the list of open workbooks.
     *
     * @param iIndex The index of the workbook to retrieve
     * @return The ExcelWorkBook at the specified index, or null if not found
     */
    public ExcelWorkBook getWorkbook(int iIndex) {
        ExcelWorkBook tempBook = null;

        try {
            tempBook = this.exlWorkBooks.get(iIndex);
            log.debug("Retrieved workbook at index: {}", iIndex);
        } catch (IndexOutOfBoundsException e) {
            log.error("Error during workbook retrieval - index out of bounds: {} - {}", iIndex, e.getMessage());
        } catch (Exception e) {
            log.error("Error during workbook retrieval: {}", e.getMessage());
        }

        return tempBook;
    }

    /**
     * Gets a workbook by its name from the list of open workbooks.
     *
     * @param sWorkBookName The name of the workbook to retrieve
     * @return The ExcelWorkBook with the specified name, or null if not found
     */
    public ExcelWorkBook getWorkbook(String sWorkBookName) {
        ExcelWorkBook tempBook = null;
        boolean fFound = false;

        if (sWorkBookName == null || sWorkBookName.trim().isEmpty()) {
            log.error("Workbook name cannot be null or empty");
            return null;
        }

        for (ExcelWorkBook excelWorkBook : this.exlWorkBooks) {
            try {
                if (excelWorkBook.getExcelBookName().equalsIgnoreCase(sWorkBookName)) {
                    fFound = true;
                    tempBook = excelWorkBook;
                    log.debug("Retrieved workbook: {}", sWorkBookName);
                    break;
                }
            } catch (Exception e) {
                log.error("Error during workbook retrieval for '{}': {}", sWorkBookName, e.getMessage());
            }
        }

        if (!fFound) {
            log.debug("Workbook not opened. Please open workbook: {}", sWorkBookName);
        }

        return tempBook;
    }

    /**
     * Closes a workbook by its index in the list of open workbooks.
     *
     * @param iIndex The index of the workbook to close
     */
    public void closeWorkBook(int iIndex) {
        ExcelWorkBook tempBook = null;
        try {
            tempBook = this.exlWorkBooks.get(iIndex);
            String bookName = tempBook.getExcelBookName();
            tempBook.closeWorkBook();
            this.exlWorkBooks.remove(iIndex);
            log.debug("Closed workbook at index {}: {}", iIndex, bookName);
        } catch (IndexOutOfBoundsException e) {
            log.error("Error closing workbook - index out of bounds: {} - {}", iIndex, e.getMessage());
        } catch (Exception e) {
            log.error("Error closing workbook at index {}: {}", iIndex, e.getMessage());
        }
    }

    /**
     * Closes a workbook by its name from the list of open workbooks.
     *
     * @param sWorkBookName The name of the workbook to close
     */
    public void closeWorkBook(String sWorkBookName) {
        int iIndex = -1;
        boolean fFound = false;

        if (sWorkBookName == null || sWorkBookName.trim().isEmpty()) {
            log.error("Workbook name cannot be null or empty");
            return;
        }

        for (ExcelWorkBook excelWorkBook : this.exlWorkBooks) {
            ++iIndex;
            try {
                if (excelWorkBook.getExcelBookName().equalsIgnoreCase(sWorkBookName)) {
                    excelWorkBook.closeWorkBook();
                    fFound = true;
                    log.debug("Found and closed workbook: {}", sWorkBookName);
                    break;
                }
            } catch (Exception e) {
                log.error("Error while trying to close workbook '{}': {}", sWorkBookName, e.getMessage());
            }
        }

        if (iIndex >= 0 && fFound) {
            this.exlWorkBooks.remove(iIndex);
            log.debug("Removed workbook from list: {}", sWorkBookName);
        } else {
            log.warn("Workbook '{}' not available for close operation", sWorkBookName);
        }
    }

    /**
     * Closes all open workbooks in the application.
     * This method attempts to close each workbook and clears the workbook list.
     */
    public void closeAllWorkBooks() {
        int closedCount = 0;

        log.debug("Closing all workbooks. Total count: {}", exlWorkBooks.size());

        for (ExcelWorkBook excelWorkBook : this.exlWorkBooks) {
            try {
                String bookName = excelWorkBook.getExcelBookName();
                excelWorkBook.closeWorkBook();
                closedCount++;
                log.debug("Closed workbook: {}", bookName);
            } catch (Exception e) {
                log.warn("Workbook {} could not be closed: {}", excelWorkBook.getExcelBookName(), e.getMessage());
            }
        }

        this.exlWorkBooks.clear();
        log.info("All workbooks closed and list cleared. Successfully closed: {}", closedCount);
    }

}
