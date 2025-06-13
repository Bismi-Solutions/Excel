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
        log.info("Attempting to create excel workbook: {}", sCompleteFileName);
        if (sCompleteFileName == null || sCompleteFileName.trim().isEmpty()) {
            log.error("Workbook file name is null or empty. Cannot create workbook.");
            return null;
        }
        ExcelWorkBook newWorkBook = new ExcelWorkBook(); // Temporary instance to call the method
        ExcelWorkBook createdWorkBook = newWorkBook.createWorkBook(sCompleteFileName); // This method in ExcelWorkBook returns the instance

        if (createdWorkBook != null && createdWorkBook.getWb() != null) {
            exlWorkBooks.add(createdWorkBook); // Add the returned and validated instance
            log.info("Successfully created and added workbook: {}", sCompleteFileName);
            return createdWorkBook;
        } else {
            log.error("Failed to create workbook or workbook object (wb) is null for: {}", sCompleteFileName);
            // createdWorkBook might be non-null but its internal wb could be null if creation failed.
            return null;
        }
    }

    /**
     * Opens an existing Excel workbook with the specified file name.
     *
     * @param sCompleteFileName The complete file path and name of the workbook to open
     * @return The opened ExcelWorkBook object
     */
    public ExcelWorkBook openWorkbook(String sCompleteFileName) {
        log.info("Attempting to open excel workbook: {}", sCompleteFileName);
        if (sCompleteFileName == null || sCompleteFileName.trim().isEmpty()) {
            log.error("Workbook file name is null or empty. Cannot open workbook.");
            return null;
        }
        ExcelWorkBook newWorkBook = new ExcelWorkBook(); // Temporary instance
        ExcelWorkBook openedWorkBook = newWorkBook.openWorkbook(sCompleteFileName); // This method in ExcelWorkBook returns the instance

        if (openedWorkBook != null && openedWorkBook.getWb() != null) {
            exlWorkBooks.add(openedWorkBook);
            log.info("Successfully opened and added workbook: {}", sCompleteFileName);
            return openedWorkBook;
        } else {
            log.error("Failed to open workbook or workbook object (wb) is null for: {}", sCompleteFileName);
            return null;
        }
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
            log.error("Error during workbook retrieval - index {} out of bounds (size {}): " + e.getMessage(), iIndex, this.exlWorkBooks.size(), e);
        } catch (Exception e) {
            log.error("Error during workbook retrieval at index {}: " + e.getMessage(), iIndex, e);
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
                String currentBookName = excelWorkBook.getExcelBookName();
                if (currentBookName != null && currentBookName.equalsIgnoreCase(sWorkBookName)) {
                    fFound = true;
                    tempBook = excelWorkBook;
                    log.debug("Retrieved workbook: {}", sWorkBookName);
                    break;
                } else if (currentBookName == null) {
                    log.warn("Found a workbook with a null name while searching for: {}", sWorkBookName);
                }
            } catch (Exception e) {
                // Log exception if getExcelBookName() or other operations fail for a specific workbook
                log.error("Error accessing workbook details while searching for '{}': " + e.getMessage(), sWorkBookName, e);
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
        ExcelWorkBook tempBook;
        try {
            if (iIndex < 0 || iIndex >= this.exlWorkBooks.size()) {
                log.error("Error closing workbook - index {} is out of bounds (size {}).", iIndex, this.exlWorkBooks.size());
                return;
            }
            tempBook = this.exlWorkBooks.get(iIndex);
            if (tempBook != null) {
                String bookName = tempBook.getExcelBookName() != null ? tempBook.getExcelBookName() : "Unknown (name was null)";
                tempBook.closeWorkBook(); // closeWorkBook in ExcelWorkBook should handle if its wb is null
                this.exlWorkBooks.remove(iIndex);
                log.debug("Closed and removed workbook at index {}: {}", iIndex, bookName);
            } else {
                log.warn("Workbook at index {} was null. Removing from list.", iIndex);
                this.exlWorkBooks.remove(iIndex); // Remove if it's a null entry for some reason
            }
        } catch (IndexOutOfBoundsException e) { // Should be caught by the preliminary check, but as safeguard
            log.error("Error closing workbook - index {} out of bounds: " + e.getMessage(), iIndex, e);
        } catch (Exception e) {
            log.error("Error closing workbook at index {}: " + e.getMessage(), iIndex, e);
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
            ++iIndex; // Current index of excelWorkBook in the loop
            try {
                String currentBookName = excelWorkBook.getExcelBookName();
                if (currentBookName != null && currentBookName.equalsIgnoreCase(sWorkBookName)) {
                    excelWorkBook.closeWorkBook(); // closeWorkBook in ExcelWorkBook handles null wb
                    fFound = true; // Mark that we found and attempted to close the workbook
                    log.debug("Found and closed workbook: {}", sWorkBookName);
                    // We will remove it outside the loop to avoid ConcurrentModificationException
                    break;
                } else if (currentBookName == null) {
                    log.warn("Encountered a workbook with a null name while trying to close: {}", sWorkBookName);
                }
            } catch (Exception e) {
                log.error("Error while processing workbook for closing operation (target: '{}', current: '{}'): " + e.getMessage(),
                          sWorkBookName, excelWorkBook.getExcelBookName(), e);
                // Decide if we should try to remove 'excelWorkBook' if it's problematic
            }
        }

        // Remove the workbook by index if found and closed.
        // iIndex will be the index of the found workbook.

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

        // Iterate backwards to safely remove items from a list
        for (int i = this.exlWorkBooks.size() - 1; i >= 0; i--) {
            ExcelWorkBook excelWorkBook = this.exlWorkBooks.get(i);
            if (excelWorkBook == null) { // Should not happen if list is managed well
                this.exlWorkBooks.remove(i);
                log.warn("Removed a null workbook entry from the list during closeAllWorkBooks.");
                continue;
            }
            try {
                String bookName = excelWorkBook.getExcelBookName();
                if (bookName == null || bookName.trim().isEmpty()) {
                    bookName = "Unnamed Workbook at index " + i;
                }
                excelWorkBook.closeWorkBook(); // ExcelWorkBook.closeWorkBook() is responsible for null checks on its own 'wb'
                closedCount++;
                log.debug("Closed workbook: {}", bookName);
            } catch (Exception e) {
                String bookNameToLog = excelWorkBook.getExcelBookName() != null ? excelWorkBook.getExcelBookName() : "Unknown Workbook at index " + i;
                log.warn("Workbook {} could not be closed cleanly during closeAllWorkBooks: " + e.getMessage(), bookNameToLog, e);
            }
        }

        this.exlWorkBooks.clear();
        log.info("All workbooks closed and list cleared. Successfully closed: {}", closedCount);
    }

}
