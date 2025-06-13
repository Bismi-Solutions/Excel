package solutions.bismi.excel;

import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.file.InvalidPathException;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;

/**
 * Represents an Excel workbook.
 * Provides methods for creating, opening, and manipulating Excel workbooks and their sheets.
 */
@NoArgsConstructor
@AllArgsConstructor
public class ExcelWorkBook {

    // List to keep currently available sheets in excel
    @Getter
    @Setter
    protected List<ExcelWorkSheet> excelSheets = new ArrayList<>();

    @Getter
    @Setter
    Workbook wb;

    @Getter
    @Setter
    String excelBookName = "";

    @Getter
    @Setter
    String excelBookPath = "";

    @Getter
    @Setter
    FileInputStream inputStream = null;

    @Getter
    @Setter
    FileOutputStream outputStream = null;

    @Getter
    @Setter
    File xlFile = null;

    @Getter
    @Setter
    String xlFileExtension = null;

    @Getter
    @Setter
    private Logger log = LogManager.getLogger(ExcelWorkBook.class);


    /**
     * Constructor with essential parameters.
     *
     * @param wb            The workbook
     * @param excelBookName The name of the Excel book
     * @param excelBookPath The path of the Excel book
     */
    public ExcelWorkBook(Workbook wb, String excelBookName, String excelBookPath) {
        if (wb == null) {
            // Allow wb to be null initially, e.g. before createWorkbook or openWorkbook is called
            // However, methods using this.wb will need to check for null.
            log.warn("ExcelWorkBook constructed with a null Workbook object. Name: '{}', Path: '{}'", excelBookName, excelBookPath);
        }
        this.wb = wb;
        this.excelBookName = excelBookName;
        this.excelBookPath = excelBookPath;
        this.excelSheets = new ArrayList<>();
    }

    /**
     * Closes the workbook and releases resources.
     *
     * @return true if the workbook was closed successfully, false otherwise
     */
    protected boolean closeWorkBook() {
        try {
            if (this.wb != null) {
                this.wb.close();
                this.wb = null; // Indicate that workbook is closed
                log.debug("Closed workbook: {}", this.excelBookName);
                return true;
            } else {
                log.warn("Attempted to close a workbook that is already null or was never initialized: {}", this.excelBookName);
                return false; // Or true if 'already closed' is considered a success
            }
        } catch (IOException e) {
            log.error("Error in closing workbook '{}': " + e.getMessage(), this.excelBookName, e);
            return false;
        }
    }

    /**
     * Creates a new Excel workbook at the specified file path.
     *
     * @param sCompleteFileName The complete file path and name for the new workbook
     * @return The newly created ExcelWorkBook object
     */
    protected ExcelWorkBook createWorkBook(String sCompleteFileName) {
        ExcelWorkBook xlWbook = null; // Initialize to null
        if (sCompleteFileName == null || sCompleteFileName.trim().isEmpty()) {
            log.error("Cannot create workbook, the complete file name is null or empty.");
            return null;
        }
        try {
            // isValidPath can modify this.excelBookPath, so call it before constructing the final ExcelWorkBook object
            boolean pathIsValid = isValidPath(sCompleteFileName);
            if (pathIsValid) {
                this.xlFile = new File(sCompleteFileName);
                // getFileName also calls isValidPath, which is redundant if called above.
                // Consider refactoring getFileName or isValidPath to avoid this.
                // For now, ensure getFileName is robust.
                this.excelBookName = getFileName(sCompleteFileName); // Sets this.excelBookName and this.xlFileExtension

                String[] fileExtn = this.excelBookName.split("[.]");

                if (fileExtn.length > 1) {
                    String extension = fileExtn[fileExtn.length - 1].toLowerCase();

                    if ("xlsx".equals(extension) || "xlsm".equals(extension)) {
                        // Ensure wb is not null after this block
                        if ("xlsx".equals(extension) || "xlsm".equals(extension)) {
                            this.wb = new XSSFWorkbook();
                            createInitialSheet(sCompleteFileName, "XLSX"); // this.wb should be set before calling this
                        } else if ("xls".equals(extension)) {
                            this.wb = new HSSFWorkbook();
                            createInitialSheet(sCompleteFileName, "XLS"); // this.wb should be set before calling this
                        } else {
                            log.warn("File {} has an unsupported extension: {}", this.excelBookName, extension);
                            // this.wb remains null or as previously set
                        }
                    } else {
                        log.warn("File {} has no extension or extension could not be determined.", this.excelBookName);
                        // this.wb remains null or as previously set
                    }
                } else {
                    log.warn("File {} has no extension", this.excelBookName);
                    // this.wb remains null or as previously set
                }
            } else {
                log.error("Invalid path provided for workbook creation: {}", sCompleteFileName);
            }
        } catch (IOException e) {
            log.error("IO error creating workbook '{}': " + e.getMessage(), sCompleteFileName, e);
        } catch (Exception e) {
            log.error("Error creating workbook '{}': " + e.getMessage(), sCompleteFileName, e);
        }

        // Construct with potentially null wb if creation failed.
        xlWbook = new ExcelWorkBook(this.wb, this.excelBookName, this.excelBookPath);
        // If this.wb is null, xlWbook will also have a null wb. Callers should check.
        return xlWbook;
    }

    /**
     * Creates an initial sheet in the workbook and saves it
     *
     * @param fileName The file path to save the workbook to
     * @param fileType The type of file being created (for logging)
     * @throws IOException If an I/O error occurs
     */
    private void createInitialSheet(String fileName, String fileType) throws IOException {
        if (this.wb == null) {
            log.error("Workbook (wb) is null, cannot create initial sheet for {}.", fileName);
            throw new IOException("Workbook not initialized before creating initial sheet.");
        }
        if (fileName == null || fileName.trim().isEmpty()) {
            log.error("File name is null or empty, cannot create initial sheet.");
            throw new IOException("File name is null or empty for initial sheet creation.");
        }
        try (OutputStream fileOut = new FileOutputStream(fileName)) {
            this.wb.createSheet("Sheet1"); // Creates and returns the sheet, but we don't capture it here.
            this.wb.write(fileOut);
            updateSheetList(); // updateSheetList needs this.wb to be non-null.
            log.info("Created {} file: {}", fileType, this.excelBookName);
        }
    }

    /**
     * Extracts the file name from a complete file path.
     *
     * @param sCompleteFilePath The complete file path
     * @return The file name, or an empty string if an error occurs
     */
    private String getFileName(String sCompleteFilePath) {
        if (sCompleteFilePath == null || sCompleteFilePath.trim().isEmpty()) {
            log.warn("File path is null or empty in getFileName. Returning empty string.");
            return "";
        }

        try {
            // isValidPath also tries to create directories, which might not be desired for just getting a name.
            // For now, we assume it's okay or rely on its robustness.
            // However, direct parsing of sCompleteFilePath might be safer if isValidPath has side effects.
            // Paths.get(sCompleteFilePath).getFileName().toString() is a more direct way if path is valid.

            // The existing isValidPath call is problematic here if sCompleteFilePath is just a name like "file.xlsx"
            // as it would try to create a directory "file.xlsx" if it doesn't exist.
            // For now, let's bypass isValidPath here for name extraction if it causes issues.
            // A simpler approach:
            File tempFile = new File(sCompleteFilePath);
            String fileName = tempFile.getName();

            if (fileName.isEmpty()) {
                log.warn("Extracted file name is empty for path: {}", sCompleteFilePath);
                return "";
            }
            this.excelBookName = fileName; // Set the book name here

            // Check if file has an extension before extracting it
            int lastDotIndex = fileName.lastIndexOf(".");
            if (lastDotIndex > 0 && lastDotIndex < fileName.length() - 1) {
                this.xlFileExtension = fileName.substring(lastDotIndex);
            } else {
                this.xlFileExtension = ""; // No extension or dot is at the start
            }

            log.debug("Extracted file name: {}, extension: {}", fileName, this.xlFileExtension);
            return fileName;

            /* Original logic preserved for reference, but seems problematic:
            if (isValidPath(sCompleteFilePath)) { // This call can have side effects (dir creation)
                File file = new File(sCompleteFilePath); // This file object might be a directory due to isValidPath
                String fileName = file.getName(); // This could be a directory name
            }
            */
        } catch (StringIndexOutOfBoundsException e) {
            log.error("Error extracting file extension from '{}': " + e.getMessage(), sCompleteFilePath, e);
        } catch (Exception e) {
            log.error("Error extracting file name from '{}': " + e.getMessage(), sCompleteFilePath, e);
        }

        return ""; // Return empty string on error
    }

    /**
     * Validates a file path and creates any missing directories.
     *
     * @param path The file path to validate
     * @return true if the path is valid, false otherwise
     */
    private boolean isValidPath(String path) {
        if (path == null || path.trim().isEmpty()) {
            log.error("Path is null or empty, cannot validate.");
            return false;
        }
        try {
            java.nio.file.Path inputPath = Paths.get(path).normalize(); // Normalize to handle "." and ".."
            java.nio.file.Path parentNioPath = inputPath.getParent();

            if (parentNioPath != null) {
                if (!Files.exists(parentNioPath)) {
                    try {
                        Files.createDirectories(parentNioPath);
                        log.info("Created parent directory/directories: {}", parentNioPath.toString());
                    } catch (IOException e) {
                        log.warn("Failed to create parent directories for {}: {}. This could be due to permission issues or an invalid path segment.", parentNioPath.toString(), e.getMessage(), e);
                        return false; // Indicate failure if directory creation fails
                    }
                } else if (!Files.isDirectory(parentNioPath)) {
                    log.warn("Invalid path: The parent part of the path {} is a file, not a directory.", inputPath.toString());
                    return false; // Parent part exists but is not a directory
                }
            }
            // else: Path has no explicit parent (e.g., "file.xlsx" in current dir, or it's a root component).
            // excelBookPath will be set based on resolved absolute parent below.

            // Consistent assignment of this.excelBookPath to the absolute parent directory
            java.nio.file.Path absoluteInputPath = inputPath.toAbsolutePath();
            java.nio.file.Path resolvedParentPath = absoluteInputPath.getParent();

            if (resolvedParentPath != null) {
                this.excelBookPath = resolvedParentPath.toString();
            } else {
                // This case implies the input path is a root directory itself (e.g. "C:/", "/").
                // Or, if inputPath was something like "file.txt", absoluteInputPath would be "/current/working/dir/file.txt",
                // and its parent would be "/current/working/dir".
                // If resolvedParentPath is null even after toAbsolutePath(), it's an unusual scenario.
                // Fallback to current working directory as a sensible default for where the "book path" might be.
                this.excelBookPath = Paths.get(".").toAbsolutePath().normalize().toString();
                log.debug("Resolved parent path for '{}' (absolute: '{}') was null, defaulting excelBookPath to CWD: {}",
                          path, absoluteInputPath.toString(), this.excelBookPath);
            }
            log.debug("Path validated successfully for '{}'. excelBookPath set to: {}", path, this.excelBookPath);
            return true;
        } catch (InvalidPathException ipe) {
            log.warn("Invalid path structure for '{}': {}", path, ipe.getMessage(), ipe);
            return false;
        } catch (SecurityException se) {
            log.error("SecurityException during path validation or directory creation for '{}': {}", path, se.getMessage(), se);
            return false;
        } catch (Exception e) { // Catch any other unexpected exception during path operations
            log.error("Unexpected error during path validation for '{}': {}", path, e.getMessage(), e);
            return false;
        }
    }


    /**
     * Updates the list of worksheets in the workbook.
     * This method clears the existing list and repopulates it with all sheets from the workbook.
     */
    protected void updateSheetList() {
        excelSheets.clear();
        if (this.wb == null) {
            log.warn("Workbook (wb) is null. Cannot update sheet list.");
            return;
        }
        int sheetCount = this.wb.getNumberOfSheets();

        try {
            for (int i = 0; i < sheetCount; i++) {
                String sheetName = this.wb.getSheetName(i);
                Sheet sheet = this.wb.getSheet(sheetName);
                ExcelWorkSheet sheetTemp = createWorkSheet(sheet); // createWorkSheet already checks for sheet==null
                if (sheetTemp != null) {
                    excelSheets.add(sheetTemp);
                }
            }
            log.debug("Updated sheet list in workbook '{}'. Total sheets: {}", this.excelBookName, sheetCount);
        } catch (Exception e) {
            log.error("Error updating sheet list for workbook '{}': " + e.getMessage(), this.excelBookName, e);
        }
    }

    /**
     * Opens an existing Excel workbook from the specified file path.
     *
     * @param sCompleteFileName The complete file path and name of the workbook to open
     * @return The opened ExcelWorkBook object, or null if an error occurs
     */
    protected ExcelWorkBook openWorkbook(String sCompleteFileName) {
        ExcelWorkBook xlWbook = null;
        log.debug("Opening Excel workbook: {}", sCompleteFileName);

        if (sCompleteFileName == null || sCompleteFileName.trim().isEmpty()) {
            log.error("Cannot open workbook, the complete file name is null or empty.");
            return null;
        }

        try {
            if (isValidPath(sCompleteFileName)) { // isValidPath sets this.excelBookPath
                this.xlFile = new File(sCompleteFileName);
                this.excelBookName = getFileName(sCompleteFileName); // getFileName sets this.excelBookName and this.xlFileExtension
                xlWbook = openWorkbookFromFile();
            } else {
                log.error("Invalid path provided for opening workbook: {}", sCompleteFileName);
            }
        } catch (IOException e) {
            log.error("IO error opening workbook '{}': " + e.getMessage(), sCompleteFileName, e);
        } catch (Exception e) {
            log.error("Error opening workbook '{}': " + e.getMessage(), sCompleteFileName, e);
        }
        return xlWbook;
    }

    private ExcelWorkBook openWorkbookFromFile() throws IOException {
        ExcelWorkBook xlWbook = null;
        if (this.xlFile == null) {
            log.error("Excel file (xlFile) is null. Cannot open from file.");
            throw new IOException("Excel file object is null.");
        }
        if (!this.xlFile.exists()) {
            log.error("Excel file does not exist: {}", this.xlFile.getAbsolutePath());
            throw new FileNotFoundException("Excel file not found: " + this.xlFile.getAbsolutePath());
        }

        try (FileInputStream fileIn = new FileInputStream(this.xlFile)) {
            this.wb = WorkbookFactory.create(fileIn);
            // this.inputStream = null; // inputStream member is not used after this, so setting to null is fine.
                                     // However, the member `inputStream` itself seems to be of little use in this class.

            if (this.wb == null) {
                log.error("WorkbookFactory.create returned null for file: {}", this.xlFile.getAbsolutePath());
                throw new IOException("Failed to create workbook from file.");
            }

            // Get the active sheet index before creating the ExcelWorkBook
            // Active sheet index might be -1 if no sheet is explicitly set active (e.g. for a new workbook from Apache POI < 3.17)
            // Or 0 for newer versions.
            int activeSheetIndex = 0;
            try {
                activeSheetIndex = this.wb.getActiveSheetIndex();
            } catch (Exception e) {
                log.warn("Could not get active sheet index for workbook '{}'. Defaulting to 0. Error: {}", this.excelBookName, e.getMessage());
                // Default to 0 if there's an issue (e.g. workbook has no sheets yet, though WorkbookFactory usually handles this)
            }

            xlWbook = new ExcelWorkBook(this.wb, this.excelBookName, this.excelBookPath);
            xlWbook.updateSheetList(); // updateSheetList needs xlWbook.wb to be non-null.

            // Ensure the active sheet is properly set
            if (this.wb.getNumberOfSheets() > 0) {
                if (activeSheetIndex < 0 || activeSheetIndex >= this.wb.getNumberOfSheets()) {
                    log.warn("Active sheet index {} is out of bounds for workbook '{}' with {} sheets. Setting to 0.", activeSheetIndex, this.excelBookName, this.wb.getNumberOfSheets());
                    activeSheetIndex = 0; // Default to first sheet if index is invalid
                }
                this.wb.setActiveSheet(activeSheetIndex);
                log.debug("Set active sheet to: {}", this.wb.getSheetName(activeSheetIndex));
            } else {
                log.warn("Workbook '{}' has no sheets. Cannot set active sheet.", this.excelBookName);
            }
        }
        return xlWbook;
    }

    /**
     * Saves the workbook to the file.
     *
     * @param logMessage Message to log on successful save
     * @return true if successful, false otherwise
     */
    private boolean saveWorkbookToFile(String logMessage) {
        // inputStream is a member of ExcelWorkBook but seems unused for saving.
        // It's primarily for opening. It's set to null in openWorkbookFromFile.
        // No need to close this.inputStream here as it's likely already null or managed by try-with-resources where it was used.

        if (this.xlFile == null) {
            log.error("Excel file (xlFile) is null. Cannot save workbook.");
            return false;
        }
        if (this.wb == null) {
            log.error("Workbook (wb) is null. Cannot save workbook: {}", this.excelBookName);
            return false;
        }

        try (FileOutputStream fileOut = new FileOutputStream(this.xlFile)) {
            this.wb.write(fileOut);
            log.debug(logMessage);
            return true;
        } catch (IOException e) {
            log.error("IO error saving workbook '{}': " + e.getMessage(), this.excelBookName, e);
            return false;
        } catch (Exception e) {
            log.error("Error saving workbook '{}': " + e.getMessage(), this.excelBookName, e);
            return false;
        }
    }

    /**
     * Adds a new worksheet to the workbook.
     *
     * @param sSheetName The name of the new worksheet
     * @return The newly created ExcelWorkSheet object
     */
    public ExcelWorkSheet addSheet(String sSheetName) {
        if (this.wb == null) {
            log.error("Workbook (wb) is null. Cannot add sheet: {}", sSheetName);
            return null;
        }
        if (sSheetName == null || sSheetName.trim().isEmpty()) {
            log.error("Sheet name cannot be null or empty.");
            return null;
        }

        // The old ExcelWorkSheet.addSheet was problematic.
        // Directly create sheet in this workbook:
        Sheet newPoiSheet = this.wb.createSheet(sSheetName);
        if (newPoiSheet == null) {
            log.error("Failed to create sheet '{}' in workbook '{}'.", sSheetName, this.excelBookName);
            return null;
        }

        if (saveWorkbookToFile("Added new sheet: " + sSheetName + " to workbook " + this.excelBookName)) {
            updateSheetList(); // Refresh the list of ExcelWorkSheet objects
            return getExcelSheet(sSheetName); // Retrieve the newly created sheet wrapper
        } else {
            log.error("Failed to save workbook after adding sheet: {}. Sheet was added to in-memory model but not persisted.", sSheetName);
            // Depending on desired atomicity, we might want to remove the sheet from the in-memory model if save fails.
            // For now, just returning null and logging.
            // Consider this: if (newPoiSheet != null) { this.wb.removeSheetAt(this.wb.getSheetIndex(newPoiSheet)); }
            // However, newPoiSheet is not directly accessible here without passing its name or index.
            // A simpler approach for now is to document that the in-memory model might be ahead of the saved state.
            int sheetIndexToRemove = this.wb.getSheetIndex(sSheetName);
            if (sheetIndexToRemove != -1) {
                 this.wb.removeSheetAt(sheetIndexToRemove);
                 log.info("Removed sheet '{}' from in-memory model due to save failure.", sSheetName);
            }
            updateSheetList(); // Refresh list after potential removal
            return null;
        }
    }

    /**
     * Adds multiple worksheets to the workbook.
     *
     * @param strArrSheets Array of sheet names to add
     */
    public void addSheets(String[] strArrSheets) {
        if (this.wb == null) {
            log.error("Workbook (wb) is null. Cannot add multiple sheets.");
            return;
        }
        if (strArrSheets == null || strArrSheets.length == 0) {
            log.warn("Sheet array is null or empty. No sheets to add.");
            return;
        }

        int sheetsAddedCount = 0;
        for (String sSheetName : strArrSheets) {
            if (sSheetName != null && !sSheetName.trim().isEmpty()) {
                if (this.wb.getSheet(sSheetName) == null) { // Avoid creating duplicate sheets
                    this.wb.createSheet(sSheetName);
                    sheetsAddedCount++;
                } else {
                    log.warn("Sheet '{}' already exists in workbook '{}'. Skipping.", sSheetName, this.excelBookName);
                }
            } else {
                log.warn("Null or empty sheet name encountered in addSheets array. Skipping.");
            }
        }

        if (sheetsAddedCount > 0) {
            boolean saveSuccess = saveWorkbookToFile("Added " + sheetsAddedCount + " new sheets to workbook " + this.excelBookName);
            if (saveSuccess) {
                updateSheetList();
            } else {
                log.error("Failed to save workbook after adding {} sheets. In-memory model was modified but changes were not persisted.", sheetsAddedCount);
                // As with addSheet, consider if sheets added in memory should be rolled back.
                // This is more complex for multiple sheets. For now, logging the discrepancy.
            }
        } else {
            log.info("No new sheets were actually added to workbook {}.", this.excelBookName);
        }
    }


    /**
     * Gets the full path of the Excel file.
     *
     * @return The full path of the Excel file
     */
    private String getFullPath() {
        String bookPath = (this.excelBookPath != null) ? this.excelBookPath : "";
        String bookName = (this.excelBookName != null) ? this.excelBookName : "";

        if (bookPath.isEmpty() && bookName.isEmpty()) {
            log.warn("excelBookPath and excelBookName are both empty/null. Full path is indeterminate.");
            return ""; // Or throw an IllegalStateException
        }

        if (bookPath.isEmpty()) {
            // If only bookName is present, assume it's a relative path or just a name.
            return bookName;
        }
        if (bookName.isEmpty()) {
            // If only bookPath is present, this is unusual but return it.
            return bookPath;
        }

        // Use Paths.get for robust, cross-platform path construction.
        try {
            return Paths.get(bookPath, bookName).toString();
        } catch (InvalidPathException e) {
            log.error("Error constructing full path from bookPath '{}' and bookName '{}': {}", bookPath, bookName, e.getMessage(), e);
            // Fallback to simpler concatenation if Paths.get fails, though this indicates a deeper issue with path/name.
            // This fallback might still produce an invalid path but is better than crashing.
            return bookPath + File.separator + bookName;
        }
    }

    /**
     * Creates an ExcelWorkSheet object from a Sheet object.
     *
     * @param sheet The Sheet object to create an ExcelWorkSheet from
     * @return The created ExcelWorkSheet object, or null if the sheet is null
     */
    private ExcelWorkSheet createWorkSheet(Sheet sheet) {
        if (sheet == null) {
            log.warn("Input sheet is null. Cannot create ExcelWorkSheet.");
            return null;
        }
        if (this.wb == null) {
            log.warn("Workbook (this.wb) is null. Cannot create ExcelWorkSheet for sheet: {}", sheet.getSheetName());
            return null;
        }

        String fullPath = getFullPath();
        if (fullPath.isEmpty()) {
            log.error("Full path is empty for workbook. Cannot create ExcelWorkSheet for sheet: {}. Ensure excelBookPath and excelBookName are set.", sheet.getSheetName());
            // Depending on how ExcelWorkSheet uses sCompleteFileName, this might be critical.
            // If ExcelWorkSheet requires a valid file path for saving, this is an issue.
            // For now, we pass it, but ExcelWorkSheet constructor might throw if it's invalid.
        }

        try {
            // The inputStream and outputStream members of ExcelWorkBook seem to be misapplied here.
            // inputStream is typically for reading the whole workbook, not for individual sheets.
            // outputStream is not managed by ExcelWorkBook for writing; sheets are saved via wb.write().
            // Passing null for these if they are not genuinely used by ExcelWorkSheet might be cleaner.
            // For now, passing the instance members as per original logic.
            return new ExcelWorkSheet(sheet, sheet.getSheetName(), this.wb, this.inputStream, this.outputStream, fullPath);
        } catch (Exception e) {
            log.error("Error creating ExcelWorkSheet wrapper for sheet '{}': " + e.getMessage(), sheet.getSheetName(), e);
            return null;
        }
    }

    /**
     * Gets a worksheet by name.
     *
     * @param sSheetName The name of the worksheet to get
     * @return The worksheet with the specified name, or null if not found
     */
    public ExcelWorkSheet getExcelSheet(String sSheetName) {
        if (this.wb == null) {
            log.warn("Workbook (wb) is null. Cannot get sheet: {}", sSheetName);
            return null;
        }
        if (sSheetName == null || sSheetName.trim().isEmpty()) {
            log.warn("Sheet name is null or empty. Cannot get sheet.");
            return null;
        }
        try {
            Sheet sheet = this.wb.getSheet(sSheetName);
            if (sheet == null) {
                log.warn("Sheet '{}' not found in workbook '{}'.", sSheetName, this.excelBookName);
                return null;
            }
            return createWorkSheet(sheet);
        } catch (Exception e) {
            log.error("Error in getting sheet '{}' from workbook '{}': " + e.getMessage(), sSheetName, this.excelBookName, e);
            return null;
        }
    }

    /**
     * Gets a worksheet by index.
     *
     * @param iIndex The index of the worksheet to get
     * @return The worksheet at the specified index, or null if not found
     */
    public ExcelWorkSheet getExcelSheet(int iIndex) {
        if (this.wb == null) {
            log.warn("Workbook (wb) is null. Cannot get sheet at index: {}", iIndex);
            return null;
        }
        try {
            if (iIndex < 0 || iIndex >= this.wb.getNumberOfSheets()) {
                log.warn("Sheet index {} is out of bounds for workbook '{}' with {} sheets.", iIndex, this.excelBookName, this.wb.getNumberOfSheets());
                return null;
            }
            Sheet sheet = this.wb.getSheetAt(iIndex);
            if (sheet == null) { // Should not happen if index is valid, but as a safeguard
                log.warn("Sheet at index {} is null in workbook '{}'.", iIndex, this.excelBookName);
                return null;
            }
            return createWorkSheet(sheet);
        } catch (Exception e) { // Catching generic Exception is broad, consider specific POI exceptions if known
            log.error("Error in getting sheet at index {} from workbook '{}': " + e.getMessage(), iIndex, this.excelBookName, e);
            return null;
        }
    }

    /**
     * Gets the active worksheet.
     *
     * @return The active worksheet, or null if not found
     */
    public ExcelWorkSheet getActiveSheet() {
        if (this.wb == null) {
            log.warn("Workbook (wb) is null. Cannot get active sheet.");
            return null;
        }
        try {
            int iIndex = getActiveSheetIndex(); // This already checks for wb null
            if (iIndex < 0 || iIndex >= this.wb.getNumberOfSheets()) { // Check bounds after getting index
                 log.warn("Active sheet index {} is out of bounds for workbook '{}'. Number of sheets: {}", iIndex, this.excelBookName, this.wb.getNumberOfSheets());
                 return null;
            }
            Sheet sheet = this.wb.getSheetAt(iIndex);
            if (sheet == null) {
                log.warn("Active sheet at index {} is null in workbook '{}'.", iIndex, this.excelBookName);
                return null;
            }
            return createWorkSheet(sheet);
        } catch (Exception e) {
            log.error("Error in getting active sheet from workbook '{}': " + e.getMessage(), this.excelBookName, e);
            return null;
        }
    }


    public int getSheetCount() {
        if (this.wb == null) {
            log.warn("Workbook (wb) is null. Returning sheet count 0 from excelSheets list size.");
            // excelSheets list might be stale if wb became null after list was populated.
            // However, direct access to wb.getNumberOfSheets() is better if wb is available.
            return excelSheets.size(); // Fallback to list size
        }
        return this.wb.getNumberOfSheets(); // More accurate if wb is available
    }

    public String getActiveSheetName() {
        if (this.wb == null) {
            log.warn("Workbook (wb) is null. Cannot get active sheet name.");
            return null;
        }
        try {
            return this.wb.getSheetName(this.wb.getActiveSheetIndex());
        } catch (Exception e) {
            log.error("Error getting active sheet name for workbook '{}': " + e.getMessage(), this.excelBookName, e);
            return null;
        }
    }

    /**
     * @return
     */
    public int getActiveSheetIndex() {
        if (this.wb == null) {
            log.warn("Workbook (wb) is null. Cannot get active sheet index. Returning -1 or a default.");
            return -1; // Or handle as per application's error strategy, e.g., throw IllegalStateException
        }
        return this.wb.getActiveSheetIndex();
    }


    /**
     * Saves the workbook to the file
     *
     * @return true if successful, false otherwise
     */
    public boolean saveWorkbook() {
        // Ensure excelBookName is sensible for the log message, even if xlFile is primary determinant for path
        String bookNameToLog = (this.excelBookName != null && !this.excelBookName.isEmpty()) ? this.excelBookName : (this.xlFile != null ? this.xlFile.getName() : "UnknownWorkbook");
        return saveWorkbookToFile("Workbook saved successfully: " + bookNameToLog);
    }


}
