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
            this.wb.close();
            this.wb = null;
            log.debug("Closed workbook: {}", this.excelBookName);
            return true;
        } catch (IOException e) {
            log.error("Error in closing workbook: {}", e.getMessage());
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
        ExcelWorkBook xlWbook;
        try {
            if (isValidPath(sCompleteFileName)) {
                this.xlFile = new File(sCompleteFileName);
                this.excelBookName = getFileName(sCompleteFileName);
                String[] fileExtn = this.excelBookName.split("[.]");

                if (fileExtn.length > 1) {
                    String extension = fileExtn[fileExtn.length - 1].toLowerCase();

                    if ("xlsx".equals(extension) || "xlsm".equals(extension)) {
                        wb = new XSSFWorkbook();
                        createInitialSheet(sCompleteFileName, "XLSX");
                    } else if ("xls".equals(extension)) {
                        wb = new HSSFWorkbook();
                        createInitialSheet(sCompleteFileName, "XLS");
                    }
                } else {
                    log.warn("File {} has no extension", this.excelBookName);
                }
            }
        } catch (IOException e) {
            log.error("IO error creating workbook: {}", e.getMessage());
        } catch (Exception e) {
            log.error("Error creating workbook: {}", e.getMessage());
        }

        xlWbook = new ExcelWorkBook(this.wb, this.excelBookName, this.excelBookPath);
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
        try (OutputStream fileOut = new FileOutputStream(fileName)) {
            wb.createSheet("Sheet1");
            wb.write(fileOut);
            updateSheetList();
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
            log.error("File path is null or empty");
            return "";
        }

        try {
            if (isValidPath(sCompleteFilePath)) {
                File file = new File(sCompleteFilePath);
                String fileName = file.getName();
                this.excelBookName = fileName;

                // Check if file has an extension before extracting it
                int lastDotIndex = fileName.lastIndexOf(".");
                if (lastDotIndex > 0 && lastDotIndex < fileName.length() - 1) {
                    this.xlFileExtension = fileName.substring(lastDotIndex);
                } else {
                    this.xlFileExtension = "";
                }

                log.debug("Extracted file name: {}", fileName);
                return fileName;
            }
        } catch (StringIndexOutOfBoundsException e) {
            log.error("Error extracting file extension: {}", e.getMessage());
        } catch (Exception e) {
            log.error("Error extracting file name: {}", e.getMessage());
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
        try {
            Paths.get(path);
            log.debug("Verified path is valid: {}", path);
            File file = new File(path);

            if (!file.isDirectory()) {
                file = file.getParentFile();
                new File(file.getPath()).mkdirs();
                log.debug("Created missing directories");
                this.excelBookPath = file.getPath();
            } else {
                new File(file.getPath()).mkdirs();
                this.excelBookPath = file.getPath();
            }
            return true;
        } catch (InvalidPathException | NullPointerException ex) {
            log.error("Invalid path: {}", ex.getMessage());
            return false;
        }
    }


    /**
     * Updates the list of worksheets in the workbook.
     * This method clears the existing list and repopulates it with all sheets from the workbook.
     */
    protected void updateSheetList() {
        excelSheets.clear();
        int sheetCount = this.wb.getNumberOfSheets();

        try {
            for (int i = 0; i < sheetCount; i++) {
                String sheetName = this.wb.getSheetName(i);
                Sheet sheet = this.wb.getSheet(sheetName);
                ExcelWorkSheet sheetTemp = createWorkSheet(sheet);
                if (sheetTemp != null) {
                    excelSheets.add(sheetTemp);
                }
            }
            log.debug("Updated sheet list in workbook. Total sheets: {}", sheetCount);
        } catch (Exception e) {
            log.error("Error updating sheet list: {}", e.getMessage());
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
        try {
            if (isValidPath(sCompleteFileName)) {
                this.xlFile = new File(sCompleteFileName);
                this.excelBookName = getFileName(sCompleteFileName);
                xlWbook = openWorkbookFromFile();
            }
        } catch (IOException e) {
            log.error("IO error opening workbook: {}", e.getMessage());
        } catch (Exception e) {
            log.error("Error opening workbook: {}", e.getMessage());
        }
        return xlWbook;
    }

    private ExcelWorkBook openWorkbookFromFile() throws IOException {
        ExcelWorkBook xlWbook = null;
        try (FileInputStream fileIn = new FileInputStream(this.xlFile)) {
            this.wb = WorkbookFactory.create(fileIn);
            this.inputStream = null; // We don't need to store the input stream anymore
            // Get the active sheet index before creating the ExcelWorkBook
            int activeSheetIndex = this.wb.getActiveSheetIndex();
            xlWbook = new ExcelWorkBook(this.wb, this.excelBookName, this.excelBookPath);
            updateSheetList();
            // Ensure the active sheet is properly set
            if (activeSheetIndex >= 0 && activeSheetIndex < this.wb.getNumberOfSheets()) {
                this.wb.setActiveSheet(activeSheetIndex);
                log.debug("Set active sheet to: {}", this.wb.getSheetName(activeSheetIndex));
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
        try {
            // Close input stream if open
            if (inputStream != null) {
                try {
                    inputStream.close();
                    inputStream = null;
                } catch (IOException e) {
                    log.warn("Error closing input stream: {}", e.getMessage());
                }
            }

            // Use try-with-resources for output stream
            if (this.xlFile != null && this.wb != null) {
                try (FileOutputStream fileOut = new FileOutputStream(this.xlFile)) {
                    this.wb.write(fileOut);
                    log.debug(logMessage);
                    return true;
                }
            }
        } catch (IOException e) {
            log.error("IO error saving workbook: {}", e.getMessage());
            return false;
        } catch (Exception e) {
            log.error("Error saving workbook: {}", e.getMessage());
            return false;
        }
        return false;
    }

    /**
     * Adds a new worksheet to the workbook.
     *
     * @param sSheetName The name of the new worksheet
     * @return The newly created ExcelWorkSheet object
     */
    public ExcelWorkSheet addSheet(String sSheetName) {
        ExcelWorkSheet wSheet = new ExcelWorkSheet();
        wSheet.wb = this.wb;
        ExcelWorkSheet tempSheet = wSheet.addSheet(sSheetName);

        if (tempSheet != null) {
            saveWorkbookToFile("Added new sheet: " + sSheetName);
            excelSheets.add(tempSheet);
            return getExcelSheet(sSheetName);
        }
        return null;
    }

    /**
     * Adds multiple worksheets to the workbook.
     *
     * @param strArrSheets Array of sheet names to add
     */
    public void addSheets(String[] strArrSheets) {
        ExcelWorkSheet wSheet = new ExcelWorkSheet();
        wSheet.wb = this.wb;
        wSheet.addSheets(strArrSheets);

        saveWorkbookToFile("Added " + strArrSheets.length + " new sheets");
        updateSheetList();
    }


    /**
     * Gets the full path of the Excel file.
     *
     * @return The full path of the Excel file
     */
    private String getFullPath() {
        return excelBookPath + "\\" + excelBookName;
    }

    /**
     * Creates an ExcelWorkSheet object from a Sheet object.
     *
     * @param sheet The Sheet object to create an ExcelWorkSheet from
     * @return The created ExcelWorkSheet object, or null if the sheet is null
     */
    private ExcelWorkSheet createWorkSheet(Sheet sheet) {
        if (sheet == null) {
            return null;
        }

        try {
            return new ExcelWorkSheet(sheet, sheet.getSheetName(), this.wb, this.inputStream, this.outputStream, getFullPath());
        } catch (Exception e) {
            log.debug("Error creating worksheet: {}", e.getMessage());
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
        try {
            Sheet sheet = this.wb.getSheet(sSheetName);
            return createWorkSheet(sheet);
        } catch (Exception e) {
            log.debug("Error in getting sheet {}", sSheetName);
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
        try {
            Sheet sheet = this.wb.getSheetAt(iIndex);
            return createWorkSheet(sheet);
        } catch (Exception e) {
            log.debug("Error in getting sheet index {}", iIndex);
            return null;
        }
    }

    /**
     * Gets the active worksheet.
     *
     * @return The active worksheet, or null if not found
     */
    public ExcelWorkSheet getActiveSheet() {
        try {
            int iIndex = getActiveSheetIndex();
            Sheet sheet = this.wb.getSheetAt(iIndex);
            return createWorkSheet(sheet);
        } catch (Exception e) {
            log.debug("Error in getting active sheet: {}", e.getMessage());
            return null;
        }
    }


    public int getSheetCount() {

        return excelSheets.size();
    }

    public String getActiveSheetName() {

        return this.wb.getSheetName(this.wb.getActiveSheetIndex());
    }

    /**
     * @return
     */
    public int getActiveSheetIndex() {
        return this.wb.getActiveSheetIndex();
    }


    /**
     * Saves the workbook to the file
     *
     * @return true if successful, false otherwise
     */
    public boolean saveWorkbook() {
        return saveWorkbookToFile("Workbook saved successfully: " + this.excelBookName);
    }


}
