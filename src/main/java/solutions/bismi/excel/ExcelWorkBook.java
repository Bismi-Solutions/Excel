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
import java.nio.file.Files;
import java.nio.file.InvalidPathException;
import java.nio.file.Path;
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

    /* ========== State ========== */

    @Getter
    @Setter
    private List<ExcelWorkSheet> excelSheets = new ArrayList<>();

    @Getter
    @Setter
    private Workbook wb;
    @Getter
    @Setter
    private String excelBookName = "";
    @Getter
    @Setter
    private String excelBookPath = "";
    @Getter
    @Setter
    private FileInputStream inputStream = null;
    @Getter
    @Setter
    private FileOutputStream outputStream = null;
    @Getter
    @Setter
    private File xlFile = null;
    @Getter
    @Setter
    private String xlFileExtension = null;

    @Getter
    @Setter
    private Logger log = LogManager.getLogger(ExcelWorkBook.class);

    /* ========== Constructors ========== */

    public ExcelWorkBook(Workbook wb, String excelBookName, String excelBookPath) {
        this.wb = wb;
        this.excelBookName = excelBookName;
        this.excelBookPath = excelBookPath;
        this.excelSheets = new ArrayList<>();
    }

    /* ========== Lifecycle ========== */

    protected boolean closeWorkBook() {
        try {
            wb.close();
            wb = null;
            log.debug("Closed workbook: {}", excelBookName);
            return true;
        } catch (IOException e) {
            log.error("Error closing workbook: {}", e.getMessage());
            return false;
        }
    }

    /* ========== Creation / Opening ========== */

    protected ExcelWorkBook createWorkBook(String fullFileName) {
        try {
            if (isValidPath(fullFileName)) {
                xlFile = new File(fullFileName);
                excelBookName = getFileName(fullFileName);

                String[] parts = excelBookName.split("\\.");
                if (parts.length > 1) {
                    String ext = parts[parts.length - 1].toLowerCase();
                    if ("xlsx".equals(ext) || "xlsm".equals(ext)) {
                        wb = new XSSFWorkbook();
                        createInitialSheet(fullFileName, "XLSX");
                    } else if ("xls".equals(ext)) {
                        wb = new HSSFWorkbook();
                        createInitialSheet(fullFileName, "XLS");
                    }
                } else {
                    log.warn("File {} has no extension", excelBookName);
                }
            }
        } catch (IOException e) {
            log.error("IO error creating workbook: {}", e.getMessage());
        } catch (Exception e) {
            log.error("Error creating workbook: {}", e.getMessage());
        }

        return new ExcelWorkBook(wb, excelBookName, excelBookPath);
    }

    private void createInitialSheet(String fileName, String fileType) throws IOException {
        try (OutputStream fileOut = new FileOutputStream(fileName)) {
            wb.createSheet("Sheet1");
            wb.write(fileOut);
            updateSheetList();
            log.info("Created {} file: {}", fileType, excelBookName);
        }
    }

    /* ========== Helpers: Path & Filename ========== */

    /**
     * Extracts just the file name from a full path.
     */
    private String getFileName(String fullPath) {
        if (fullPath == null || fullPath.trim().isEmpty()) {
            log.error("File path is null or empty");
            return "";
        }

        try {
            File f = new File(fullPath);
            String name = f.getName();

            int dot = name.lastIndexOf('.');
            xlFileExtension = (dot > 0 && dot < name.length() - 1) ? name.substring(dot) : "";

            log.debug("Extracted file name: {}", name);
            return name;
        } catch (Exception e) {
            log.error("Error extracting file name: {}", e.getMessage());
            return "";
        }
    }

    /**
     * Validates a file path, creates missing parent directories, and initialises {@code excelBookPath}.
     *
     * @return {@code true} if the path is usable, {@code false} otherwise.
     */
    private boolean isValidPath(String path) {
        try {
            Path p = Paths.get(path).normalize();
            Path parent = Files.isDirectory(p) ? p : p.getParent();
            if (parent != null) {
                Files.createDirectories(parent);   // mkdir -p
                excelBookPath = parent.toString();
            }
            log.debug("Verified/created path: {}", excelBookPath);
            return true;
        } catch (InvalidPathException | NullPointerException ex) {
            log.error("Invalid path '{}': {}", path, ex.getMessage());
            return false;
        } catch (IOException io) {
            log.error("Cannot create directories for '{}': {}", path, io.getMessage());
            return false;
        }
    }

    /**
     * Builds the absolute workbook path in an OS-neutral way.
     */
    private String getFullPath() {
        return Paths.get(excelBookPath).resolve(excelBookName).normalize().toString();
    }

    /* ========== Sheet List Maintenance ========== */

    protected void updateSheetList() {
        excelSheets.clear();
        int sheetCount = wb.getNumberOfSheets();

        try {
            for (int i = 0; i < sheetCount; i++) {
                Sheet sheet = wb.getSheetAt(i);
                ExcelWorkSheet ew = createWorkSheet(sheet);
                if (ew != null) {
                    excelSheets.add(ew);
                }
            }
            log.debug("Updated sheet list. Total sheets: {}", sheetCount);
        } catch (Exception e) {
            log.error("Error updating sheet list: {}", e.getMessage());
        }
    }

    /* ========== Opening Existing Workbooks ========== */

    protected ExcelWorkBook openWorkbook(String fullFileName) {
        try {
            if (isValidPath(fullFileName)) {
                xlFile = new File(fullFileName);
                excelBookName = getFileName(fullFileName);
                return openWorkbookFromFile();
            }
        } catch (IOException e) {
            log.error("IO error opening workbook: {}", e.getMessage());
        } catch (Exception e) {
            log.error("Error opening workbook: {}", e.getMessage());
        }
        return null;
    }

    private ExcelWorkBook openWorkbookFromFile() throws IOException {
        try (FileInputStream fileIn = new FileInputStream(xlFile)) {
            wb = WorkbookFactory.create(fileIn);
            int activeIdx = wb.getActiveSheetIndex();
            ExcelWorkBook res = new ExcelWorkBook(wb, excelBookName, excelBookPath);
            updateSheetList();
            if (activeIdx >= 0 && activeIdx < wb.getNumberOfSheets()) {
                wb.setActiveSheet(activeIdx);
            }
            return res;
        }
    }

    /* ========== Saving ========== */

    private boolean saveWorkbookToFile(String msg) {
        try {
            if (inputStream != null) {
                inputStream.close();
                inputStream = null;
            }

            try (FileOutputStream fileOut = new FileOutputStream(xlFile)) {
                wb.write(fileOut);
                log.debug(msg);
                return true;
            }
        } catch (IOException e) {
            log.error("IO error saving workbook: {}", e.getMessage());
        } catch (Exception e) {
            log.error("Error saving workbook: {}", e.getMessage());
        }
        return false;
    }

    /* ========== Public Sheet Manipulation API ========== */

    public ExcelWorkSheet addSheet(String sheetName) {
        ExcelWorkSheet ew = new ExcelWorkSheet();
        ew.wb = wb;
        ExcelWorkSheet newS = ew.addSheet(sheetName);

        if (newS != null) {
            saveWorkbookToFile("Added new sheet: " + sheetName);
            excelSheets.add(newS);
            return getExcelSheet(sheetName);
        }
        return null;
    }

    public void addSheets(String[] names) {
        ExcelWorkSheet ew = new ExcelWorkSheet();
        ew.wb = wb;
        ew.addSheets(names);
        saveWorkbookToFile("Added " + names.length + " new sheets");
        updateSheetList();
    }

    public ExcelWorkSheet getExcelSheet(String sheetName) {
        try {
            Sheet s = wb.getSheet(sheetName);
            return createWorkSheet(s);
        } catch (Exception e) {
            log.debug("Error getting sheet '{}': {}", sheetName, e.getMessage());
            return null;
        }
    }

    public ExcelWorkSheet getExcelSheet(int index) {
        try {
            Sheet s = wb.getSheetAt(index);
            return createWorkSheet(s);
        } catch (Exception e) {
            log.debug("Error getting sheet index {}: {}", index, e.getMessage());
            return null;
        }
    }

    public ExcelWorkSheet getActiveSheet() {
        try {
            return getExcelSheet(wb.getActiveSheetIndex());
        } catch (Exception e) {
            log.debug("Error getting active sheet: {}", e.getMessage());
            return null;
        }
    }

    /* ========== Misc. Getters ========== */

    public int getSheetCount() {
        return excelSheets.size();
    }

    public String getActiveSheetName() {
        return wb.getSheetName(wb.getActiveSheetIndex());
    }

    public int getActiveSheetIndex() {
        return wb.getActiveSheetIndex();
    }

    public boolean saveWorkbook() {
        return saveWorkbookToFile("Workbook saved successfully: " + excelBookName);
    }

    /* ========== Private Helpers ========== */

    private ExcelWorkSheet createWorkSheet(Sheet s) {
        if (s == null) return null;
        return new ExcelWorkSheet(s, s.getSheetName(), wb, inputStream, outputStream, getFullPath());
    }
}
