package solutions.bismi.excel;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Comparator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

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

    /* ===================================================================== */
    /*  Layout conveniences                                                  */
    /* ===================================================================== */

    /**
     * Freezes rows above and columns to the left of the given cell so they stay
     * visible while scrolling. Pass {@code (2, 1)} to freeze the header row.
     *
     * @param row 1-based row — rows above (exclusive) stay frozen
     * @param col 1-based column — columns to the left (exclusive) stay frozen
     * @return this sheet for chaining
     */
    public ExcelWorkSheet freezePane(int row, int col) {
        try {
            sheet.createFreezePane(Math.max(0, col - 1), Math.max(0, row - 1));
            saveWorkBook();
        } catch (Exception e) {
            log.error("Error setting freeze pane: {}", e.getMessage());
        }
        return this;
    }

    /**
     * Enables Excel's dropdown filters on a header row.
     *
     * @param headerRow 1-based header row
     * @param firstCol  1-based first column of the filter range
     * @param lastCol   1-based last column of the filter range (inclusive)
     * @return this sheet for chaining
     */
    public ExcelWorkSheet enableAutoFilter(int headerRow, int firstCol, int lastCol) {
        try {
            sheet.setAutoFilter(new CellRangeAddress(headerRow - 1, headerRow - 1, firstCol - 1, lastCol - 1));
            saveWorkBook();
        } catch (Exception e) {
            log.error("Error enabling auto-filter: {}", e.getMessage());
        }
        return this;
    }

    /**
     * Sets the width of a column in characters (roughly the Excel units × 256).
     *
     * @param col   1-based column
     * @param width width in character units (e.g. 20)
     */
    public ExcelWorkSheet setColumnWidth(int col, int width) {
        try {
            sheet.setColumnWidth(col - 1, width * 256);
        } catch (Exception e) {
            log.error("Error setting column width: {}", e.getMessage());
        }
        return this;
    }

    /**
     * Auto-sizes every populated column to fit its widest cell content.
     * Safe to call after all data is written.
     */
    public ExcelWorkSheet autoSizeAllColumns() {
        try {
            int lastCol = getColNumber();
            for (int c = 0; c < lastCol; c++) {
                sheet.autoSizeColumn(c);
            }
        } catch (Exception e) {
            log.error("Error auto-sizing columns: {}", e.getMessage());
        }
        return this;
    }

    /**
     * Protects the sheet with a password. Password may be null to enable protection
     * without one.
     */
    public ExcelWorkSheet protect(String password) {
        try {
            sheet.protectSheet(password);
            saveWorkBook();
        } catch (Exception e) {
            log.error("Error protecting sheet: {}", e.getMessage());
        }
        return this;
    }

    /* ===================================================================== */
    /*  Collection framework — write                                         */
    /* ===================================================================== */

    /**
     * Writes a list of maps as a table. Keys of the first map become the headers
     * (in iteration order — use {@link LinkedHashMap} to preserve column order).
     * Subsequent rows pull values by key.
     *
     * <pre>
     * List&lt;Map&lt;String,Object&gt;&gt; rows = List.of(
     *     Map.of("Item","Apple", "Qty",10, "Price",1.20),
     *     Map.of("Item","Pear",  "Qty",8,  "Price",1.80));
     * sheet.writeMaps(rows);
     * </pre>
     *
     * @param rows the data rows (null/empty is a no-op)
     * @return this sheet for chaining
     */
    public ExcelWorkSheet writeMaps(List<? extends Map<String, ?>> rows) {
        if (rows == null || rows.isEmpty()) return this;
        List<String> headers = new ArrayList<>(rows.get(0).keySet());
        String[] headerArr = headers.toArray(new String[0]);
        row(1).setRowValues(headerArr);
        for (int i = 0; i < rows.size(); i++) {
            Object[] values = new Object[headers.size()];
            Map<String, ?> src = rows.get(i);
            for (int c = 0; c < headers.size(); c++) {
                values[c] = src.get(headers.get(c));
            }
            row(i + 2).setValues(values);
        }
        saveWorkBook();
        return this;
    }

    /**
     * Writes a list of beans annotated with {@link ExcelColumn}. Only annotated fields
     * are exported; ordering comes from {@code order()}, then declaration order. Field
     * {@code format()} is applied to the column's data rows.
     *
     * @param beans list of beans (null/empty is a no-op)
     * @param <T>   bean type
     * @return this sheet for chaining
     */
    public <T> ExcelWorkSheet writeBeans(List<T> beans) {
        if (beans == null || beans.isEmpty()) return this;
        Class<?> clazz = beans.get(0).getClass();
        List<Field> fields = getExcelFields(clazz);
        if (fields.isEmpty()) {
            log.warn("Class {} has no @ExcelColumn fields; nothing written", clazz.getSimpleName());
            return this;
        }
        String[] headers = new String[fields.size()];
        for (int i = 0; i < fields.size(); i++) {
            ExcelColumn ann = fields.get(i).getAnnotation(ExcelColumn.class);
            headers[i] = ann.name().isEmpty() ? fields.get(i).getName() : ann.name();
        }
        row(1).setRowValues(headers);
        for (int r = 0; r < beans.size(); r++) {
            Object[] values = new Object[fields.size()];
            T bean = beans.get(r);
            for (int c = 0; c < fields.size(); c++) {
                try {
                    fields.get(c).setAccessible(true);
                    values[c] = fields.get(c).get(bean);
                } catch (Exception e) {
                    log.error("Error reading field {}: {}", fields.get(c).getName(), e.getMessage());
                }
            }
            row(r + 2).setValues(values);
        }
        for (int c = 0; c < fields.size(); c++) {
            String fmt = fields.get(c).getAnnotation(ExcelColumn.class).format();
            if (!fmt.isEmpty()) {
                for (int r = 2; r <= beans.size() + 1; r++) {
                    cell(r, c + 1).setNumberFormat(fmt);
                }
            }
        }
        saveWorkBook();
        return this;
    }

    /* ===================================================================== */
    /*  Collection framework — read                                          */
    /* ===================================================================== */

    /**
     * Reads the sheet into a list of maps, using row 1 as headers.
     *
     * @return a list of {@link LinkedHashMap}s — empty if sheet is empty
     */
    public List<Map<String, String>> readAsMaps() {
        List<Map<String, String>> result = new ArrayList<>();
        try {
            int lastRow = sheet.getLastRowNum();
            Row header = sheet.getRow(0);
            if (header == null) return result;
            int lastCol = header.getLastCellNum();
            String[] headers = new String[lastCol];
            for (int c = 0; c < lastCol; c++) {
                Cell cell = header.getCell(c);
                headers[c] = cell == null ? "" : cell.toString();
            }
            for (int r = 1; r <= lastRow; r++) {
                Row row = sheet.getRow(r);
                if (row == null) continue;
                Map<String, String> map = new LinkedHashMap<>();
                for (int c = 0; c < lastCol; c++) {
                    Cell cell = row.getCell(c);
                    map.put(headers[c], stringifyCell(cell));
                }
                result.add(map);
            }
        } catch (Exception e) {
            log.error("Error reading sheet as maps: {}", e.getMessage());
        }
        return result;
    }

    /**
     * Reads the sheet into a list of beans using {@link ExcelColumn} annotations to map
     * columns to fields. Requires a public no-arg constructor on {@code T}.
     *
     * @param clazz bean class
     * @param <T>   bean type
     * @return list of populated beans
     */
    public <T> List<T> readAsBeans(Class<T> clazz) {
        List<T> result = new ArrayList<>();
        List<Field> fields = getExcelFields(clazz);
        if (fields.isEmpty()) return result;
        Map<String, Field> byHeader = new LinkedHashMap<>();
        for (Field f : fields) {
            ExcelColumn ann = f.getAnnotation(ExcelColumn.class);
            byHeader.put(ann.name().isEmpty() ? f.getName() : ann.name(), f);
        }
        List<Map<String, String>> rows = readAsMaps();
        for (Map<String, String> rowMap : rows) {
            T bean;
            try {
                bean = clazz.getDeclaredConstructor().newInstance();
            } catch (Exception e) {
                log.error("Cannot instantiate {} (needs a no-arg constructor): {}", clazz.getName(), e.getMessage());
                return result;
            }
            for (Map.Entry<String, Field> entry : byHeader.entrySet()) {
                String raw = rowMap.get(entry.getKey());
                if (raw == null) continue;
                Field f = entry.getValue();
                try {
                    f.setAccessible(true);
                    Object coerced = coerce(raw, f.getType());
                    if (coerced != null) {
                        f.set(bean, coerced);
                    }
                } catch (Exception e) {
                    log.error("Cannot set field {} from value '{}': {}", f.getName(), raw, e.getMessage());
                }
            }
            result.add(bean);
        }
        return result;
    }

    /* ===================================================================== */
    /*  Internal helpers                                                     */
    /* ===================================================================== */

    private static List<Field> getExcelFields(Class<?> clazz) {
        List<Field> result = new ArrayList<>();
        for (Class<?> c = clazz; c != null && c != Object.class; c = c.getSuperclass()) {
            for (Field f : c.getDeclaredFields()) {
                if (f.isAnnotationPresent(ExcelColumn.class)) {
                    result.add(f);
                }
            }
        }
        result.sort(Comparator.comparingInt(f -> f.getAnnotation(ExcelColumn.class).order()));
        return result;
    }

    private static String stringifyCell(Cell cell) {
        if (cell == null) return "";
        switch (cell.getCellType()) {
            case STRING:  return cell.getStringCellValue();
            case BOOLEAN: return String.valueOf(cell.getBooleanCellValue());
            case FORMULA: return cell.getCellFormula();
            case BLANK:   return "";
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    // ISO-8601 so readAsBeans can parse back into LocalDate/LocalDateTime/Date
                    java.util.Date d = cell.getDateCellValue();
                    java.time.LocalDateTime ldt = d.toInstant()
                            .atZone(java.time.ZoneId.systemDefault()).toLocalDateTime();
                    if (ldt.getHour() == 0 && ldt.getMinute() == 0 && ldt.getSecond() == 0) {
                        return ldt.toLocalDate().toString();   // yyyy-MM-dd
                    }
                    return ldt.toString();                       // yyyy-MM-ddTHH:mm:ss
                }
                double n = cell.getNumericCellValue();
                if (n == Math.floor(n) && !Double.isInfinite(n)) {
                    return String.valueOf((long) n);
                }
                return String.valueOf(n);
            default: return cell.toString();
        }
    }

    @SuppressWarnings({"unchecked", "rawtypes"})
    private static Object coerce(String raw, Class<?> target) {
        try {
            if (target == String.class)                            return raw;
            if (raw.isEmpty())                                     return null;
            String t = raw.trim();
            if (target == int.class     || target == Integer.class) return Integer.parseInt(t);
            if (target == long.class    || target == Long.class)    return Long.parseLong(t);
            if (target == double.class  || target == Double.class)  return Double.parseDouble(t);
            if (target == float.class   || target == Float.class)   return Float.parseFloat(t);
            if (target == boolean.class || target == Boolean.class) return Boolean.parseBoolean(t);
            if (target == java.time.LocalDate.class)                return java.time.LocalDate.parse(t);
            if (target == java.time.LocalDateTime.class)            return java.time.LocalDateTime.parse(t);
            if (target == java.util.Date.class) {
                java.time.LocalDateTime ldt = t.contains("T")
                        ? java.time.LocalDateTime.parse(t)
                        : java.time.LocalDate.parse(t).atStartOfDay();
                return java.util.Date.from(ldt.atZone(java.time.ZoneId.systemDefault()).toInstant());
            }
            if (target.isEnum())                                   return Enum.valueOf((Class<Enum>) target, t);
        } catch (Exception ignored) {
            // fall through — return raw String
        }
        return raw;
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
