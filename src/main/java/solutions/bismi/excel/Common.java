package solutions.bismi.excel;
import java.lang.reflect.Field;
import java.util.Arrays;
import java.util.HashMap;
import java.util.Map;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.IndexedColors;

/**
 * Common utility class for Excel operations.
 * Contains methods for color code handling, font property copying, and enum value retrieval.
 */
public abstract class Common {
    private static final Logger log = LogManager.getLogger(Common.class);

    private static IndexedColors[] values = null;
    private static Map<String, Short> map = new HashMap<>();

    /**
     * Helper method to copy all properties from one font to another.
     * 
     * @param source The source font to copy properties from
     * @param target The target font to copy properties to
     */
    public static void copyFontProperties(org.apache.poi.ss.usermodel.Font source, org.apache.poi.ss.usermodel.Font target) {
        if (source == null) {
            log.warn("Source font is null in copyFontProperties. Cannot copy properties.");
            return;
        }
        if (target == null) {
            log.warn("Target font is null in copyFontProperties. Cannot copy properties.");
            return;
        }
        target.setBold(source.getBold());
        target.setItalic(source.getItalic());
        target.setStrikeout(source.getStrikeout());
        target.setUnderline(source.getUnderline());
        target.setFontHeightInPoints(source.getFontHeightInPoints());
        target.setFontName(source.getFontName());

        // For XSSF specific fonts, we could copy more properties if needed
        // But these are the core properties available in both HSSF and XSSF
    }

    /**
     * Updates the indexed color values if not already initialized.
     * Populates a map of color names to their corresponding index values.
     */
    public static void updateIndexedValues() {
        if (values == null) {
            E[] localValues = getEnumValues(IndexedColors.class);
            if (localValues != null) { // Check if getEnumValues succeeded
                values = (IndexedColors[]) localValues; // Assign only if successful
                map.clear(); // Clear map before repopulating, in case of partial previous load or re-entry
                for (IndexedColors ele : values) {
                    if (ele != null) { // Guard against null elements if reflection somehow returns such an array
                        map.put(ele.toString(), ele.getIndex());
                    }
                }
            } else {
                log.error("Failed to retrieve enum values for IndexedColors. Color mapping will be incomplete.");
                // values remains null, map might be empty or hold previous (possibly stale) data.
                // Consider explicitly setting values to an empty array or map to empty to avoid repeated attempts if it's a permanent failure.
                // For now, it will retry on next call if values is still null.
            }
        }
    }
    /**
     * Retrieves the color code for a given color name.
     * 
     * @param color The name of the color to retrieve the code for
     * @return The color code as a short value, or 0 if the color is not found
     */
    public static short getColorCode(String color) {
        if (color == null || color.trim().isEmpty()) {
            log.warn("Color string is null or empty. Cannot get color code.");
            return 0; // Or a default color index like IndexedColors.AUTOMATIC.getIndex()
        }

        updateIndexedValues(); // Ensures 'values' and 'map' are populated if possible

        if (values == null) { // Check if values array is still null after updateIndexedValues attempt
            log.error("IndexedColors values array is null. Cannot perform color lookup for: {}", color);
            return 0; // Or a default color index
        }

        String availableColors = Arrays.toString(values); // Safe now due to null check above
        String processedColor = color.toUpperCase().trim();

        if (map.containsKey(processedColor)) {
            return map.get(processedColor);
        } else {
            log.warn("Color '{}' not found in standard IndexedColors. Supported string constant colors are: {}. Hex code can also be passed for XLSX.", color, availableColors);
            // Consider returning a specific default or error indicator if strict color matching is required.
            return 0; // Defaulting to black or an error-indicating index might be an option.
        }
    }

    /**
     * Retrieves the enum values for a given enum class using reflection.
     * 
     * @param <E> The enum type
     * @param enumClass The class of the enum to retrieve values for
     * @return An array of enum values, or null if an error occurs
     */
    private static <E extends Enum<E>> E[] getEnumValues(Class<E> enumClass) {
        if (enumClass == null) {
            log.error("Enum class provided to getEnumValues is null.");
            return null;
        }
        try {
            // Using enumClass.getEnumConstants() is the standard and safer way
            E[] enumConstants = enumClass.getEnumConstants();
            if (enumConstants == null) {
                // This can happen if enumClass is not an enum type, or has no constants.
                log.error("Failed to get enum constants for class: {}. It might not be an enum or has no values.", enumClass.getName());
            }
            return enumConstants;
        } catch (Exception e) { // Catching general exception as getEnumConstants() theoretically shouldn't throw checked ones here.
            log.error("Error in getting enum constants for {}: " + e.getMessage(), enumClass.getName(), e);
            return null;
        }
    }

    /**
     * Checks if a row is empty or contains only blank cells.
     *
     * @param row The row to check
     * @return true if the row is empty or null, false otherwise
     */
    public static boolean checkIfRowIsEmpty(org.apache.poi.ss.usermodel.Row row) {
        try {
            if (row == null || row.getLastCellNum() <= 0) {
                return true;
            }

            for (int cellNum = row.getFirstCellNum(); cellNum < row.getLastCellNum(); cellNum++) {
                org.apache.poi.ss.usermodel.Cell cell = row.getCell(cellNum);
                if (cell != null && cell.getCellType() != org.apache.poi.ss.usermodel.CellType.BLANK && 
                    org.apache.commons.lang3.StringUtils.isNotBlank(cell.toString())) {
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
    public static boolean checkIfCellIsEmpty(org.apache.poi.ss.usermodel.Cell cell) {
        return cell == null || cell.getCellType() == org.apache.poi.ss.usermodel.CellType.BLANK || 
               !org.apache.commons.lang3.StringUtils.isNotBlank(cell.toString());
    }

    /**
     * Saves the workbook to the specified file.
     *
     * @param wb The workbook to save
     * @param filePath The complete file path and name to save to
     * @return true if save was successful, false otherwise
     */
    public static boolean saveWorkBook(org.apache.poi.ss.usermodel.Workbook wb, String filePath) {
        if (wb == null) {
            log.error("Workbook object is null. Cannot save.");
            return false;
        }
        if (filePath == null || filePath.trim().isEmpty()) {
            log.error("File path is null or empty. Cannot save workbook.");
            return false;
        }
        try (java.io.OutputStream fileOut = new java.io.FileOutputStream(filePath)) {
            wb.write(fileOut);
            log.debug("Workbook saved successfully to: {}", filePath);
            return true;
        } catch (java.io.FileNotFoundException e) {
            log.error("Error saving workbook to '{}': File not found or path is a directory. " + e.getMessage(), filePath, e);
        } catch (java.io.IOException e) {
            log.error("IO error saving workbook to '{}': " + e.getMessage(), filePath, e);
        } catch (Exception e) { // Catch any other unexpected exceptions
            log.error("Unexpected error saving workbook to '{}': " + e.getMessage(), filePath, e);
        }
        return false;
    }
}
