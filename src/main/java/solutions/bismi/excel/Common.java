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
            values = getEnumValues(IndexedColors.class);
            for (IndexedColors ele : values) {
                map.put(ele.toString(), ele.getIndex());
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
        updateIndexedValues();

        String availableColors = Arrays.toString(values);
        if (map.containsKey(color.toUpperCase().trim())) {
            return map.get(color.toUpperCase().trim());
        } else {
            log.debug("Supported string constant colors are: {}", availableColors);
            log.debug("Hex code can also be passed if you need more colors");
            return 0;
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
        try {
            Field fld = enumClass.getDeclaredField("$VALUES");
            fld.setAccessible(true);
            Object o = fld.get(null);
            return (E[]) o;
        } catch (Exception e) {
            log.error("Error in getting color enumerators", e);
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
     */
    public static void saveWorkBook(org.apache.poi.ss.usermodel.Workbook wb, String filePath) {
        try (java.io.OutputStream fileOut = new java.io.FileOutputStream(filePath)) {
            wb.write(fileOut);
            log.debug("Workbook saved successfully to: {}", filePath);
        } catch (Exception e) {
            log.error("Error in saving workbook: {}", e.getMessage());
        }
    }
}
