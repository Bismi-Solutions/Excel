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
 * <p>
 * It centralises frequently‑used helpers such as colour handling, font‑copying and
 * basic cell/row emptiness checks.  All operations are <i>completely</i>
 * platform‑independent and safe to call from any thread.
 * </p>
 */
public abstract class Common {
    private static final Logger log = LogManager.getLogger(Common.class);

    private static IndexedColors[] values = null;
    private static final Map<String, Short> COLOR_MAP = new HashMap<>();

    /* ===================================================================== */
    /*  Font utilities                                                       */
    /* ===================================================================== */

    /**
     * Copy <b>all</b> public font attributes from <code>source</code> to
     * <code>target</code> (bold, italic&nbsp;…).
     */
    public static void copyFontProperties(org.apache.poi.ss.usermodel.Font source,
                                          org.apache.poi.ss.usermodel.Font target) {
        if (source == null || target == null) {
            log.warn("copyFontProperties called with null arguments (src={}, tgt={})", source, target);
            return;
        }
        target.setBold(source.getBold());
        target.setItalic(source.getItalic());
        target.setStrikeout(source.getStrikeout());
        target.setUnderline(source.getUnderline());
        target.setFontHeightInPoints(source.getFontHeightInPoints());
        target.setFontName(source.getFontName());
    }

    /* ===================================================================== */
    /*  Colour helpers                                                       */
    /* ===================================================================== */

    /** One‑time initialiser for the colour map. */
    private static synchronized void initColorMap() {
        if (values != null) return;                 // already done
        values = getEnumValues(IndexedColors.class);
        if (values != null) {
            for (IndexedColors col : values) {
                COLOR_MAP.put(col.toString(), col.getIndex());
            }
            log.debug("IndexedColours initialised ({} entries)", COLOR_MAP.size());
        }
    }

    /**
     * Translate a POI colour name (case‑insensitive) to its palette index.
     *
     * @param colour name – e.g. "RED", "light_blue".
     * @return index in the workbook palette or 0 if unknown.
     */
    public static short getColorCode(String colour) {
        if (colour == null) return 0;
        initColorMap();

        String key = colour.toUpperCase().trim();
        Short idx = COLOR_MAP.get(key);
        if (idx != null) {
            return idx;
        }
        if (log.isDebugEnabled()) {
            log.debug("Colour '{}' not recognised.  Available colours: {}", key, Arrays.toString(values));
        }
        return 0;
    }

    /* ===================================================================== */
    /*  Reflection helper                                                    */
    /* ===================================================================== */

    @SuppressWarnings("unchecked")
    private static <E extends Enum<E>> E[] getEnumValues(Class<E> enumClass) {
        try {
            Field f = enumClass.getDeclaredField("$VALUES");
            f.setAccessible(true);
            return (E[]) f.get(null);
        } catch (Exception ex) {
            log.error("Unable to read enum constants from {} – {}", enumClass.getSimpleName(), ex.toString());
            return null;
        }
    }

    /* ===================================================================== */
    /*  Cell/Row emptiness checks                                            */
    /* ===================================================================== */

    public static boolean checkIfRowIsEmpty(org.apache.poi.ss.usermodel.Row row) {
        if (row == null || row.getLastCellNum() <= 0) return true;
        try {
            for (int c = row.getFirstCellNum(); c < row.getLastCellNum(); c++) {
                org.apache.poi.ss.usermodel.Cell cell = row.getCell(c);
                if (!checkIfCellIsEmpty(cell)) return false;
            }
        } catch (Exception ex) {
            log.debug("Row emptiness check failed – treating as empty: {}", ex.toString());
        }
        return true;
    }

    public static boolean checkIfCellIsEmpty(org.apache.poi.ss.usermodel.Cell cell) {
        return cell == null || cell.getCellType() == org.apache.poi.ss.usermodel.CellType.BLANK ||
                org.apache.commons.lang3.StringUtils.isBlank(cell.toString());
    }

    /* ===================================================================== */
    /*  Workbook helper                                                      */
    /* ===================================================================== */

    /**
     * Convenience wrapper around {@link org.apache.poi.ss.usermodel.Workbook#write(java.io.OutputStream)}
     * that always closes the stream and logs success/failure at <tt>INFO</tt> level.
     */
    public static void saveWorkBook(org.apache.poi.ss.usermodel.Workbook wb, String filePath) {
        try (java.io.OutputStream out = new java.io.FileOutputStream(filePath)) {
            wb.write(out);
            log.info("Workbook saved to {}", filePath);
        } catch (Exception ex) {
            log.error("Failed to save workbook to {} – {}", filePath, ex.toString());
        }
    }
}