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
 * Contains methods for color code handling and enum value retrieval.
 */
public abstract class Common {
    private static final Logger log = LogManager.getLogger(Common.class);

    private static IndexedColors[] values = null;
    private static Map<String, Short> map = new HashMap<>();

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

}
