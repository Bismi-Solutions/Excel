package solutions.bismi.excel;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Marks a field for Excel serialization. Use with {@code writeBeans} / {@code readAsBeans}
 * on {@link ExcelWorkSheet}.
 *
 * <pre>
 * public class Product {
 *     &#064;ExcelColumn(name = "Item", order = 1)          String name;
 *     &#064;ExcelColumn(name = "Price", order = 2, format = "$#,##0.00") double price;
 *     &#064;ExcelColumn(name = "Launched", order = 3, format = "dd-MMM-yyyy") java.util.Date launchedOn;
 * }
 * </pre>
 *
 * <p>Unannotated fields are ignored. If {@code name} is empty, the field name is used.
 * If {@code order} is unset, fields appear in declaration order after ordered ones.</p>
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface ExcelColumn {

    /** Header label written to row 1. Defaults to the field name. */
    String name() default "";

    /** Column position (1-based). Lower values appear first. Defaults to declaration order. */
    int order() default Integer.MAX_VALUE;

    /**
     * Excel number/date format pattern applied to this column's data rows.
     * Examples: {@code "#,##0.00"}, {@code "0.00%"}, {@code "dd-MMM-yyyy"}, {@code "@"}.
     */
    String format() default "";
}
