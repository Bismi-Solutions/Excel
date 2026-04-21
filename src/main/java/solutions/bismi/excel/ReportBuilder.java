package solutions.bismi.excel;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/**
 * High-level, fluent helper for producing business-ready reports without touching
 * {@code CellStyle}, {@code XSSFWorkbook}, or indexed rows. Typical usage:
 *
 * <pre>
 * ReportBuilder.on(sheet)
 *     .title("Q3 Sales")
 *     .headers("Region", "Units", "Revenue")
 *     .rows(salesData)                 // List&lt;Object[]&gt; or List&lt;Map&gt; or List&lt;Bean&gt;
 *     .zebraStripes(true)
 *     .freezeHeader(true)
 *     .autoFilter(true)
 *     .autoSize(true)
 *     .headerStyle(ExcelStyle.header())
 *     .render();
 * </pre>
 *
 * <p>Everything above is a one-shot call — {@link #render()} writes the whole table
 * and saves the workbook.</p>
 */
public final class ReportBuilder {

    private final ExcelWorkSheet sheet;

    private String title;
    private ExcelStyle titleStyle;
    private String[] headers;
    private ExcelStyle headerStyle = ExcelStyle.header();
    private List<Object[]> rows = new ArrayList<>();
    private ExcelStyle rowStyle;
    private ExcelStyle stripeStyle = ExcelStyle.zebraStripe();
    private boolean zebra;
    private boolean freezeHeader;
    private boolean autoFilter;
    private boolean autoSize = true;
    private Map<Integer, ExcelStyle> columnStyles = new LinkedHashMap<>();
    private Map<Integer, Integer> columnWidths = new LinkedHashMap<>();

    private ReportBuilder(ExcelWorkSheet sheet) {
        this.sheet = sheet;
    }

    /** Creates a builder bound to the given sheet. */
    public static ReportBuilder on(ExcelWorkSheet sheet) {
        return new ReportBuilder(sheet);
    }

    /** Adds an optional merged title row above the headers. */
    public ReportBuilder title(String title) {
        this.title = title;
        return this;
    }

    /** Custom style for the title row (defaults to bold + centered). */
    public ReportBuilder titleStyle(ExcelStyle style) {
        this.titleStyle = style;
        return this;
    }

    /** Sets the column headers. */
    public ReportBuilder headers(String... headers) {
        this.headers = headers;
        return this;
    }

    /** Overrides the default header style. */
    public ReportBuilder headerStyle(ExcelStyle style) {
        this.headerStyle = style;
        return this;
    }

    /** Adds a single row (mixed types allowed). */
    public ReportBuilder row(Object... values) {
        this.rows.add(values);
        return this;
    }

    /** Adds rows from a list of arrays. */
    public ReportBuilder rows(List<Object[]> rows) {
        if (rows != null) this.rows.addAll(rows);
        return this;
    }

    /**
     * Adds rows from a list of maps — headers are inferred from the first map's keys
     * if not already set via {@link #headers(String...)}.
     */
    public ReportBuilder rowsFromMaps(List<? extends Map<String, ?>> data) {
        if (data == null || data.isEmpty()) return this;
        if (this.headers == null) {
            this.headers = data.get(0).keySet().toArray(new String[0]);
        }
        for (Map<String, ?> m : data) {
            Object[] values = new Object[headers.length];
            for (int i = 0; i < headers.length; i++) {
                values[i] = m.get(headers[i]);
            }
            this.rows.add(values);
        }
        return this;
    }

    /**
     * Adds rows from a list of beans annotated with {@link ExcelColumn}. If headers
     * haven't been set, they're derived from the annotations. Column formats from
     * the annotations are applied as {@link #columnStyle(int, ExcelStyle)}.
     */
    public <T> ReportBuilder rowsFromBeans(List<T> beans) {
        if (beans == null || beans.isEmpty()) return this;
        List<Field> fields = getExcelFields(beans.get(0).getClass());
        if (fields.isEmpty()) return this;
        if (this.headers == null) {
            String[] hs = new String[fields.size()];
            for (int i = 0; i < fields.size(); i++) {
                ExcelColumn ann = fields.get(i).getAnnotation(ExcelColumn.class);
                hs[i] = ann.name().isEmpty() ? fields.get(i).getName() : ann.name();
            }
            this.headers = hs;
        }
        for (int i = 0; i < fields.size(); i++) {
            String fmt = fields.get(i).getAnnotation(ExcelColumn.class).format();
            if (!fmt.isEmpty()) {
                columnStyle(i + 1, ExcelStyle.builder().numberFormat(fmt).build());
            }
        }
        for (T bean : beans) {
            Object[] values = new Object[fields.size()];
            for (int i = 0; i < fields.size(); i++) {
                try {
                    fields.get(i).setAccessible(true);
                    values[i] = fields.get(i).get(bean);
                } catch (Exception ignored) { /* leave null */ }
            }
            this.rows.add(values);
        }
        return this;
    }

    /** Applies a base style to every data row. */
    public ReportBuilder rowStyle(ExcelStyle style) {
        this.rowStyle = style;
        return this;
    }

    /** Overlays a different fill on every other row (readability). */
    public ReportBuilder zebraStripes(boolean enabled) {
        this.zebra = enabled;
        return this;
    }

    /** Overrides the default zebra-stripe style (defaults to grey_25_percent fill). */
    public ReportBuilder stripeStyle(ExcelStyle style) {
        this.stripeStyle = style;
        return this;
    }

    /**
     * Applies a style to a specific 1-based column in every data row. Useful for
     * number formatting (e.g. currency on column 3).
     */
    public ReportBuilder columnStyle(int col, ExcelStyle style) {
        columnStyles.put(col, style);
        return this;
    }

    /** Sets an explicit width (character units) for a 1-based column. */
    public ReportBuilder columnWidth(int col, int width) {
        columnWidths.put(col, width);
        return this;
    }

    /** Freezes the header row so it stays visible while scrolling. */
    public ReportBuilder freezeHeader(boolean enabled) {
        this.freezeHeader = enabled;
        return this;
    }

    /** Enables Excel's dropdown filter on the header row. */
    public ReportBuilder autoFilter(boolean enabled) {
        this.autoFilter = enabled;
        return this;
    }

    /** Auto-sizes all columns after render. Defaults to {@code true}. */
    public ReportBuilder autoSize(boolean enabled) {
        this.autoSize = enabled;
        return this;
    }

    /**
     * Writes everything configured to the sheet and saves the workbook.
     *
     * @return the bound {@link ExcelWorkSheet} for further chaining
     */
    public ExcelWorkSheet render() {
        if (headers == null || headers.length == 0) {
            throw new IllegalStateException("ReportBuilder: headers() is required before render()");
        }
        int cursor = 1;
        if (title != null && !title.isEmpty()) {
            sheet.cell(cursor, 1).setText(title);
            sheet.mergeCells(cursor, 1, cursor, headers.length);
            ExcelStyle effective = titleStyle != null
                    ? titleStyle
                    : ExcelStyle.builder().bold(true).horizontalAlignment("CENTER").build();
            sheet.row(cursor).applyStyle(effective, 1, headers.length + 1);
            cursor++;
        }

        int headerRow = cursor;
        sheet.row(headerRow).setRowValues(headers);
        sheet.row(headerRow).applyStyle(headerStyle, 1, headers.length + 1);
        cursor++;

        for (int i = 0; i < rows.size(); i++) {
            int r = cursor + i;
            Object[] values = rows.get(i);
            sheet.row(r).setValues(values);
            if (rowStyle != null) {
                sheet.row(r).applyStyle(rowStyle, 1, headers.length + 1);
            }
            if (zebra && i % 2 == 1) {
                sheet.row(r).applyStyle(stripeStyle, 1, headers.length + 1);
            }
            for (Map.Entry<Integer, ExcelStyle> entry : columnStyles.entrySet()) {
                sheet.cell(r, entry.getKey()).applyStyle(entry.getValue());
            }
        }

        for (Map.Entry<Integer, Integer> entry : columnWidths.entrySet()) {
            sheet.setColumnWidth(entry.getKey(), entry.getValue());
        }

        if (autoSize && columnWidths.isEmpty()) {
            sheet.autoSizeAllColumns();
        }
        if (freezeHeader) {
            sheet.freezePane(headerRow + 1, 1);
        }
        if (autoFilter) {
            sheet.enableAutoFilter(headerRow, 1, headers.length);
        }
        return sheet;
    }

    private static List<Field> getExcelFields(Class<?> clazz) {
        List<Field> result = new ArrayList<>();
        for (Class<?> c = clazz; c != null && c != Object.class; c = c.getSuperclass()) {
            Field[] declared = c.getDeclaredFields();
            Arrays.stream(declared)
                    .filter(f -> f.isAnnotationPresent(ExcelColumn.class))
                    .forEach(result::add);
        }
        result.sort((a, b) -> Integer.compare(
                a.getAnnotation(ExcelColumn.class).order(),
                b.getAnnotation(ExcelColumn.class).order()));
        return result;
    }
}
