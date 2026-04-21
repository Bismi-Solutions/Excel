package solutions.bismi.excel;

/**
 * Reusable, immutable style definition. Build once, apply to many cells — avoids
 * the repetition of calling {@code setFontColor / setFillColor / setFullBorder}
 * on every cell and sidesteps Apache POI's {@code CellStyle} quota (64,000 per workbook)
 * by letting callers reuse the same style object.
 *
 * <pre>
 * ExcelStyle header = ExcelStyle.builder()
 *     .fillColor("grey_50_percent")
 *     .fontColor("white")
 *     .bold(true)
 *     .horizontalAlignment("CENTER")
 *     .fullBorder("black")
 *     .build();
 *
 * sheet.row(1).applyStyle(header);
 * </pre>
 */
public final class ExcelStyle {

    private final String fontColor;
    private final String fillColor;
    private final String borderColor;
    private final String horizontalAlignment;
    private final String verticalAlignment;
    private final String numberFormat;
    private final boolean bold;
    private final boolean italic;
    private final boolean underline;

    private ExcelStyle(Builder b) {
        this.fontColor = b.fontColor;
        this.fillColor = b.fillColor;
        this.borderColor = b.borderColor;
        this.horizontalAlignment = b.horizontalAlignment;
        this.verticalAlignment = b.verticalAlignment;
        this.numberFormat = b.numberFormat;
        this.bold = b.bold;
        this.italic = b.italic;
        this.underline = b.underline;
    }

    public static Builder builder() {
        return new Builder();
    }

    public String getFontColor() { return fontColor; }
    public String getFillColor() { return fillColor; }
    public String getBorderColor() { return borderColor; }
    public String getHorizontalAlignment() { return horizontalAlignment; }
    public String getVerticalAlignment() { return verticalAlignment; }
    public String getNumberFormat() { return numberFormat; }
    public boolean isBold() { return bold; }
    public boolean isItalic() { return italic; }
    public boolean isUnderline() { return underline; }

    /**
     * Applies this style to the given cell. Font-style flags (bold/italic/underline)
     * are always applied; other properties are applied only when set.
     */
    public void applyTo(ExcelCell cell) {
        if (cell == null) return;
        if (bold || italic || underline) {
            cell.setFontStyle(bold, italic, underline);
        }
        if (fontColor != null)            cell.setFontColor(fontColor);
        if (fillColor != null)            cell.setFillColor(fillColor);
        if (borderColor != null)          cell.setFullBorder(borderColor);
        if (horizontalAlignment != null)  cell.setHorizontalAlignment(horizontalAlignment);
        if (verticalAlignment != null)    cell.setVerticalAlignment(verticalAlignment);
        if (numberFormat != null)         cell.setNumberFormat(numberFormat);
    }

    public static final class Builder {
        private String fontColor;
        private String fillColor;
        private String borderColor;
        private String horizontalAlignment;
        private String verticalAlignment;
        private String numberFormat;
        private boolean bold;
        private boolean italic;
        private boolean underline;

        public Builder fontColor(String c)            { this.fontColor = c;            return this; }
        public Builder fillColor(String c)            { this.fillColor = c;            return this; }
        public Builder fullBorder(String c)           { this.borderColor = c;          return this; }
        public Builder horizontalAlignment(String a)  { this.horizontalAlignment = a;  return this; }
        public Builder verticalAlignment(String a)    { this.verticalAlignment = a;    return this; }
        public Builder numberFormat(String f)         { this.numberFormat = f;         return this; }
        public Builder bold(boolean v)                { this.bold = v;                 return this; }
        public Builder italic(boolean v)              { this.italic = v;               return this; }
        public Builder underline(boolean v)           { this.underline = v;            return this; }

        public ExcelStyle build() { return new ExcelStyle(this); }
    }

    /* ------------------------------------------------------------------ */
    /*  Ready-made presets                                                 */
    /* ------------------------------------------------------------------ */

    /** Dark header with white bold centered text and black border — typical report header. */
    public static ExcelStyle header() {
        return builder()
                .fillColor("grey_50_percent")
                .fontColor("white")
                .bold(true)
                .horizontalAlignment("CENTER")
                .fullBorder("black")
                .build();
    }

    /** Light-grey alternating row fill for zebra-striped tables. */
    public static ExcelStyle zebraStripe() {
        return builder().fillColor("grey_25_percent").build();
    }

    /** Currency preset — right-aligned, {@code $#,##0.00}. */
    public static ExcelStyle currency() {
        return builder().numberFormat("$#,##0.00").horizontalAlignment("RIGHT").build();
    }

    /** Percentage preset — right-aligned, {@code 0.00%}. */
    public static ExcelStyle percent() {
        return builder().numberFormat("0.00%").horizontalAlignment("RIGHT").build();
    }

    /** Plain date preset — {@code dd-MMM-yyyy}. */
    public static ExcelStyle date() {
        return builder().numberFormat("dd-MMM-yyyy").build();
    }
}
