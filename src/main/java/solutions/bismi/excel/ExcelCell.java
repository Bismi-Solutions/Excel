package solutions.bismi.excel;

import lombok.AccessLevel;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;
import lombok.AllArgsConstructor;
import lombok.Builder;
import org.apache.commons.lang3.StringUtils;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.HashMap;
import java.util.Map;

/**
 * Represents a cell in an Excel worksheet.
 * Provides methods for manipulating cell content, formatting, and styling.
 */
@NoArgsConstructor(access = AccessLevel.PRIVATE)
@AllArgsConstructor(access = AccessLevel.PRIVATE)
@Builder
public class ExcelCell {

    @Getter(AccessLevel.PRIVATE) @Setter(AccessLevel.PRIVATE)
    Sheet sheet;

    @Getter(AccessLevel.PRIVATE) @Setter(AccessLevel.PRIVATE)
    Workbook wb;

    @Getter(AccessLevel.PRIVATE) @Setter(AccessLevel.PRIVATE)
    int row;

    @Getter(AccessLevel.PRIVATE) @Setter(AccessLevel.PRIVATE)
    int col;

    @Getter(AccessLevel.PRIVATE) @Setter(AccessLevel.PRIVATE)
    FileInputStream inputStream = null;

    @Getter(AccessLevel.PRIVATE) @Setter(AccessLevel.PRIVATE)
    FileOutputStream outputStream = null;

    @Getter(AccessLevel.PRIVATE) @Setter(AccessLevel.PRIVATE)
    String sCompleteFileName = null;

    @Getter(AccessLevel.PRIVATE) @Setter(AccessLevel.PRIVATE)
    Cell cell = null;

    @Getter(AccessLevel.PRIVATE) @Setter(AccessLevel.PRIVATE)
    String xlFileExtension = null;

    private final Logger log = LogManager.getLogger(ExcelCell.class);

    /**
     * Creates a new ExcelCell with the specified parameters.
     * This constructor initializes the cell and creates it if it doesn't exist.
     *
     * @param sheet             The worksheet containing the cell
     * @param wb                The workbook containing the sheet
     * @param row               The row number (1-based)
     * @param col               The column number (1-based)
     * @param inputStream       The input stream for the workbook file
     * @param outputStream      The output stream for the workbook file
     * @param sCompleteFileName The complete file path and name of the workbook
     */
    protected ExcelCell(Sheet sheet, Workbook wb, int row, int col, FileInputStream inputStream, FileOutputStream outputStream, String sCompleteFileName) {
        this.sheet = sheet;
        this.wb = wb;
        this.row = row;
        this.col = col;
        this.inputStream = inputStream;
        this.outputStream = outputStream;
        this.sCompleteFileName = sCompleteFileName;
        this.xlFileExtension = getFileExtension(sCompleteFileName);

        try {
            Row cRow = getOrCreateRow(sheet, this.row - 1);
            Cell cell1 = getOrCreateCell(cRow, this.col - 1);
            this.setCell(cell1);
        } catch (Exception e) {
            log.error("Error in cell constructor creation: {}", e.getMessage());
        }
    }

    private Row getOrCreateRow(Sheet sheet, int rowIndex) {
        Row cRow = null;
        try {
            cRow = sheet.getRow(rowIndex);
        } catch (Exception e) {
            // Row doesn't exist
        }
        if (checkIfRowIsEmpty(cRow)) {
            cRow = sheet.createRow(rowIndex);
        }
        return cRow;
    }

    private Cell getOrCreateCell(Row row, int colIndex) {
        Cell localCell = null;
        try {
            localCell = row.getCell(colIndex);
        } catch (Exception e) {
            // Cell doesn't exist
        }
        if (checkIfCellIsEmpty(localCell)) {
            localCell = row.createCell(colIndex);
        }
        return localCell;
    }

    public ExcelCell setText(String sText) {
        return setText(sText, false);
    }

    public ExcelCell setCellValue(String sText) {
        return setCellValue(sText, false);
    }

    /**
     * Sets a plain text value in the cell.
     *
     * @param sText           The text value to set
     * @param autoSizeColoumn Whether to auto-size the column
     * @return this cell for chaining
     */
    public ExcelCell setCellValue(String sText, boolean autoSizeColoumn) {
        try {
            this.getCell().setCellValue(sText);
            if (autoSizeColoumn) {
                this.sheet.autoSizeColumn(col - 1);
            }
        } catch (Exception e) {
            log.error("Error in setting cell value: {}", e.getMessage());
        }
        return this;
    }

    /**
     * Sets a numeric value in the cell.
     *
     * @param value          The numeric value to set
     * @param autoSizeColumn Whether to auto-size the column
     * @return this cell for chaining
     */
    public ExcelCell setNumericValue(double value, boolean autoSizeColumn) {
        try {
            this.getCell().setCellValue(value);
            if (autoSizeColumn) {
                this.sheet.autoSizeColumn(col - 1);
            }
        } catch (Exception e) {
            log.error("Error in setting numeric value: {}", e.getMessage());
        }
        return this;
    }

    public ExcelCell setNumericValue(double value) {
        return setNumericValue(value, false);
    }

    /**
     * Sets a date value in the cell.
     *
     * @param date           The date value to set
     * @param autoSizeColumn Whether to auto-size the column
     * @return this cell for chaining
     */
    public ExcelCell setDateValue(java.util.Date date, boolean autoSizeColumn) {
        try {
            this.getCell().setCellValue(date);
            if (autoSizeColumn) {
                this.sheet.autoSizeColumn(col - 1);
            }
        } catch (Exception e) {
            log.error("Error in setting date value: {}", e.getMessage());
        }
        return this;
    }

    public ExcelCell setDateValue(java.util.Date date) {
        return setDateValue(date, false);
    }

    /**
     * Sets a formula in the cell. Leading {@code =} is stripped if present.
     *
     * @param formula        The formula to set
     * @param autoSizeColumn Whether to auto-size the column
     * @return this cell for chaining
     */
    public ExcelCell setFormula(String formula, boolean autoSizeColumn) {
        try {
            if (formula != null && formula.startsWith("=")) {
                formula = formula.substring(1);
            }
            this.getCell().setCellFormula(formula);
            if (autoSizeColumn) {
                this.sheet.autoSizeColumn(col - 1);
            }
        } catch (Exception e) {
            log.error("Error in setting formula: {}", e.getMessage());
        }
        return this;
    }

    public ExcelCell setFormula(String formula) {
        return setFormula(formula, false);
    }

    /**
     * Sets a text value in the cell with explicit "@" text formatting.
     *
     * @param sText           The text value to set
     * @param autoSizeColoumn Whether to auto-size the column
     * @return this cell for chaining
     */
    public ExcelCell setText(String sText, boolean autoSizeColoumn) {
        try {
            Cell cCell = this.getCell();

            Workbook workbook = cCell.getSheet().getWorkbook();
            CellStyle originalStyle = cCell.getCellStyle();
            CellStyle newStyle = workbook.createCellStyle();
            newStyle.cloneStyleFrom(originalStyle);

            DataFormat format = workbook.createDataFormat();
            newStyle.setDataFormat(format.getFormat("@"));

            if (sText != null && !sText.isEmpty()) {
                RichTextString richText = workbook.getCreationHelper().createRichTextString(sText);
                cCell.setCellValue(richText);
            } else {
                cCell.setCellValue("");
            }

            cCell.setCellStyle(newStyle);

            if (autoSizeColoumn) {
                this.sheet.autoSizeColumn(col - 1);
            }
        } catch (Exception e) {
            log.error("Error in setting text value: {}", e.getMessage());
        }
        return this;
    }

    private void setHexColor(String hexColor, Font newFont) {
        try {
            byte[] rgb = new byte[] {
                (byte) Integer.parseInt(hexColor.substring(0, 2), 16),
                (byte) Integer.parseInt(hexColor.substring(2, 4), 16),
                (byte) Integer.parseInt(hexColor.substring(4, 6), 16)
            };
            XSSFFont xssfFont = (XSSFFont) newFont;
            xssfFont.setColor(new XSSFColor(rgb, null));
        } catch (Exception e) {
            log.error("Error parsing hex color: {}", e.getMessage());
            newFont.setColor(IndexedColors.BLACK.getIndex());
        }
    }

    private void setHSSFHexColor(String hexColor, Font newFont) {
        try {
            int r = Integer.parseInt(hexColor.substring(0, 2), 16);
            int g = Integer.parseInt(hexColor.substring(2, 4), 16);
            int b = Integer.parseInt(hexColor.substring(4, 6), 16);
            short closestIndex = getClosestIndexedColor(r, g, b);
            newFont.setColor(closestIndex);
        } catch (Exception e) {
            log.error("Error parsing hex color for HSSF: {}", e.getMessage());
            newFont.setColor(org.apache.poi.ss.usermodel.IndexedColors.BLACK.getIndex());
        }
    }

    /**
     * Sets the font color for the cell while preserving existing formatting.
     * <p>
     * Supported values:
     * <ul>
     *   <li>Named colors (see {@link NamedColor}) for both XLS and XLSX</li>
     *   <li>Hex color codes (e.g., "#FF0000") for XLSX only</li>
     * </ul>
     * <b>Note:</b> Hex color codes are ignored for XLS files due to format limitations.<br>
     * <b>Named colors supported:</b> See {@link NamedColor}
     * </p>
     * @param fontColor Named color (see NamedColor enum) or hex color (XLSX only)
     * @return this ExcelCell instance for method chaining
     */
    public ExcelCell setFontColor(String fontColor) {
        try {
            Cell cCell = this.getCell();
            Workbook workbook = cCell.getSheet().getWorkbook();
            CellStyle originalStyle = cCell.getCellStyle();
            CellStyle newStyle = workbook.createCellStyle();
            newStyle.cloneStyleFrom(originalStyle);
            Font existingFont = workbook.getFontAt(originalStyle.getFontIndex());
            Font newFont = workbook.createFont();
            Common.copyFontProperties(existingFont, newFont);

            if (fontColor != null && fontColor.startsWith("#")) {
                String hexColor = fontColor.substring(1);
                if (workbook instanceof XSSFWorkbook) {
                    // XLSX: true hex color
                    setHexColor(hexColor, newFont);
                } else {
                    // HSSF: fallback to closest IndexedColors
                    setHSSFHexColor(hexColor, newFont);
                }
            } else {
                // Use the existing color mapping for named colors
                newFont.setColor(Common.getColorCode(fontColor));
            }

            newStyle.setFont(newFont);
            cCell.setCellStyle(newStyle);

            return this;
        } catch (Exception e) {
            log.error("Error in setting font color: {}", e.getMessage());
            return this;
        }
    }

    // Helper: Find closest IndexedColors for HSSF hex fallback
    private short getClosestIndexedColor(int r, int g, int b) {
        short closestIndex = org.apache.poi.ss.usermodel.IndexedColors.BLACK.getIndex();
        double minDistance = Double.MAX_VALUE;
        for (org.apache.poi.ss.usermodel.IndexedColors color : org.apache.poi.ss.usermodel.IndexedColors.values()) {
            short idx = color.getIndex();
            int[] rgb = getIndexedColorRGB(idx);
            double distance = Math.pow((double)r - rgb[0], 2) + Math.pow((double)g - rgb[1], 2) + Math.pow((double)b - rgb[2], 2);
            if (distance < minDistance) {
                minDistance = distance;
                closestIndex = idx;
            }
        }
        return closestIndex;
    }

    // Helper: Get RGB for IndexedColors (HSSF)
    private int[] getIndexedColorRGB(short idx) {
        try {
            java.util.Map<Integer, org.apache.poi.hssf.util.HSSFColor> colorMap = org.apache.poi.hssf.util.HSSFColor.getIndexHash();
            org.apache.poi.hssf.util.HSSFColor hssfColor = colorMap.get((int) idx);
            if (hssfColor != null) {
                short[] triplet = hssfColor.getTriplet();
                return new int[] { triplet[0], triplet[1], triplet[2] };
            }
        } catch (Exception e) {
            // fallback
        }
        return new int[] {0, 0, 0};
    }

    /**
     * Sets the background fill color for the cell.
     * <p>
     * Accepts named colours (see {@link NamedColor}) or hex codes ({@code "#RRGGBB"}).
     * Hex is rendered exactly on {@code .xlsx} workbooks; on {@code .xls} it falls back
     * to the closest indexed colour.
     *
     * @param fillColor named color or hex code
     * @return this cell for chaining
     */
    public ExcelCell setFillColor(String fillColor) {
        try {
            Cell cCell = this.getCell();

            Workbook workbook = cCell.getSheet().getWorkbook();
            CellStyle originalStyle = cCell.getCellStyle();
            CellStyle newStyle = workbook.createCellStyle();
            newStyle.cloneStyleFrom(originalStyle);

            newStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            if (fillColor != null && fillColor.startsWith("#") && fillColor.length() >= 7) {
                String hex = fillColor.substring(1);
                if (workbook instanceof XSSFWorkbook) {
                    setXSSFHexFillColor(hex, newStyle);
                } else {
                    setHSSFHexFillColor(hex, newStyle);
                }
            } else {
                newStyle.setFillForegroundColor(Common.getColorCode(fillColor));
            }

            cCell.setCellStyle(newStyle);
        } catch (Exception e) {
            log.error("Error in setting fill color: {}", e.getMessage());
        }
        return this;
    }

    private void setXSSFHexFillColor(String hex, CellStyle style) {
        try {
            byte[] rgb = new byte[] {
                (byte) Integer.parseInt(hex.substring(0, 2), 16),
                (byte) Integer.parseInt(hex.substring(2, 4), 16),
                (byte) Integer.parseInt(hex.substring(4, 6), 16)
            };
            XSSFCellStyle xs = (XSSFCellStyle) style;
            xs.setFillForegroundColor(new XSSFColor(rgb, null));
        } catch (Exception e) {
            log.error("Error parsing hex fill color: {}", e.getMessage());
            style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        }
    }

    private void setHSSFHexFillColor(String hex, CellStyle style) {
        try {
            int r = Integer.parseInt(hex.substring(0, 2), 16);
            int g = Integer.parseInt(hex.substring(2, 4), 16);
            int b = Integer.parseInt(hex.substring(4, 6), 16);
            style.setFillForegroundColor(getClosestIndexedColor(r, g, b));
        } catch (Exception e) {
            log.error("Error parsing hex fill color for HSSF: {}", e.getMessage());
            style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        }
    }

    /**
     * Sets the horizontal alignment of the cell content.
     *
     * @param alignment LEFT / CENTER / RIGHT / JUSTIFY / FILL / GENERAL
     * @return this cell for chaining
     */
    public ExcelCell setHorizontalAlignment(String alignment) {
        try {
            Cell cCell = this.getCell();
            Workbook workbook = cCell.getSheet().getWorkbook();
            CellStyle originalStyle = cCell.getCellStyle();
            CellStyle newStyle = workbook.createCellStyle();
            newStyle.cloneStyleFrom(originalStyle);

            HorizontalAlignment hAlign;
            switch (alignment.toUpperCase()) {
                case "LEFT":    hAlign = HorizontalAlignment.LEFT;    break;
                case "CENTER":  hAlign = HorizontalAlignment.CENTER;  break;
                case "RIGHT":   hAlign = HorizontalAlignment.RIGHT;   break;
                case "JUSTIFY": hAlign = HorizontalAlignment.JUSTIFY; break;
                case "FILL":    hAlign = HorizontalAlignment.FILL;    break;
                default:        hAlign = HorizontalAlignment.GENERAL;
            }
            newStyle.setAlignment(hAlign);
            cCell.setCellStyle(newStyle);
        } catch (Exception e) {
            log.error("Error in setting horizontal alignment: {}", e.getMessage());
        }
        return this;
    }

    /**
     * Sets the vertical alignment of the cell content.
     *
     * @param alignment TOP / CENTER / BOTTOM / JUSTIFY / DISTRIBUTED
     * @return this cell for chaining
     */
    public ExcelCell setVerticalAlignment(String alignment) {
        try {
            Cell cCell = this.getCell();
            Workbook workbook = cCell.getSheet().getWorkbook();
            CellStyle originalStyle = cCell.getCellStyle();
            CellStyle newStyle = workbook.createCellStyle();
            newStyle.cloneStyleFrom(originalStyle);

            VerticalAlignment vAlign;
            switch (alignment.toUpperCase()) {
                case "TOP":         vAlign = VerticalAlignment.TOP;         break;
                case "CENTER":      vAlign = VerticalAlignment.CENTER;      break;
                case "BOTTOM":      vAlign = VerticalAlignment.BOTTOM;      break;
                case "JUSTIFY":     vAlign = VerticalAlignment.JUSTIFY;     break;
                case "DISTRIBUTED": vAlign = VerticalAlignment.DISTRIBUTED; break;
                default:            vAlign = VerticalAlignment.CENTER;
            }
            newStyle.setVerticalAlignment(vAlign);
            cCell.setCellStyle(newStyle);
        } catch (Exception e) {
            log.error("Error in setting vertical alignment: {}", e.getMessage());
        }
        return this;
    }

    /**
     * Sets the number format for the cell.
     *
     * @param formatPattern e.g. "#,##0.00", "dd/mm/yyyy", "$#,##0.00"
     * @return this cell for chaining
     */
    public ExcelCell setNumberFormat(String formatPattern) {
        try {
            Cell cCell = this.getCell();
            Workbook workbook = cCell.getSheet().getWorkbook();
            CellStyle originalStyle = cCell.getCellStyle();
            CellStyle newStyle = workbook.createCellStyle();
            newStyle.cloneStyleFrom(originalStyle);

            DataFormat format = workbook.createDataFormat();
            newStyle.setDataFormat(format.getFormat(formatPattern));

            cCell.setCellStyle(newStyle);
        } catch (Exception e) {
            log.error("Error in setting number format: {}", e.getMessage());
        }
        return this;
    }

    /**
     * Sets the font style (bold, italic, underline).
     *
     * @return this cell for chaining
     */
    public ExcelCell setFontStyle(boolean bold, boolean italic, boolean underline) {
        try {
            Cell cCell = this.getCell();
            Font font = wb.createFont();
            font.setBold(bold);
            font.setItalic(italic);
            if (underline) {
                font.setUnderline(org.apache.poi.ss.usermodel.FontUnderline.SINGLE.getByteValue());
            }
            CellUtil.setFont(cCell, font);
        } catch (Exception e) {
            log.error("Error in setting font style: {}", e.getMessage());
        }
        return this;
    }

    /**
     * Adds a medium-weight border (all four sides) around the cell. Accepts a
     * {@linkplain NamedColor named colour} or a hex code ({@code "#RRGGBB"}).
     * Hex is rendered exactly on {@code .xlsx}; on {@code .xls} (HSSF) it falls
     * back to the closest indexed colour — mirroring how {@link #setFillColor}
     * and {@link #setFontColor} behave, so a single hex palette works for both.
     *
     * @param borderColor named or hex colour
     * @return this cell for chaining
     */
    public ExcelCell setFullBorder(String borderColor) {
        try {
            Cell cCell = this.getCell();
            Workbook workbook = cCell.getSheet().getWorkbook();
            CellStyle originalStyle = cCell.getCellStyle();
            CellStyle newStyle = workbook.createCellStyle();
            newStyle.cloneStyleFrom(originalStyle);

            newStyle.setBorderLeft(BorderStyle.MEDIUM);
            newStyle.setBorderRight(BorderStyle.MEDIUM);
            newStyle.setBorderTop(BorderStyle.MEDIUM);
            newStyle.setBorderBottom(BorderStyle.MEDIUM);

            boolean isHex = borderColor != null
                    && borderColor.startsWith("#")
                    && borderColor.length() >= 7;

            if (isHex && workbook instanceof XSSFWorkbook) {
                // XLSX: apply the exact hex on all four sides via XSSFColor.
                XSSFColor color = hexToXSSFColor(borderColor.substring(1));
                XSSFCellStyle xs = (XSSFCellStyle) newStyle;
                xs.setLeftBorderColor(color);
                xs.setRightBorderColor(color);
                xs.setTopBorderColor(color);
                xs.setBottomBorderColor(color);
            } else {
                // HSSF (hex → nearest indexed) or a named colour.
                short colorCode = isHex
                        ? closestIndexedForHex(borderColor.substring(1))
                        : Common.getColorCode(borderColor);
                newStyle.setLeftBorderColor(colorCode);
                newStyle.setRightBorderColor(colorCode);
                newStyle.setTopBorderColor(colorCode);
                newStyle.setBottomBorderColor(colorCode);
            }

            cCell.setCellStyle(newStyle);
        } catch (Exception e) {
            log.error("Error in setting border: {}", e.getMessage());
        }
        return this;
    }

    private XSSFColor hexToXSSFColor(String hex) {
        byte[] rgb = new byte[] {
                (byte) Integer.parseInt(hex.substring(0, 2), 16),
                (byte) Integer.parseInt(hex.substring(2, 4), 16),
                (byte) Integer.parseInt(hex.substring(4, 6), 16)
        };
        return new XSSFColor(rgb, null);
    }

    private short closestIndexedForHex(String hex) {
        int r = Integer.parseInt(hex.substring(0, 2), 16);
        int g = Integer.parseInt(hex.substring(2, 4), 16);
        int b = Integer.parseInt(hex.substring(4, 6), 16);
        return getClosestIndexedColor(r, g, b);
    }


    /**
     * Gets the value of the cell as a string.
     *
     * @return The cell value as a string
     */
    public String getValue() {
        return getTextValue();
    }

    /**
     * Polymorphic value setter — dispatches to the correct typed setter based on runtime type.
     * Accepts {@link Number}, {@link Boolean}, {@link java.util.Date}, {@link java.time.LocalDate},
     * {@link java.time.LocalDateTime}, {@link CharSequence}, or {@code null}. Anything else
     * is written via {@code toString()}.
     *
     * @param value the value to write
     * @return this cell for chaining
     */
    public ExcelCell setValue(Object value) {
        try {
            Cell c = this.getCell();
            if (value == null) {
                c.setBlank();
            } else if (value instanceof Number) {
                c.setCellValue(((Number) value).doubleValue());
            } else if (value instanceof Boolean) {
                c.setCellValue((Boolean) value);
            } else if (value instanceof java.util.Date) {
                c.setCellValue((java.util.Date) value);
            } else if (value instanceof java.time.LocalDate) {
                c.setCellValue((java.time.LocalDate) value);
            } else if (value instanceof java.time.LocalDateTime) {
                c.setCellValue((java.time.LocalDateTime) value);
            } else if (value instanceof CharSequence) {
                c.setCellValue(value.toString());
            } else {
                c.setCellValue(value.toString());
            }
        } catch (Exception e) {
            log.error("Error in setValue: {}", e.getMessage());
        }
        return this;
    }

    /**
     * Applies a reusable {@link ExcelStyle} to this cell.
     *
     * @param style the style to apply (no-op if null)
     * @return this cell for chaining
     */
    public ExcelCell applyStyle(ExcelStyle style) {
        if (style != null) {
            style.applyTo(this);
        }
        return this;
    }

    /**
     * Attaches a clickable hyperlink to this cell (URL, email, or file path).
     *
     * @param url       target URL (e.g. {@code https://example.com} or {@code mailto:a@b.com})
     * @param displayText text shown in the cell (if null, the URL is used)
     * @return this cell for chaining
     */
    public ExcelCell setHyperlink(String url, String displayText) {
        try {
            Cell c = this.getCell();
            Workbook workbook = c.getSheet().getWorkbook();
            CreationHelper helper = workbook.getCreationHelper();
            HyperlinkType type = HyperlinkType.URL;
            if (url != null && url.toLowerCase().startsWith("mailto:")) {
                type = HyperlinkType.EMAIL;
            } else if (url != null && (url.startsWith("file:") || url.matches("^[A-Za-z]:\\\\.*") || url.startsWith("/"))) {
                type = HyperlinkType.FILE;
            }
            Hyperlink link = helper.createHyperlink(type);
            link.setAddress(url);
            c.setHyperlink(link);
            c.setCellValue(displayText != null ? displayText : url);

            Font blue = workbook.createFont();
            blue.setColor(IndexedColors.BLUE.getIndex());
            blue.setUnderline(Font.U_SINGLE);
            CellStyle linkStyle = workbook.createCellStyle();
            linkStyle.cloneStyleFrom(c.getCellStyle());
            linkStyle.setFont(blue);
            c.setCellStyle(linkStyle);
        } catch (Exception e) {
            log.error("Error setting hyperlink: {}", e.getMessage());
        }
        return this;
    }

    /**
     * Attaches a note / comment to this cell (appears on hover in Excel).
     *
     * @param text   comment body
     * @param author author label (null falls back to empty)
     * @return this cell for chaining
     */
    public ExcelCell setComment(String text, String author) {
        try {
            Cell c = this.getCell();
            Sheet s = c.getSheet();
            Workbook workbook = s.getWorkbook();
            Drawing<?> drawing = s.createDrawingPatriarch();
            CreationHelper helper = workbook.getCreationHelper();
            ClientAnchor anchor = helper.createClientAnchor();
            anchor.setCol1(c.getColumnIndex());
            anchor.setCol2(c.getColumnIndex() + 3);
            anchor.setRow1(c.getRowIndex());
            anchor.setRow2(c.getRowIndex() + 3);
            Comment comment = drawing.createCellComment(anchor);
            comment.setString(helper.createRichTextString(text != null ? text : ""));
            comment.setAuthor(author != null ? author : "");
            c.setCellComment(comment);
        } catch (Exception e) {
            log.error("Error setting cell comment: {}", e.getMessage());
        }
        return this;
    }

    /**
     * Gets the text value of the cell by reading from the workbook file.
     *
     * @return The text value of the cell, or an empty string if not found or on error
     */
    public String getTextValue() {
        String val = "";
        try {
            Cell cell = this.getCell();
            if (cell != null) {
                switch (cell.getCellType()) {
                    case STRING:
                        if (cell.getRichStringCellValue() != null) {
                            val = cell.getRichStringCellValue().getString();
                        } else {
                            val = cell.getStringCellValue();
                        }
                        break;
                    case NUMERIC:
                        if (DateUtil.isCellDateFormatted(cell)) {
                            val = cell.getDateCellValue().toString();
                        } else {
                            val = String.valueOf(cell.getNumericCellValue());
                        }
                        break;
                    case BOOLEAN:
                        val = String.valueOf(cell.getBooleanCellValue());
                        break;
                    case FORMULA:
                        val = cell.getCellFormula();
                        break;
                    case BLANK:
                        val = "";
                        break;
                    default:
                        val = cell.toString();
                }
            }
        } catch (Exception e) {
            log.error("Error in getting text value: {}", e.getMessage());
        }
        return val;
    }


    /**
     * Checks if a row is empty or contains only blank cells.
     * Delegates to Common utility method.
     *
     * @param row The row to check
     * @return true if the row is empty or null, false otherwise
     */
    private boolean checkIfRowIsEmpty(Row row) {
        return Common.checkIfRowIsEmpty(row);
    }

    /**
     * Checks if a cell is empty or blank.
     * Delegates to Common utility method.
     *
     * @param cell The cell to check
     * @return true if the cell is empty, null, or blank, false otherwise
     */
    private boolean checkIfCellIsEmpty(Cell cell) {
        return Common.checkIfCellIsEmpty(cell);
    }

    /**
     * Gets the file extension from a complete file path.
     *
     * @param sCompleteFilePath The complete file path
     * @return The file extension including the dot, or an empty string on error
     */
    private String getFileExtension(String sCompleteFilePath) {
        try {
            return sCompleteFilePath.substring(sCompleteFilePath.lastIndexOf("."));
        } catch (Exception e) {
            log.error("Error in getting file extension: {}", e.getMessage());
        }

        return ""; // Return empty string on error
    }

    /**
     * Enum for supported named colors for intellisense and documentation.
     * These correspond to org.apache.poi.ss.usermodel.IndexedColors.
     */
    public enum NamedColor {
        AQUA, AUTOMATIC, BLACK, BLACK1, BLUE, BLUE1, BLUE_GREY, BRIGHT_GREEN, BRIGHT_GREEN1, BROWN, CORAL, CORNFLOWER_BLUE,
        DARK_BLUE, DARK_GREEN, DARK_RED, DARK_TEAL, DARK_YELLOW, GOLD, GREEN, GREY_25_PERCENT, GREY_40_PERCENT, GREY_50_PERCENT, GREY_80_PERCENT,
        INDIGO, LAVENDER, LEMON_CHIFFON, LIGHT_BLUE, LIGHT_CORNFLOWER_BLUE, LIGHT_GREEN, LIGHT_ORANGE, LIGHT_TURQUOISE, LIGHT_TURQUOISE1,
        LIGHT_YELLOW, LIME, MAROON, OLIVE_GREEN, ORANGE, ORCHID, PALE_BLUE, PINK, PINK1, PLUM, RED, RED1, ROSE, ROYAL_BLUE, SEA_GREEN, SKY_BLUE,
        TAN, TEAL, TURQUOISE, TURQUOISE1, VIOLET, WHITE, WHITE1, YELLOW, YELLOW1
    }

}
