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
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellUtil;

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
        if (sheet == null) {
            log.error("Sheet cannot be null in ExcelCell constructor");
            throw new IllegalArgumentException("Sheet cannot be null");
        }
        if (wb == null) {
            log.error("Workbook cannot be null in ExcelCell constructor");
            throw new IllegalArgumentException("Workbook cannot be null");
        }
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
            if (cell1 == null) {
                // Log the error, this.cell will remain null.
                // Subsequent getCell() calls will return null, which should be handled by callers.
                log.error("Cell object (cell1) is null after getOrCreateCell for row {}, col {}. Cell will not be set.", this.row, this.col);
            }
            this.setCell(cell1);
        } catch (Exception e) {
            log.error("Error in cell constructor creation: " + e.getMessage(), e);
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

    public boolean setText(String sText) {


        return setText(sText, false);

    }

    public boolean setCellValue(String sText) {


        return setCellValue(sText, false);

    }

    /**
     * Sets a text value in the cell.
     *
     * @param sText           The text value to set
     * @param autoSizeColoumn Whether to auto-size the column
     * @return true if successful, false otherwise
     */
    public boolean setCellValue(String sText, boolean autoSizeColoumn) {
        try {
            Cell currentCell = this.getCell();
            if (currentCell == null) {
                log.error("Cannot set cell value, cell is null (row {}, col {})", this.row, this.col);
                return false;
            }
            currentCell.setCellValue(sText);
            if (autoSizeColoumn) {
                // Ensure sheet is not null, though it should be if currentCell is not.
                if (this.sheet == null) {
                    log.warn("Sheet is null, cannot auto-size column for cell (row {}, col {})", this.row, this.col);
                    return false; // Or handle as appropriate
                }
                this.sheet.autoSizeColumn(col - 1);
            }
            return true;
        } catch (Exception e) {
            log.error("Error in setting cell value: " + e.getMessage(), e);
            return false;
        }
    }

    /**
     * Sets a numeric value in the cell
     *
     * @param value          The numeric value to set
     * @param autoSizeColumn Whether to auto-size the column
     * @return true if successful, false otherwise
     */
    public boolean setNumericValue(double value, boolean autoSizeColumn) {
        try {
            Cell currentCell = this.getCell();
            if (currentCell == null) {
                log.error("Cannot set numeric value, cell is null (row {}, col {})", this.row, this.col);
                return false;
            }
            currentCell.setCellValue(value);
            if (autoSizeColumn) {
                if (this.sheet == null) {
                    log.warn("Sheet is null, cannot auto-size column for cell (row {}, col {})", this.row, this.col);
                    return false;
                }
                this.sheet.autoSizeColumn(col - 1);
            }
            return true;
        } catch (Exception e) {
            log.error("Error in setting numeric value: " + e.getMessage(), e);
            return false;
        }
    }

    /**
     * Sets a numeric value in the cell
     *
     * @param value The numeric value to set
     * @return true if successful, false otherwise
     */
    public boolean setNumericValue(double value) {
        return setNumericValue(value, false);
    }

    /**
     * Sets a date value in the cell
     *
     * @param date           The date value to set
     * @param autoSizeColumn Whether to auto-size the column
     * @return true if successful, false otherwise
     */
    public boolean setDateValue(java.util.Date date, boolean autoSizeColumn) {
        try {
            Cell currentCell = this.getCell();
            if (currentCell == null) {
                log.error("Cannot set date value, cell is null (row {}, col {})", this.row, this.col);
                return false;
            }
            currentCell.setCellValue(date);
            if (autoSizeColumn) {
                if (this.sheet == null) {
                    log.warn("Sheet is null, cannot auto-size column for cell (row {}, col {})", this.row, this.col);
                    return false;
                }
                this.sheet.autoSizeColumn(col - 1);
            }
            return true;
        } catch (Exception e) {
            log.error("Error in setting date value: " + e.getMessage(), e);
            return false;
        }
    }

    /**
     * Sets a date value in the cell
     *
     * @param date The date value to set
     * @return true if successful, false otherwise
     */
    public boolean setDateValue(java.util.Date date) {
        return setDateValue(date, false);
    }

    /**
     * Sets a formula in the cell
     *
     * @param formula        The formula to set (without the leading '=')
     * @param autoSizeColumn Whether to auto-size the column
     * @return true if successful, false otherwise
     */
    public boolean setFormula(String formula, boolean autoSizeColumn) {
        try {
            Cell currentCell = this.getCell();
            if (currentCell == null) {
                log.error("Cannot set formula, cell is null (row {}, col {})", this.row, this.col);
                return false;
            }
            // Remove leading equals sign if present
            if (formula != null && formula.startsWith("=")) {
                formula = formula.substring(1);
            }
            currentCell.setCellFormula(formula);
            if (autoSizeColumn) {
                if (this.sheet == null) {
                    log.warn("Sheet is null, cannot auto-size column for cell (row {}, col {})", this.row, this.col);
                    return false;
                }
                this.sheet.autoSizeColumn(col - 1);
            }
            return true;
        } catch (Exception e) {
            log.error("Error in setting formula: " + e.getMessage(), e);
            return false;
        }
    }

    /**
     * Sets a formula in the cell
     *
     * @param formula The formula to set (without the leading '=')
     * @return true if successful, false otherwise
     */
    public boolean setFormula(String formula) {
        return setFormula(formula, false);
    }

    /**
     * Sets a text value in the cell with text formatting.
     *
     * @param sText           The text value to set
     * @param autoSizeColoumn Whether to auto-size the column
     * @return true if successful, false otherwise
     */
    public boolean setText(String sText, boolean autoSizeColoumn) {
        Cell cCell = null;

        try {
            cCell = this.getCell();
            if (cCell == null) {
                log.error("Cannot set text, cell is null (row {}, col {})", this.row, this.col);
                return false;
            }

            // Get the current cell style
            Workbook workbook = cCell.getSheet().getWorkbook(); // NPE if cCell.getSheet() is null, but POI normally ensures sheet exists for a cell
            CellStyle originalStyle = cCell.getCellStyle();
            CellStyle newStyle = workbook.createCellStyle();
            if (originalStyle != null) {
                newStyle.cloneStyleFrom(originalStyle);
            }
            // If originalStyle is null, newStyle remains a fresh, default style.

            // Set the data format to Text
            DataFormat format = workbook.createDataFormat();
            newStyle.setDataFormat(format.getFormat("@"));  // "@" is the format code for Text

            // Use RichTextString for long text to ensure proper handling
            if (sText != null && !sText.isEmpty()) {
                RichTextString richText = workbook.getCreationHelper().createRichTextString(sText);
                cCell.setCellValue(richText);
            } else {
                cCell.setCellValue("");
            }

            cCell.setCellStyle(newStyle);

            if (autoSizeColoumn) {
                if (this.sheet == null) {
                    log.warn("Sheet is null, cannot auto-size column for cell (row {}, col {})", this.row, this.col);
                    // Decide if this should be a failure or just a warning
                } else {
                    this.sheet.autoSizeColumn(col - 1);
                }
            }

            return true;
        } catch (Exception e) {
            log.error("Error in setting text value: " + e.getMessage(), e);
            return false;
        }
    }

    private void setHexColor(String hexColor, Font newFont) {
        try {
            byte[] rgb = new byte[] {
                (byte) Integer.parseInt(hexColor.substring(0, 2), 16),
                (byte) Integer.parseInt(hexColor.substring(2, 4), 16),
                (byte) Integer.parseInt(hexColor.substring(4, 6), 16)
            };
            org.apache.poi.xssf.usermodel.XSSFFont xssfFont = (org.apache.poi.xssf.usermodel.XSSFFont) newFont;
            xssfFont.setColor(new org.apache.poi.xssf.usermodel.XSSFColor(rgb, null));
        } catch (Exception e) {
            log.error("Error parsing hex color: " + e.getMessage(), e);
            if (newFont != null) { // newFont itself could be null if workbook.createFont() failed, though unlikely
                newFont.setColor(org.apache.poi.ss.usermodel.IndexedColors.BLACK.getIndex());
            }
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
            log.error("Error parsing hex color for HSSF: " + e.getMessage(), e);
            if (newFont != null) {
                newFont.setColor(org.apache.poi.ss.usermodel.IndexedColors.BLACK.getIndex());
            }
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
            if (cCell == null) {
                log.error("Cannot set font color, cell is null (row {}, col {})", this.row, this.col);
                return this;
            }
            // Ensure workbook (wb) is available
            if (this.wb == null) {
                log.error("Workbook (wb) is null, cannot set font color for cell (row {}, col {})", this.row, this.col);
                return this;
            }

            Workbook workbook = this.wb; // Use this.wb for consistency and null-safety
            CellStyle originalStyle = cCell.getCellStyle();
            CellStyle newStyle = workbook.createCellStyle();
            Font newFont = workbook.createFont(); // Create font before trying to copy

            if (originalStyle != null) {
                newStyle.cloneStyleFrom(originalStyle);
                // In POI, getFontIndex() is short, getFontAt(short) is fine.
                // Let's use getFontIndexAsInt() if available, otherwise getFontIndex().
                // For safety, check if font index is valid or if getFontAt returns null.
                Font existingFont = null;
                try {
                    // Attempting with getFontIndexAsInt(), common in newer POI versions
                    existingFont = workbook.getFontAt(originalStyle.getFontIndexAsInt());
                } catch (NoSuchMethodError nsme) {
                    // Fallback for older POI versions
                    existingFont = workbook.getFontAt(originalStyle.getFontIndex());
                }

                if (existingFont != null) {
                    Common.copyFontProperties(existingFont, newFont);
                } else {
                    // existingFont is null, newFont will keep its default properties.
                    // This might happen if the font index in the style is somehow invalid.
                    log.warn("Existing font not found for cell (row {}, col {}), using default font properties.", this.row, this.col);
                }
            }
            // If originalStyle is null, newStyle is a fresh style, and newFont has defaults.

            if (fontColor != null && fontColor.startsWith("#")) {
                String hexColor = fontColor.substring(1);
                if (workbook instanceof org.apache.poi.xssf.usermodel.XSSFWorkbook) {
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
            log.error("Error in setting font color: " + e.getMessage(), e);
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
     *
     * @param fillColor The color name to set for the cell background
     */
    public void setFillColor(String fillColor) {
        try {
            Cell cCell = this.getCell();
            if (cCell == null) {
                log.error("Cannot set fill color, cell is null (row {}, col {})", this.row, this.col);
                return;
            }
            // Ensure workbook (wb) is available
            if (this.wb == null) {
                log.error("Workbook (wb) is null, cannot set fill color for cell (row {}, col {})", this.row, this.col);
                return;
            }

            Workbook workbook = this.wb;
            CellStyle originalStyle = cCell.getCellStyle();
            CellStyle newStyle = workbook.createCellStyle();

            if (originalStyle != null) {
                newStyle.cloneStyleFrom(originalStyle);
            }

            // Set the fill pattern and color
            newStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            newStyle.setFillForegroundColor(Common.getColorCode(fillColor));

            // Apply the new style to the cell
            cCell.setCellStyle(newStyle);
        } catch (Exception e) {
            log.error("Error in setting fill color: " + e.getMessage(), e);
        }
    }

    /**
     * Sets the horizontal alignment of the cell content
     *
     * @param alignment The alignment to set (LEFT, CENTER, RIGHT, etc.)
     * @return true if successful, false otherwise
     */
    public boolean setHorizontalAlignment(String alignment) {
        try {
            Cell cCell = this.getCell();
            if (cCell == null) {
                log.error("Cannot set horizontal alignment, cell is null (row {}, col {})", this.row, this.col);
                return false;
            }
            if (this.wb == null) {
                log.error("Workbook (wb) is null, cannot set horizontal alignment for cell (row {}, col {})", this.row, this.col);
                return false;
            }

            Workbook workbook = this.wb;
            CellStyle originalStyle = cCell.getCellStyle();
            CellStyle newStyle = workbook.createCellStyle();
            if (originalStyle != null) {
                newStyle.cloneStyleFrom(originalStyle);
            }

            HorizontalAlignment hAlign;
            if (alignment == null) { // Guard against NPE from alignment.toUpperCase()
                log.warn("Alignment string is null, using default (GENERAL) for cell (row {}, col {})", this.row, this.col);
                hAlign = HorizontalAlignment.GENERAL;
            } else {
                switch (alignment.toUpperCase()) {
                case "LEFT":
                    hAlign = HorizontalAlignment.LEFT;
                    break;
                case "CENTER":
                    hAlign = HorizontalAlignment.CENTER;
                    break;
                case "RIGHT":
                    hAlign = HorizontalAlignment.RIGHT;
                    break;
                case "JUSTIFY":
                    hAlign = HorizontalAlignment.JUSTIFY;
                    break;
                case "FILL":
                    hAlign = HorizontalAlignment.FILL;
                    break;
                default:
                    hAlign = HorizontalAlignment.GENERAL;
            }

            // Set the horizontal alignment
            newStyle.setAlignment(hAlign);

                }
            }

            // Set the horizontal alignment
            newStyle.setAlignment(hAlign);

            // Apply the new style to the cell
            cCell.setCellStyle(newStyle);
            return true;
        } catch (Exception e) {
            log.error("Error in setting horizontal alignment: " + e.getMessage(), e);
            return false;
        }
    }

    /**
     * Sets the vertical alignment of the cell content
     *
     * @param alignment The alignment to set (TOP, CENTER, BOTTOM, etc.)
     * @return true if successful, false otherwise
     */
    public boolean setVerticalAlignment(String alignment) {
        try {
            Cell cCell = this.getCell();
            if (cCell == null) {
                log.error("Cannot set vertical alignment, cell is null (row {}, col {})", this.row, this.col);
                return false;
            }
            if (this.wb == null) {
                log.error("Workbook (wb) is null, cannot set vertical alignment for cell (row {}, col {})", this.row, this.col);
                return false;
            }

            Workbook workbook = this.wb;
            CellStyle originalStyle = cCell.getCellStyle();
            CellStyle newStyle = workbook.createCellStyle();
            if (originalStyle != null) {
                newStyle.cloneStyleFrom(originalStyle);
            }

            VerticalAlignment vAlign;
            if (alignment == null) { // Guard against NPE from alignment.toUpperCase()
                log.warn("Alignment string is null, using default (CENTER) for cell (row {}, col {})", this.row, this.col);
                vAlign = VerticalAlignment.CENTER;
            } else {
                switch (alignment.toUpperCase()) {
                case "TOP":
                    vAlign = VerticalAlignment.TOP;
                    break;
                case "CENTER":
                    vAlign = VerticalAlignment.CENTER;
                    break;
                case "BOTTOM":
                    vAlign = VerticalAlignment.BOTTOM;
                    break;
                case "JUSTIFY":
                    vAlign = VerticalAlignment.JUSTIFY;
                    break;
                case "DISTRIBUTED":
                    vAlign = VerticalAlignment.DISTRIBUTED;
                    break;
                default:
                    vAlign = VerticalAlignment.CENTER;
            }

            // Set the vertical alignment
            newStyle.setVerticalAlignment(vAlign);

                }
            }

            // Set the vertical alignment
            newStyle.setVerticalAlignment(vAlign);

            // Apply the new style to the cell
            cCell.setCellStyle(newStyle);
            return true;
        } catch (Exception e) {
            log.error("Error in setting vertical alignment: " + e.getMessage(), e);
            return false;
        }
    }

    /**
     * Sets the number format for the cell
     *
     * @param formatPattern The format pattern to use (e.g., "#,##0.00", "dd/mm/yyyy")
     * @return true if successful, false otherwise
     */
    public boolean setNumberFormat(String formatPattern) {
        try {
            Cell cCell = this.getCell();
            if (cCell == null) {
                log.error("Cannot set number format, cell is null (row {}, col {})", this.row, this.col);
                return false;
            }
            if (this.wb == null) {
                log.error("Workbook (wb) is null, cannot set number format for cell (row {}, col {})", this.row, this.col);
                return false;
            }
            if (formatPattern == null) {
                log.error("Format pattern is null for cell (row {}, col {})", this.row, this.col);
                return false;
            }

            Workbook workbook = this.wb;
            CellStyle originalStyle = cCell.getCellStyle();
            CellStyle newStyle = workbook.createCellStyle();
            if (originalStyle != null) {
                newStyle.cloneStyleFrom(originalStyle);
            }

            // Set the number format
            DataFormat format = workbook.createDataFormat();
            newStyle.setDataFormat(format.getFormat(formatPattern));

            // Apply the new style to the cell
            cCell.setCellStyle(newStyle);
            return true;
        } catch (Exception e) {
            log.error("Error in setting number format: " + e.getMessage(), e);
            return false;
        }
    }

    /**
     * Sets the font style (bold, italic, etc.)
     *
     * @param bold      true to make the font bold
     * @param italic    true to make the font italic
     * @param underline true to underline the text
     * @return true if successful, false otherwise
     */
    public boolean setFontStyle(boolean bold, boolean italic, boolean underline) {
        try {
            Cell cCell = this.getCell();
            if (cCell == null) {
                log.error("Cannot set font style, cell is null (row {}, col {})", this.row, this.col);
                return false;
            }
            if (this.wb == null) { // Check this.wb instead of just wb (which is a local var here if not careful)
                log.error("Workbook (this.wb) is null, cannot set font style for cell (row {}, col {})", this.row, this.col);
                return false;
            }
            Font font = this.wb.createFont(); // Use this.wb
            font.setBold(bold);
            font.setItalic(italic);
            if (underline) {
                font.setUnderline(org.apache.poi.ss.usermodel.FontUnderline.SINGLE.getByteValue());
            }
            CellUtil.setFont(cCell, font);
            return true;
        } catch (Exception e) {
            log.error("Error in setting font style: " + e.getMessage(), e);
            return false;
        }
    }

    /**
     * Sets a full border around the cell with the specified color.
     *
     * @param borderColor The color name to set for the border
     */
    public void setFullBorder(String borderColor) {
        try {
            Cell cCell = this.getCell();
            if (cCell == null) {
                log.error("Cannot set full border, cell is null (row {}, col {})", this.row, this.col);
                return;
            }
            if (this.wb == null) {
                log.error("Workbook (wb) is null, cannot set full border for cell (row {}, col {})", this.row, this.col);
                return;
            }
            if (borderColor == null) {
                log.warn("Border color is null, using default for cell (row {}, col {})", this.row, this.col);
                // Potentially set a default color or return, for now proceeding with Common.getColorCode which might handle null
            }

            Workbook workbook = this.wb;
            CellStyle originalStyle = cCell.getCellStyle();
            CellStyle newStyle = workbook.createCellStyle();
            if (originalStyle != null) {
                newStyle.cloneStyleFrom(originalStyle);
            }

            // Set border styles directly
            short colorCode = Common.getColorCode(borderColor); // Common.getColorCode should handle null borderColor gracefully

            newStyle.setBorderLeft(BorderStyle.MEDIUM);
            newStyle.setBorderRight(BorderStyle.MEDIUM);
            newStyle.setBorderTop(BorderStyle.MEDIUM);
            newStyle.setBorderBottom(BorderStyle.MEDIUM);

            newStyle.setLeftBorderColor(colorCode);
            newStyle.setRightBorderColor(colorCode);
            newStyle.setTopBorderColor(colorCode);
            newStyle.setBottomBorderColor(colorCode);

            // Apply the new style to the cell
            cCell.setCellStyle(newStyle);
        } catch (Exception e) {
            log.error("Error in setting border: " + e.getMessage(), e);
        }
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
                        val = cell.getCellFormula(); // Consider: evaluate formula if needed using FormulaEvaluator
                        break;
                    case BLANK:
                        val = "";
                        break;
                    default:
                        val = cell.toString(); // Fallback, might not always be desirable
                }
            }
        } catch (Exception e) {
            log.error("Error in getting text value for cell (row {}, col {}): {}", this.row, this.col, e.getMessage(), e);
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
            if (sCompleteFilePath != null && sCompleteFilePath.contains(".")) {
                return sCompleteFilePath.substring(sCompleteFilePath.lastIndexOf("."));
            } else {
                log.debug("Could not determine file extension for path: {}", sCompleteFilePath); // Changed to debug as it might be normal for non-file paths
            }
        } catch (Exception e) {
            log.error("Error in getting file extension for '" + sCompleteFilePath + "': " + e.getMessage(), e);
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
