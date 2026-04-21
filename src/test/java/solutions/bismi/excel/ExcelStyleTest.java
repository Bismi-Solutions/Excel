package solutions.bismi.excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.MethodOrderer;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.TestMethodOrder;

@TestMethodOrder(MethodOrderer.MethodName.class)
class ExcelStyleTest {

    @Test
    void aBuilderProducesExpectedFields() {
        ExcelStyle style = ExcelStyle.builder()
                .fontColor("white")
                .fillColor("grey_50_percent")
                .fullBorder("black")
                .horizontalAlignment("CENTER")
                .verticalAlignment("TOP")
                .numberFormat("#,##0.00")
                .bold(true)
                .italic(true)
                .underline(true)
                .build();

        Assertions.assertEquals("white", style.getFontColor());
        Assertions.assertEquals("grey_50_percent", style.getFillColor());
        Assertions.assertEquals("black", style.getBorderColor());
        Assertions.assertEquals("CENTER", style.getHorizontalAlignment());
        Assertions.assertEquals("TOP", style.getVerticalAlignment());
        Assertions.assertEquals("#,##0.00", style.getNumberFormat());
        Assertions.assertTrue(style.isBold());
        Assertions.assertTrue(style.isItalic());
        Assertions.assertTrue(style.isUnderline());
    }

    @Test
    void bPresetsAreNonNull() {
        Assertions.assertNotNull(ExcelStyle.header());
        Assertions.assertNotNull(ExcelStyle.greyHeader());
        Assertions.assertNotNull(ExcelStyle.zebraStripe());
        Assertions.assertNotNull(ExcelStyle.currency());
        Assertions.assertNotNull(ExcelStyle.percent());
        Assertions.assertNotNull(ExcelStyle.date());
        Assertions.assertEquals("$#,##0.00", ExcelStyle.currency().getNumberFormat());
        Assertions.assertEquals("0.00%", ExcelStyle.percent().getNumberFormat());
    }

    @Test
    void bReportPresetsCarryExpectedColours() {
        ExcelStyle title = ExcelStyle.title();
        Assertions.assertEquals("#0f2d5c", title.getFillColor());
        Assertions.assertEquals("white",   title.getFontColor());
        Assertions.assertTrue(title.isBold());
        Assertions.assertEquals("CENTER",  title.getHorizontalAlignment());

        ExcelStyle totals = ExcelStyle.totals();
        Assertions.assertEquals("#e6eefc", totals.getFillColor());
        Assertions.assertEquals("#0f2d5c", totals.getFontColor());
        Assertions.assertEquals("#2d6cdf", totals.getBorderColor());
        Assertions.assertTrue(totals.isBold());

        // Status pills — fill, font, border all set; bold; centred.
        ExcelStyle active = ExcelStyle.statusActive();
        Assertions.assertEquals("#d1f4dc", active.getFillColor());
        Assertions.assertEquals("#1a7f37", active.getFontColor());
        Assertions.assertEquals("#1a7f37", active.getBorderColor());

        Assertions.assertEquals("#bc4c00", ExcelStyle.statusReview().getFontColor());
        Assertions.assertEquals("#cf222e", ExcelStyle.statusClosed().getFontColor());

        // Custom pill via statusPill(fill, text)
        ExcelStyle custom = ExcelStyle.statusPill("#e0e7ff", "#1d4ed8");
        Assertions.assertEquals("#e0e7ff", custom.getFillColor());
        Assertions.assertEquals("#1d4ed8", custom.getBorderColor());
    }

    @Test
    void fHexBorderColourRendersOnXlsx() {
        ExcelApplication app = new ExcelApplication();
        ExcelWorkBook wb = app.createWorkBook("./resources/testdata/styleHexBorder.xlsx");
        ExcelWorkSheet sh = wb.addSheet("S");

        sh.cell(1, 1).setValue("Totals").applyStyle(ExcelStyle.totals());
        sh.saveWorkBook();

        Cell raw = wb.getWb().getSheet("S").getRow(0).getCell(0);
        org.apache.poi.xssf.usermodel.XSSFCellStyle xs =
                (org.apache.poi.xssf.usermodel.XSSFCellStyle) raw.getCellStyle();
        org.apache.poi.xssf.usermodel.XSSFColor topBorder = xs.getTopBorderXSSFColor();
        Assertions.assertNotNull(topBorder);
        byte[] rgb = topBorder.getRGB();
        // #2d6cdf → 0x2D, 0x6C, 0xDF (bytes may be sign-extended)
        Assertions.assertEquals((byte) 0x2D, rgb[0]);
        Assertions.assertEquals((byte) 0x6C, rgb[1]);
        Assertions.assertEquals((byte) 0xDF, rgb[2]);

        app.closeAllWorkBooks();
    }

    @Test
    void cApplyToNullCellIsNoOp() {
        Assertions.assertDoesNotThrow(() -> ExcelStyle.header().applyTo(null));
    }

    @Test
    void dApplyToCellActuallyChangesStyle() {
        ExcelApplication app = new ExcelApplication();
        ExcelWorkBook wb = app.createWorkBook("./resources/testdata/styleApply.xlsx");
        ExcelWorkSheet sh = wb.addSheet("S");

        ExcelStyle money = ExcelStyle.builder()
                .fillColor("yellow")
                .numberFormat("$#,##0.00")
                .horizontalAlignment("RIGHT")
                .build();

        sh.cell(1, 1).setValue(123.45).applyStyle(money);
        sh.saveWorkBook();

        Cell raw = wb.getWb().getSheet("S").getRow(0).getCell(0);
        CellStyle cs = raw.getCellStyle();
        Assertions.assertEquals(FillPatternType.SOLID_FOREGROUND, cs.getFillPattern());
        Assertions.assertTrue(cs.getDataFormatString().contains("$"));

        app.closeAllWorkBooks();
    }

    @Test
    void eNullStyleIsNoOp() {
        ExcelApplication app = new ExcelApplication();
        ExcelWorkBook wb = app.createWorkBook("./resources/testdata/styleNull.xlsx");
        ExcelWorkSheet sh = wb.addSheet("S");
        Assertions.assertDoesNotThrow(() -> sh.cell(1, 1).applyStyle(null));
        Assertions.assertDoesNotThrow(() -> sh.row(1).applyStyle(null));
        app.closeAllWorkBooks();
    }

    @Test
    void gHexBorderColourFallsBackOnXls() {
        // On .xls (HSSF) there is no true-colour palette for borders. The library
        // must fall back to the closest indexed colour instead of silently writing
        // AUTOMATIC (which is what Common.getColorCode returns for an unknown name).
        ExcelApplication app = new ExcelApplication();
        ExcelWorkBook wb = app.createWorkBook("./resources/testdata/styleHexBorder.xls");
        ExcelWorkSheet sh = wb.addSheet("S");

        sh.cell(1, 1).setValue("Totals").applyStyle(ExcelStyle.totals());
        sh.saveWorkBook();

        Cell raw = wb.getWb().getSheet("S").getRow(0).getCell(0);
        CellStyle cs = raw.getCellStyle();
        short indexed = cs.getTopBorderColor();
        // AUTOMATIC (0x40) is the "no colour picked" sentinel — returned by
        // Common.getColorCode for unknown names. A working nearest-indexed
        // fallback always picks a real palette index.
        Assertions.assertNotEquals((short) 0x40, indexed,
                "hex border on .xls fell through to AUTOMATIC — nearest-indexed fallback missing");

        app.closeAllWorkBooks();
    }
}
