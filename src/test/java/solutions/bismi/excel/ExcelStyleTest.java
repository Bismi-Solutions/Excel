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
        Assertions.assertNotNull(ExcelStyle.zebraStripe());
        Assertions.assertNotNull(ExcelStyle.currency());
        Assertions.assertNotNull(ExcelStyle.percent());
        Assertions.assertNotNull(ExcelStyle.date());
        Assertions.assertEquals("$#,##0.00", ExcelStyle.currency().getNumberFormat());
        Assertions.assertEquals("0.00%", ExcelStyle.percent().getNumberFormat());
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
}
