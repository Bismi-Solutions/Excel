package solutions.bismi.excel;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPane;
import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.MethodOrderer;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.TestMethodOrder;

import java.util.Arrays;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

@TestMethodOrder(MethodOrderer.MethodName.class)
class ReportBuilderTest {

    public static class Sale {
        @ExcelColumn(name = "Region", order = 1)                          String region;
        @ExcelColumn(name = "Units",  order = 2, format = "#,##0")        int units;
        @ExcelColumn(name = "Revenue",order = 3, format = "$#,##0.00")    double revenue;

        public Sale() {}
        public Sale(String region, int units, double revenue) {
            this.region = region; this.units = units; this.revenue = revenue;
        }
    }

    @Test
    void aRenderFromBeansWritesHeaderAndRows() {
        String path = "./resources/testdata/reportBeans.xlsx";
        ExcelApplication app = new ExcelApplication();
        ExcelWorkBook wb = app.createWorkBook(path);
        ExcelWorkSheet sh = wb.addSheet("Sales");

        List<Sale> sales = Arrays.asList(
                new Sale("North", 120, 1500.00),
                new Sale("South",  85,  920.50),
                new Sale("East",  210, 2750.00));

        ReportBuilder.on(sh)
                .title("Q3 Sales")
                .rowsFromBeans(sales)
                .zebraStripes(true)
                .freezeHeader(true)
                .autoFilter(true)
                .render();

        Sheet raw = wb.getWb().getSheet("Sales");
        Assertions.assertEquals("Q3 Sales", raw.getRow(0).getCell(0).getStringCellValue());
        Assertions.assertEquals("Region",   raw.getRow(1).getCell(0).getStringCellValue());
        Assertions.assertEquals("North",    raw.getRow(2).getCell(0).getStringCellValue());
        Assertions.assertEquals(120d,       raw.getRow(2).getCell(1).getNumericCellValue(), 0.0001);
        Assertions.assertEquals(1500.00,    raw.getRow(2).getCell(2).getNumericCellValue(), 0.0001);

        XSSFSheet xs = (XSSFSheet) raw;
        CTPane pane = xs.getCTWorksheet().getSheetViews().getSheetViewArray(0).getPane();
        Assertions.assertNotNull(pane);
        // Title row + header row → freeze below row 2, so ySplit == 2
        Assertions.assertEquals(2.0, pane.getYSplit(), 0.0001);

        app.closeAllWorkBooks();
    }

    @Test
    void bRenderFromMapsInfersHeaders() {
        String path = "./resources/testdata/reportMaps.xlsx";
        ExcelApplication app = new ExcelApplication();
        ExcelWorkBook wb = app.createWorkBook(path);
        ExcelWorkSheet sh = wb.addSheet("M");

        List<Map<String, Object>> rows = Arrays.asList(
                ordered("Name", "A", "Score", 10),
                ordered("Name", "B", "Score", 20));
        ReportBuilder.on(sh).rowsFromMaps(rows).render();

        Sheet raw = wb.getWb().getSheet("M");
        Assertions.assertEquals("Name", raw.getRow(0).getCell(0).getStringCellValue());
        Assertions.assertEquals("Score", raw.getRow(0).getCell(1).getStringCellValue());
        Assertions.assertEquals(20d, raw.getRow(2).getCell(1).getNumericCellValue(), 0.0001);

        app.closeAllWorkBooks();
    }

    @Test
    void cRenderFromArrayRows() {
        String path = "./resources/testdata/reportArrays.xlsx";
        ExcelApplication app = new ExcelApplication();
        ExcelWorkBook wb = app.createWorkBook(path);
        ExcelWorkSheet sh = wb.addSheet("A");

        ReportBuilder.on(sh)
                .headers("Col1", "Col2")
                .row("x", 1)
                .row("y", 2)
                .columnStyle(2, ExcelStyle.builder().numberFormat("#,##0").build())
                .columnWidth(1, 15)
                .render();

        Sheet raw = wb.getWb().getSheet("A");
        Assertions.assertEquals("Col1", raw.getRow(0).getCell(0).getStringCellValue());
        Assertions.assertEquals("x",    raw.getRow(1).getCell(0).getStringCellValue());
        Assertions.assertEquals(2d,     raw.getRow(2).getCell(1).getNumericCellValue(), 0.0001);
        Assertions.assertEquals(15 * 256, raw.getColumnWidth(0));

        app.closeAllWorkBooks();
    }

    @Test
    void dRenderWithoutHeadersThrows() {
        ExcelApplication app = new ExcelApplication();
        ExcelWorkBook wb = app.createWorkBook("./resources/testdata/reportNoHeaders.xlsx");
        ExcelWorkSheet sh = wb.addSheet("X");

        ReportBuilder rb = ReportBuilder.on(sh);
        Assertions.assertThrows(IllegalStateException.class, rb::render);
        app.closeAllWorkBooks();
    }

    @Test
    void eAutoFilterAppliedOnHeader() {
        String path = "./resources/testdata/reportFilter.xlsx";
        ExcelApplication app = new ExcelApplication();
        ExcelWorkBook wb = app.createWorkBook(path);
        ExcelWorkSheet sh = wb.addSheet("F");

        ReportBuilder.on(sh)
                .headers("A", "B", "C")
                .row(1, 2, 3)
                .autoFilter(true)
                .render();

        org.apache.poi.xssf.usermodel.XSSFSheet xs =
                (org.apache.poi.xssf.usermodel.XSSFSheet) wb.getWb().getSheet("F");
        Assertions.assertNotNull(xs.getCTWorksheet().getAutoFilter());
        Assertions.assertEquals("A1:C1", xs.getCTWorksheet().getAutoFilter().getRef());

        app.closeAllWorkBooks();
    }

    @Test
    void fRowsFromMapsNullIsNoOp() {
        ExcelApplication app = new ExcelApplication();
        ExcelWorkBook wb = app.createWorkBook("./resources/testdata/reportNullMaps.xlsx");
        ExcelWorkSheet sh = wb.addSheet("N");
        Assertions.assertDoesNotThrow(() ->
                ReportBuilder.on(sh).headers("x").rowsFromMaps(null).render());
        app.closeAllWorkBooks();
    }

    private static Map<String, Object> ordered(Object... kv) {
        LinkedHashMap<String, Object> m = new LinkedHashMap<>();
        for (int i = 0; i < kv.length; i += 2) {
            m.put((String) kv[i], kv[i + 1]);
        }
        return m;
    }
}
