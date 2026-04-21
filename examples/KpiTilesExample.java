import solutions.bismi.excel.ExcelApplication;
import solutions.bismi.excel.ExcelStyle;
import solutions.bismi.excel.ExcelWorkBook;
import solutions.bismi.excel.ExcelWorkSheet;

/**
 * Beginner-friendly dashboard: just four colour-coded KPI tiles on one sheet.
 * No beans, no annotations, no loops — ~15 meaningful lines of code.
 *
 * <p>For the full three-sheet version with a data table and raw data, see
 * {@code DashboardExample}.</p>
 */
public class KpiTilesExample {

    public static void main(String[] args) {
        ExcelApplication app = new ExcelApplication();
        ExcelWorkBook wb = app.createWorkBook("./resources/testdata/kpiTiles.xlsx");
        ExcelWorkSheet sh = wb.addSheet("KPIs");
        sh.activate();

        sh.cell(1, 1).setText("Q3 Snapshot");
        sh.mergeCells(1, 1, 1, 4);
        sh.row(1).applyStyle(ExcelStyle.header(), 1, 5);

        kpi(sh, 1, "Revenue",       "$1,248,500", "#1a7f37");   // green
        kpi(sh, 2, "Orders",        "4,812",      "#2d6cdf");   // blue
        kpi(sh, 3, "New Customers", "1,203",      "#bc4c00");   // orange
        kpi(sh, 4, "Return Rate",   "2.4%",       "#cf222e");   // red

        wb.saveWorkbook();
        app.closeAllWorkBooks();
        System.out.println("KPI tiles written to ./resources/testdata/kpiTiles.xlsx");
    }

    private static void kpi(ExcelWorkSheet sh, int col, String label, String value, String colour) {
        ExcelStyle labelStyle = ExcelStyle.builder()
                .fillColor(colour).fontColor("white").bold(true)
                .horizontalAlignment("CENTER").build();
        ExcelStyle valueStyle = ExcelStyle.builder()
                .bold(true).horizontalAlignment("CENTER").fullBorder("black").build();

        sh.cell(3, col).setText(label).applyStyle(labelStyle);
        sh.cell(4, col).setText(value).applyStyle(valueStyle);
        sh.setColumnWidth(col, 18);
    }
}
