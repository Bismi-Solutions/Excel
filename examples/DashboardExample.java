import solutions.bismi.excel.ExcelApplication;
import solutions.bismi.excel.ExcelColumn;
import solutions.bismi.excel.ExcelStyle;
import solutions.bismi.excel.ExcelWorkBook;
import solutions.bismi.excel.ExcelWorkSheet;
import solutions.bismi.excel.ReportBuilder;

import java.util.Arrays;
import java.util.List;

/**
 * Executive dashboard — KPI tiles on one sheet, a top-products table on another,
 * and raw data on a third. Demonstrates:
 *
 * <ul>
 *   <li>Multiple {@link ExcelStyle} reused for KPI tiles (coloured backgrounds, centred)</li>
 *   <li>Cell merging to build "tile" layouts</li>
 *   <li>{@link ReportBuilder} with zebra / freeze / filter on the top-products sheet</li>
 *   <li>Column-style overrides for currency and percent</li>
 * </ul>
 */
public class DashboardExample {

    public static class ProductPerf {
        @ExcelColumn(name = "Product", order = 1)                                 String product;
        @ExcelColumn(name = "Region",  order = 2)                                 String region;
        @ExcelColumn(name = "Revenue", order = 3, format = "$#,##0.00")           double revenue;
        @ExcelColumn(name = "Margin",  order = 4, format = "0.00%")               double margin;

        public ProductPerf() {}
        public ProductPerf(String product, String region, double revenue, double margin) {
            this.product = product; this.region = region;
            this.revenue = revenue; this.margin = margin;
        }
    }

    public static void main(String[] args) {
        ExcelApplication app = new ExcelApplication();
        ExcelWorkBook wb = app.createWorkBook("./resources/testdata/dashboard.xlsx");

        buildSummary(wb);
        buildTopProducts(wb);
        buildRawData(wb);

        wb.saveWorkbook();
        app.closeAllWorkBooks();
        System.out.println("Dashboard written to ./resources/testdata/dashboard.xlsx");
    }

    private static void buildSummary(ExcelWorkBook wb) {
        ExcelWorkSheet sh = wb.addSheet("Summary");
        sh.activate();

        ExcelStyle title = ExcelStyle.builder()
                .bold(true).fontColor("white").fillColor("#0d4ba1")
                .horizontalAlignment("CENTER").build();

        sh.cell(1, 1).setText("Q3 Executive Dashboard");
        sh.mergeCells(1, 1, 1, 8);
        sh.row(1).applyStyle(title, 1, 9);

        kpiTile(sh, 3, 1, "REVENUE",       "$1,248,500", "#1a7f37");
        kpiTile(sh, 3, 3, "ORDERS",        "4,812",      "#2d6cdf");
        kpiTile(sh, 3, 5, "NEW CUSTOMERS", "1,203",      "#bc4c00");
        kpiTile(sh, 3, 7, "RETURN RATE",   "2.4%",       "#cf222e");

        sh.cell(7, 1).setText("Highlights").setFontStyle(true, false, false);
        sh.cell(8, 1).setText("• Widget Pro led revenue at $18,375");
        sh.cell(9, 1).setText("• North region grew 12.4% QoQ");
        sh.cell(10, 1).setText("• Legacy Kit flagged for review");

        for (int c = 1; c <= 8; c++) sh.setColumnWidth(c, 14);
    }

    private static void kpiTile(ExcelWorkSheet sh, int row, int col, String label, String value, String color) {
        ExcelStyle labelStyle = ExcelStyle.builder()
                .bold(true).fontColor("white").fillColor(color)
                .horizontalAlignment("CENTER").build();
        ExcelStyle valueStyle = ExcelStyle.builder()
                .bold(true).fillColor("grey_25_percent")
                .horizontalAlignment("CENTER").fullBorder("black").build();

        sh.cell(row, col).setText(label);
        sh.mergeCells(row, col, row, col + 1);
        sh.row(row).applyStyle(labelStyle, col, col + 2);

        sh.cell(row + 1, col).setText(value);
        sh.mergeCells(row + 1, col, row + 2, col + 1);
        sh.row(row + 1).applyStyle(valueStyle, col, col + 2);
        sh.row(row + 2).applyStyle(valueStyle, col, col + 2);
    }

    private static void buildTopProducts(ExcelWorkBook wb) {
        ExcelWorkSheet sh = wb.addSheet("TopProducts");

        List<ProductPerf> data = Arrays.asList(
                new ProductPerf("Widget Pro",  "North",  18375.00, 0.284),
                new ProductPerf("Gadget Max",  "East",   24934.00, 0.221),
                new ProductPerf("Contraption", "West",    9504.00, 0.157),
                new ProductPerf("Gizmo Lite",  "South",   6210.50, 0.312),
                new ProductPerf("Flagship X",  "North",  31240.00, 0.402));

        ReportBuilder.on(sh)
                .title("Top Products by Revenue")
                .rowsFromBeans(data)
                .zebraStripes(true)
                .freezeHeader(true)
                .autoFilter(true)
                .render();
    }

    private static void buildRawData(ExcelWorkBook wb) {
        ExcelWorkSheet sh = wb.addSheet("RawData");
        String[] headers = {"Order#", "Date", "Product", "Qty", "Amount"};
        sh.row(1).setRowValues(headers);
        sh.row(1).applyStyle(ExcelStyle.header());

        Object[][] orders = {
                {"ORD-1001", java.time.LocalDate.of(2026, 1, 5),  "Widget Pro",  12, 1800.00},
                {"ORD-1002", java.time.LocalDate.of(2026, 1, 7),  "Gadget Max",   5, 1200.00},
                {"ORD-1003", java.time.LocalDate.of(2026, 1, 11), "Flagship X",   2, 7800.00},
                {"ORD-1004", java.time.LocalDate.of(2026, 1, 14), "Gizmo Lite",  40,  800.00},
                {"ORD-1005", java.time.LocalDate.of(2026, 1, 18), "Contraption",  3,  900.00},
        };
        for (int i = 0; i < orders.length; i++) {
            sh.row(i + 2).setValues(orders[i]);
            sh.cell(i + 2, 2).setNumberFormat("dd-MMM-yyyy");
            sh.cell(i + 2, 5).applyStyle(ExcelStyle.currency());
        }
        sh.freezePane(2, 1).enableAutoFilter(1, 1, headers.length).autoSizeAllColumns();
    }
}
