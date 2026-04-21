import solutions.bismi.excel.ExcelApplication;
import solutions.bismi.excel.ExcelColumn;
import solutions.bismi.excel.ExcelStyle;
import solutions.bismi.excel.ExcelWorkBook;
import solutions.bismi.excel.ExcelWorkSheet;
import solutions.bismi.excel.ReportBuilder;

import java.time.LocalDate;
import java.util.Arrays;
import java.util.List;

/**
 * Reproduces the styled report rendered by {@code docs/report-preview.svg} — a
 * polished Q3 regional breakdown with a navy title, blue gradient header, zebra
 * stripes, coloured growth arrows, status pills, a comment, and a totals row.
 *
 * <p>Layout (sheet <b>Summary</b>):</p>
 * <pre>
 *   Row 1 : merged A1:G1 — dark-navy bar "Q3 Sales Report — Regional Breakdown"
 *   Row 2 : blue header  — Product · Units · Revenue · Growth · Launched · Owner · Status
 *   Row 3-7: data rows with zebra stripes
 *           Growth column turns green (▲) on positive, red (▼) on negative.
 *           Status column renders as a coloured pill (Active / Review / Closed).
 *   Row 8 : pale-blue totals row with navy bold text
 * </pre>
 *
 * <p>The workbook also carries three empty sibling tabs (Products, Regions,
 * Raw Data) so the Excel tab strip matches the preview.</p>
 */
public class SalesReport {

    /** Bean with {@link ExcelColumn} annotations — drives the bulk of the report. */
    public static class Product {
        @ExcelColumn(name = "Product",  order = 1)                                 String name;
        @ExcelColumn(name = "Units",    order = 2, format = "#,##0")               int units;
        @ExcelColumn(name = "Revenue",  order = 3, format = "$#,##0.00")           double revenue;
        @ExcelColumn(name = "Growth",   order = 4, format = "\"▲ \"0.00%;\"▼ \"0.00%") double growth;
        @ExcelColumn(name = "Launched", order = 5, format = "dd-MMM-yyyy")         LocalDate launched;
        @ExcelColumn(name = "Owner",    order = 6)                                 String owner;
        @ExcelColumn(name = "Status",   order = 7)                                 String status;

        public Product() {}
        public Product(String name, int units, double revenue, double growth,
                       LocalDate launched, String owner, String status) {
            this.name = name; this.units = units; this.revenue = revenue;
            this.growth = growth; this.launched = launched;
            this.owner = owner; this.status = status;
        }
    }

    public static void main(String[] args) {
        ExcelApplication app = new ExcelApplication();
        ExcelWorkBook wb = app.createWorkBook("./resources/testdata/salesReport.xlsx");

        buildSummary(wb);
        wb.addSheet("Products");
        wb.addSheet("Regions");
        wb.addSheet("Raw Data");

        wb.saveWorkbook();
        app.closeAllWorkBooks();
        System.out.println("Sales report written to ./resources/testdata/salesReport.xlsx");
    }

    private static void buildSummary(ExcelWorkBook wb) {
        ExcelWorkSheet sh = wb.addSheet("Summary");
        sh.activate();

        List<Product> products = Arrays.asList(
                new Product("Widget Pro",  2450, 18375.00,  0.1240, LocalDate.of(2026, 1, 12), "A. Rahman", "Active"),
                new Product("Gadget Max",  1820, 24934.00,  0.0875, LocalDate.of(2026, 2,  3), "J. Smith",  "Active"),
                new Product("Contraption",  960,  9504.00, -0.0320, LocalDate.of(2026, 2, 21), "S. Ali",    "Review"),
                new Product("Gizmo Lite",  3105,  6210.50,  0.2210, LocalDate.of(2026, 3,  5), "D. Wu",     "Active"),
                new Product("Legacy Kit",   410,  2050.00, -0.1500, LocalDate.of(2024,11, 12), "R. Khan",   "Closed"));

        // The 5-line core the README advertises — title / headers / zebra / freeze / filter
        // all come from the builder. titleStyle() swaps the default centred-bold title for
        // the navy bar the preview uses.
        ReportBuilder.on(sh)
                .title("Q3 Sales Report — Regional Breakdown")
                .titleStyle(ExcelStyle.title())
                .rowsFromBeans(products)
                .zebraStripes(true)
                .freezeHeader(true)
                .autoFilter(true)
                .columnWidth(1, 18).columnWidth(2, 10).columnWidth(3, 14)
                .columnWidth(4, 12).columnWidth(5, 14).columnWidth(6, 14)
                .columnWidth(7, 12)
                .render();

        // Per-cell tweaks ReportBuilder can't infer: growth colour tracks sign,
        // status colour tracks category. Rows 3..7 (title=1, header=2).
        for (int i = 0; i < products.size(); i++) {
            int r = i + 3;
            Product p = products.get(i);
            sh.cell(r, 4).applyStyle(growthStyle(p.growth));
            sh.cell(r, 7).applyStyle(statusStyle(p.status));
        }

        // Comment on the "Contraption" row explaining the Review flag.
        sh.cell(5, 1).setComment(
                "Margin slipped below threshold this quarter — flagged for review.",
                "Q3 Review");

        // Totals row — row 8. Growth here is always positive, so we pin it green.
        int totalsRow = products.size() + 3;
        sh.row(totalsRow).setValues(new Object[]{"TOTALS", 8745, 61073.50, 0.0501, null, "5 owners", "4 / 5"});
        sh.row(totalsRow).applyStyle(ExcelStyle.totals(), 1, 8);
        sh.cell(totalsRow, 2).setNumberFormat("#,##0");
        sh.cell(totalsRow, 3).setNumberFormat("$#,##0.00");
        sh.cell(totalsRow, 4).setNumberFormat("\"▲ \"0.00%;\"▼ \"0.00%");
        sh.cell(totalsRow, 4).setFontColor("#1a7f37");
        sh.cell(totalsRow, 5).setText("—");
        sh.cell(totalsRow, 7).setHorizontalAlignment("CENTER");
    }

    private static ExcelStyle growthStyle(double growth) {
        return ExcelStyle.builder()
                .fontColor(growth >= 0 ? "#1a7f37" : "#cf222e")
                .bold(true)
                .horizontalAlignment("RIGHT")
                .numberFormat("\"▲ \"0.00%;\"▼ \"0.00%")
                .build();
    }

    private static ExcelStyle statusStyle(String status) {
        switch (status) {
            case "Active": return ExcelStyle.statusActive();
            case "Review": return ExcelStyle.statusReview();
            case "Closed": return ExcelStyle.statusClosed();
            default:       return ExcelStyle.builder().horizontalAlignment("CENTER").build();
        }
    }
}
