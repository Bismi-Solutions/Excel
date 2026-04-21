import solutions.bismi.excel.ExcelApplication;
import solutions.bismi.excel.ExcelColumn;
import solutions.bismi.excel.ExcelStyle;
import solutions.bismi.excel.ExcelWorkBook;
import solutions.bismi.excel.ExcelWorkSheet;
import solutions.bismi.excel.ReportBuilder;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/**
 * End-to-end demo of collection-driven report generation.
 *
 * <p>Shows three user-friendly paths to a styled workbook:
 * <ol>
 *   <li>Sheet 1 — {@link ReportBuilder} driven by a {@code List<Bean>} with
 *       {@link ExcelColumn} annotations</li>
 *   <li>Sheet 2 — {@link ReportBuilder} driven by a {@code List<Map>} with
 *       zebra striping, frozen header, auto-filter</li>
 *   <li>Sheet 3 — Direct {@link ExcelWorkSheet#writeMaps} for the simplest case</li>
 * </ol>
 * None of the user code touches {@code CellStyle}, {@code XSSFWorkbook}, or
 * raw POI constants.</p>
 */
public class CollectionReportExample {

    public static class Product {
        @ExcelColumn(name = "SKU",      order = 1)                               String sku;
        @ExcelColumn(name = "Item",     order = 2)                               String name;
        @ExcelColumn(name = "Units",    order = 3, format = "#,##0")             int units;
        @ExcelColumn(name = "Price",    order = 4, format = "$#,##0.00")         double price;
        @ExcelColumn(name = "Discount", order = 5, format = "0.00%")             double discount;

        public Product() { }
        public Product(String sku, String name, int units, double price, double discount) {
            this.sku = sku; this.name = name; this.units = units;
            this.price = price; this.discount = discount;
        }
    }

    public static void main(String[] args) {
        ExcelApplication app = new ExcelApplication();
        ExcelWorkBook wb = app.createWorkBook("./resources/testdata/collectionReport.xlsx");

        // ---------- Sheet 1: annotated beans ----------
        List<Product> products = Arrays.asList(
                new Product("A-100", "Widget",   120, 12.50, 0.05),
                new Product("A-101", "Gadget",    85, 24.00, 0.10),
                new Product("A-102", "Gizmo",    210,  6.75, 0.00),
                new Product("A-103", "Contraption", 42, 99.99, 0.15));

        ExcelWorkSheet catalog = wb.addSheet("Catalog");
        catalog.activate();   // show this sheet by default when the file is opened
        ReportBuilder.on(catalog)
                .title("Product Catalog — Q3")
                .rowsFromBeans(products)
                .zebraStripes(true)
                .freezeHeader(true)
                .autoFilter(true)
                .render();

        // ---------- Sheet 2: List<Map> + custom column styling ----------
        List<Map<String, Object>> sales = new ArrayList<>();
        sales.add(row("North", 120, 1500.00));
        sales.add(row("South",  85,  920.50));
        sales.add(row("East",  210, 2750.00));
        sales.add(row("West",   42,  499.95));

        ExcelWorkSheet salesSheet = wb.addSheet("Sales");
        ReportBuilder.on(salesSheet)
                .title("Regional Sales Summary")
                .rowsFromMaps(sales)
                .columnStyle(3, ExcelStyle.currency())
                .zebraStripes(true)
                .freezeHeader(true)
                .autoFilter(true)
                .render();

        // ---------- Sheet 3: bare-minimum writeMaps ----------
        ExcelWorkSheet raw = wb.addSheet("Raw");
        raw.writeMaps(sales)
           .freezePane(2, 1)
           .autoSizeAllColumns();

        wb.saveWorkbook();
        app.closeAllWorkBooks();
        System.out.println("Report written to ./resources/testdata/collectionReport.xlsx");
    }

    private static Map<String, Object> row(String region, int units, double revenue) {
        Map<String, Object> m = new LinkedHashMap<>();
        m.put("Region", region);
        m.put("Units", units);
        m.put("Revenue", revenue);
        return m;
    }
}
