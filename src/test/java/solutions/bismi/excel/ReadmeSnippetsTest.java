package solutions.bismi.excel;

import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.MethodOrderer;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.TestMethodOrder;

import java.util.Arrays;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/**
 * Every code block visible in the README is reproduced here verbatim and executed.
 * If a snippet drifts from the real API, this class stops compiling or fails —
 * making README code a build-time invariant, not a hope.
 */
@TestMethodOrder(MethodOrderer.MethodName.class)
class ReadmeSnippetsTest {

    public static class Product {
        @ExcelColumn(name = "SKU",   order = 1)                      String sku;
        @ExcelColumn(name = "Item",  order = 2)                      String name;
        @ExcelColumn(name = "Units", order = 3, format = "#,##0")    int    units;
        @ExcelColumn(name = "Price", order = 4, format = "$#,##0.00") double price;

        public Product() {}
        public Product(String sku, String name, int units, double price) {
            this.sku = sku; this.name = name; this.units = units; this.price = price;
        }
    }

    // README — "Hook" block: ReportBuilder on beans
    @Test
    void aHookSnippetRendersFromBeans() {
        List<Product> products = Arrays.asList(
                new Product("A-100", "Widget",  120, 12.50),
                new Product("A-101", "Gadget",   85, 24.00));

        ExcelApplication app = new ExcelApplication();
        ExcelWorkBook wb = app.createWorkBook("./resources/testdata/readmeHook.xlsx");
        ExcelWorkSheet sheet = wb.addSheet("Sales");

        // --- snippet begin ---
        ReportBuilder.on(sheet)
            .title("Q3 Sales Report")
            .rowsFromBeans(products)
            .zebraStripes(true).freezeHeader(true).autoFilter(true)
            .render();
        // --- snippet end ---

        Assertions.assertEquals("Q3 Sales Report",
                wb.getWb().getSheet("Sales").getRow(0).getCell(0).getStringCellValue());
        app.closeAllWorkBooks();
    }

    // README — "1 · Hello, styled workbook (6 lines)"
    @Test
    void bHelloStyledWorkbookSixLines() {
        // --- snippet begin ---
        ExcelApplication app = new ExcelApplication();
        ExcelWorkBook wb = app.createWorkBook("./resources/testdata/readmeHello.xlsx");
        ExcelWorkSheet sh = wb.addSheet("Summary");
        sh.cell(1,1).setText("Hello World").applyStyle(ExcelStyle.header());
        wb.saveWorkbook();
        app.closeAllWorkBooks();
        // --- snippet end ---

        ExcelApplication a2 = new ExcelApplication();
        ExcelWorkBook w2 = a2.openWorkbook("./resources/testdata/readmeHello.xlsx");
        Assertions.assertEquals("Hello World",
                w2.getWb().getSheet("Summary").getRow(0).getCell(0).getStringCellValue());
        a2.closeAllWorkBooks();
    }

    // README — "2 · Bean → styled report (one call)"
    @Test
    void cBeanToStyledReportOneCall() {
        List<Product> productList = Arrays.asList(new Product("X", "Thing", 1, 1.00));
        ExcelApplication app = new ExcelApplication();
        ExcelWorkBook wb = app.createWorkBook("./resources/testdata/readmeBean.xlsx");
        ExcelWorkSheet sheet = wb.addSheet("Catalog");

        // --- snippet begin ---
        ReportBuilder.on(sheet)
            .title("Catalog")
            .rowsFromBeans(productList)
            .zebraStripes(true).freezeHeader(true).autoFilter(true)
            .render();
        // --- snippet end ---

        Assertions.assertEquals("Catalog",
                wb.getWb().getSheet("Catalog").getRow(0).getCell(0).getStringCellValue());
        app.closeAllWorkBooks();
    }

    // README — "3 · List<Map> → spreadsheet"
    @Test
    void dListMapToSpreadsheet() {
        ExcelApplication app = new ExcelApplication();
        ExcelWorkBook wb = app.createWorkBook("./resources/testdata/readmeMaps.xlsx");
        ExcelWorkSheet sheet = wb.addSheet("M");

        // --- snippet begin ---
        List<Map<String,Object>> rows = List.of(
            ordered("Item","Apple", "Qty",10, "Price",1.20),
            ordered("Item","Pear",  "Qty", 8, "Price",1.80));

        sheet.writeMaps(rows).freezePane(2,1).autoSizeAllColumns();
        // --- snippet end ---

        Assertions.assertEquals("Apple",
                wb.getWb().getSheet("M").getRow(1).getCell(0).getStringCellValue());
        app.closeAllWorkBooks();
    }

    // README — "4 · Read Excel → List<Bean>"
    @Test
    void eReadExcelToListOfBeans() {
        String path = "./resources/testdata/readmeRead.xlsx";
        ExcelApplication app = new ExcelApplication();
        ExcelWorkBook wb = app.createWorkBook(path);
        ExcelWorkSheet sheet = wb.addSheet("P");
        sheet.writeBeans(Arrays.asList(
                new Product("A", "Alpha", 5, 10.00),
                new Product("B", "Beta",  6, 20.00)));
        app.closeAllWorkBooks();

        ExcelApplication app2 = new ExcelApplication();
        ExcelWorkBook wb2 = app2.openWorkbook(path);
        ExcelWorkSheet sheet2 = wb2.getExcelSheet("P");

        // --- snippet begin ---
        List<Product> products = sheet2.readAsBeans(Product.class);
        // --- snippet end ---

        Assertions.assertEquals(2, products.size());
        Assertions.assertEquals("Alpha", products.get(0).name);
        Assertions.assertEquals(5, products.get(0).units);
        app2.closeAllWorkBooks();
    }

    // README — "5 · Reusable style, applied 1,000 times"
    @Test
    void fReusableStyleApplied1000Times() {
        ExcelApplication app = new ExcelApplication();
        ExcelWorkBook wb = app.createWorkBook("./resources/testdata/readmeReuse.xlsx");
        ExcelWorkSheet sheet = wb.addSheet("R");
        double[] revenue = new double[1000];
        for (int i = 0; i < 1000; i++) revenue[i] = 100.0 + i;

        // --- snippet begin ---
        ExcelStyle money = ExcelStyle.builder()
                .numberFormat("$#,##0.00").horizontalAlignment("RIGHT").fullBorder("black").build();

        for (int r = 2; r <= 1001; r++) {
            sheet.cell(r, 3).setValue(revenue[r-2]).applyStyle(money);
        }
        // --- snippet end ---

        Assertions.assertEquals(100.0,
                wb.getWb().getSheet("R").getRow(1).getCell(2).getNumericCellValue(), 0.0001);
        app.closeAllWorkBooks();
    }

    // README — "6 · Hyperlinks & comments"
    @Test
    void gHyperlinksAndComments() {
        ExcelApplication app = new ExcelApplication();
        ExcelWorkBook wb = app.createWorkBook("./resources/testdata/readmeHyperlinks.xlsx");
        ExcelWorkSheet sheet = wb.addSheet("H");

        // --- snippet begin ---
        sheet.cell(1,1).setHyperlink("https://bismi.solutions", "Bismi Solutions");
        sheet.cell(1,1).setComment("Official site", "Release notes");
        // --- snippet end ---

        org.apache.poi.ss.usermodel.Cell c = wb.getWb().getSheet("H").getRow(0).getCell(0);
        Assertions.assertNotNull(c.getHyperlink());
        Assertions.assertNotNull(c.getCellComment());
        app.closeAllWorkBooks();
    }

    // README — "Rows from arrays — one call writes the whole row"
    @Test
    void iRowsFromArrays() {
        ExcelApplication app = new ExcelApplication();
        ExcelWorkBook wb = app.createWorkBook("./resources/testdata/readmeRowArrays.xlsx");
        ExcelWorkSheet sh = wb.addSheet("S");

        // --- snippet begin ---
        // Header row from a String[]
        String[] headers = {"Item", "Qty", "Unit", "Aisle"};
        sh.row(2).setRowValues(headers);

        // Data rows from Object[] — mixed types get routed automatically
        sh.row(3).setValues(new Object[]{"Apples",  6, "pcs", "Produce"});
        sh.row(4).setValues(new Object[]{"Milk",    2, "L",   "Dairy"});
        // --- snippet end ---

        org.apache.poi.ss.usermodel.Sheet raw = wb.getWb().getSheet("S");
        Assertions.assertEquals("Item",    raw.getRow(1).getCell(0).getStringCellValue());
        Assertions.assertEquals("Aisle",   raw.getRow(1).getCell(3).getStringCellValue());
        Assertions.assertEquals("Apples",  raw.getRow(2).getCell(0).getStringCellValue());
        Assertions.assertEquals(6.0,       raw.getRow(2).getCell(1).getNumericCellValue(), 0.0001);
        Assertions.assertEquals("Produce", raw.getRow(2).getCell(3).getStringCellValue());

        sh.saveWorkBook();
        app.closeAllWorkBooks();
    }

    // README — "Style reuse at a glance"
    @Test
    void hStyleReuseAtAGlance() {
        ExcelApplication app = new ExcelApplication();
        ExcelWorkBook wb = app.createWorkBook("./resources/testdata/readmeStyleReuse.xlsx");
        ExcelWorkSheet sheet = wb.addSheet("S");

        // --- snippet begin ---
        ExcelStyle header = ExcelStyle.header();
        ExcelStyle zebra  = ExcelStyle.zebraStripe();
        ExcelStyle money  = ExcelStyle.currency();
        ExcelStyle pct    = ExcelStyle.percent();
        ExcelStyle when   = ExcelStyle.date();

        ExcelStyle custom = ExcelStyle.builder()
                .fontColor("white").fillColor("#0d4ba1")
                .bold(true).horizontalAlignment("CENTER").fullBorder("black")
                .build();

        sheet.row(1).applyStyle(header);
        sheet.row(5).applyStyle(money, 2, 6);            // cells 2..5 of row 5
        sheet.cell(10,3).applyStyle(money);              // single cell
        // --- snippet end ---

        Assertions.assertNotNull(header);
        Assertions.assertNotNull(zebra);
        Assertions.assertNotNull(pct);
        Assertions.assertNotNull(when);
        Assertions.assertNotNull(custom);
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
