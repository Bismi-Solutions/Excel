import solutions.bismi.excel.ExcelApplication;
import solutions.bismi.excel.ExcelStyle;
import solutions.bismi.excel.ExcelWorkBook;
import solutions.bismi.excel.ExcelWorkSheet;

/**
 * Beginner example showing how to write a whole row from an array in ONE call,
 * instead of reaching for each cell individually.
 *
 * <p>Two patterns are demonstrated:</p>
 * <ul>
 *   <li>{@code row(n).setRowValues(String[])} — when every cell is a string
 *       (headers are the classic use case)</li>
 *   <li>{@code row(n).setValues(Object[])} — when the row mixes strings, numbers,
 *       dates, booleans, etc.</li>
 * </ul>
 *
 * <p>Both methods handle cell creation, so you don't have to loop or track columns.</p>
 */
public class RowFromArrayExample {

    public static void main(String[] args) {
        ExcelApplication app = new ExcelApplication();
        ExcelWorkBook wb = app.createWorkBook("./resources/testdata/rowFromArray.xlsx");
        ExcelWorkSheet sh = wb.addSheet("Shopping");
        sh.activate();

        sh.cell(1, 1).setText("Weekly Shopping List");
        sh.mergeCells(1, 1, 1, 4);
        sh.row(1).applyStyle(ExcelStyle.header(), 1, 5);

        // 1. Header row from a String[] — one call writes all four columns.
        String[] headers = {"Item", "Qty", "Unit", "Aisle"};
        sh.row(2).setRowValues(headers);
        sh.row(2).applyStyle(ExcelStyle.header());

        // 2. Data rows from Object[] — mixed types (String, int, String, String)
        //    get routed to the right Excel cell type automatically.
        sh.row(3).setValues(new Object[]{"Apples",  6, "pcs",  "Produce"});
        sh.row(4).setValues(new Object[]{"Milk",    2, "L",    "Dairy"});
        sh.row(5).setValues(new Object[]{"Bread",   1, "loaf", "Bakery"});
        sh.row(6).setValues(new Object[]{"Coffee",  1, "kg",   "Grocery"});
        sh.row(7).setValues(new Object[]{"Eggs",   12, "pcs",  "Dairy"});

        sh.autoSizeAllColumns();

        wb.saveWorkbook();
        app.closeAllWorkBooks();
        System.out.println("Rows from arrays written to ./resources/testdata/rowFromArray.xlsx");
    }
}
