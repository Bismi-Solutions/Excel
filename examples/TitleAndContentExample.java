import solutions.bismi.excel.ExcelApplication;
import solutions.bismi.excel.ExcelStyle;
import solutions.bismi.excel.ExcelWorkBook;
import solutions.bismi.excel.ExcelWorkSheet;

/**
 * The simplest possible styled workbook — a title and a few lines of content.
 * No beans, no annotations, no loops. Ideal starting point for first-time users.
 *
 * <p>Shows just three things:</p>
 * <ol>
 *   <li>Merge cells to make a title band</li>
 *   <li>Apply a built-in {@link ExcelStyle} preset ({@code header})</li>
 *   <li>Write plain text rows below it</li>
 * </ol>
 */
public class TitleAndContentExample {

    public static void main(String[] args) {
        ExcelApplication app = new ExcelApplication();
        ExcelWorkBook wb = app.createWorkBook("./resources/testdata/titleContent.xlsx");
        ExcelWorkSheet sh = wb.addSheet("Notes");
        sh.activate();

        sh.cell(1, 1).setText("Project Kickoff — 2026 Roadmap");
        sh.mergeCells(1, 1, 1, 4);
        sh.row(1).applyStyle(ExcelStyle.header(), 1, 5);

        ExcelStyle bold = ExcelStyle.builder().bold(true).build();

        sh.cell(3, 1).setText("Goals").applyStyle(bold);
        sh.cell(4, 1).setText("• Migrate all clients to the v2 API");
        sh.cell(5, 1).setText("• Cut average latency by 30%");
        sh.cell(6, 1).setText("• Launch mobile SDK in Q3");

        sh.cell(8, 1).setText("Owners").applyStyle(bold);
        sh.cell(9, 1).setText("• Engineering — John Smith");
        sh.cell(10, 1).setText("• Product — Amina Khan");
        sh.cell(11, 1).setText("• Design — John Jones");

        sh.autoSizeAllColumns();

        wb.saveWorkbook();
        app.closeAllWorkBooks();
        System.out.println("Title + content written to ./resources/testdata/titleContent.xlsx");
    }
}
