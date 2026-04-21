import solutions.bismi.excel.ExcelApplication;
import solutions.bismi.excel.ExcelStyle;
import solutions.bismi.excel.ExcelWorkBook;
import solutions.bismi.excel.ExcelWorkSheet;

/**
 * Produces a one-page invoice styled with merged header, bill-to block, line items,
 * subtotal / tax / total. Demonstrates:
 *
 * <ul>
 *   <li>Merged title across columns</li>
 *   <li>Reusable {@link ExcelStyle} for labels, currency, totals</li>
 *   <li>Mixed-type row values via {@code setValues(Object[])}</li>
 *   <li>Formulas for line totals and subtotal</li>
 * </ul>
 */
public class InvoiceExample {

    public static void main(String[] args) {
        ExcelApplication app = new ExcelApplication();
        ExcelWorkBook wb = app.createWorkBook("./resources/testdata/invoice.xlsx");
        ExcelWorkSheet sh = wb.addSheet("Invoice");

        ExcelStyle titleStyle = ExcelStyle.builder()
                .bold(true).fontColor("white").fillColor("#0d4ba1")
                .horizontalAlignment("CENTER").build();
        ExcelStyle labelStyle = ExcelStyle.builder()
                .bold(true).fontColor("#0d4ba1").build();
        ExcelStyle addressStyle = ExcelStyle.builder()
                .fillColor("#ebeef2").build();
        ExcelStyle headerStyle = ExcelStyle.greyHeader();
        ExcelStyle money = ExcelStyle.currency();
        ExcelStyle totalRow = ExcelStyle.builder()
                .bold(true).fillColor("#e6eefc").numberFormat("$#,##0.00")
                .horizontalAlignment("RIGHT").fullBorder("black").build();

        sh.cell(1, 1).setText("INVOICE");
        sh.mergeCells(1, 1, 1, 5);
        sh.row(1).applyStyle(titleStyle, 1, 6);

        sh.cell(3, 1).setText("Invoice #:").applyStyle(labelStyle);
        sh.cell(3, 2).setText("INV-2026-0042");
        sh.cell(3, 4).setText("Date:").applyStyle(labelStyle);
        sh.cell(3, 5).setValue(java.time.LocalDate.of(2026, 4, 21))
                      .setNumberFormat("dd-MMM-yyyy");

        sh.cell(5, 1).setText("From").applyStyle(labelStyle);
        sh.cell(6, 1).setText("Bismi Solutions").applyStyle(addressStyle);
        sh.cell(7, 1).setText("Infopark, Phase II").applyStyle(addressStyle);
        sh.cell(8, 1).setText("Kochi, India 682042").applyStyle(addressStyle);

        sh.cell(5, 4).setText("Bill To").applyStyle(labelStyle);
        sh.cell(6, 4).setText("Acme Corp.").applyStyle(addressStyle);
        sh.cell(7, 4).setText("123 Market Street").applyStyle(addressStyle);
        sh.cell(8, 4).setText("New York, NY 10001").applyStyle(addressStyle);

        String[] headers = {"Qty", "Description", "Unit Price", "Total"};
        sh.row(10).setRowValues(headers);
        sh.row(10).applyStyle(headerStyle, 1, 5);

        Object[][] items = {
                {5,  "Excel library — enterprise licence",  299.00},
                {12, "Premium support — hours",              150.00},
                {1,  "Onsite training day",                 1200.00},
                {3,  "Custom reporting setup",               450.00},
        };
        int startRow = 11;
        for (int i = 0; i < items.length; i++) {
            int r = startRow + i;
            sh.row(r).setValues(new Object[]{items[i][0], items[i][1], items[i][2]});
            sh.cell(r, 4).setFormula("A" + r + "*C" + r);
            sh.cell(r, 3).applyStyle(money);
            sh.cell(r, 4).applyStyle(money);
        }

        int totalsRow = startRow + items.length + 1;
        sh.cell(totalsRow,     3).setText("Subtotal").applyStyle(labelStyle);
        sh.cell(totalsRow,     4).setFormula("SUM(D" + startRow + ":D" + (totalsRow - 2) + ")")
                                  .applyStyle(money);
        sh.cell(totalsRow + 1, 3).setText("Tax (8%)").applyStyle(labelStyle);
        sh.cell(totalsRow + 1, 4).setFormula("D" + totalsRow + "*0.08").applyStyle(money);
        sh.cell(totalsRow + 2, 3).setText("TOTAL").applyStyle(totalRow);
        sh.cell(totalsRow + 2, 4).setFormula("D" + totalsRow + "+D" + (totalsRow + 1))
                                  .applyStyle(totalRow);

        sh.setColumnWidth(1, 12);
        sh.setColumnWidth(2, 40);
        sh.setColumnWidth(3, 16);
        sh.setColumnWidth(4, 16);
        sh.setColumnWidth(5, 14);

        wb.saveWorkbook();
        app.closeAllWorkBooks();
        System.out.println("Invoice written to ./resources/testdata/invoice.xlsx");
    }
}
