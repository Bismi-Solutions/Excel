import solutions.bismi.excel.*;

import java.time.LocalDate;
import java.util.Arrays;
import java.util.List;

/**
 * A comprehensive example demonstrating Excel library features
 * Creates a multi-sheet sales report with formatting, formulas, and charts
 */
public class SalesReport {
    public static void main(String[] args) {
        ExcelApplication app = new ExcelApplication();
        ExcelWorkBook wb = app.createWorkBook("SalesReport.xlsx");
        
        // Create Summary sheet
        ExcelWorkSheet summary = wb.addSheet("Summary");
        summary.activate();
        createSummarySheet(summary);
        
        // Create Sales Data sheet
        ExcelWorkSheet sales = wb.addSheet("Sales Data");
        createSalesDataSheet(sales);
        
        // Create Charts sheet
        ExcelWorkSheet charts = wb.addSheet("Charts");
        createChartsSheet(charts, sales);
        
        // Save and close
        wb.saveWorkbook();
        app.closeAllWorkBooks();
    }
    
    private static void createSummarySheet(ExcelWorkSheet sheet) {
        // Title
        sheet.cell(1, 1).setText("Q1 2024 Sales Report");
        sheet.cell(1, 1).setFontStyle(true, false, false);
        sheet.cell(1, 1).setHorizontalAlignment("CENTER");
        sheet.mergeCells(1, 1, 1, 5);
        
        // Summary metrics
        String[] metrics = {"Total Sales", "Average Order", "Top Product", "Growth"};
        String[] values = {"=SUM('Sales Data'!B2:B31)", "=AVERAGE('Sales Data'!B2:B31)", 
                          "=INDEX('Sales Data'!A2:A31,MATCH(MAX('Sales Data'!B2:B31),'Sales Data'!B2:B31,0))",
                          "=((SUM('Sales Data'!B2:B31)-SUM('Sales Data'!C2:C31))/SUM('Sales Data'!C2:C31))"};
        
        for (int i = 0; i < metrics.length; i++) {
            sheet.cell(3, i+1).setText(metrics[i]);
            sheet.cell(3, i+1).setFontStyle(true, false, false);
            sheet.cell(3, i+1).setFillColor("grey_25_percent");
            sheet.cell(4, i+1).setFormula(values[i]);
            if (i < 2) {
                sheet.cell(4, i+1).setNumberFormat("$#,##0.00");
            } else if (i == 3) {
                sheet.cell(4, i+1).setNumberFormat("0.00%");
            } else {
                sheet.cell(4, i+1).setNumberFormat("@");
            }
        }
        
        // Format summary section
        sheet.cell(3, 1).setFullBorder("black");
        sheet.cell(3, 4).setFullBorder("black");
        sheet.cell(4, 1).setFullBorder("black");
        sheet.cell(4, 4).setFullBorder("black");
    }
    
    private static void createSalesDataSheet(ExcelWorkSheet sheet) {
        // Headers
        String[] headers = {"Product", "Current Sales", "Previous Sales", "Growth", "Category"};
        for (int i = 0; i < headers.length; i++) {
            sheet.cell(1, i+1).setText(headers[i]);
            sheet.cell(1, i+1).setFontStyle(true, false, false);
            sheet.cell(1, i+1).setFillColor("grey_25_percent");
            sheet.cell(1, i+1).setHorizontalAlignment("CENTER");
        }
        
        // Sample data
        List<Product> products = Arrays.asList(
            new Product("Laptop Pro", 125000, 100000, "Electronics"),
            new Product("Smartphone X", 150000, 120000, "Electronics"),
            new Product("Office Chair", 45000, 40000, "Furniture"),
            new Product("Desk Lamp", 25000, 20000, "Furniture"),
            new Product("Wireless Mouse", 35000, 30000, "Accessories")
        );
        
        // Add data rows
        for (int i = 0; i < products.size(); i++) {
            Product p = products.get(i);
            int row = i + 2;
            
            sheet.cell(row, 1).setText(p.name);
            sheet.cell(row, 2).setNumericValue(p.currentSales);
            sheet.cell(row, 2).setNumberFormat("$#,##0.00");
            sheet.cell(row, 3).setNumericValue(p.previousSales);
            sheet.cell(row, 3).setNumberFormat("$#,##0.00");
            sheet.cell(row, 4).setFormula(String.format("=(B%d-C%d)/C%d", row, row, row));
            sheet.cell(row, 4).setNumberFormat("0.00%");
            sheet.cell(row, 5).setText(p.category);
        }
        
        // Add totals row
        int lastRow = products.size() + 2;
        sheet.cell(lastRow, 1).setText("Total");
        sheet.cell(lastRow, 1).setFontStyle(true, false, false);
        sheet.cell(lastRow, 2).setFormula(String.format("=SUM(B2:B%d)", lastRow-1));
        sheet.cell(lastRow, 2).setNumberFormat("$#,##0.00");
        sheet.cell(lastRow, 2).setFontStyle(true, false, false);
        sheet.cell(lastRow, 3).setFormula(String.format("=SUM(C2:C%d)", lastRow-1));
        sheet.cell(lastRow, 3).setNumberFormat("$#,##0.00");
        sheet.cell(lastRow, 3).setFontStyle(true, false, false);
        
        // Format data section
        sheet.cell(1, 1).setFullBorder("black");
        sheet.cell(lastRow, 5).setFullBorder("black");
    }
    
    private static void createChartsSheet(ExcelWorkSheet sheet, ExcelWorkSheet data) {
        // Title
        sheet.cell(1, 1).setText("Sales Analysis");
        sheet.cell(1, 1).setFontStyle(true, false, false);
        
        // Add chart (Note: This is a placeholder as chart creation would require additional implementation)
        sheet.cell(3, 1).setText("Sales by Product");
        sheet.cell(3, 1).setFontStyle(true, false, false);
        // Chart implementation would go here
        
        sheet.cell(3, 4).setText("Growth Analysis");
        sheet.cell(3, 4).setFontStyle(true, false, false);
        // Chart implementation would go here
    }
    
    private static class Product {
        final String name;
        final double currentSales;
        final double previousSales;
        final String category;
        
        Product(String name, double currentSales, double previousSales, String category) {
            this.name = name;
            this.currentSales = currentSales;
            this.previousSales = previousSales;
            this.category = category;
        }
    }
}
