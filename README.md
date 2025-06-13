# Excel

<div align="center">
  <em>A powerful and easy-to-use Java library for Excel file manipulation</em>
  <br><br>

  [![CI & Release](https://github.com/Bismi-Solutions/Excel/actions/workflows/ci.yml/badge.svg)](https://github.com/Bismi-Solutions/Excel/actions/workflows/ci.yml)
  [![codecov](https://codecov.io/gh/Bismi-Solutions/Excel/branch/master/graph/badge.svg)](https://codecov.io/gh/Bismi-Solutions/Excel)
  [![Quality Gate Status](https://sonarcloud.io/api/project_badges/measure?project=Bismi-Solutions_Excel&metric=alert_status)](https://sonarcloud.io/project/overview?id=Bismi-Solutions_Excel)
  [![Known Vulnerabilities](https://snyk.io/test/github/Bismi-Solutions/Excel/badge.svg?targetFile=pom.xml)](https://snyk.io/test/github/Bismi-Solutions/Excel?targetFile=pom.xml)
  [![Maven Central](https://img.shields.io/maven-central/v/solutions.bismi.excel/excel.svg)](https://search.maven.org/artifact/solutions.bismi.excel/excel)
  [![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
  [![Java Version](https://img.shields.io/badge/Java-17%2B-blue)](https://openjdk.java.net/)
</div>

## Table of Contents
- [Overview](#overview)
- [Features](#features)
- [Benefits](#benefits)
- [Requirements](#requirements)
- [Platform Considerations](#platform-considerations)
  - [macOS Path Permissions](#macos-path-permissions)
  - [Cross-Platform Path Handling](#cross-platform-path-handling)
  - [Relative File Paths](#relative-file-paths)
- [Installation](#installation)
- [Quick Start](#quick-start)
- [Usage Examples](#usage-examples)
  - [Basic Workbook Operations](#basic-workbook-operations)
  - [Cell Operations](#cell-operations)
  - [Row Operations](#row-operations)
  - [Cell Merging](#cell-merging)
- [Library Demonstration](#library-demonstration)
  - [Running the Demo](#running-the-demo)
  - [Demo Java Code](#demo-java-code)
  - [Expected Output Screenshots](#expected-output-screenshots)
  - [Example Workflow GIF](#example-workflow-gif)
- [Advanced Example: Creating a Sales Report](#advanced-example-creating-a-sales-report)
- [Comparison with Other Libraries](#comparison-with-other-libraries)
- [Library Architecture](#library-architecture)
- [Contributing](#contributing)
- [Documentation](#documentation)
- [License](#license)

## Overview

This library provides a simplified API similar to Microsoft Excel COM model, making it easy to create, read, and modify Excel files. It internally uses Apache POI for file processing but offers a much more intuitive interface. Whether you're building reports, data processing applications, or need to automate Excel operations, this library provides a robust solution with minimal learning curve.

## Features

<div align="center">
  <table>
    <tr>
      <td align="center"><strong>📊</strong><br>Create, open, and save Excel workbooks<br>(.xlsx and .xls formats)</td>
      <td align="center"><strong>📑</strong><br>Add, access, and modify<br>worksheets</td>
      <td align="center"><strong>📝</strong><br>Set cell values<br>(text, numbers, dates, formulas)</td>
    </tr>
    <tr>
      <td align="center"><strong>🎨</strong><br>Apply cell formatting<br>(colors, borders, alignment, number formats)</td>
      <td align="center"><strong>🔄</strong><br>Merge and<br>unmerge cells</td>
      <td align="center"><strong>📋</strong><br>Row operations<br>(set values, apply formatting)</td>
    </tr>
  </table>
</div>

### Key Features in Detail

- **Workbook Management**
  - Create new workbooks
  - Open existing Excel files
  - Save workbooks in multiple formats
  - Handle multiple workbooks simultaneously

- **Worksheet Operations**
  - Add, rename, and delete sheets
  - Navigate between sheets
  - Copy and move sheets
  - Sheet protection and visibility control

- **Cell Manipulation**
  - Set various data types (text, numbers, dates, formulas)
  - Apply rich formatting (fonts, colors, borders)
  - Cell merging and splitting
  - Conditional formatting support

- **Row and Column Operations**
  - Insert and delete rows/columns
  - Set row heights and column widths
  - Apply formatting to entire rows/columns
  - Auto-fit functionality

## Benefits

1. **Developer-Friendly API**
   - Intuitive method names and parameters
   - Fluent interface design
   - Comprehensive error handling
   - Extensive documentation

2. **Performance Optimized**
   - Efficient memory management
   - Batch operations support
   - Optimized for large datasets
   - Minimal overhead

3. **Enterprise Ready**
   - Thread-safe operations
   - Robust error handling
   - Comprehensive logging
   - Production-grade reliability

4. **Maintenance Benefits**
   - Clean, maintainable code
   - Regular updates and bug fixes
   - Active community support
   - Extensive test coverage

## Requirements

- Java 17 or higher
- Maven 3.6 or higher (for building)

## Platform Considerations

### macOS Path Permissions
On macOS (and other Unix-like systems), ensure that the application has the necessary write permissions for the directory where you intend to create or save Excel files. If you encounter issues saving files, please verify the permissions of the target path.

### Cross-Platform Path Handling
The library handles file paths in a cross-platform manner using standard Java APIs like `java.nio.file.Paths` and `File.separator` to ensure consistent behavior on Windows, macOS, and Linux.

### Relative File Paths

The library supports the use of relative file paths for creating and opening workbooks. For example, you can use paths like:
- `"my_workbook.xlsx"` (resolved in the current working directory)
- `"./output_files/report.xlsx"` (resolved relative to the current working directory)
- `"../archived_reports/old_data.xls"` (resolved relative to the current working directory)

These paths are resolved based on the application's current working directory when the library functions are called. If a path includes subdirectories that do not yet exist (e.g., specifying `"data/new_folder/my_report.xlsx"` when `data/new_folder` is not present), the library will attempt to create these directories.

## Installation

Add the following Maven dependency to your project:

```xml
<dependency>
    <groupId>solutions.bismi.excel</groupId>
    <artifactId>excel</artifactId>
    <version>1.1.11</version>
</dependency>
```

## Quick Start

Get up and running with Excel library in minutes:

1. Add the Maven dependency to your project
2. Create an Excel application instance
3. Create or open a workbook
4. Manipulate sheets and cells
5. Save and close

```java
// Quick example: Create a simple Excel file with formatted data
ExcelApplication app = new ExcelApplication();
ExcelWorkBook workbook = app.createWorkBook("quickstart.xlsx");
ExcelWorkSheet sheet = workbook.addSheet("Report");

// Add a title
sheet.cell(1, 1).setCellValue("Sales Report");
sheet.cell(1, 1).setFontStyle(true, false, false);
sheet.cell(1, 1).setFillColor("LIGHT_BLUE");

// Add some data
sheet.cell(3, 1).setCellValue("Product");
sheet.cell(3, 2).setCellValue("Quantity");
sheet.cell(3, 3).setCellValue("Price");

sheet.cell(4, 1).setCellValue("Widget A");
sheet.cell(4, 2).setNumericValue(5);
sheet.cell(4, 3).setNumericValue(19.99);

sheet.cell(5, 1).setCellValue("Widget B");
sheet.cell(5, 2).setNumericValue(3);
sheet.cell(5, 3).setNumericValue(29.99);

// Save and close
sheet.saveWorkBook();
app.closeAllWorkBooks();
```

## Usage Examples

### Basic Workbook Operations

```java
// Create a new Excel application
ExcelApplication xlApp = new ExcelApplication();

// Create a new workbook
ExcelWorkBook xlbook = xlApp.createWorkBook("path/to/file.xlsx");

// Get sheet count
int sheetCount = xlbook.getSheetCount();

// Add a new sheet
ExcelWorkSheet sheet = xlbook.addSheet("MySheet");

// Add multiple sheets at once
String[] sheetNames = {"Sheet1", "Sheet2", "Sheet3"};
xlbook.addSheets(sheetNames);

// Activate the sheet
sheet.activate();

// Save the workbook
sheet.saveWorkBook();

// Close all workbooks
xlApp.closeAllWorkBooks();
```

### Cell Operations

```java
// Set cell values with different data types
sheet.cell(1, 1).setCellValue("Hello World");
sheet.cell(2, 1).setNumericValue(123.45);
sheet.cell(3, 1).setDateValue(new java.util.Date());

// Set formulas
sheet.cell(4, 1).setFormula("SUM(A2:A3)");

// Apply formatting
sheet.cell(1, 1).setFontColor("BLUE");
sheet.cell(1, 1).setFillColor("YELLOW");
sheet.cell(1, 1).setFullBorder("RED");

// Set text alignment
sheet.cell(1, 1).setHorizontalAlignment("CENTER");
sheet.cell(1, 1).setVerticalAlignment("CENTER");

// Set number format
sheet.cell(2, 1).setNumberFormat("#,##0.00");

// Set font style (bold, italic, underline)
sheet.cell(1, 1).setFontStyle(true, false, false);

// Use hex color codes
sheet.cell(5, 5).setFontColor("#FF0000");  // Red
sheet.cell(6, 6).setFillColor("#00FF00");  // Green
```

### Row Operations

```java
// Set row values
String[] rowData = {"A", "B", "C", "D", "E"};
sheet.row(5).setRowValues(rowData);

// Set row values with column offset
sheet.row(6).setRowValues(rowData, 2);  // Start from column 2

// Set row values with auto-size
sheet.row(7).setRowValues(rowData, true);

// Apply formatting to row
sheet.row(5).setFontColor("RED", 1, 3);  // Apply to columns 1-3
sheet.row(5).setFillColor("GREEN");      // Apply to entire row
sheet.row(5).setFullBorder("BLUE");      // Apply to entire row

// Format specific columns in a row
sheet.row(8).setFillColor("YELLOW", 2, 4);  // Fill columns 2-4
```

### Cell Merging

```java
// Merge cells
sheet.mergeCells(1, 3, 1, 5);  // Merge cells from A1 to E3

// Check if a cell is merged
boolean isMerged = sheet.isCellMerged(2, 3);

// Unmerge cells
sheet.unmergeCells(1, 3, 1, 5);
```

## Library Demonstration

### Running the Demo
To run the demo, compile and execute the `solutions.bismi.excel.ExcelLibraryDemo` class (typically found in `src/test/java`). It will generate a file named `demo_workbook.xlsx` in the project's root directory. You can modify the output path within the demo code if needed.

### Demo Java Code
```java
package solutions.bismi.excel;

import java.io.File;
import java.util.Date;

public class ExcelLibraryDemo {

    public static void main(String[] args) {
        ExcelApplication xlApp = new ExcelApplication();
        String outputFileName = "demo_workbook.xlsx";
        String outputFilePath = System.getProperty("user.dir") + File.separator + outputFileName;

        System.out.println("Attempting to create demo workbook at: " + outputFilePath);

        ExcelWorkBook xlBook = xlApp.createWorkBook(outputFilePath);

        if (xlBook == null || xlBook.getWb() == null) {
            System.err.println("Failed to create the demo workbook. Please check permissions and path.");
            return;
        }

        // Sheet 1: Data Showcase
        try {
            ExcelWorkSheet sheet1 = xlBook.addSheet("Data Showcase");
            if (sheet1 == null) {
                System.err.println("Failed to create sheet 'Data Showcase'");
            } else {
                sheet1.activate();

                // Title
                sheet1.cell(1, 1).setText("Data Showcase");
                sheet1.cell(1, 1).setFontStyle(true, false, false).setFontColor("DARK_BLUE"); // Bold, Dark Blue
                sheet1.cell(1,1).getCell().getCellStyle().getFont().setFontHeightInPoints((short)16); // Example direct POI for font size

                // Text examples
                sheet1.cell(3, 1).setText("Simple Text:");
                sheet1.cell(3, 2).setText("Hello, Excel World!");
                sheet1.cell(3, 2).setFontColor("GREEN");

                sheet1.cell(4, 1).setText("Styled Text:");
                sheet1.cell(4, 2).setText("Bold & Italic");
                sheet1.cell(4, 2).setFontStyle(true, true, false).setFontColor("RED");

                // Number examples
                sheet1.cell(6, 1).setText("Integer:");
                sheet1.cell(6, 2).setNumericValue(12345);
                sheet1.cell(6, 2).setFillColor("LIGHT_YELLOW");

                sheet1.cell(7, 1).setText("Decimal:");
                sheet1.cell(7, 2).setNumericValue(678.90);
                sheet1.cell(7, 2).setFillColor("LIGHT_GREEN");

                // Date example
                sheet1.cell(9, 1).setText("Current Date:");
                sheet1.cell(9, 2).setDateValue(new Date());
                sheet1.cell(9, 2).setNumberFormat("dd-mmm-yyyy hh:mm AM/PM"); // Date format

                // Alignment and Wrapping
                sheet1.cell(11, 1).setText("Alignment & Wrap Test:");
                ExcelCell cellForWrap = sheet1.cell(11, 2);
                cellForWrap.setText("This is a long piece of text that should wrap within the cell to demonstrate text wrapping capabilities of the library.");
                // Text wrapping needs to be set on CellStyle directly via POI for now
                org.apache.poi.ss.usermodel.CellStyle wrapStyle = cellForWrap.getCell().getCellStyle();
                if (wrapStyle == null) wrapStyle = xlBook.getWb().createCellStyle();
                wrapStyle.setWrapText(true);
                cellForWrap.getCell().setCellStyle(wrapStyle);
                // Adjust row height manually for wrapped text to be visible
                sheet1.getSheet().getRow(10).setHeightInPoints((short)(3 * sheet1.getSheet().getDefaultRowHeightInPoints()));


                sheet1.cell(12, 1).setText("Top-Left");
                sheet1.cell(12, 1).setHorizontalAlignment("LEFT").setVerticalAlignment("TOP").setFillColor("LIGHT_CORNFLOWER_BLUE");
                sheet1.cell(12, 2).setText("Center-Center");
                sheet1.cell(12, 2).setHorizontalAlignment("CENTER").setVerticalAlignment("CENTER").setFillColor("LIGHT_CORNFLOWER_BLUE");
                sheet1.cell(12, 3).setText("Bottom-Right");
                sheet1.cell(12, 3).setHorizontalAlignment("RIGHT").setVerticalAlignment("BOTTOM").setFillColor("LIGHT_CORNFLOWER_BLUE");

                // Column width for wrapped text
                sheet1.getSheet().setColumnWidth(1, 25 * 256); // Approx 25 characters width for column B (index 1)

                System.out.println("'Data Showcase' sheet created.");
            }
        } catch (Exception e) {
            System.err.println("Error creating 'Data Showcase' sheet: " + e.getMessage());
            e.printStackTrace();
        }

        // Sheet 2: Formulas and Structures
        try {
            ExcelWorkSheet sheet2 = xlBook.addSheet("Formulas and Structures");
            if (sheet2 == null) {
                System.err.println("Failed to create sheet 'Formulas and Structures'");
            } else {
                sheet2.activate();

                // Title
                sheet2.cell(1, 1).setText("Sales Report Q1");
                sheet2.cell(1, 1).setFontStyle(true, false, false).setFontColor("DARK_RED");
                sheet2.cell(1,1).getCell().getCellStyle().getFont().setFontHeightInPoints((short)14);
                sheet2.mergeCells(1, 1, 1, 4); // Merge R1C1 to R1C4
                sheet2.cell(1,1).setHorizontalAlignment("CENTER");


                // Data table
                sheet2.cell(3, 1).setText("Month").setFontStyle(true, false, false);
                sheet2.cell(3, 2).setText("Product A Sales").setFontStyle(true, false, false);
                sheet2.cell(3, 3).setText("Product B Sales").setFontStyle(true, false, false);
                sheet2.cell(3, 4).setText("Monthly Total").setFontStyle(true, false, false);

                sheet2.cell(4, 1).setText("January");
                sheet2.cell(4, 2).setNumericValue(1500);
                sheet2.cell(4, 3).setNumericValue(1200);
                sheet2.cell(4, 4).setFormula("B4+C4");

                sheet2.cell(5, 1).setText("February");
                sheet2.cell(5, 2).setNumericValue(1800);
                sheet2.cell(5, 3).setNumericValue(1100);
                sheet2.cell(5, 4).setFormula("B5+C5");

                sheet2.cell(6, 1).setText("March");
                sheet2.cell(6, 2).setNumericValue(2100);
                sheet2.cell(6, 3).setNumericValue(1350);
                sheet2.cell(6, 4).setFormula("B6+C6");

                // Totals and Average
                sheet2.cell(8, 1).setText("Total Sales").setFontStyle(true, false, false);
                sheet2.cell(8, 2).setFormula("SUM(B4:B6)");
                sheet2.cell(8, 3).setFormula("SUM(C4:C6)");
                sheet2.cell(8, 4).setFormula("SUM(D4:D6)");

                sheet2.cell(9, 1).setText("Average Sales").setFontStyle(true, false, false);
                sheet2.cell(9, 2).setFormula("AVERAGE(B4:B6)");
                sheet2.cell(9, 3).setFormula("AVERAGE(C4:C6)");

                // Apply borders
                for(int r=3; r<=9; r++) {
                    for(int c=1; c<=4; c++) {
                        if (sheet2.cell(r,c).getTextValue() != null && !sheet2.cell(r,c).getTextValue().isEmpty()) {
                             sheet2.cell(r,c).setFullBorder("GREY_50_PERCENT");
                        } else if (sheet2.cell(r,c).getValue() != null && !sheet2.cell(r,c).getValue().isEmpty()){
                             sheet2.cell(r,c).setFullBorder("GREY_50_PERCENT");
                        }
                    }
                }
                 System.out.println("'Formulas and Structures' sheet created.");
            }
        } catch (Exception e) {
            System.err.println("Error creating 'Formulas and Structures' sheet: " + e.getMessage());
            e.printStackTrace();
        }

        // Sheet 3: Advanced Formatting
        try {
            ExcelWorkSheet sheet3 = xlBook.addSheet("Advanced Formatting");
            if (sheet3 == null) {
                System.err.println("Failed to create sheet 'Advanced Formatting'");
            } else {
                sheet3.activate();

                sheet3.cell(1,1).setText("Formatting Showcase").setFontStyle(true, false, false);

                sheet3.cell(3, 1).setText("Currency (USD):");
                sheet3.cell(3, 2).setNumericValue(12345.67);
                sheet3.cell(3, 2).setNumberFormat("$#,##0.00");

                sheet3.cell(4, 1).setText("Currency (EUR):");
                sheet3.cell(4, 2).setNumericValue(12345.67);
                sheet3.cell(4, 2).setNumberFormat("€#,##0.00");


                sheet3.cell(6, 1).setText("Percentage:");
                sheet3.cell(6, 2).setNumericValue(0.8567);
                sheet3.cell(6, 2).setNumberFormat("0.00%");

                sheet3.cell(8, 1).setText("Scientific Notation:");
                sheet3.cell(8, 2).setNumericValue(1234567890.123);
                sheet3.cell(8, 2).setNumberFormat("0.00E+00");

                sheet3.cell(10, 1).setText("Custom Date:");
                sheet3.cell(10, 2).setDateValue(new Date());
                sheet3.cell(10, 2).setNumberFormat("ddd, mmm d, yyyy");

                // Column Widths
                sheet3.cell(12, 1).setText("Wide Column (A)");
                sheet3.getSheet().setColumnWidth(0, 30 * 256); // Column A (index 0) width 30 chars

                sheet3.cell(12, 2).setText("Narrow Column (B)");
                 sheet3.getSheet().setColumnWidth(1, 10 * 256); // Column B (index 1) width 10 chars

                // Text Rotation - requires direct POI manipulation if not in ExcelCell API
                ExcelCell rotationCell = sheet3.cell(14, 1);
                rotationCell.setText("Rotated Text");
                org.apache.poi.ss.usermodel.CellStyle rotationStyle = rotationCell.getCell().getCellStyle();
                if (rotationStyle == null) rotationStyle = xlBook.getWb().createCellStyle();
                rotationStyle.setRotation((short)45); // 45 degrees
                rotationCell.getCell().setCellStyle(rotationStyle);

                System.out.println("'Advanced Formatting' sheet created.");
            }
        } catch (Exception e) {
            System.err.println("Error creating 'Advanced Formatting' sheet: " + e.getMessage());
            e.printStackTrace();
        }

        // Save and Close
        if (xlBook.saveWorkbook()) {
            System.out.println("Demo workbook saved successfully to: " + outputFilePath);
        } else {
            System.err.println("Failed to save demo workbook to: " + outputFilePath);
        }
        xlApp.closeAllWorkBooks();
        System.out.println("All workbooks closed.");
    }
}
```

### Expected Output Screenshots

**Sheet 1: Data Showcase**
`[Image: Screenshot of Sheet 1 - Data Showcase]`
This sheet demonstrates various data types (text, numbers, dates), font styling (colors, bold, italic), cell background fills, text wrapping, and alignments.

**Sheet 2: Formulas and Structures**
`[Image: Screenshot of Sheet 2 - Formulas and Structures]`
This sheet showcases the use of formulas (SUM, AVERAGE), cell merging for titles, and different border styles around a data table.

**Sheet 3: Advanced Formatting**
`[Image: Screenshot of Sheet 3 - Advanced Formatting]`
This sheet highlights advanced number formatting (currency, percentage, scientific notation), custom column widths, and text rotation.

### Example Workflow GIF
`[GIF: Library Workflow Demo]`
A brief animation demonstrating the creation of a simple Excel file, adding data and basic formatting to a few cells, and saving the file using the library.


## Advanced Example: Creating a Sales Report with Formulas

```java
ExcelApplication xlApp = new ExcelApplication();
ExcelWorkBook xlbook = xlApp.createWorkBook("SalesReport.xlsx");
ExcelWorkSheet sheet = xlbook.addSheet("Sales");
sheet.activate();

// Add title
sheet.cell(1, 1).setCellValue("Sales Report");
sheet.cell(1, 1).setFontStyle(true, false, false);
sheet.cell(1, 1).setHorizontalAlignment("CENTER");
sheet.mergeCells(1, 1, 1, 5);

// Add headers
String[] headers = {"Item", "Quantity", "Price", "Total"};
for (int i = 0; i < headers.length; i++) {
    sheet.cell(3, i+1).setCellValue(headers[i]);
    sheet.cell(3, i+1).setFontStyle(true, false, false);
    sheet.cell(3, i+1).setFillColor("GREY_25_PERCENT");
}

// Add data with formulas
sheet.cell(4, 1).setCellValue("Product A");
sheet.cell(4, 2).setNumericValue(10);
sheet.cell(4, 3).setNumericValue(25.50);
sheet.cell(4, 4).setFormula("B4*C4");

sheet.cell(5, 1).setCellValue("Product B");
sheet.cell(5, 2).setNumericValue(5);
sheet.cell(5, 3).setNumericValue(34.99);
sheet.cell(5, 4).setFormula("B5*C5");

// Add total with formula
sheet.cell(6, 1).setCellValue("Total");
sheet.cell(6, 4).setFormula("SUM(D4:D5)");
sheet.cell(6, 4).setFontStyle(true, false, false);

// Format numbers
for (int row = 4; row <= 6; row++) {
    sheet.cell(row, 3).setNumberFormat("$#,##0.00");
    sheet.cell(row, 4).setNumberFormat("$#,##0.00");
}

sheet.saveWorkBook();
xlApp.closeAllWorkBooks();
```

### Error Handling Example

```java
ExcelApplication xlApp = new ExcelApplication();
ExcelWorkBook xlbook = xlApp.createWorkBook("ErrorHandling.xlsx");
ExcelWorkSheet sheet = xlbook.addSheet("ErrorTest");

try {
    // Test with invalid formula
    sheet.cell(1, 1).setFormula("INVALID_FORMULA(A1)");
    
    // Test with invalid color name
    sheet.cell(2, 1).setFontColor("NONEXISTENT_COLOR");
    
    // Test with invalid cell indices
    sheet.mergeCells(-1, 1, 1, 2);
    
    // Test with empty cell
    String value = sheet.cell(3, 3).getValue();
    // Empty cell returns empty string
    
} catch (Exception e) {
    // Handle exceptions appropriately
    System.err.println("Error: " + e.getMessage());
} finally {
    sheet.saveWorkBook();
    xlApp.closeAllWorkBooks();
}
```

## Comparison with Other Libraries

| Feature | Excel Library | Apache POI | JExcel |
|---------|--------------|------------|--------|
| API Simplicity | ★★★★★ | ★★★ | ★★★★ |
| Excel Features Support | ★★★★ | ★★★★★ | ★★★ |
| Performance | ★★★★ | ★★★★ | ★★★★ |
| Learning Curve | Easy | Steep | Moderate |
| Active Development | Yes | Yes | Limited |
| Memory Efficiency | ★★★★ | ★★★ | ★★★★ |
| Documentation | ★★★★★ | ★★★ | ★★★ |
| Community Support | ★★★★ | ★★★★★ | ★★★ |

## Library Architecture

The Excel library follows a simple object hierarchy that mimics the Microsoft Excel object model:

```
ExcelApplication
    └── ExcelWorkBook
        └── ExcelWorkSheet
            ├── ExcelRow
            └── ExcelCell
```

## Contributing

Contributions are welcome! Here's how you can contribute:

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add some amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

Please make sure to update tests as appropriate.

## Documentation

For more detailed information about the API, please refer to the Javadoc documentation generated during the build process.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
