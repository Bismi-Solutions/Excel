# Excel

<div align="center">
  <em>A powerful and easy-to-use Java library for Excel file manipulation</em>
  <br><br>

  [![CI & Release](https://github.com/Bismi-Solutions/Excel/actions/workflows/ci.yml/badge.svg)](https://github.com/Bismi-Solutions/Excel/actions/workflows/ci.yml)
  [![codecov](https://codecov.io/gh/Bismi-Solutions/Excel/branch/main/graph/badge.svg)](https://codecov.io/gh/Bismi-Solutions/Excel)
  [![Quality Gate Status](https://sonarcloud.io/api/project_badges/measure?project=Bismi-Solutions_Excel&metric=alert_status)](https://sonarcloud.io/project/overview?id=Bismi-Solutions_Excel)
  [![Known Vulnerabilities](https://snyk.io/test/github/Bismi-Solutions/Excel/badge.svg?targetFile=pom.xml)](https://snyk.io/test/github/Bismi-Solutions/Excel?targetFile=pom.xml)
  [![Maven Central](https://img.shields.io/maven-central/v/solutions.bismi.excel/excel.svg)](https://search.maven.org/artifact/solutions.bismi.excel/excel)
  [![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
  [![Java Version](https://img.shields.io/badge/Java-17%2B-blue)](https://openjdk.java.net/)
</div>

## Table of Contents
- [Overview](#overview)
- [Features](#features)
- [Requirements](#requirements)
- [Installation](#installation)
- [Quick Start](#quick-start)
- [Usage Examples](#usage-examples)
  - [Basic Workbook Operations](#basic-workbook-operations)
  - [Cell Operations](#cell-operations)
  - [Row Operations](#row-operations)
  - [Cell Merging](#cell-merging)
- [Advanced Example: Creating a Sales Report](#advanced-example-creating-a-sales-report)
- [Comparison with Other Libraries](#comparison-with-other-libraries)
- [Library Architecture](#library-architecture)
- [Contributing](#contributing)
- [Documentation](#documentation)
- [License](#license)

## Overview

This library provides a simplified API similar to Microsoft Excel COM model, making it easy to create, read, and modify Excel files. It internally uses Apache POI for file processing but offers a much more intuitive interface.

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

## Requirements

- Java 17 or higher
- Maven 3.6 or higher (for building)

## Installation

Add the following Maven dependency to your project:

```xml
<dependency>
    <groupId>solutions.bismi.excel</groupId>
    <artifactId>excel</artifactId>
    <version>1.1.5</version>
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

// Activate the sheet
sheet.activate();

// Save the workbook
sheet.saveWorkBook();

// Close all workbooks
xlApp.closeAllWorkBooks();
```

### Cell Operations

```java
// Set cell values
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
```

### Row Operations

```java
// Set row values
String[] rowData = {"A", "B", "C", "D", "E"};
sheet.row(5).setRowValues(rowData);

// Apply formatting to row
sheet.row(5).setFontColor("RED", 1, 3);  // Apply to columns 1-3
sheet.row(5).setFillColor("GREEN");      // Apply to entire row
sheet.row(5).setFullBorder("BLUE");      // Apply to entire row
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

## Advanced Example: Creating a Sales Report

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

// Add data
sheet.cell(4, 1).setCellValue("Product A");
sheet.cell(4, 2).setNumericValue(10);
sheet.cell(4, 3).setNumericValue(25.50);
sheet.cell(4, 4).setFormula("B4*C4");

sheet.cell(5, 1).setCellValue("Product B");
sheet.cell(5, 2).setNumericValue(5);
sheet.cell(5, 3).setNumericValue(34.99);
sheet.cell(5, 4).setFormula("B5*C5");

// Add total
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

## Comparison with Other Libraries

| Feature | Excel Library | Apache POI | JExcel |
|---------|--------------|------------|--------|
| API Simplicity | ★★★★★ | ★★★ | ★★★★ |
| Excel Features Support | ★★★★ | ★★★★★ | ★★★ |
| Performance | ★★★★ | ★★★★ | ★★★★ |
| Learning Curve | Easy | Steep | Moderate |
| Active Development | Yes | Yes | Limited |

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
