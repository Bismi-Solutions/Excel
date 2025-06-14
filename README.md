# Excel 📈

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

## 🚀  Why *Excel*?

Apache POI is wonderfully complete—but you pay in verbosity. *Excel* wraps POI with a **fluent, COM‑style API** so you can:

- Build a workbook in **two lines**:
  ```java
  ExcelApplication app = new ExcelApplication();
  ExcelWorkBook wb = app.createWorkBook("hello.xlsx");
  ExcelWorkSheet sh = wb.addSheet("Hi");
  sh.cell(1,1).setText("👋");
  sh.saveWorkBook();
  app.closeAllWorkBooks();
  ```
- Use named or hex colours, borders, number‑formats, merges & formulas without memorising POI constants.
- Keep full access to the underlying POI objects when you need edge‑case power.

Works unchanged on Windows 🪟, macOS 🍎 and Linux 🐧 (Java 17 +).

---

## 📦 Installation (copy & go)

Choose the snippet for your build tool and you're done.

### Maven

```xml
<dependency>
  <groupId>solutions.bismi.excel</groupId>
  <artifactId>excel</artifactId>
  <version>1.1.12</version>
</dependency>
```

### Gradle (Groovy DSL)

```groovy
dependencies {
  implementation "solutions.bismi.excel:excel:1.1.12"
}
```

### Gradle (Kotlin DSL)

```kotlin
dependencies {
  implementation("solutions.bismi.excel:excel:1.1.12")
}
```

### Scala

```scala
libraryDependencies += "solutions.bismi.excel" % "excel" % "1.1.12"
```

---

## 📑 Index

- [Why *Excel*?](#🚀--why-excel)
- [Features ☑️](#features-☑️)
- [Quick Start](#quick-start-)
- [Usage Snippets](#usage-snippets)
  - [Workbook Basics](#workbook-basics)
  - [Cell Goodies](#cell-goodies)
  - [Rows & Columns](#rows--columns)
  - [Merging Cells](#merging-cells)
- [Advanced Demo 🛠️](#advanced-demo-)
- [Architecture](#architecture)
- [Comparison 🆚 POI / JExcel](#comparison--🆚-poi--jexcel)
- [Contributing 🤝](#contributing-)
- [License](#license)

---

## Features ☑️

|                      |                                                                     |
| -------------------- | ------------------------------------------------------------------- |
| 📑 **Workbook**      | create, open, save, multi‑format (.xlsx / .xls)                     |
| 📄 **Sheets**        | add / rename / activate / protect                                   |
| 📝 **Cells**         | text, numbers, dates, formulas, conditional formatting              |
| 🎨 **Styling**       | fonts, colours (named ✨ or hex), borders, alignment, number formats |
| 📋 **Rows/Cols**     | bulk value setting, auto‑fit, insert/delete                         |
| 🔗 **Merge/Unmerge** | succinct helpers & intersection safety checks                       |

Fast: streams large datasets (100 k+ rows) & reuses POI styles to keep memory low.

---

## Quick Start 🎉

```java
ExcelApplication app = new ExcelApplication();
ExcelWorkBook wb = app.createWorkBook("demo.xlsx");
ExcelWorkSheet sh = wb.addSheet("Summary");
sh.activate();
sh.cell(1,1).setText("Hello World");
sh.cell(1,1).setFontStyle(true, false, false);
sh.cell(1,1).setFillColor("LIGHT_GREEN");
wb.saveWorkbook();
app.closeAllWorkBooks();
```

Runs on Java 17+; produces a styled workbook in seconds.

---

## Usage Snippets

### Workbook Basics

```java
ExcelApplication app = new ExcelApplication();
ExcelWorkBook wb = app.openWorkbook("existing.xls");
wb.addSheet("2025-Q2");
System.out.println(wb.getSheetCount());
```

### Cell Goodies

```java
ExcelCell c = wb.getActiveSheet().cell(2,3);
c.setNumericValue(42);
c.setNumberFormat("#,##0.00");
c.setFontColor("#FF9900");
c.setHorizontalAlignment("RIGHT");
```

### Rows & Columns

```java
String[] header = {"Item","Qty","Price","Total"};
ExcelRow row = wb.getActiveSheet().row(5);
row.setRowValues(header);
row.setFontColor("white");
row.setFillColor("grey_50_percent");
row.setFullBorder("black");
```

### Merging Cells

```java
ExcelWorkSheet s = wb.getActiveSheet();
s.mergeCells(1,1,1,5);                 // A1:E1

// later…
if (s.isCellMerged(1,3)) s.unmergeCells(1,1,1,5);
```

---

## Advanced Demo 🛠️

*See **`examples/SalesReport.java`** — multi‑sheet, fully‑formatted workbook with charts in under 80 LOC.*

---

## Architecture

```
ExcelApplication
 └─ ExcelWorkBook
     └─ ExcelWorkSheet
         ├─ ExcelRow
         └─ ExcelCell
```

Each object only exposes context‑appropriate methods, making code completion your best documentation.

---

## Comparison 🆚 POI / JExcel

|  Capability              |  Excel (lib) | Apache POI | JExcel API    |
| ------------------------- | ------------ | ---------- | ------------- |
| Fluent API                | ⭐⭐⭐⭐⭐        | ⭐⭐         | ⭐⭐⭐           |
| Large‑data memory profile | ⭐⭐⭐⭐         | ⭐⭐⭐        | ⭐⭐⭐⭐          |
| XLSX support              | ✅            | ✅          | ⚠ (read‑only) |
| Styling shorthand         | Yes          | No         | Limited       |
| Active development        | ✅            | ✅          | ❌             |

---

## Contributing 🤝

```bash
git clone https://github.com/Bismi-Solutions/Excel.git
cd Excel
mvn test   # green? great – pick an issue and hack away!
```

Pull requests are welcome. Please include unit tests and follow the log‑level convention:

- **info** → user‑visible events (file created, sheet saved)
- **debug** → flow diagnostics
- **warn/error** → exceptional situations

---

## License

MIT – *use it, fork it, profit.*  See [LICENSE](LICENSE).
