# Excel 📈

<div align="center">

<h3><em>Ship a styled Excel report in 5 lines — skip the 60 lines of Apache POI boilerplate.</em></h3>

[![CI & Release](https://github.com/Bismi-Solutions/Excel/actions/workflows/ci.yml/badge.svg)](https://github.com/Bismi-Solutions/Excel/actions/workflows/ci.yml)
[![codecov](https://codecov.io/gh/Bismi-Solutions/Excel/branch/master/graph/badge.svg)](https://codecov.io/gh/Bismi-Solutions/Excel)
[![Quality Gate Status](https://sonarcloud.io/api/project_badges/measure?project=Bismi-Solutions_Excel&metric=alert_status)](https://sonarcloud.io/project/overview?id=Bismi-Solutions_Excel)
[![OpenSSF Scorecard](https://api.securityscorecards.dev/projects/github.com/Bismi-Solutions/Excel/badge)](https://scorecard.dev/viewer/?uri=github.com/Bismi-Solutions/Excel)
[![Maven Central](https://img.shields.io/maven-central/v/solutions.bismi.excel/excel.svg)](https://search.maven.org/artifact/solutions.bismi.excel/excel)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Java Version](https://img.shields.io/badge/Java-17%2B-blue)](https://openjdk.java.net/)

</div>

```java
ReportBuilder.on(sheet)
    .title("Q3 Sales Report")
    .rowsFromBeans(products)      // ← your List<Bean>
    .zebraStripes(true).freezeHeader(true).autoFilter(true)
    .render();
```

**Result:**

<p align="center">
  <img src="docs/report-preview.svg" alt="Styled Excel report generated in 5 lines" width="100%"/>
</p>

> The workbook shown above is produced by [`examples/SalesReport.java`](examples/SalesReport.java) — beans + `ReportBuilder`, plus the new `ExcelStyle.title()` / `totals()` / `statusActive|Review|Closed()` presets for the navy title, the pale-blue totals row and the coloured status pills.

---

## 🤔 Why *Excel*, not Apache POI directly?

POI is powerful — and verbose. Every styled cell forces you through the same choreography: `CreationHelper` → `CellStyle.cloneStyleFrom` → `DataFormat` → `Font` → manual type dispatch (`setCellValue(String)` vs `setCellValue(double)` vs `setCellValue(Date)`) → `CellRangeAddress` → `createFreezePane` → `setAutoFilter`.

Most business-Excel tasks boil down to **"take a list of objects and make it look like a report."** *Excel* gives you **one call** for that, while still exposing the POI workbook for edge cases.

### What you skip

| You no longer juggle | Because *Excel* handles it |
|---|---|
| Creating/cloning `CellStyle` for every cell | `ExcelStyle` — build once, apply everywhere |
| Hitting POI's **64K-CellStyle quota** in big reports | Reused immutable styles by design |
| Casting between `HSSFWorkbook` / `XSSFWorkbook` for hex colours | Auto-detects format; hex falls back to nearest indexed on `.xls` |
| `FileInputStream` / `FileOutputStream` lifecycle | Opened and closed internally |
| `setCellValue` type dispatch | `setValue(Object)` — accepts `String`, `Number`, `Boolean`, `Date`, `LocalDate`, `null` |
| 0-based POI indexes | 1-based public API (matches Excel UI) |
| Bean-to-sheet and sheet-to-bean loops | `@ExcelColumn` + `writeBeans` / `readAsBeans` |

### ⏱️ Time saved (measured in lines of code)

| Task | Apache POI | **Excel** | Ratio |
|---|---:|---:|---:|
| Create `.xlsx` + write "Hello World" with a style | ~15 | **5** | **3×** |
| Write `List<Bean>` as a styled, filtered, frozen table | ~60 | **5** | **12×** |
| Apply one reused style to 1,000 cells | ~8 / cell (repeat loops) | **1** | ≫10× |
| Read sheet into `List<Map<String,String>>` | ~30 | **1** | **30×** |
| Round-trip `List<Bean>` → file → `List<Bean>` | ~80 | **2** | **40×** |
| Freeze header + auto-filter + column widths | ~20 | **3** | **7×** |

---

## 🚀 Jump-away examples

### 1 · Hello, styled workbook (6 lines)

```java
ExcelApplication app = new ExcelApplication();
ExcelWorkBook wb = app.createWorkBook("demo.xlsx");
ExcelWorkSheet sh = wb.addSheet("Summary");
sh.cell(1,1).setText("Hello World").applyStyle(ExcelStyle.header());
wb.saveWorkbook();
app.closeAllWorkBooks();
```

### 2 · Bean → styled report (one call)

```java
public class Product {
    @ExcelColumn(name = "SKU",   order = 1)                      String sku;
    @ExcelColumn(name = "Item",  order = 2)                      String name;
    @ExcelColumn(name = "Units", order = 3, format = "#,##0")    int    units;
    @ExcelColumn(name = "Price", order = 4, format = "$#,##0.00") double price;
}

ReportBuilder.on(sheet)
    .title("Catalog")
    .rowsFromBeans(productList)
    .zebraStripes(true).freezeHeader(true).autoFilter(true)
    .render();
```

### 3 · `List<Map>` → spreadsheet

```java
List<Map<String,Object>> rows = List.of(
    Map.of("Item","Apple", "Qty",10, "Price",1.20),
    Map.of("Item","Pear",  "Qty", 8, "Price",1.80));

sheet.writeMaps(rows).freezePane(2,1).autoSizeAllColumns();
```

### 4 · Read Excel → `List<Bean>`

```java
List<Product> products = sheet.readAsBeans(Product.class);   // headers → fields
```

### 5 · Reusable style, applied 1,000 times

```java
ExcelStyle money = ExcelStyle.builder()
        .numberFormat("$#,##0.00").horizontalAlignment("RIGHT").fullBorder("black").build();

for (int r = 2; r <= 1001; r++) {
    sheet.cell(r, 3).setValue(revenue[r-2]).applyStyle(money);   // one style, many cells
}
```

### 6 · Hyperlinks & comments

```java
sheet.cell(1,1).setHyperlink("https://bismi.solutions", "Bismi Solutions");
sheet.cell(1,1).setComment("Official site", "Release notes");
```

---

## 🖼️ Runnable examples (each image is produced by the linked file)

Every example below is a real file under [`examples/`](examples) that you can run with
`mvn compile exec:java -Dexec.mainClass=<className>`. Each screenshot is a faithful
mock of the workbook it produces — colours, zebra, freeze pane, auto-filter, and all.

### 🟢 For beginners — no beans, no annotations

These two are the simplest possible starting points — a title, a few cells, and one built-in style preset.

#### 📝 Title and content — the simplest possible styled sheet

<p align="center">
  <img src="docs/title-content-preview.svg" alt="Title + content produced by TitleAndContentExample.java" width="80%"/>
</p>

> Source: [`examples/TitleAndContentExample.java`](examples/TitleAndContentExample.java) ·
> Showcases: a merged title row · the `ExcelStyle.header()` preset · plain text rows · `autoSizeAllColumns()`.
> **No beans, no loops, no annotations — about 15 lines of logic.**

---

#### 📊 KPI tiles — 4 colour-coded summary tiles on one sheet

<p align="center">
  <img src="docs/kpi-preview.svg" alt="KPIs produced by KpiTilesExample.java" width="90%"/>
</p>

> Source: [`examples/KpiTilesExample.java`](examples/KpiTilesExample.java) ·
> Showcases: custom fill colours per tile (green / blue / orange / red) · a small helper method for reuse.
> A beginner-sized slice of what [`DashboardExample`](examples/DashboardExample.java) does on sheet 1.

---

#### 🛒 Rows from arrays — one call writes the whole row

<p align="center">
  <img src="docs/row-array-preview.svg" alt="Shopping list produced by RowFromArrayExample.java" width="90%"/>
</p>

> Source: [`examples/RowFromArrayExample.java`](examples/RowFromArrayExample.java) ·
> Showcases:
>
> ```java
> // Header row from a String[]
> String[] headers = {"Item", "Qty", "Unit", "Aisle"};
> sh.row(2).setRowValues(headers);
>
> // Data rows from Object[] — mixed types get routed automatically
> sh.row(3).setValues(new Object[]{"Apples",  6, "pcs", "Produce"});
> sh.row(4).setValues(new Object[]{"Milk",    2, "L",   "Dairy"});
> ```
>
> No per-cell loop, no column-index bookkeeping. Use `setRowValues(String[])` when
> every cell is text (headers are the classic case) and `setValues(Object[])` when
> the row mixes strings, numbers, dates, booleans, etc.

---

### 🟠 For intermediate + advanced users

The rest use `@ExcelColumn` bean mapping, `ReportBuilder`, formulas, hyperlinks, round-trip reads, etc.

### 📦 Invoice — merged title · address blocks · line items · formulas · totals

<p align="center">
  <img src="docs/invoice-preview.svg" alt="Invoice produced by InvoiceExample.java" width="90%"/>
</p>

> Source: [`examples/InvoiceExample.java`](examples/InvoiceExample.java) ·
> Showcases: cell merging · reusable label/address/currency styles · formulas (`A*C`, `SUM`, tax) · bordered totals row · column widths.

---

### 📊 Executive Dashboard — KPI tiles · chart-ready sheet · raw data

<p align="center">
  <img src="docs/dashboard-preview.svg" alt="Dashboard produced by DashboardExample.java" width="95%"/>
</p>

> Source: [`examples/DashboardExample.java`](examples/DashboardExample.java) ·
> Showcases: 4 colour-coded KPI tiles (green/blue/orange/red) built from merged cells ·
> second sheet with `ReportBuilder` top-products table · third sheet with mixed-type raw data · frozen header + auto-filter.

---

### 👥 Employee Directory — beans + hyperlinks + comments + round-trip read

<p align="center">
  <img src="docs/employee-preview.svg" alt="Directory produced by EmployeeDirectoryExample.java" width="98%"/>
</p>

> Source: [`examples/EmployeeDirectoryExample.java`](examples/EmployeeDirectoryExample.java) ·
> Showcases: `@ExcelColumn` beans · zebra stripes · frozen header · auto-filter ·
> `mailto:` hyperlinks on the email column · greyed-italic style for inactive rows ·
> cell comment with author · **round-trip read back into `List<Employee>` via `readAsBeans`**.

---

### 🗂️ Three-in-one Collection Report — beans, maps, raw

> Source: [`examples/CollectionReportExample.java`](examples/CollectionReportExample.java) ·
> Sheet 1 uses `ReportBuilder.rowsFromBeans(List<Product>)`, sheet 2 uses `rowsFromMaps(List<Map>)`
> with a currency override on column 3, sheet 3 uses the bare-minimum `sheet.writeMaps(...)`.

---

## 📦 Installation

### Maven
```xml
<dependency>
  <groupId>solutions.bismi.excel</groupId>
  <artifactId>excel</artifactId>
  <version>1.3.0</version>
</dependency>
```

### Gradle (Kotlin DSL)
```kotlin
implementation("solutions.bismi.excel:excel:1.3.0")
```

### Gradle (Groovy DSL)
```groovy
implementation "solutions.bismi.excel:excel:1.3.0"
```

**Requires:** Java 17+  ·  Works on Windows · macOS · Linux.

---

## ☑️ Features at a glance

| Area | What's in the box |
|---|---|
| 📑 **Workbook** | create · open · save · `.xlsx` + `.xls` |
| 📄 **Sheets** | add · rename · activate · freeze panes · auto-filter · protect (password) |
| 📝 **Cells** | text · numbers · dates · formulas · **polymorphic `setValue(Object)`** · hyperlinks · comments |
| 🎨 **Styling** | fonts · 52 named colours ✨ · hex colours (XLSX) · borders · alignment · number formats · **reusable `ExcelStyle`** + presets |
| 📋 **Rows/Cols** | bulk values (mixed types) · auto-fit · column width · range styling |
| 🔗 **Merge/Unmerge** | succinct helpers · overlap safety checks |
| 🗂️ **Collections** | `writeMaps` · `writeBeans` · `readAsMaps` · `readAsBeans` · `@ExcelColumn` annotation |
| 🏗️ **Reports** | `ReportBuilder` — title · headers · zebra · freeze · auto-filter · per-column formats · column widths |
| 🧪 **Quality** | 97 unit tests (incl. every README snippet) · CI · Codecov · Sonar · OpenSSF Scorecard |

---

## 🗺️ Architecture

```mermaid
graph TD
    A[ExcelApplication<br/><i>lifecycle & open workbooks</i>] --> B[ExcelWorkBook<br/><i>file I/O, sheets</i>]
    B --> C[ExcelWorkSheet<br/><i>data, layout, filters</i>]
    C --> D[ExcelRow<br/><i>bulk values & styling</i>]
    C --> E[ExcelCell<br/><i>values, style, hyperlink, comment</i>]
    C -.uses.-> F[ReportBuilder<br/><i>one-call styled tables</i>]
    D -.applies.-> G[ExcelStyle<br/><i>reusable style defs</i>]
    E -.applies.-> G
    C -.reads/writes.-> H[List&lt;Map&gt; · List&lt;Bean&gt;<br/><i>@ExcelColumn</i>]

    style A fill:#2d6cdf,color:#fff,stroke:#1e4fb2
    style B fill:#2d9fdf,color:#fff,stroke:#1e7ba2
    style C fill:#36b37e,color:#fff,stroke:#278658
    style F fill:#ff9f43,color:#fff,stroke:#c77c2e
    style G fill:#9b59b6,color:#fff,stroke:#6d3a87
    style H fill:#e74c3c,color:#fff,stroke:#a33325
```

Each class exposes only methods appropriate to its scope — code completion *is* your documentation.

### Data-flow for a collection-driven report

```mermaid
flowchart LR
    subgraph Your_Data
      M["List&lt;Map&lt;String,Object&gt;&gt;"]
      B["List&lt;Bean&gt;<br/>@ExcelColumn"]
      A["Object[] rows"]
    end
    subgraph Excel_Lib
      RB[ReportBuilder]
      WS[ExcelWorkSheet]
      ST[ExcelStyle]
    end
    Out[(workbook.xlsx)]

    M --> RB
    B --> RB
    A --> RB
    M --> WS
    B --> WS
    RB -->|title · headers · data · zebra · freeze · filter| WS
    ST -.applies.-> WS
    WS --> Out

    style M fill:#fdf6e3,stroke:#b58900
    style B fill:#fdf6e3,stroke:#b58900
    style A fill:#fdf6e3,stroke:#b58900
    style RB fill:#ff9f43,color:#fff,stroke:#c77c2e
    style WS fill:#36b37e,color:#fff,stroke:#278658
    style ST fill:#9b59b6,color:#fff,stroke:#6d3a87
    style Out fill:#2d6cdf,color:#fff,stroke:#1e4fb2
```

---

## 🧱 Style reuse at a glance

```java
ExcelStyle header = ExcelStyle.header();         // blue fill · white bold · centred · bordered
ExcelStyle zebra  = ExcelStyle.zebraStripe();    // light-grey fill
ExcelStyle money  = ExcelStyle.currency();       // right-aligned $#,##0.00
ExcelStyle pct    = ExcelStyle.percent();        // right-aligned 0.00%
ExcelStyle when   = ExcelStyle.date();           // dd-MMM-yyyy

// v1.3 presets — pair with ReportBuilder for polished reports
ExcelStyle title  = ExcelStyle.title();          // navy merged title bar (pairs with header())
ExcelStyle totals = ExcelStyle.totals();         // pale-blue totals row, navy bold text
ExcelStyle active = ExcelStyle.statusActive();   // green pill — e.g. "Active", "OK"
ExcelStyle review = ExcelStyle.statusReview();   // amber pill — e.g. "Review", "Pending"
ExcelStyle closed = ExcelStyle.statusClosed();   // red pill   — e.g. "Closed", "Error"
ExcelStyle pill   = ExcelStyle.statusPill("#e0e7ff", "#1d4ed8");  // roll your own

ExcelStyle custom = ExcelStyle.builder()
        .fontColor("white").fillColor("#0d4ba1")
        .bold(true).horizontalAlignment("CENTER").fullBorder("black")
        .build();
```

Apply to anything:

```java
sheet.row(1).applyStyle(header);
sheet.row(5).applyStyle(money, 2, 6);            // cells 2..5 of row 5
sheet.cell(10,3).applyStyle(money);              // single cell
```

---

## 🤝 Contributing

```bash
git clone https://github.com/Bismi-Solutions/Excel.git
cd Excel
mvn test        # 97 tests — all green is the baseline
```

PRs welcome. Please include unit tests and follow the log-level convention:

- **info** → user-visible events (file created, sheet saved)
- **debug** → flow diagnostics
- **warn/error** → exceptional situations

Area ideas that need love: charts · conditional formatting · data validation (dropdowns) · named ranges · images · pivot tables.

---

## 📄 License

MIT — *use it, fork it, profit.* See [LICENSE](LICENSE).
