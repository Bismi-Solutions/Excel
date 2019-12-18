# Excel
[![build_status](https://travis-ci.com/Bismi-Solutions/Excel.svg?branch=master)](https://travis-ci.com/Bismi-Solutions/Excel)    [![Known Vulnerabilities](https://snyk.io/test/github/Bismi-Solutions/Excel/badge.svg?targetFile=pom.xml)](https://snyk.io/test/github/Bismi-Solutions/Excel?targetFile=pom.xml)

Repository to read and write excel file seamlessly. This repo internally uses apache poi java library for file processing.





Add Maven dependency:
```
<!-- https://mvnrepository.com/artifact/solutions.bismi.excel/Excel -->
<dependency>
    <groupId>solutions.bismi.excel</groupId>
    <artifactId>Excel</artifactId>
    <version>1.0.8</version>
</dependency>

```

Sample program to use with excel cell

```
        ExcelApplication xlApp =new ExcelApplication();
        ExcelWorkBook  xlbook=xlApp.createWorkBook(strCompleteFileName);
        int cnt=0;
        cnt=xlbook.getSheetCount();
        assertEquals(1,cnt);
        ExcelWorkSheet sh1 = xlbook.addSheet("Bismi1");
        sh1.activate();
        sh1.cell(10,10).setText("TestColor");
        sh1.cell(10,10).setFontColor("blue");
        sh1.cell(10,10).setFillColor("yellow");
        sh1.cell(1,1).setFillColor("GREEN");
        sh1.cell(3,17).setFullBorder("Red");
        sh1.saveWorkBook();
        xlApp.closeAllWorkBooks();

```

Sample Excel row program

```
ExcelApplication xlApp =new ExcelApplication();
        ExcelWorkBook  xlbook=xlApp.createWorkBook(strCompleteFileName);
        int cnt=0;
        cnt=xlbook.getSheetCount();
        assertEquals(1,cnt);
        ExcelWorkSheet sh1 = xlbook.addSheet("Bismi1");
        sh1.activate();

        String[] arrRow= {"A","B","C","D","E"};
        sh1.row(11).setRowValues(arrRow);
        sh1.row(11).setFontColor("Red",1,3);
        sh1.row(11).setFontColor("green",3,7);
        sh1.row(2).setRowValues(arrRow);
        sh1.row(2).setFontColor("White");
        sh1.row(2).setFillColor("Green");
        sh1.row(2).setFullBorder("Red");
        sh1.row(11).setFullBorder("blue");
        sh1.row(5).setFullBorder("blue",1,10);
        sh1.saveWorkBook();
        xlApp.closeAllWorkBooks();
```







