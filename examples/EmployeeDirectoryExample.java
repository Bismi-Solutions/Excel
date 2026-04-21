import solutions.bismi.excel.ExcelApplication;
import solutions.bismi.excel.ExcelColumn;
import solutions.bismi.excel.ExcelStyle;
import solutions.bismi.excel.ExcelWorkBook;
import solutions.bismi.excel.ExcelWorkSheet;
import solutions.bismi.excel.ReportBuilder;

import java.time.LocalDate;
import java.util.Arrays;
import java.util.List;

/**
 * Staff directory built from a {@code List<Employee>}. Demonstrates:
 *
 * <ul>
 *   <li>{@link ExcelColumn} annotations for header, order, format</li>
 *   <li>{@link ReportBuilder} — zebra striping, frozen header, auto-filter, per-column currency</li>
 *   <li>Post-render enrichment: clickable email hyperlinks and a cell comment</li>
 *   <li>Round-trip: read the saved file back into a {@code List<Employee>} via {@code readAsBeans}</li>
 * </ul>
 */
public class EmployeeDirectoryExample {

    public static class Employee {
        @ExcelColumn(name = "Id",      order = 1)                              int        id;
        @ExcelColumn(name = "Name",    order = 2)                              String     name;
        @ExcelColumn(name = "Role",    order = 3)                              String     role;
        @ExcelColumn(name = "Email",   order = 4)                              String     email;
        @ExcelColumn(name = "Joined",  order = 5, format = "dd-MMM-yyyy")      LocalDate  joined;
        @ExcelColumn(name = "Salary",  order = 6, format = "$#,##0")           double     salary;
        @ExcelColumn(name = "Active",  order = 7)                              boolean    active;

        public Employee() {}
        public Employee(int id, String name, String role, String email, LocalDate joined, double salary, boolean active) {
            this.id = id; this.name = name; this.role = role; this.email = email;
            this.joined = joined; this.salary = salary; this.active = active;
        }
    }

    public static void main(String[] args) {
        String path = "./resources/testdata/employees.xlsx";

        ExcelApplication app = new ExcelApplication();
        ExcelWorkBook wb = app.createWorkBook(path);
        ExcelWorkSheet sh = wb.addSheet("Directory");
        sh.activate();   // show this sheet by default when the file is opened

        List<Employee> staff = Arrays.asList(
                new Employee(101, "Amina Khan",    "Engineer",         "amina.khan@example.com",   LocalDate.of(2021, 7, 12),  95000, true),
                new Employee(102, "Rahul Menon",   "Senior Engineer",  "rahul.menon@example.com",  LocalDate.of(2019, 3, 4), 135000, true),
                new Employee(103, "Sofia Gomez",   "Product Manager",  "sofia.gomez@example.com",  LocalDate.of(2022, 11, 21),110000, true),
                new Employee(104, "Dmitri Ivanov", "QA Lead",          "dmitri.ivanov@example.com",LocalDate.of(2018, 1, 9), 120000, false),
                new Employee(105, "John Jones",    "Designer",         "john.jones@example.com",   LocalDate.of(2023, 6, 15),  88000, true));

        // No title() — keeps row 1 as the header so readAsBeans works cleanly later.
        ReportBuilder.on(sh)
                .rowsFromBeans(staff)
                .zebraStripes(true)
                .freezeHeader(true)
                .autoFilter(true)
                .render();

        int emailCol = 4;
        for (int r = 0; r < staff.size(); r++) {
            int row = r + 2;
            sh.cell(row, emailCol).setHyperlink("mailto:" + staff.get(r).email, staff.get(r).email);
        }

        sh.cell(2, 2).setComment("Go-to for onboarding questions", "HR Team");

        ExcelStyle inactive = ExcelStyle.builder()
                .fontColor("grey_50_percent").italic(true).build();
        for (int r = 0; r < staff.size(); r++) {
            if (!staff.get(r).active) {
                int row = r + 2;
                sh.row(row).applyStyle(inactive, 1, 8);
            }
        }

        wb.saveWorkbook();
        app.closeAllWorkBooks();

        ExcelApplication app2 = new ExcelApplication();
        ExcelWorkBook wb2 = app2.openWorkbook(path);
        List<Employee> read = wb2.getExcelSheet("Directory").readAsBeans(Employee.class);
        System.out.printf("Read back %d employees; first is %s (%s)%n",
                read.size(), read.get(0).name, read.get(0).role);
        app2.closeAllWorkBooks();
    }
}
