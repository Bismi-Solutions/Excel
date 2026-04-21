package solutions.bismi.excel;

import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.MethodOrderer;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.TestMethodOrder;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

@TestMethodOrder(MethodOrderer.MethodName.class)
class ExcelCollectionsTest {

    public static class Person {
        @ExcelColumn(name = "Id",   order = 1) int    id;
        @ExcelColumn(name = "Name", order = 2) String name;
        @ExcelColumn(name = "Age",  order = 3) int    age;
        @ExcelColumn(name = "Pro",  order = 4) boolean pro;

        public Person() {}
        public Person(int id, String name, int age, boolean pro) {
            this.id = id; this.name = name; this.age = age; this.pro = pro;
        }
    }

    @Test
    void aWriteMapsThenReadAsMapsRoundTrip() {
        String path = "./resources/testdata/writeMaps.xlsx";
        ExcelApplication app = new ExcelApplication();
        ExcelWorkBook wb = app.createWorkBook(path);
        ExcelWorkSheet sh = wb.addSheet("Maps");

        List<Map<String, Object>> rows = new ArrayList<>();
        rows.add(ordered("Item", "Apple", "Qty", 10, "Price", 1.2));
        rows.add(ordered("Item", "Pear",  "Qty",  8, "Price", 1.8));
        sh.writeMaps(rows);
        app.closeAllWorkBooks();

        ExcelApplication app2 = new ExcelApplication();
        ExcelWorkBook wb2 = app2.openWorkbook(path);
        ExcelWorkSheet sh2 = wb2.getExcelSheet("Maps");
        List<Map<String, String>> read = sh2.readAsMaps();

        Assertions.assertEquals(2, read.size());
        Assertions.assertEquals("Apple", read.get(0).get("Item"));
        Assertions.assertEquals("10",    read.get(0).get("Qty"));
        Assertions.assertEquals("Pear",  read.get(1).get("Item"));
        app2.closeAllWorkBooks();
    }

    @Test
    void bWriteMapsEmptyListIsNoOp() {
        ExcelApplication app = new ExcelApplication();
        ExcelWorkBook wb = app.createWorkBook("./resources/testdata/writeMapsEmpty.xlsx");
        ExcelWorkSheet sh = wb.addSheet("Empty");
        Assertions.assertDoesNotThrow(() -> sh.writeMaps(null));
        Assertions.assertDoesNotThrow(() -> sh.writeMaps(new ArrayList<>()));
        app.closeAllWorkBooks();
    }

    @Test
    void cWriteBeansThenReadAsBeansRoundTrip() {
        String path = "./resources/testdata/writeBeans.xlsx";
        ExcelApplication app = new ExcelApplication();
        ExcelWorkBook wb = app.createWorkBook(path);
        ExcelWorkSheet sh = wb.addSheet("People");

        List<Person> people = Arrays.asList(
                new Person(1, "Alice", 30, true),
                new Person(2, "Bob",   25, false),
                new Person(3, "Carol", 40, true));
        sh.writeBeans(people);
        app.closeAllWorkBooks();

        ExcelApplication app2 = new ExcelApplication();
        ExcelWorkBook wb2 = app2.openWorkbook(path);
        ExcelWorkSheet sh2 = wb2.getExcelSheet("People");
        List<Person> read = sh2.readAsBeans(Person.class);

        Assertions.assertEquals(3, read.size());
        Assertions.assertEquals(1,       read.get(0).id);
        Assertions.assertEquals("Alice", read.get(0).name);
        Assertions.assertEquals(30,      read.get(0).age);
        Assertions.assertTrue(read.get(0).pro);
        Assertions.assertFalse(read.get(1).pro);
        app2.closeAllWorkBooks();
    }

    @Test
    void dReadAsMapsEmptySheetReturnsEmpty() {
        String path = "./resources/testdata/emptyMaps.xlsx";
        ExcelApplication app = new ExcelApplication();
        ExcelWorkBook wb = app.createWorkBook(path);
        ExcelWorkSheet sh = wb.addSheet("Blank");
        Assertions.assertTrue(sh.readAsMaps().isEmpty());
        app.closeAllWorkBooks();
    }

    @Test
    void eExcelColumnAnnotationMetadata() throws Exception {
        java.lang.reflect.Field idField = Person.class.getDeclaredField("id");
        ExcelColumn ann = idField.getAnnotation(ExcelColumn.class);
        Assertions.assertNotNull(ann);
        Assertions.assertEquals("Id", ann.name());
        Assertions.assertEquals(1, ann.order());
        Assertions.assertEquals("", ann.format());
    }

    @Test
    void fWriteBeansWithNoAnnotationsIsNoOp() {
        ExcelApplication app = new ExcelApplication();
        ExcelWorkBook wb = app.createWorkBook("./resources/testdata/beansNoAnn.xlsx");
        ExcelWorkSheet sh = wb.addSheet("X");
        List<String> plainStrings = Arrays.asList("a", "b");
        Assertions.assertDoesNotThrow(() -> sh.writeBeans(plainStrings));
        app.closeAllWorkBooks();
    }

    private static Map<String, Object> ordered(Object... kv) {
        LinkedHashMap<String, Object> m = new LinkedHashMap<>();
        for (int i = 0; i < kv.length; i += 2) {
            m.put((String) kv[i], kv[i + 1]);
        }
        return m;
    }
}
