package solutions.bismi.excel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.*;
import org.junit.jupiter.api.io.TempDir;

import java.io.File;
import java.io.IOException;
import java.nio.file.Path;
import java.util.UUID;

import static org.junit.jupiter.api.Assertions.*;

class CommonTest {

    @TempDir
    Path temp;

    private String createTestFile(String prefix) {
        String id = UUID.randomUUID().toString().substring(0, 8);
        return temp.resolve(prefix + id + ".xlsx").toString();
    }

    @Test
    void copyFontProperties() throws IOException {
        try (Workbook wb = new XSSFWorkbook()) {
            Font src = wb.createFont();
            Font tgt = wb.createFont();

            src.setBold(true);
            src.setItalic(true);
            src.setStrikeout(true);
            src.setUnderline(Font.U_SINGLE);
            src.setFontHeightInPoints((short) 14);
            src.setFontName("Arial");

            Common.copyFontProperties(src, tgt);

            assertAll(
                    () -> assertTrue(tgt.getBold()),
                    () -> assertTrue(tgt.getItalic()),
                    () -> assertTrue(tgt.getStrikeout()),
                    () -> assertEquals(Font.U_SINGLE, tgt.getUnderline()),
                    () -> assertEquals(14, tgt.getFontHeightInPoints()),
                    () -> assertEquals("Arial", tgt.getFontName())
            );
        }
    }

    @Test
    void getColorCode() {
        assertEquals(IndexedColors.RED.getIndex(),   Common.getColorCode("RED"));
        assertEquals(IndexedColors.GREEN.getIndex(), Common.getColorCode("green"));
        assertEquals(0, Common.getColorCode("NO_SUCH_COLOUR"));
    }

    @Test
    void rowAndCellChecks() throws IOException {
        try (Workbook wb = new XSSFWorkbook()) {
            Sheet sh = wb.createSheet();
            Row row0 = sh.createRow(0);
            Cell populated = row0.createCell(0);
            populated.setCellValue("x");

            assertFalse(Common.checkIfRowIsEmpty(row0));
            assertFalse(Common.checkIfCellIsEmpty(populated));
        }
    }

    @Test
    void saveWorkBook() throws IOException {
        try (Workbook wb = new XSSFWorkbook()) {
            wb.createSheet("T").createRow(0).createCell(0).setCellValue("data");
            String path = createTestFile("save");
            Common.saveWorkBook(wb, path);
            assertTrue(new File(path).exists());
        }
    }
}