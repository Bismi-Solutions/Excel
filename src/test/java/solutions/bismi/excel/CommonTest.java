package solutions.bismi.excel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.UUID;

import static org.junit.jupiter.api.Assertions.*;

class CommonTest {

    private static final String TEST_DATA_DIR = "./resources/testdata";

    @BeforeEach
    void setUp() throws IOException {
        // Create test data directory if it doesn't exist
        Files.createDirectories(Path.of(TEST_DATA_DIR));
    }

    private String createTestFile(String prefix)  {
        String uniqueId = UUID.randomUUID().toString().substring(0, 8);
        String fileName = prefix + uniqueId + ".XLSX";
        return Path.of(TEST_DATA_DIR, fileName).toString();
    }

    @Test
    void testCopyFontProperties() throws IOException {
        try (Workbook workbook = new XSSFWorkbook()) {
            Font sourceFont = workbook.createFont();
            Font targetFont = workbook.createFont();

            // Set source font properties
            sourceFont.setBold(true);
            sourceFont.setItalic(true);
            sourceFont.setStrikeout(true);
            sourceFont.setUnderline(Font.U_SINGLE);
            sourceFont.setFontHeightInPoints((short) 14);
            sourceFont.setFontName("Arial");

            // Copy properties
            Common.copyFontProperties(sourceFont, targetFont);

            // Verify properties were copied correctly
            assertTrue(targetFont.getBold());
            assertTrue(targetFont.getItalic());
            assertTrue(targetFont.getStrikeout());
            assertEquals(Font.U_SINGLE, targetFont.getUnderline());
            assertEquals(14, targetFont.getFontHeightInPoints());
            assertEquals("Arial", targetFont.getFontName());
        }
    }

    @Test
    void testGetColorCode() {
        // Test valid color names
        assertEquals(IndexedColors.RED.getIndex(), Common.getColorCode("RED"));
        assertEquals(IndexedColors.BLUE.getIndex(), Common.getColorCode("BLUE"));
        assertEquals(IndexedColors.GREEN.getIndex(), Common.getColorCode("GREEN"));

        // Test case insensitivity
        assertEquals(IndexedColors.RED.getIndex(), Common.getColorCode("red"));
        assertEquals(IndexedColors.RED.getIndex(), Common.getColorCode("Red"));

        // Test invalid color name
        assertEquals(0, Common.getColorCode("INVALID_COLOR"));
    }

    @Test
    void testCheckIfRowIsEmpty() throws IOException {
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet();
            
            // Test null row
            assertTrue(Common.checkIfRowIsEmpty(null));

            // Test empty row
            Row emptyRow = sheet.createRow(0);
            assertTrue(Common.checkIfRowIsEmpty(emptyRow));

            // Test row with empty cells
            Row rowWithEmptyCells = sheet.createRow(1);
            rowWithEmptyCells.createCell(0);
            rowWithEmptyCells.createCell(1);
            assertTrue(Common.checkIfRowIsEmpty(rowWithEmptyCells));

            // Test row with non-empty cell
            Row nonEmptyRow = sheet.createRow(2);
            Cell cell = nonEmptyRow.createCell(0);
            cell.setCellValue("Test");
            assertFalse(Common.checkIfRowIsEmpty(nonEmptyRow));
        }
    }

    @Test
    void testCheckIfCellIsEmpty() throws IOException {
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet();
            Row row = sheet.createRow(0);

            // Test null cell
            assertTrue(Common.checkIfCellIsEmpty(null));

            // Test blank cell
            Cell blankCell = row.createCell(0);
            assertTrue(Common.checkIfCellIsEmpty(blankCell));

            // Test cell with value
            Cell nonEmptyCell = row.createCell(1);
            nonEmptyCell.setCellValue("Test");
            assertFalse(Common.checkIfCellIsEmpty(nonEmptyCell));

            // Test cell with number
            Cell numberCell = row.createCell(2);
            numberCell.setCellValue(123);
            assertFalse(Common.checkIfCellIsEmpty(numberCell));
        }
    }

    @Test
    void testSaveWorkBook() throws IOException {
        try (Workbook workbook = new XSSFWorkbook()) {
            // Create a test sheet and add some data
            Sheet sheet = workbook.createSheet("Test");
            Row row = sheet.createRow(0);
            row.createCell(0).setCellValue("Test Data");

            // Save the workbook
            String filePath = createTestFile("testSaveWorkbook");
            Common.saveWorkBook(workbook, filePath);

            // Verify the file was created
            File savedFile = new File(filePath);
            assertTrue(savedFile.exists());
            assertTrue(savedFile.length() > 0);

            // Verify the saved workbook can be opened
            try (Workbook savedWorkbook = WorkbookFactory.create(savedFile)) {
                Sheet savedSheet = savedWorkbook.getSheet("Test");
                assertNotNull(savedSheet);
                assertEquals("Test Data", savedSheet.getRow(0).getCell(0).getStringCellValue());
            }
        }
    }
} 