package solutions.bismi.excel;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.Closeable;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

/**
 * Streaming writer for large Excel files using Apache POI's SXSSFWorkbook.
 * <p>
 * This class allows writing rows to a workbook without keeping the entire
 * document in memory, which is useful when dealing with very large datasets.
 * </p>
 */
public class StreamingExcelWriter implements Closeable {

    private final SXSSFWorkbook workbook;
    private final Map<String, SXSSFSheet> sheets = new HashMap<>();
    private final FileOutputStream out;

    /**
     * Creates a new streaming writer with the default row window size.
     *
     * @param fileName the target file name
     * @throws IOException if the file cannot be created
     */
    public StreamingExcelWriter(String fileName) throws IOException {
        this(fileName, 100);
    }

    /**
     * Creates a new streaming writer.
     *
     * @param fileName            the target file name
     * @param rowAccessWindowSize the number of rows to keep in memory
     * @throws IOException if the file cannot be created
     */
    public StreamingExcelWriter(String fileName, int rowAccessWindowSize) throws IOException {
        this.workbook = new SXSSFWorkbook(rowAccessWindowSize);
        this.out = new FileOutputStream(fileName);
    }

    /**
     * Writes a row of data to the specified sheet. Sheets are created on demand.
     *
     * @param sheetName the sheet name
     * @param values    the values to write
     */
    public void writeRow(String sheetName, String[] values) {
        SXSSFSheet sheet = sheets.computeIfAbsent(sheetName, workbook::createSheet);
        Row row = sheet.createRow(sheet.getLastRowNum() + 1);
        for (int i = 0; i < values.length; i++) {
            row.createCell(i).setCellValue(values[i]);
        }
    }

    /**
     * Flushes the current contents to disk. The workbook remains usable after
     * calling this method.
     *
     * @throws IOException if an IO error occurs
     */
    public void flush() throws IOException {
        workbook.write(out);
        out.flush();
    }

    /**
     * Flushes all data and closes the underlying streams. After calling this
     * method the writer cannot be used any further.
     */
    @Override
    public void close() throws IOException {
        flush();
        workbook.dispose();
        workbook.close();
        out.close();
    }
}

