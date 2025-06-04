package solutions.bismi.excel;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.xml.sax.Attributes;
import org.xml.sax.InputSource;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;
import org.xml.sax.helpers.XMLReaderFactory;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

/**
 * Simple streaming reader for XLSX files.
 * <p>
 * Parses sheets using a SAX handler and supplies each row to a caller provided
 * {@link RowHandler}. This avoids loading the entire workbook into memory.
 * </p>
 */
public class StreamingExcelReader {

    /** Handler for row data. */
    @FunctionalInterface
    public interface RowHandler {
        void handle(int rowNum, List<String> rowValues);
    }

    /**
     * Reads all sheets in the given XLSX file and sends each row to the provided
     * handler. Only string cell values are returned; other types are converted to
     * their string representation.
     *
     * @param fileName the file to read
     * @param handler  callback for each row
     * @throws Exception if reading fails
     */
    public static void read(String fileName, RowHandler handler) throws Exception {
        try (OPCPackage pkg = OPCPackage.open(fileName)) {
            XSSFReader reader = new XSSFReader(pkg);
            SharedStringsTable sst = reader.getSharedStringsTable();
            XMLReader parser = XMLReaderFactory.createXMLReader();
            parser.setContentHandler(new SheetHandler(sst, handler));
            Iterator<InputStream> sheets = reader.getSheetsData();
            while (sheets.hasNext()) {
                try (InputStream sheet = sheets.next()) {
                    parser.parse(new InputSource(sheet));
                }
            }
        }
    }

    private static class SheetHandler extends DefaultHandler {
        private final SharedStringsTable sst;
        private final RowHandler handler;
        private boolean nextIsString;
        private String lastContents = "";
        private List<String> row = new ArrayList<>();
        private int rowNum = 0;

        SheetHandler(SharedStringsTable sst, RowHandler handler) {
            this.sst = sst;
            this.handler = handler;
        }

        @Override
        public void startElement(String uri, String localName, String qName, Attributes attributes) {
            if ("row".equals(qName)) {
                row = new ArrayList<>();
            } else if ("c".equals(qName)) {
                String cellType = attributes.getValue("t");
                nextIsString = cellType != null && cellType.equals("s");
            }
            lastContents = "";
        }

        @Override
        public void endElement(String uri, String localName, String qName) {
            if (nextIsString && "v".equals(qName)) {
                int idx = Integer.parseInt(lastContents);
                lastContents = sst.getItemAt(idx).getString();
                nextIsString = false;
            }
            if ("v".equals(qName) || "t".equals(qName)) {
                row.add(lastContents);
            } else if ("row".equals(qName)) {
                handler.handle(rowNum++, row);
            }
        }

        @Override
        public void characters(char[] ch, int start, int length) {
            lastContents += new String(ch, start, length);
        }
    }
}

