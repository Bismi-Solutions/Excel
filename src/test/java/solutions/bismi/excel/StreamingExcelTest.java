package solutions.bismi.excel;

import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

class StreamingExcelTest {

    @Test
    void testWriteAndReadStreaming() throws Exception {
        String fileName = "./resources/testdata/streaming.xlsx";
        File f = new File(fileName);
        if (f.exists()) {
            f.delete();
        }

        try (StreamingExcelWriter writer = new StreamingExcelWriter(fileName)) {
            for (int i = 0; i < 10; i++) {
                writer.writeRow("Sheet1", new String[]{"Row" + i, String.valueOf(i)});
            }
        }

        List<List<String>> rows = new ArrayList<>();
        StreamingExcelReader.read(fileName, (rowNum, values) -> rows.add(new ArrayList<>(values)));

        Assertions.assertEquals(10, rows.size());
        Assertions.assertEquals("Row0", rows.get(0).get(0));
        Assertions.assertEquals("9", rows.get(9).get(1));
    }
}
