import excel.ExcelProcessor;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.AfterAll;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import static org.junit.jupiter.api.Assertions.*;

class ExcelProcessorTest {

    private static File tempExcelFile;

    @BeforeAll
    static void setUp() throws IOException {
        tempExcelFile = File.createTempFile("test", ".xlsx");
        try (Workbook workbook = new XSSFWorkbook(); FileOutputStream fos = new FileOutputStream(tempExcelFile)) {
            Sheet sheet = workbook.createSheet();

            // header
            Row headerRow = sheet.createRow(0);
            headerRow.createCell(0).setCellValue("Index");
            headerRow.createCell(1).setCellValue("Number");

            // data
            Object[][] data = {{1, 2},    // prime
                    {2, 4},    // not prime
                    {3, 17},   // prime
                    {4, 20},   // not prime
                    {5, 23},   // prime
                    {6, "Invalid"}, // invalid
                    {7, -7}    // invalid
            };

            for (Object[] rowData : data) {
                Row row = sheet.createRow((int) rowData[0]);
                row.createCell(0).setCellValue((int) rowData[0]);
                if (rowData[1] instanceof Integer) {
                    row.createCell(1).setCellValue((int) rowData[1]);
                } else {
                    row.createCell(1).setCellValue((String) rowData[1]);
                }
            }

            workbook.write(fos);
        }
    }

    @AfterAll
    static void tearDown() {
        if (tempExcelFile.exists()) {
            tempExcelFile.delete();
        }
    }

    @Test
    void testIsPrime() {
        ExcelProcessor processor = new ExcelProcessor(tempExcelFile.getAbsolutePath());

        assertTrue(processor.isPrime(2));
        assertTrue(processor.isPrime(17));
        assertFalse(processor.isPrime(4));
        assertFalse(processor.isPrime(20));
        assertFalse(processor.isPrime(1));
        assertFalse(processor.isPrime(-7));
    }

    @Test
    void testIsCellValid() throws IOException, InvalidFormatException {
        ExcelProcessor processor = new ExcelProcessor(tempExcelFile.getAbsolutePath());

        try (Workbook workbook = new XSSFWorkbook(tempExcelFile)) {
            Sheet sheet = workbook.getSheetAt(0);

            assertTrue(processor.isCellValid(sheet.getRow(1).getCell(1)), "Cell with value 2 should be valid");
            assertFalse(processor.isCellValid(sheet.getRow(6).getCell(1)), "Cell with 'Invalid' should be invalid");
            assertFalse(processor.isCellValid(sheet.getRow(7).getCell(1)), "Cell with value -7 should be invalid");
        }
    }

    @Test
    void testProcessRow() throws IOException, InvalidFormatException {
        ExcelProcessor processor = new ExcelProcessor(tempExcelFile.getAbsolutePath());

        try (Workbook workbook = new XSSFWorkbook(tempExcelFile)) {
            Sheet sheet = workbook.getSheetAt(0);

            assertEquals(2L, processor.processRow(sheet.getRow(1), 1), "Row with value 2 should return prime 2");
            assertEquals(-1L, processor.processRow(sheet.getRow(2), 1), "Row with value 4 should return -1");
            assertEquals(17L, processor.processRow(sheet.getRow(3), 1), "Row with value 17 should return prime 17");
            assertEquals(-1L, processor.processRow(sheet.getRow(6), 1), "Row with value 'Invalid' should return -1");
            assertEquals(-1L, processor.processRow(sheet.getRow(7), 1), "Row with value -7 should return -1");
        }
    }
}
