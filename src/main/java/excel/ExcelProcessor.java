package excel;

import lombok.RequiredArgsConstructor;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;

/**
 * This class takes on the responsibility of finding prime numbers within an Excel file
 * Data are expected to be in the second column (B)
 */
@RequiredArgsConstructor
public class ExcelProcessor {
    private static final Logger log = LogManager.getLogger(ExcelProcessor.class);

    private final String filePath;

    /**
     * Processes numbers in the second column of the data file (B) and logs the prime numbers.
     * We don't store anything into the memory as the size of the data file could be massive.
     *
     * @throws RuntimeException if the file cannot be found or read for some reason
     */
    public void processPrimeNumbers(int sheetIndex, int columnIndex) throws RuntimeException {
        try (FileInputStream fis = new FileInputStream(filePath); Workbook workbook = new XSSFWorkbook(fis)) {
            Sheet sheet = workbook.getSheetAt(sheetIndex);
            int rowCount = sheet.getPhysicalNumberOfRows();
            for (int i = 0; i < rowCount; i++) {
                Row row = sheet.getRow(i);
                long rowResult = processRow(row, columnIndex);
                if (rowResult != -1) {
                    log.info("FOUND PRIME: {}", rowResult);
                }
            }
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    /**
     * Processes a single row in the sheet, including validation.
     *
     * @param row       - A single row in the sheet to be processed
     * @param cellIndex - Index of the cell to process
     * @return a number if the number is valid and prime, -1 in all other cases
     */
    public long processRow(Row row, int cellIndex) {
        if (row == null) {
            return -1;
        }

        int rowNumber = row.getRowNum();
        Cell targetCell = row.getCell(cellIndex); // data are in the second column
        if (!isCellValid(targetCell)) {
            log.debug("Invalid cell value, skipping row: {} (Possibly header or a non-numeric value)", rowNumber);
            return -1;
        }

        try {
            long number = switch (targetCell.getCellType()) {
                case NUMERIC -> (long) targetCell.getNumericCellValue();
                case STRING -> Long.parseLong(targetCell.getStringCellValue().trim());
                default -> -1;
            };
            if (number != -1 && isPrime(number)) {
                return number;
            }
        } catch (NumberFormatException e) {
            log.debug("Failed to parse number at row: {}, skipping this row", rowNumber);
            return -1;
        }

        return -1;
    }

    /**
     * Check if the cell contains only a positive integer value
     *
     * @param cell - Cell to check the validity of
     * @return true if the cell is valid, else false
     */
    public boolean isCellValid(Cell cell) {
        if (cell == null) {
            return false;
        }
        log.debug("Processing cell of type: {}", cell.getCellType());
        return switch (cell.getCellType()) {
            case NUMERIC -> isPositiveInteger(cell.getNumericCellValue());
            case STRING -> {
                String cellStringValue = cell.getStringCellValue().trim();
                // Here we could also iterate over the string and for each char check if it's a number using Character.isDigit()
                // But the regex is much simpler to implement, read, and it's faster in this case
                if (cellStringValue.matches("\\d+")) {
                    yield isPositiveInteger(Double.parseDouble(cellStringValue));
                }
                yield false;
            }
            default -> false; // reject any other type
        };
    }

    /**
     * Checks if a given number is a positive whole number
     *
     * @param number - the number to be checked
     * @return true if the number is a positive integer, otherwise false
     */
    private boolean isPositiveInteger(double number) {
        return number > 0 && number == Math.floor(number);
    }

    /**
     * Checks if a number is prime = only divisible by 1 and itself.
     * Every factor of a prime number is not a prime number
     * - This is because such number would be divisible by itself, 1, and the prime number
     * - therefore we only need to check up to the √number
     *
     * @param number - number to be checked
     * @return true if the given number is prime, false otherwise
     */
    public boolean isPrime(long number) {
        if (number <= 1) {
            return false;
        }

        for (long i = 2; i <= Math.sqrt(number); i++) {
            if (number % i == 0) {
                return false;
            }
        }

        return true;
    }
}
