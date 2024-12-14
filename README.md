# XLSX Prime Finder

## Description

A simple Java application which finds prime numbers among a column of numbers within an Excel sheet in XLSX format.

## Usage

- Provide the path to the data file as a runtime argument
  - Application only accepts this single argument
- In the main method you can specify which sheet and column the prime numbers should be checked for

```java
public static void main(String[] args) {
    if (!areArgumentsValid(args)) {
        log.error("Invalid arguments were given");
        return;
    }
    String filePath = args[0];
    ExcelProcessor excelProcessor = new ExcelProcessor(filePath);
    // First argument is the sheet, second is the column
    // Below we check for prime numbers in the first sheet (0-indexed) and the second column
    excelProcessor.processPrimeNumbers(0, 1);
}
```