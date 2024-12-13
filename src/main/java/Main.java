import excel.ExcelProcessor;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;


public class Main {
    private static final Logger log = LogManager.getLogger(Main.class);

    public static void main(String[] args) {
        if (!areArgumentsValid(args)) {
            log.error("Invalid arguments were given");
            return;
        }
        String filePath = args[0];
        ExcelProcessor excelProcessor = new ExcelProcessor(filePath);
        excelProcessor.processPrimeNumbers(0, 1);
    }

    /**
     * Checks if the passed arguments are valid, we only expect a single argument - the data file path
     *
     * @param args - Array of provided arguments by the user
     * @return true if exactly 1 non-empty argument is passed by the user
     */
    private static boolean areArgumentsValid(String[] args) {
        return args.length == 1 && !args[0].isEmpty();
    }
}
