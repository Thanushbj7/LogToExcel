import java.io.BufferedReader;
import java.io.FileReader;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

public class LogExcelUpdater {

    // Define the log file path
    static String logFilePath = "";

    // Define the Excel file path
    static String excelFilePath = "extract-excel.xlsx";

    // Regular expression for extracting date and time
    static String datetimePattern = "(\\d{4}-\\d{2}-\\d{2} \\d{2}:\\d{2}:\\d{2})";

    // Regular expression for extracting a number before the word 'successful'
    static String numberPattern = "(\\d+)\\s+successful";

    // Function to extract data from the log file based on TABLE_NAME column in Excel
    static List<String> extractDataFromLog(String logFilePath, String excelFilePath) {
        List<String> extractedData = new ArrayList<>();
        try {
            BufferedReader reader = new BufferedReader(new FileReader(logFilePath));
            String line;
            String firstRecord = null;
            String lastRecord = null;
            String decryptingRecord = null;
            String lastSuccessful = null;
            while ((line = reader.readLine()) != null) {
                Pattern pattern = Pattern.compile(datetimePattern);
                Matcher matcher = pattern.matcher(line);
                if (matcher.find()) {
                    if (firstRecord == null) {
                        firstRecord = matcher.group(0);
                    }
                    lastRecord = matcher.group(0);
                }
                if (line.contains("Decrypting...")) {
                    line = reader.readLine(); // read the next line
                    matcher = pattern.matcher(line);
                    if (matcher.find()) {
                        decryptingRecord = matcher.group(0);
                    }
                }
                pattern = Pattern.compile(numberPattern);
                matcher = pattern.matcher(line);
                if (matcher.find()) {
                    lastSuccessful = matcher.group(1);
                }
            }
            reader.close();
            extractedData.add(firstRecord);
            extractedData.add(lastRecord);
            extractedData.add(decryptingRecord);
            extractedData.add(lastSuccessful);
        } catch (IOException e) {
            e.printStackTrace();
        }
        return extractedData;
    }

    // Function to update Excel file with extracted data
    static void updateExcelWithData(int rowNumber, String excelFilePath, List<String> extractedData) {
        try {
            FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
            Workbook workbook = new XSSFWorkbook(inputStream);
            Sheet sheet = workbook.getSheetAt(0);
            Row row = sheet.getRow(rowNumber);
            if (row == null) {
                row = sheet.createRow(rowNumber);
            }
            Cell cell1 = row.createCell(0);
            cell1.setCellValue(extractedData.get(0));
            Cell cell2 = row.createCell(6);
            cell2.setCellValue(extractedData.get(1));
            Cell cell3 = row.createCell(5);
            cell3.setCellValue(extractedData.get(2));
            Cell cell4 = row.createCell(4);
            cell4.setCellValue(extractedData.get(3));
            inputStream.close();
            FileOutputStream outputStream = new FileOutputStream(excelFilePath);
            workbook.write(outputStream);
            workbook.close();
            outputStream.close();
            System.out.println("Data successfully updated in Excel file.");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static void main(String[] args) {
        String[] tableNames = { "REG_MAP", "asa", "start_map2" };
        int rowNumber = 2;
        for (String tableName : tableNames) {
            if (tableName.equals("REG_MAP")) {
                logFilePath = "example.log";
            } else if (tableName.equals("asa")) {
                logFilePath = "example1.log";
                rowNumber = 3;
            } else if (tableName.equals("start_map2")) {
                logFilePath = "example2.log";
                rowNumber = 4;
            } else {
                continue;
            }
            List<String> extractedData = extractDataFromLog(logFilePath, excelFilePath);
            if (!extractedData.isEmpty()) {
                updateExcelWithData(rowNumber, excelFilePath, extractedData);
            }
        }
    }
}
