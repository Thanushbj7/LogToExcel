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





    import java.io.*;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;

public class AutomatedUnzipAndReadLogFiles {

    public static void main(String[] args) {
        // Directory containing zip files
        String zipDirectory = "path/to/your/zip/files";

        try {
            File dir = new File(zipDirectory);
            File[] zipFiles = dir.listFiles((dir1, name) -> name.endsWith(".zip"));

            if (zipFiles != null) {
                for (File zipFile : zipFiles) {
                    System.out.println("Unzipping: " + zipFile.getName());

                    try (FileInputStream fis = new FileInputStream(zipFile);
                         ZipInputStream zis = new ZipInputStream(fis)) {

                        ZipEntry entry;
                        while ((entry = zis.getNextEntry()) != null) {
                            if (entry.getName().endsWith(".log")) {
                                System.out.println("Reading log file: " + entry.getName());

                                // Read the log file content here
                                // You can use a BufferedReader to read the content line by line
                                BufferedReader reader = new BufferedReader(new InputStreamReader(zis));
                                String line;
                                while ((line = reader.readLine()) != null) {
                                    // Process each line of the log file
                                    System.out.println(line);
                                }
                            }
                        }
                    }
                }
            } else {
                System.out.println("No zip files found in the directory.");
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}

















import java.io.*;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CombinedUnzipAndLogExcelUpdater {

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
        // Directory containing zip files
        String zipDirectory = "path/to/your/zip/files";

        try {
            File dir = new File(zipDirectory);
            File[] zipFiles = dir.listFiles((dir1, name) -> name.endsWith(".zip"));

            if (zipFiles != null) {
                for (File zipFile : zipFiles) {
                    System.out.println("Unzipping: " + zipFile.getName());

                    try (FileInputStream fis = new FileInputStream(zipFile);
                         ZipInputStream zis = new ZipInputStream(fis)) {

                        ZipEntry entry;
                        while ((entry = zis.getNextEntry()) != null) {
                            if (entry.getName().endsWith(".log")) {
                                System.out.println("Reading log file: " + entry.getName());

                                // Read the log file content here
                                // You can use a BufferedReader to read the content line by line
                                BufferedReader reader = new BufferedReader(new InputStreamReader(zis));
                                String line;
                                while ((line = reader.readLine()) != null) {
                                    // Process each line of the log file
                                    System.out.println(line);
                                }
                            }
                        }
                    }
                }
            } else {
                System.out.println("No zip files found in the directory.");
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}


The current manual process involves efforts from both the Salesforce AM Team and Salesforce AD Team, and creates about 20 additional hours of work every month. In order to eliminate this resource draw, Salesforce AD will create an automated process to retrieve the copy stats and input them directly into the SW_SFDC_DATACOPY_CONTROL table in Smartworks 1.0.
 
Requirements:
•	The following Copy Table logs must be queried for the data requested in Requirement #2: 
o	SBR_Account_History_SFDC
o	SBR_Reg_Member_History_SFDC
o	SBR_Registration_History_SFDC
o	SBR_REG_LETTER_LOG_SFDC
o	SBR_REG_LETTER_LOG_T2_SFDC
•	
o	SFDC_W_TR_REGISTRATION_MAP
o	SBR_W_REG_LETTER_LOG_REL_SFDC
o	SFDC_W_SPONSOR_NAMES
•	
o	SFDC_W_CLIENT_DISCLOSURE
o	SFDC_W_CLIENT
o	SFDC_W_PORTFOLIO_REVIEW
o	SFDC_W_BENEFICIARY
o	SFDC_W_REGISTRATION
o	SFDC_W_REGISTRATION_MEMBERS
•	
o	SFDC_EBLOTTER
o	SFDC_TRADES
o	SFDC_CHECKS
•	
o	SFDC_USER
•	  
o	SFDC_W_FINANCIAL_ACCOUNT
o	SFDC_W_FINANCIAL_ACCOUNT_TEAM
 
•	The following values must be extracted from each of the logs for the above Copy Tables: 
o	RUN_DATE 
	This is the first timestamp in the log
o	RUN_CYCLE 
	This is 1,2,or 3 depending on the run time for the log. For example, the 2am batches are Run Cycle 1, the subsequent set of batches are Run Cycle 2, etc.
o	TABLE_NAME 
o	WEEKLY_DAILY 
	This is the frequency of the job. All jobs are "D" for daily, except SFDC_W_FINANCIAL ACCOUNT and SFDC_W_FINANCIAL_ACCOUNT_TEAM, which are "W" for weekly
o	RECORD_COUNT 
	This is the number of successful extractions captured in the last line of the log file.
o	START_DATE 
	This is the date/time value captured in the first line after “decrypting” in the log file. 
o	END_DATE 
	This is the date/time value captured in the line of the log file that is affiliated to  “The operation has fully completed” 
•	The values should be mapped to the SW_SFDC_DATACOPY_CONTROL table in SmartWorks 1.0 before 10:30am ET daily.

