import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Enumeration;
import java.util.Iterator;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.nio.file.Paths;

public class Final {

    // Regular expression for extracting date and time
    // Assuming datetime format is YYYY-MM-DD HH:MM:SS
    private static final String datetimePattern = "(\\d{4}-\\d{2}-\\d{2} \\d{2}:\\d{2}:\\d{2})";

    // Regular expression for extracting a number before the word 'successful'
    // Assuming the number is followed by ' successful'
    private static final String numberPattern = "(\\d+)\\s+successful";

    // Function to extract data from the log file based on TABLE_NAME column in Excel
    private static String[] extractDataFromLog(String logFilePath, String excelFilePath) throws IOException {
        BufferedReader reader = new BufferedReader(new FileReader(logFilePath));
        List<String> data = new ArrayList<>();
        String line;
        while ((line = reader.readLine()) != null) {
            data.add(line);
        }
        reader.close();

        Pattern datetimePattern = Pattern.compile(Final.datetimePattern);
        Pattern numberPattern = Pattern.compile(Final.numberPattern);

        Matcher firstRecordMatcher = datetimePattern.matcher(data.get(0));
        String firstDatetime = firstRecordMatcher.group(0);

        Matcher lastRecordMatcher = datetimePattern.matcher(data.get(data.size() - 1));
        String lastDatetime = lastRecordMatcher.group(0);

        List<Integer> decryptingIndexes = new ArrayList<>();
        for (int i = 0; i < data.size(); i++) {
            if (data.get(i).contains("Decrypting...")) {
                decryptingIndexes.add(i);
            }
        }

        String decryptingRecord = null;
        for (int idx : decryptingIndexes) {
            if (idx + 1 < data.size()) {
                Matcher match = datetimePattern.matcher(data.get(idx + 1));
                if (match.find()) {
                    decryptingRecord = match.group(0);
                    break;
                }
            }
        }

        String lastSuccessful = null;
        for (int i = data.size() - 1; i >= 0; i--) {
            Matcher match = numberPattern.matcher(data.get(i));
            if (match.find()) {
                lastSuccessful = match.group(1);
                break;
            }
        }

        return new String[]{firstDatetime, lastDatetime, decryptingRecord, lastSuccessful};
    }

    // Function to update Excel file with extracted data
    public static void updateExcelWithData(String table, int[] rowNumbers, String excelFilePath, String firstDatetime,
            String lastDatetime, String decryptingDatetime, String lastSuccessful, int result) throws IOException {
        FileInputStream fileInputStream = new FileInputStream(new File(excelFilePath));
        Workbook workbook = new XSSFWorkbook(fileInputStream);
        Sheet sheet = workbook.getSheetAt(0);

        for (int rowNo : rowNumbers) {
            Row row = sheet.getRow(rowNo - 1);
            if (row == null) {
                row = sheet.createRow(rowNo - 1);
            }

            // Update cell values
            row.createCell(0).setCellValue(firstDatetime);
            row.createCell(7).setCellValue(lastDatetime);
            row.createCell(6).setCellValue(decryptingDatetime);
            row.createCell(5).setCellValue(lastSuccessful);
            row.createCell(1).setCellValue(result);

            if ("SFDC_W_FINANCIAL ACCOUNT".equals(table) || "SFDC_W_FINANCIAL_ACCOUNT_TEAM".equals(table)) {
                row.createCell(4).setCellValue("W");
            } else {
                row.createCell(4).setCellValue("D");
            }
        }

        // Save the changes back to the Excel file
        try (FileOutputStream fileOutputStream = new FileOutputStream(excelFilePath)) {
            workbook.write(fileOutputStream);
        }

        System.out.println("Data successfully updated in Excel file.");
    }


    // Function to count occurrences in zip file names
    private static int countOccurrencesInZipFileNames(String folderPath, String targetString) {
        int occurrencesCount = 0;

        // Iterate over each file in the folder
        File folder = new File(folderPath);
        for (File file : folder.listFiles()) {
            // Check if the file is a zip file
            if (file.isFile() && file.getName().endsWith(".zip")) {
                // Check if the target string is present in the zip file name
                if (file.getName().contains(targetString)) {
                    occurrencesCount++;
                }
            }
        }

        return occurrencesCount;
    }

    // Function to extract zip file
    private static void extractZip(String zipFilePath, String extractFolder) throws IOException {
        try (ZipFile zipFile = new ZipFile(zipFilePath)) {
            Enumeration<? extends ZipEntry> entries = zipFile.entries();

            // Create the extraction folder if it doesn't exist
            Path folderPath = Paths.get(extractFolder);
            Files.createDirectories(folderPath);

            // Unzip the folder
            while (entries.hasMoreElements()) {
                ZipEntry entry = entries.nextElement();
                Path entryPath = folderPath.resolve(entry.getName());

                if (!entry.isDirectory()) {
                    Files.createDirectories(entryPath.getParent());

                    try (InputStream in = zipFile.getInputStream(entry);
                         OutputStream out = new FileOutputStream(entryPath.toFile())) {
                        byte[] buffer = new byte[1024];
                        int bytesRead;
                        while ((bytesRead = in.read(buffer)) != -1) {
                            out.write(buffer, 0, bytesRead);
                        }
                    }
                }
            }
        }
    }

    // Function to read files
    private static List<TableInfo> readFiles(String tableName) throws IOException {
        List<TableInfo> result = new ArrayList<>();

        // Specify the path to the zip file
        String extractFolder = "C:/Users/Windows/Downloads/zipfiles";

        // Iterate over each file in the folder
        File extractFolderFile = new File(extractFolder);
        for (File file : extractFolderFile.listFiles()) {
            if (file.isFile() && file.getName().endsWith(".zip")) {
                // Extract nested zip folder
                String nestedExtractFolder = "./zipfiles/zipfiles/" + file.getName().replace(".zip", "");
                System.out.println("Zip within zip: " + nestedExtractFolder);
                extractZip("./zipfiles/zipfiles/" + file.getName(), nestedExtractFolder);

                // Iterate over files in the nested zip folder
                File nestedFolder = new File(nestedExtractFolder);
                for (File nestedFile : nestedFolder.listFiles()) {
                    String nestedFileName = nestedFile.getName();

                    // Check conditions based on the nested file name and table name
                    if (nestedFileName.contains("sfdcRegistrationMapExtractProcess") && tableName.equals("SFDC_W_TR_REGISTRATION_MAP")) {
                        result.add(new TableInfo(nestedFile.getAbsolutePath(), new int[]{2, 21, 30}));
                    } else if (nestedFileName.contains("sfdcSbrLetterLogRelExtractProcess") && tableName.equals("SBR_W_REG_LETTER_LOG_REL_SFDC")) {
                        result.add(new TableInfo(nestedFile.getAbsolutePath(), new int[]{3, 22, 31}));
                    } else if (nestedFileName.contains("sfdcSponserNamesExtractProcess") && tableName.equals("SFDC_W_SPONSOR_NAMES")) {
                        result.add(new TableInfo(nestedFile.getAbsolutePath(), new int[]{4, 23, 32}));
                    }
                    
                    // Example:
                    // else if (nestedFileName.contains("exampleProcess") && tableName.equals("ExampleTable")) {
                    //     result.add(new TableInfo(nestedFile.getAbsolutePath(), new int[]{/* column numbers */}));
                    // }
                }
            }
        }
		return result;
    }
    private static List<String> getColumnData(Sheet sheet, int columnIndex) {
        List<String> columnData = new ArrayList<>();
        Iterator<Row> rowIterator = sheet.iterator();

        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            Cell cell = row.getCell(columnIndex - 1); // Adjust index to 0-based

            if (cell != null) {
                columnData.add(cell.getStringCellValue());
            } else {
                // Handle null cells if needed
                columnData.add("");
            }
        }

        return columnData;
    }

    public static void main(String[] args) throws IOException {
        // Define the Excel file path
        String excelFilePath = "C:/Users/Windows/Downloads/Uzipthezip.xlsx";

        // Load the workbook
        Workbook wb = WorkbookFactory.create(new File(excelFilePath));
        Sheet sheet = wb.getSheetAt(0);

        // Get table names from Excel sheet
        List<String> tableNames = getColumnData(sheet, 2);

        // Iterate through table names
     // ...

        for (String table : tableNames) {
            String logFilePath = "";
            List<TableInfo> tableInfoList = readFiles(table);

            System.out.println(logFilePath);

            String[] extractedData = extractDataFromLog(logFilePath, excelFilePath);
            String firstDatetime = extractedData[0];
            String lastDatetime = extractedData[1];
            String decryptingDatetime = extractedData[2];
            String lastSuccessful = extractedData[3];

            // Update the Excel file with the extracted data
            if (firstDatetime != null && lastDatetime != null && decryptingDatetime != null && lastSuccessful != null) {
                // Replace 'folderPath' with the path to your folder containing zip files
                String folderPath = "zipfiles/zipfiles";

                String targetString = "";
                List<String> tablesWithList = Arrays.asList("SFDC_W_TR_REGISTRATION_MAP", "SBR_W_REG_LETTER_LOG_REL_SFDC", "SFDC_W_SPONSOR_NAMES");
                List<String> tablesWithArrays = Arrays.asList("SFDC_W_CLIENT", "SFDC_W_REGISTRATION", "SFDC_W_REGISTRATION_MEMBERS", "SFDC_W_BENEFICIARY", "SFDC_W_CLIENT_DISCLOSURE", "SFDC_W_PORTFOLIO_REVIEW");
                List<String> tablesWithEblotter = Arrays.asList("SFDC_EBLOTTER");
                List<String> tablesWithHistory = Arrays.asList("SBR_ACCOUNT_HISTORY_SFDC", "SBR_REG_MEMBER_HISTORY_SFDC", "SBR_REGISTRATION_HISTORY_SFDC", "SBR_REG_LETTER_LOG_SFDC", "SBR_REG_LETTER_LOG_T2_SFDC");
                List<String> tablesWithChecksAndTrades = Arrays.asList("SFDC_CHECKS", "SFDC_TRADES");

                if (tablesWithList.contains(table)) {
                    targetString = "Copy_Tables_extract_log_files";
                } else if (tablesWithArrays.contains(table)) {
                    targetString = "trade_review_extract_log_files";
                } else if (tablesWithEblotter.contains(table)) {
                    targetString = "EblotterExtract__log_files";
                } else if (tablesWithHistory.contains(table)) {
                    targetString = "SFDC_History";
                } else if (tablesWithChecksAndTrades.contains(table)) {
                    targetString = "sfdcEBlotter";
                }

                // Count the occurrences of the target string in zip file names
                int result = countOccurrencesInZipFileNames(folderPath, targetString);

                System.out.printf("The string \"%s\" appears in the names of %d zip files in the folder.%n", targetString, result);

                for (TableInfo tableInfo : tableInfoList) {
                    updateExcelWithData(table, tableInfo.getRowNumbers(), excelFilePath, firstDatetime, lastDatetime, decryptingDatetime, lastSuccessful, result);
                }
            }
        }

    }}
