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


	
	public class UnzipAndLogExcelUpdater {

	    static String zipDirectory = "C:/Users/Windows/Documents/workspace/zipfiles";

	    static String logFilePath = "";

	    static String excelFilePath = "C:/Users/Windows/Documents/workspace/Uzipthezip";

	    static String datetimePattern = "(\\d{4}-\\d{2}-\\d{2} \\d{2}:\\d{2}:\\d{2})";

	    static String numberPattern = "(\\d+)\\s+successful";

	    // Function to extract data from the log file based on TABLE_NAME column in
	    // Excel
	    static List<String> extractDataFromLog(String zipDirectory, String excelFilePath) {
	        List<String> extractedData = new ArrayList<>();
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
	                                try(
	                                BufferedReader reader = new BufferedReader(new InputStreamReader(zis))){
	                                String line;
	                                String firstRecord = null;
	                                String lastRecord = null;
	                                String decryptingRecord = null;
	                                String lastSuccessful = null;
	                                while ((line = reader.readLine()) != null) {
	                                    Pattern pattern = Pattern.compile(datetimePattern);
	                                    Matcher matcher = pattern.matcher(line);
	                                    System.out.println(line);
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
	                                
	                               // reader.close();
	                                
	                                extractedData.add(firstRecord);
	                                extractedData.add(lastRecord);
	                                extractedData.add(decryptingRecord);
	                                extractedData.add(lastSuccessful);
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
	            Cell cell2 = row.createCell(7);
	            cell2.setCellValue(extractedData.get(1));
	            Cell cell3 = row.createCell(6);
	            cell3.setCellValue(extractedData.get(2));
	            Cell cell4 = row.createCell(5);
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

	    public static void main(String[] args) throws IOException{

	        String[] tableNames = { "SFDC_W_TR_REGISTRATION_MAP", "SBR_W_REG_LETTER_LOG_REL_SFDC", "SFDC_W_SPONSOR_NAMES",
	                "SFDC_EBLOTTER", "SBR_ACCOUNT_HISTORY_SFDC", "SBR_REG_MEMBER_HISTORY_SFDC",
	                "SBR_REGISTRATION_HISTORY_SFDC", "SBR_REG_LETTER_LOG_SFDC", "SBR_REG_LETTER_LOG_T2_SFDC",
	                "SFDC_CHECKS", "SFDC_TRADES", "SFDC_W_CLIENT", "SFDC_W_BENEFICIARY", "SFDC_W_CLIENT_DISCLOSURE",
	                "SFDC_W_REGISTRATION", "SFDC_W_REGISTRATION_MEMBERS", "SFDC_W_PORTFOLIO_REVIEW", "SFDC_USER" };
	        int rowNumber = 2;
	        for (String tableName : tableNames) {
	            if (tableName.equals("SFDC_W_TR_REGISTRATION_MAP")) {
	                logFilePath = "sfdcRegistrationMapExtractProcess20231026.20001698364809.log";
	            } else if (tableName.equals("SBR_W_REG_LETTER_LOG_REL_SFDC")) {
	                logFilePath = "sfdcSbrLetterLogRelExtractProcess20231026.20001698364809.log";
	                rowNumber = 3;
	            } else if (tableName.equals("SFDC_W_SPONSOR_NAMES")) {
	                logFilePath = "sfdcSponserNamesExtractProcess20231026.20001698364809.log";
	                rowNumber = 4;
	            } else if (tableName.equals("SFDC_W_CLIENT")) {
	                logFilePath = "sfdcClientExtractProcess20231026.20001698364837.log";
	                rowNumber = 5;
	            } else if (tableName.equals("SFDC_W_REGISTRATION")) {
	                logFilePath = "sfdcRegistrationExtractProcess20231026.20001698364837.log";
	                rowNumber = 6;
	            } else if (tableName.equals("SFDC_W_REGISTRATION_MEMBERS")) {
	                logFilePath = "sfdcRegMemberExtractProcess20231026.20001698364837.log";
	                rowNumber = 7;
	            } else if (tableName.equals("SFDC_W_BENEFICIARY")) {
	                logFilePath = "sfdcRegBeneficiaryExtractProcess20231026.20001698364837.log";
	                rowNumber = 8;
	            } else if (tableName.equals("SFDC_W_CLIENT_DISCLOSURE")) {
	                logFilePath = "sfdcClientDisclosureExtractProcess20231026.20001698364837.log";
	                rowNumber = 9;
	            } else if (tableName.equals("SFDC_W_PORTFOLIO_REVIEW")) {
	                logFilePath = "sfdcPortfolioReviewExtractProcess20231026.20001698364837.log";
	                rowNumber = 10;
	            } else if (tableName.equals("SBR_ACCOUNT_HISTORY_SFDC")) {
	                logFilePath = "SFDCHistoryAccountHistoryExtract.log";
	                rowNumber = 11;
	            } else if (tableName.equals("SBR_REG_MEMBER_HISTORY_SFDC")) {
	                logFilePath = "SFDCHistoryRegClientmemberHistoryExtract.log";
	                rowNumber = 12;
	            } else if (tableName.equals("SBR_REGISTRATION_HISTORY_SFDC")) {
	                logFilePath = "SFDCHistoryRegistrationHistoryExtract.log";
	                rowNumber = 13;
	            } else if (tableName.equals("SBR_REG_LETTER_LOG_SFDC")) {
	                logFilePath = "SFDCHistoryRegistrationLogExtract.log";
	                rowNumber = 14;
	            } else if (tableName.equals("SBR_REG_LETTER_LOG_T2_SFDC")) {
	                logFilePath = "SFDCHistoryRegistrationLogtable_T2_Extract.log";
	                rowNumber = 15;
	            } else if (tableName.equals("SFDC_EBLOTTER")) {
	                logFilePath = "oracleEblotterExtractProcess20231027.02001698386405.log";
	                rowNumber = 16;
	            } else if (tableName.equals("SFDC_EBLOTTER")) {
	                logFilePath = "oracleEblotterExtractProcess20231027.02001698386405.log";
	                rowNumber = 17;
	            } else if (tableName.equals("SFDC_CHECKS")) {
	                logFilePath = "sfdcEBlotterChecksExtractProcess20231027.03001698390009.log";
	                rowNumber = 18;
	            } else if (tableName.equals("SFDC_TRADES")) {
	                logFilePath = "sfdcEBlotterTradesExtractProcess20231027.03001698390009.log";
	                rowNumber = 19;
	            } else if (tableName.equals("SFDC_W_TR_REGISTRATION_MAP")) {
	                logFilePath = "sfdcRegistrationMapExtractProcess20231026.20001698364809.log";
	                rowNumber = 21;
	            } else if (tableName.equals("SBR_W_REG_LETTER_LOG_REL_SFDC")) {
	                logFilePath = "sfdcSbrLetterLogRelExtractProcess20231026.20001698364809.log";
	                rowNumber = 22;
	            } else if (tableName.equals("SFDC_W_SPONSOR_NAMES")) {
	                logFilePath = "sfdcSponserNamesExtractProcess20231026.20001698364809.log";
	                rowNumber = 23;
	            } else if (tableName.equals("SFDC_W_CLIENT")) {
	                logFilePath = "sfdcClientExtractProcess20231026.20001698364837.log";
	                rowNumber = 24;
	            } else if (tableName.equals("SFDC_W_REGISTRATION")) {
	                logFilePath = "sfdcRegistrationExtractProcess20231026.20001698364837.log";
	                rowNumber = 25;
	            } else if (tableName.equals("SFDC_W_REGISTRATION_MEMBERS")) {
	                logFilePath = "sfdcRegMemberExtractProcess20231026.20001698364837.log";
	                rowNumber = 26;
	            } else if (tableName.equals("SFDC_W_BENEFICIARY")) {
	                logFilePath = "sfdcRegBeneficiaryExtractProcess20231026.20001698364837.log";
	                rowNumber = 27;
	            } else if (tableName.equals("SFDC_W_CLIENT_DISCLOSURE")) {
	                logFilePath = "sfdcClientDisclosureExtractProcess20231026.20001698364837.log";
	                rowNumber = 28;
	            } else if (tableName.equals("SFDC_W_PORTFOLIO_REVIEW")) {
	                logFilePath = "sfdcPortfolioReviewExtractProcess20231026.20001698364837.log";
	                rowNumber = 29;
	            } else if (tableName.equals("SFDC_W_TR_REGISTRATION_MAP")) {
	                logFilePath = "sfdcRegistrationMapExtractProcess20231026.20001698364809.log";
	                rowNumber = 30;
	            } else if (tableName.equals("SBR_W_REG_LETTER_LOG_REL_SFDC")) {
	                logFilePath = "sfdcSbrLetterLogRelExtractProcess20231026.20001698364809.log";
	                rowNumber = 31;
	            } else if (tableName.equals("SFDC_W_SPONSOR_NAMES")) {
	                logFilePath = "sfdcSponserNamesExtractProcess20231026.20001698364809.log";
	                rowNumber = 32;
	            } else if (tableName.equals("SFDC_W_CLIENT")) {
	                logFilePath = "sfdcClientExtractProcess20231026.20001698364837.log";
	                rowNumber = 33;
	            } else if (tableName.equals("SFDC_W_REGISTRATION")) {
	                logFilePath = "sfdcRegistrationExtractProcess20231026.20001698364837.log";
	                rowNumber = 34;
	            } else if (tableName.equals("SFDC_W_REGISTRATION_MEMBERS")) {
	                logFilePath = "sfdcRegMemberExtractProcess20231026.20001698364837.log";
	                rowNumber = 35;
	            } else if (tableName.equals("SFDC_W_BENEFICIARY")) {
	                logFilePath = "sfdcRegBeneficiaryExtractProcess20231026.20001698364837.log";
	                rowNumber = 36;
	            } else if (tableName.equals("SFDC_W_CLIENT_DISCLOSURE")) {
	                logFilePath = "sfdcClientDisclosureExtractProcess20231026.20001698364837.log";
	                rowNumber = 37;
	            } else if (tableName.equals("SFDC_W_PORTFOLIO_REVIEW")) {
	                logFilePath = "sfdcPortfolioReviewExtractProcess20231026.20001698364837.log";
	                rowNumber = 38;
	            } else {
	                continue;
	            }
	            List<String> extractedData = extractDataFromLog(zipDirectory, excelFilePath);
	            if (!extractedData.isEmpty()) {
	                updateExcelWithData(rowNumber, excelFilePath, extractedData);
	            }
	        }
	    }

	}








static void updateExcelWithData(int rowNumber, String excelFilePath, List<String> extractedData, String runCycle) {
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
        Cell cell5 = row.createCell(2);
        cell5.setCellValue(runCycle); // Set runCycle value in the Excel file
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
    // ... (existing code)

    // Create a map to associate table names with row numbers
    Map<String, Integer> tableNameToRowNumber = new HashMap<>();
    tableNameToRowNumber.put("SFDC_W_TR_REGISTRATION_MAP", 21);
    tableNameToRowNumber.put("SBR_W_REG_LETTER_LOG_REL_SFDC", 22);
    tableNameToRowNumber.put("SFDC_W_SPONSOR_NAMES", 23);
    // ... (add more entries for other tables)

    List<String> logFileName = extractDataFromLog(zipDirectory, excelFilePath);

    for (String tableName : tableNames) {
        if (logFileName.contains("sfdcRegistrationMapExtractProcess")) {
            tableName.equals("SFDC_W_TR_REGISTRATION_MAP");
        } else if (logFileName.contains("sfdcSbrLetterLogRelExtractProcess")) {
            tableName.equals("SBR_W_REG_LETTER_LOG_REL_SFDC");
        } else if (logFileName.contains("sfdcSponserNamesExtractProcess")) {
            tableName.equals("SFDC_W_SPONSOR_NAMES");
        }
        // ... (add more conditions for other tables)

        // Retrieve the row number based on the table name
        Integer rowNumber = tableNameToRowNumber.get(tableName);

        if (rowNumber != null) {
            List<String> extractedData = extractDataFromLog(zipDirectory, excelFilePath);
            if (!extractedData.isEmpty()) {
                updateExcelWithData(rowNumber, excelFilePath, extractedData, runCycle);
            }
        }
    }
}




public static void main(String[] args) {
    // ... (existing code)

    // Create a map to associate table names with runCycle values
    Map<String, String> tableNameToRunCycle = new HashMap<>();
    tableNameToRunCycle.put("SFDC_W_TR_REGISTRATION_MAP", "21");
    tableNameToRunCycle.put("SBR_W_REG_LETTER_LOG_REL_SFDC", "22");
    tableNameToRunCycle.put("SFDC_W_SPONSOR_NAMES", "23");
    // ... (add more entries for other tables)

    List<String> logFileName = extractDataFromLog(zipDirectory, excelFilePath);

    for (String tableName : tableNames) {
        if (logFileName.contains("sfdcRegistrationMapExtractProcess")) {
            tableName.equals("SFDC_W_TR_REGISTRATION_MAP");
        } else if (logFileName.contains("sfdcSbrLetterLogRelExtractProcess")) {
            tableName.equals("SBR_W_REG_LETTER_LOG_REL_SFDC");
        } else if (logFileName.contains("sfdcSponserNamesExtractProcess")) {
            tableName.equals("SFDC_W_SPONSOR_NAMES");
        }
        // ... (add more conditions for other tables)

        // Retrieve the runCycle value based on the table name
        String runCycle = tableNameToRunCycle.get(tableName);

        if (runCycle != null) {
            List<String> extractedData = extractDataFromLog(zipDirectory, excelFilePath);
            if (!extractedData.isEmpty()) {
                updateExcelWithData(runCycle, excelFilePath, extractedData);
            }
        }
    }
}

static void updateExcelWithData(String runCycle, String excelFilePath, List<String> extractedData) {
    try {
        FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
        Workbook workbook = new XSSFWorkbook(inputStream);
        Sheet sheet = workbook.getSheetAt(0);

        // Find the row based on the runCycle value
        int rowIndex = findRowIndex(sheet, runCycle);

        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            row = sheet.createRow(rowIndex);
        }

        // ... (existing code to update cells)

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

static int findRowIndex(Sheet sheet, String runCycle) {
    // Implement logic to find the row index based on the runCycle value
    // For example, you can loop through the rows and check a specific column for the runCycle value
    // Return the index of the row when you find a match
    return 0; // Replace this with your actual implementation
}




import java.io.*;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class UnzipAndLogExcelUpdater {

    static String zipDirectory = "Users/i733581/Workspace/ZipFile";

    static String excelFilePath = "Users/i733581/Workspace/SFDC_COPY_TABLE_COUNT_Blank_Worksheet.xlsx";

    static String datetimePattern = "(\\d{4}-\\d{2}-\\d{2} \\d{2}:\\d{2}:\\d{2})";

    static String numberPattern = "(\\d+)\\s+successful";
    static List<String> numberOfCopyTablesextractlogfilesZipFiles = null;
    static List<String> numberOfEblotterExtractlogfilesZipFiles = null;
    static List<String> numberOfsfdcemailloadlogZipFiles = null;
    static List<String> numberOfSFDCHistoryZipFiles = null;
    static List<String> numberOfsfdcEBlotterZipFiles = null;
    static List<String> numberOftradereviewextractlogZipFiles = null;

    // Function to extract data from the log file based on TABLE_NAME column in
    // Excel
    static List<String> extractDataFromLog(String zipDirectory, String excelFilePath) {
        List<String> extractedData = new ArrayList<>();
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
                                String logFileName = entry.getName();
                                String zipFileName = zipFile.getName();

                                BufferedReader reader = new BufferedReader(new InputStreamReader(zis));
                                String line;
                                String firstRecord = null;
                                String lastRecord = null;
                                String decryptingRecord = null;
                                String lastSuccessful = null;
                                String runCycle = null;

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
                                    if (zipFileName.contains("Copy_Tables_extract_log_files")) {
                                        numberOfCopyTablesextractlogfilesZipFiles.add("Copy_Tables_extract_log_files");
                                        runCycle = Integer.toString(numberOfCopyTablesextractlogfilesZipFiles.size());
                                    }
                                    if (zipFileName.contains("EblotterExtract__log_files")) {
                                        numberOfEblotterExtractlogfilesZipFiles.add("EblotterExtract__log_files");
                                        runCycle = Integer.toString(numberOfEblotterExtractlogfilesZipFiles.size());
                                    }
                                    if (zipFileName.contains("sfdc_emailload_log")) {
                                        numberOfsfdcemailloadlogZipFiles.add("sfdc_emailload_log");
                                        runCycle = Integer.toString(numberOfsfdcemailloadlogZipFiles.size());
                                    }
                                    if (zipFileName.contains("SFDC_History")) {
                                        numberOfSFDCHistoryZipFiles.add("SFDC_History");
                                        runCycle = Integer.toString(numberOfSFDCHistoryZipFiles.size());
                                    }
                                    if (zipFileName.contains("sfdcEBlotter")) {
                                        numberOfsfdcEBlotterZipFiles.add("sfdcEBlotter");
                                        runCycle = Integer.toString(numberOfsfdcEBlotterZipFiles.size());
                                    }
                                    if (zipFileName.contains("trade_review_extract_log_files")) {
                                        numberOftradereviewextractlogZipFiles.add("trade_review_extract_log_files");
                                        runCycle = Integer.toString(numberOftradereviewextractlogZipFiles.size());
                                    }

                                }
                                reader.close();
                                extractedData.add(firstRecord);
                                extractedData.add(lastRecord);
                                extractedData.add(decryptingRecord);
                                extractedData.add(lastSuccessful);
                                // extractedData.add(4, runCycle);
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
        return extractedData;
    }

    // Function to update Excel file with extracted data
    static void updateExcelWithData(int rowNumber, String excelFilePath, List<String> extractedData, String runCycle) {
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
            Cell cell5 = row.createCell(2);
            cell5.setCellValue(runCycle); // Set runCycle value in the Excel file
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

        String[] tableNames = { "SFDC_W_TR_REGISTRATION_MAP", "SBR_W_REG_LETTER_LOG_REL_SFDC", "SFDC_W_SPONSOR_NAMES",
                "SFDC_EBLOTTER", "SBR_ACCOUNT_HISTORY_SFDC", "SBR_REG_MEMBER_HISTORY_SFDC",
                "SBR_REGISTRATION_HISTORY_SFDC", "SBR_REG_LETTER_LOG_SFDC", "SBR_REG_LETTER_LOG_T2_SFDC",
                "SFDC_CHECKS", "SFDC_TRADES", "SFDC_W_CLIENT", "SFDC_W_BENEFICIARY", "SFDC_W_CLIENT_DISCLOSURE",
                "SFDC_W_REGISTRATION", "SFDC_W_REGISTRATION_MEMBERS", "SFDC_W_PORTFOLIO_REVIEW", "SFDC_USER" };
               
                List<String> logFileName=extractDataFromLog(zipDirectory,excelFilePath);
                //logFilePath=extractDataFromLog(). entry.getName();
        int rowNumber = 2;
        // String tableNameAsLogFileName = "sfdcRegistrationMapExtractProcess";
        for (String tableName : tableNames) {
            if (logFileName.contains("sfdcRegistrationMapExtractProcess")) {
                tableName.equals("SFDC_W_TR_REGISTRATION_MAP");
            } else if (logFileName.contains("sfdcSbrLetterLogRelExtractProcess")) {
                tableName.equals("SBR_W_REG_LETTER_LOG_REL_SFDC");
                rowNumber = 3;
            } else if (logFileName.contains("sfdcSponserNamesExtractProcess")) {
                tableName.equals("SFDC_W_SPONSOR_NAMES");
                rowNumber = 4;
            } else if (logFileName.contains("sfdcClientExtractProcess")) {
                tableName.equals("SFDC_W_CLIENT");

                rowNumber = 5;
            } else if (logFileName.contains("sfdcRegistrationExtractProcess")) {
                tableName.equals("SFDC_W_REGISTRATION");

                rowNumber = 6;
            } else if (logFileName.contains("sfdcRegMemberExtractProcess")) {
                tableName.equals("SFDC_W_REGISTRATION_MEMBERS");

                rowNumber = 7;
            } else if (logFileName.contains("sfdcRegBeneficiaryExtractProcess")) {
                tableName.equals("SFDC_W_BENEFICIARY");

                rowNumber = 8;
            } else if (logFileName.contains("sfdcClientDisclosureExtractProcess")) {
                tableName.equals("SFDC_W_CLIENT_DISCLOSURE");

                rowNumber = 9;
            } else if (logFileName.contains("sfdcPortfolioReviewExtractProcess")) {
                tableName.equals("SFDC_W_PORTFOLIO_REVIEW");

                rowNumber = 10;
            } else if (logFileName.contains("SFDCHistoryAccountHistoryExtract")) {
                tableName.equals("SBR_ACCOUNT_HISTORY_SFDC");

                rowNumber = 11;
            } else if (logFileName.contains("SFDCHistoryRegClientmemberHistoryExtract")) {
                tableName.equals("SBR_REG_MEMBER_HISTORY_SFDC");

                rowNumber = 12;
            } else if (logFileName.contains("SFDCHistoryRegistrationHistoryExtract")) {
                tableName.equals("SBR_REGISTRATION_HISTORY_SFDC");

                rowNumber = 13;
            } else if (logFileName.contains("SFDCHistoryRegistrationLogExtract")) {
                tableName.equals("SBR_REG_LETTER_LOG_SFDC");

                rowNumber = 14;
            } else if (logFileName.contains("SFDCHistoryRegistrationLogtable_T2_Extract")) {
                tableName.equals("SBR_REG_LETTER_LOG_T2_SFDC");

                rowNumber = 15;
            } else if (logFileName.contains("oracleEblotterExtractProcess")) {
                tableName.equals("SFDC_EBLOTTER");

                rowNumber = 16;
            } else if (logFileName.contains("oracleEblotterExtractProcess")) {
                tableName.equals("SFDC_EBLOTTER");

                rowNumber = 17;
            } else if (logFileName.contains("sfdcEBlotterChecksExtractProcess")) {
                tableName.equals("SFDC_CHECKS");

                rowNumber = 18;
            } else if (logFileName.contains("sfdcEBlotterTradesExtractProcess")) {
                tableName.equals("SFDC_TRADES");

                rowNumber = 19;
            } else if (logFileName.contains("sfdcRegistrationMapExtractProcess")) {
                tableName.equals("SFDC_W_TR_REGISTRATION_MAP");

                rowNumber = 21;
            } else if (logFileName.contains("sfdcSbrLetterLogRelExtractProcess")) {
                tableName.equals("SBR_W_REG_LETTER_LOG_REL_SFDC");

                rowNumber = 22;
            } else if (logFileName.contains("sfdcSponserNamesExtractProcess")) {
                tableName.equals("SFDC_W_SPONSOR_NAMES");

                rowNumber = 23;
            } else if (logFileName.contains("sfdcClientExtractProcess")) {
                tableName.equals("SFDC_W_CLIENT");

                rowNumber = 24;
            } else if (logFileName.contains("sfdcRegistrationExtractProcess")) {
                tableName.equals("SFDC_W_REGISTRATION");

                rowNumber = 25;
            } else if (logFileName.contains("sfdcRegMemberExtractProcess")) {
                tableName.equals("SFDC_W_REGISTRATION_MEMBERS");

                rowNumber = 26;
            } else if (logFileName.contains("sfdcRegBeneficiaryExtractProcess")) {
                tableName.equals("SFDC_W_BENEFICIARY");

                rowNumber = 27;
            } else if (logFileName.contains("sfdcClientDisclosureExtractProcess")) {
                tableName.equals("SFDC_W_CLIENT_DISCLOSURE");

                rowNumber = 28;
            } else if (logFileName.contains("sfdcPortfolioReviewExtractProcess")) {
                tableName.equals("SFDC_W_PORTFOLIO_REVIEW");

                rowNumber = 29;
            } else if (logFileName.contains("sfdcRegistrationMapExtractProcess")) {
                tableName.equals("SFDC_W_TR_REGISTRATION_MAP");

                rowNumber = 30;
            } else if (logFileName.contains("sfdcSbrLetterLogRelExtractProcess")) {
                tableName.equals("SBR_W_REG_LETTER_LOG_REL_SFDC");

                rowNumber = 31;
            } else if (logFileName.contains("sfdcSponserNamesExtractProcess")) {
                tableName.equals("SFDC_W_SPONSOR_NAMES")
                
                rowNumber = 32;
            } else if (logFileName.contains("sfdcClientExtractProcess")) {
                tableName.equals("SFDC_W_CLIENT");
                
                rowNumber = 33;
            } else if (logFileName.contains("sfdcRegistrationExtractProcess")) {
                tableName.equals("SFDC_W_REGISTRATION");
                
                rowNumber = 34;
            } else if (logFileName.contains("sfdcRegMemberExtractProcess")) {
                tableName.equals("SFDC_W_REGISTRATION_MEMBERS");
                
                rowNumber = 35;
            } else if (logFileName.contains("sfdcRegBeneficiaryExtractProcess")) {
                tableName.equals("SFDC_W_BENEFICIARY");
                
                rowNumber = 36;
            } else if (logFileName.contains("sfdcClientDisclosureExtractProcess")) {
                tableName.equals("SFDC_W_CLIENT_DISCLOSURE");
                
                rowNumber = 37;
            } else if (logFileName.contains("sfdcPortfolioReviewExtractProcess")) {
                tableName.equals("SFDC_W_PORTFOLIO_REVIEW");
                
                rowNumber = 38;
            } else {
                continue;
            }
            List<String> extractedData = extractDataFromLog(zipDirectory, excelFilePath);
            if (!extractedData.isEmpty()) {
                updateExcelWithData(rowNumber, excelFilePath, extractedData, runCycle);
            }
        }
    }

}
