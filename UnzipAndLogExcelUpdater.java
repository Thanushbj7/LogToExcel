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
