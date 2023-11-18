	import java.io.File;
	import java.io.FileInputStream;
	import java.io.FileOutputStream;
	import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.zip.ZipEntry;
	import java.util.zip.ZipInputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
	import java.util.zip.ZipEntry;
	import java.util.zip.ZipInputStream;


public class CopyTableExtract {



	
	
	 public static void main(String[] args) {
	        // Specify the folder containing the zip files
	        String folderPath = "C:/Users/Windows/Downloads/zipfiles";
	        // Specify the path for the Excel file
	        String excelFilePath = "C:/Excel/file.xlsx";

	        // Call the method to unzip files, read logs, and create the Excel file
	        unzipAndReadLogsAndCreateExcel(folderPath, excelFilePath);
	    }

	 private static void unzipAndReadLogsAndCreateExcel(String folderPath, String excelFilePath) {
	        File folder = new File(folderPath);
	        File[] zipFiles = folder.listFiles((dir, name) -> name.toLowerCase().endsWith(".zip"));

	        if (zipFiles != null) {
	            try (Workbook workbook = new XSSFWorkbook()) {
	                // Create a new sheet
	                Sheet sheet = workbook.createSheet("Sheet1");

	                // Create headers
	                Row headerRow = sheet.createRow(0);
	                String[] headers = {"RUN_DATE", "RUN_CYCLE", "TABLE_NAME", "WEEKLY_DAILY", "RECORD_COUNT", "START_DATE", "END_DATE"};
	                for (int i = 0; i < headers.length; i++) {
	                    Cell cell = headerRow.createCell(i);
	                    cell.setCellValue(headers[i]);
	                }

	                int rowNum = 1; // Start writing data from the second row

	                for (File zipFile : zipFiles) {
	                    try (ZipInputStream zis = new ZipInputStream(new FileInputStream(zipFile))) {
	                        byte[] buffer = new byte[1024];
	                        ZipEntry zipEntry = zis.getNextEntry();

	                        while (zipEntry != null) {
	                            String entryName = zipEntry.getName();
	                            File newFile = new File(folder, entryName);

	                            // Create directories if they do not exist
	                            if (zipEntry.isDirectory()) {
	                                newFile.mkdirs();
	                            } else {
	                                // Create parent directories for non-directory entries
	                                new File(newFile.getParent()).mkdirs();

	                                // Extract the file
	                                try (FileOutputStream fos = new FileOutputStream(newFile)) {
	                                    int len;
	                                    while ((len = zis.read(buffer)) > 0) {
	                                        fos.write(buffer, 0, len);
	                                    }
	                                }

	                                // Check if the entry is a log file and does not contain "Truncate" in its name
	                                if (entryName.toLowerCase().endsWith(".log") && !entryName.toLowerCase().contains("truncate")) {
	                                    // Process the log file and extract date and time
	                                    Date runDate = extractDateFromLogFile(newFile);

	                                    // Map table name based on file name
	                                    String tableName = mapTableName(newFile.getName());

	                                    // Extract start date from the line containing "Decrypting..."
	                                    Date startDate = extractStartDate(newFile);

	                                    // Add log data to Excel file
	                                    addLogDataToExcel(sheet, rowNum++, runDate, tableName, startDate);
	                                }
	                            }

	                            // Move to the next entry
	                            zipEntry = zis.getNextEntry();
	                        }

	                        System.out.println("Unzipped and processed logs: " + zipFile.getName());

	                    } catch (IOException e) {
	                        e.printStackTrace();
	                    }
	                
	                }

	                // Write the workbook to a file
	                try (FileOutputStream fileOut = new FileOutputStream(excelFilePath)) {
	                    workbook.write(fileOut);
	                    System.out.println("Excel file created successfully at: " + excelFilePath);
	                }

	            } catch (IOException e) {
	                e.printStackTrace();
	            }
	        } else {
	            System.out.println("No zip files found in the specified folder.");
	        }
	    }

	    private static void readLogFile(File logFile) {
	        // Read and process the log file as needed
	        try (BufferedReader reader = new BufferedReader(new FileReader(logFile))) {
	            String line;
	            while ((line = reader.readLine()) != null) {
	                // Process each line of the log file
	                System.out.println(line);
	            }
	        } catch (IOException e) {
	            e.printStackTrace();
	        }
	    }
	    private static Date extractDateFromLogFile(File logFile) {
	        // Read and process the first line of the log file to extract date and time
	        try (BufferedReader reader = new BufferedReader(new FileReader(logFile))) {
	            String firstLine = reader.readLine();
	            if (firstLine != null) {
	                SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
	                try {
	                    return dateFormat.parse(firstLine);
	                } catch (ParseException e) {
	                    System.out.println("Error parsing date from log file: " + logFile.getName());
	                    e.printStackTrace();
	                }
	            }
	        } catch (IOException e) {
	            e.printStackTrace();
	        }
	        return null;
	    }
	    private static Date extractStartDate(File logFile) {
	        // Read and process the lines of the log file to extract date and time from the line containing "Decrypting..."
	        try (BufferedReader reader = new BufferedReader(new FileReader(logFile))) {
	            String line;
	            boolean decryptingFound=false;
	            while ((line = reader.readLine()) != null) {
	                if (decryptingFound) {
	                	//if(line.length()>=32){
	                		//String dateString=line.substring(0, 13).trim();
	                	//}
	                    // Extract date and time from the line
	                    SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss"); // Adjust the format based on your log
	                    try {
	                        // Assuming the date and time start at index 13 and end at index 32
	                        return dateFormat.parse(line.trim());
	                    } catch (ParseException e) {
	                        System.out.println("Error parsing start date from log file: " + logFile.getName());
	                        e.printStackTrace();
	                    }
	                }
	                if (line.contains("Decrypting...")){
	                	decryptingFound=true;
	                }
	            }
	        } catch (IOException e) {
	            e.printStackTrace();
	        }
	        return null;
	    }
	    private static String mapTableName(String fileName) {
	        // Map the table name based on the file name
	    	if (fileName.toLowerCase().contains("sfdcregistrationmapextractprocess")) {
	            return "SFDC_W_TR_REGISTRATION_MAP";
	        }else if (fileName.contains("sfdcSbrLetterLogRelExtractProcess")) {
	            return "SBR_W_REG_LETTER_LOG_REL_SFDC";
	        }else if (fileName.contains("sfdcSponserNamesExtractProcess")) {
	            return "SFDC_W_SPONSOR_NAMES";
	        }else if (fileName.contains("sfdcClientExtractProcess")) {
	            return "SFDC_W_CLIENT";
	        }else if (fileName.contains("sfdcRegistrationExtractProcess")) {
	            return "SFDC_W_REGISTRATION";
	        }else if (fileName.contains("oracleEblotterExtractProcess")) {
	            return "SFDC_EBLOTTER";
	        }else if (fileName.contains("sfdcEBlottereBlotterExtractProcess")) {
	            return "SFDC_EBLOTTER";
	        }else if (fileName.contains("sfdcRegMemberExtractProcess")) {
	            return "SFDC_W_REGISTRATION_MEMBERS";
	        }else if (fileName.contains("sfdcRegBeneficiaryExtractProcess")) {
	            return "SFDC_W_BENEFICIARY";
	        }else if (fileName.contains("sfdcClientDisclosureExtractProcess")) {
	            return "SFDC_W_CLIENT_DISCLOSURE";
	        }else if (fileName.contains("sfdcPortfolioReviewExtractProcess")) {
	            return "SFDC_W_PORTFOLIO_REVIEW";
	        }else if (fileName.contains("SFDCHistoryAccountHistoryExtract")) {
	            return "SBR_ACCOUNT_HISTORY_SFDC";
	        }else if (fileName.contains("SFDCHistoryRegClientmemberHistoryExtract")) {
	            return "SBR_REG_MEMBER_HISTORY_SFDC";
	        }else if (fileName.contains("SFDCHistoryRegistrationHistoryExtract")) {
	            return "SBR_REGISTRATION_HISTORY_SFDC";
	        }else if (fileName.contains("SFDCHistoryRegistrationLogExtract")) {
	            return "SBR_REG_LETTER_LOG_SFDC";
	        }else if (fileName.contains("SFDCHistoryRegistrationLogtable_T2_Extract")) {
	            return "SBR_REG_LETTER_LOG_T2_SFDC";
	        }else if (fileName.contains("sfdcEBlotterChecksExtractProcess")) {
	            return "SFDC_CHECKS";
	        }else if (fileName.contains("sfdcEBlotterTradesExtractProcess")) {
	            return "SFDC_TRADES";
	        }else if (fileName.toLowerCase().contains("sfdc_emailload_log")) {
	            return "SFDC_USER";
	        }else if (fileName.contains("sfdcFinancialAccountExtractProcess")) {
	            return "SFDC_W_FINANCIAL_ACCOUNT";
	        }else if (fileName.contains("sfdcFinancialAccountTeamExtractProcess")) {
	            return "SFDC_W_FINANCIAL_ACCOUNT_TEAM";
	        }
	        // Add more conditions if needed for other file names
	        // Example: if (fileName.contains("someOtherPattern")) return "SomeOtherTable";

	        // Default to an empty string if no mapping is found
	        return "";
	    }

	    private static void addLogDataToExcel(Sheet sheet, int rowNum, Date runDate, String tableName,Date startDate) {
	        Row row = sheet.createRow(rowNum);

	        // Add the extracted run date to the RUN_DATE column
	        Cell cellRunDate = row.createCell(0);
	        if (runDate != null) {
	            SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
	            cellRunDate.setCellValue(dateFormat.format(runDate));
	        }

	        // Add the mapped table name to the TABLE_NAME column
	        Cell cellTableName = row.createCell(2);
	        cellTableName.setCellValue(tableName);
	    }
	}

	


	
	    
	

