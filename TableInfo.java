

	public class TableInfo {
	    private String filePath;
	    private int[] rowNumbers;

	    public TableInfo(String filePath, int[] rowNumbers) {
	        this.filePath = filePath;
	        this.rowNumbers = rowNumbers;
	    }

	    public String getFilePath() {
	        return filePath;
	    }

	    public int[] getRowNumbers() {
	        return rowNumbers;
	    }
	}





import java.io.*;
import java.util.zip.*;
import java.io.FileWriter;
import java.util.Scanner;

public class LogFileProcessor {
    public static void main(String[] args) {
        // Provide the path to the zip file
        String zipFilePath = "path/to/your/zipfile.zip";

        // Provide the path where you want to extract the zip file
        String extractFolderPath = "path/to/extract/folder";

        // Provide the path where you want to save the Excel file
        String excelFilePath = "path/to/excel/file.xlsx";

        try {
            // Unzip the specified file
            unzipFile(zipFilePath, extractFolderPath);

            // Create an Excel file and write headers
            createExcelFile(excelFilePath);

            // Process log files
            processLogFiles(extractFolderPath, excelFilePath);

            System.out.println("Processing completed successfully.");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void unzipFile(String zipFilePath, String extractFolderPath) throws IOException {
        try (ZipInputStream zis = new ZipInputStream(new FileInputStream(zipFilePath))) {
            ZipEntry zipEntry = zis.getNextEntry();
            while (zipEntry != null) {
                String filePath = extractFolderPath + File.separator + zipEntry.getName();
                try (BufferedOutputStream bos = new BufferedOutputStream(new FileOutputStream(filePath))) {
                    byte[] bytesIn = new byte[1024];
                    int read;
                    while ((read = zis.read(bytesIn)) != -1) {
                        bos.write(bytesIn, 0, read);
                    }
                }
                zipEntry = zis.getNextEntry();
            }
        }
    }

    private static void createExcelFile(String excelFilePath) throws IOException {
        try (FileWriter writer = new FileWriter(excelFilePath)) {
            // Write headers to the Excel file
            writer.write("RUN_DATE,RUN_CYCLE,TABLE_NAME,WEEKLY_DAILY,RECORD_COUNT,START_DATE,END_DATE\n");
        }
    }

    private static void processLogFiles(String extractFolderPath, String excelFilePath) throws IOException {
        File folder = new File(extractFolderPath);
        File[] listOfFiles = folder.listFiles();

        if (listOfFiles != null) {
            for (File file : listOfFiles) {
                if (file.isFile() && file.getName().endsWith(".log")) {
                    processLogFile(file, excelFilePath);
                }
            }
        }
    }

    private static void processLogFile(File logFile, String excelFilePath) throws IOException {
        try (Scanner scanner = new Scanner(logFile);
             FileWriter writer = new FileWriter(excelFilePath, true)) {
            // Extract details from the log file
            String tableName = logFile.getName().contains("sfdcRegistrationMapExtractProcess") ? "SFDC_W_TR_REGISTRATION_MAP" : "";
            String runDate = scanner.nextLine(); // Assuming the first line contains date and time
            String startDate = "";
            String endDate = "";
            int recordCount = 0;

            while (scanner.hasNextLine()) {
                String line = scanner.nextLine();
                if (line.contains("decrypting...")) {
                    // Extract date and time below the "decrypting..." line
                    startDate = line.substring(0, line.indexOf("decrypting...")).trim();
                }
                // Extract date and time from the last line
                endDate = line;
                if (line.contains("successful")) {
                    // Extract the number before the "successful" string
                    recordCount = Integer.parseInt(line.substring(0, line.indexOf("successful")).trim());
                }
            }

            // Write the extracted details to the Excel file
            writer.write(runDate + "," + "" + "," + tableName + "," + "" + "," + recordCount + "," + startDate + "," + endDate + "\n");
        }
    }
	


}






java.io.FileNotFoundException: H:\ZipFile (Access is denied)
        at java.base/java.io.FileInputStream.open0(Native Method)
        at java.base/java.io.FileInputStream.open(FileInputStream.java:216)  
        at java.base/java.io.FileInputStream.<init>(FileInputStream.java:157)
        at java.base/java.io.FileInputStream.<init>(FileInputStream.java:111)
        at LogFileExtract.unzipFile(LogFileExtract.java:34)
        at LogFileExtract.main(LogFileExtract.java:19)
PS Microsoft.PowerShell.Core\FileSystem::\\pstrwdfs9031\Profiles\G-FR1313-NonPrivileged\i733581\Desktop\POC> 
