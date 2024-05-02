import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;
import org.apache.poi.ss.usermodel.*;

public class Helper {
    public static void main(String[] args) throws SQLException, IOException {
        String url_local = "jdbc:sqlserver://VENKAT-ASUS-VB1\\SQLEXPRESS\\localhost:1433;databaseName=employee;encrypt=true;trustServerCertificate=true";
        String user_local = "javauser";
        String password_local = "root";

        String inputFolder = "C:\\Users\\Venkat\\Downloads\\tempdocs\\test\\input";
        String processedFolder = "C:\\Users\\Venkat\\Downloads\\tempdocs\\test\\processed";
        String failedFolder = "C:\\Users\\Venkat\\Downloads\\tempdocs\\test\\failed";

        File inputDir = new File(inputFolder);
        File processedDir = new File(processedFolder);
        File failedDir = new File(failedFolder);

        // Create directories if they don't exist
        if (!processedDir.exists()) {
            processedDir.mkdirs();
        }
        if (!failedDir.exists()) {
            failedDir.mkdirs();
        }

        File[] files = inputDir.listFiles();

        if (files == null || files.length == 0) {
            System.out.println("No files found in input folder.");
            return;
        }

        for (File file : files) {
            try {
                if (file.isFile() && (file.getName().endsWith(".xlsx") || file.getName().endsWith(".xls"))) {
                    processExcel(file, processedDir);
                } else if (file.isFile() && file.getName().endsWith(".zip")) {
                    unzipFile(file, inputDir);
                }
            } catch (Exception e) {
                System.err.println("Error processing file: " + file.getName());
                e.printStackTrace();
                moveFile(file, failedDir);
            }
        }
    }

    private static void processExcel(Connection con, File file)
    throws SQLException, FileNotFoundException, IOException {
try (FileInputStream fis = new FileInputStream(file);
        Workbook workbook = WorkbookFactory.create(fis)) {

    Sheet sheet = workbook.getSheetAt(0);

    String filename = file.getName();

    System.out.println("Started Inserting Data into DB...");
    boolean success = true;
    for (int i = 0; i <= sheet.getLastRowNum(); i++) {
        Row currentRow = sheet.getRow(i);
        if (currentRow != null) {
            StringBuilder rowData = new StringBuilder();
            for (int j = 0; j < currentRow.getLastCellNum(); j++) {
                Cell cell = currentRow.getCell(j);
                Object cellValue = getCellValue(cell);
                rowData.append(cellValue).append("\t");
            }
            System.out.println("Row " + (i + 1) + ": " + rowData);
        }
    }

    String destinationFolder = "F:\\Pipline\\processed";
    String destinationFolder_local = "C:\\Users\\Venkat\\Downloads\\tempdocs\\test\\processed";

    if (!success) {
        System.out.println("Data insertion failed for file: " + filename);
        destinationFolder_local = "C:\\Users\\Venkat\\Downloads\\tempdocs\\test\\failed";
    }

    File destination = new File(destinationFolder_local, file.getName());
    if (file.renameTo(destination)) {
        System.out.println("File moved successfully to: " + destination.getAbsolutePath());
    } else {
        System.out.println("Failed to move file to: " + destination.getAbsolutePath());
    }

}
}

private static Object getCellValue(Cell cell) {
if (cell == null) {
    return null; // Return null for null cells
}
switch (cell.getCellType()) {
    case NUMERIC:
        if (DateUtil.isCellDateFormatted(cell)) {
            return cell.getDateCellValue();
        } else {
            return cell.getNumericCellValue();
        }
    case BOOLEAN:
        return cell.getBooleanCellValue();
    default:
        return cell.getStringCellValue(); // Return string for other types of cells
}
}

    private static void unzipFile(File file, File outputDir) throws IOException {
        try (ZipInputStream zis = new ZipInputStream(new FileInputStream(file))) {
            ZipEntry entry = zis.getNextEntry();
            while (entry != null) {
                String fileName = entry.getName();
                File newFile = new File(outputDir, fileName);
                if (entry.isDirectory()) {
                    newFile.mkdirs();
                } else {
                    File parentDir = newFile.getParentFile();
                    if (!parentDir.exists()) {
                        parentDir.mkdirs();
                    }
                    try (FileOutputStream fos = new FileOutputStream(newFile)) {
                        byte[] buffer = new byte[1024];
                        int len;
                        while ((len = zis.read(buffer)) > 0) {
                            fos.write(buffer, 0, len);
                        }
                    }
                }
                zis.closeEntry();
                entry = zis.getNextEntry();
            }
        }
        // Delete the original zip file after extraction
        file.delete();
    }

    private static void moveFile(File file, File destDir) throws IOException {
        File destFile = new File(destDir, file.getName());
        if (!file.renameTo(destFile)) {
            throw new IOException("Failed to move file: " + file.getName());
        }
    }
}
