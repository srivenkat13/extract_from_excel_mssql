package com.example;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;

import org.apache.poi.ss.usermodel.*;

public class App {
    public static void main(String[] args) throws SQLException, FileNotFoundException, IOException {
        String url = "jdbc:sqlserver://QINFORAPDB1\\localhost:33410;databaseName=InforEAM;encrypt=true;trustServerCertificate=true";
        String user = "svcINFOREAMq";
        String password = "4InforQ5vcsOnl!";

        String url_local = "jdbc:sqlserver://VENKAT-ASUS-VB1\\SQLEXPRESS\\localhost:1433;databaseName=employee;encrypt=true;trustServerCertificate=true";
        String user_local = "javauser";
        String password_local = "root";

        String sqlQuery = "SELECT * FROM developers";

        String inputFolder = "F:\\Pipline\\input";

        String inputFolder_local = "C:\\Users\\Venkat\\Downloads\\tempdocs\\test\\input";

        String tableName = "U5CPDATALOGGERINFO";
        try {
            Connection con = DriverManager.getConnection(url_local, user_local, password_local);
            if (con != null) {
                System.out.println("Connected to DB");
            }

            // // *PRINT from Existing Table */
            // Statement statement = con.createStatement();
            // ResultSet resultSet = statement.executeQuery(sqlQuery);
            // while (resultSet.next()) {
            // // Retrieve values from the current row
            // int id = resultSet.getInt("ID");
            // String firstname = resultSet.getString("FirstName");
            // String lastname = resultSet.getString("LastName");

            // // Do something with the retrieved values (e.g., print them)
            // System.out.println("ID: " + id + ", Name: " + firstname + ", Status:" +
            // lastname);
            // }

            // *Extract from Excel*/
            File folder = new File(inputFolder_local);
            File[] files = folder.listFiles();
            boolean processed = false;

            System.out.println(files);
            System.out.println("File names printed");

            if (files != null && files.length > 0) {
                for (File file : files) {
                    if (file.isFile() && (file.getName().endsWith(".xlsx") ||
                            file.getName().endsWith(".xls"))) {
                        processExcelFile(con, file);
                        processed = true;
                    } else if (file.isFile() && file.getName().endsWith(".zip")) {
                        // unzipFile(file, folder);
                    }
                }
            } else {
                System.out.println("No Excel Files found in the input folder");
            }

            if (processed) {
                System.out.println("Inserted data into the table:" + tableName);
            } else {
                System.out.println("No Excel files are processed");
            }

        } catch (SQLException e) {
            System.out.println(e);
        }

    }

    private static void unzipFile(File file , File outputDir) throws IOException {
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
        catch (FileNotFoundException e) {
            throw e; // Re-throw FileNotFoundException
        } catch (IOException e) {
            throw new IOException("Error unzipping file: " + file.getName(), e);
        }

        // deletes the zip file after extraction 
        // file.delete();

        //moving the extracted files back to input directory
        File extractedDir = new File(outputDir, file.getName().substring(0,file.getName().lastIndexOf('.')));
        if(extractedDir.exists()){
            File[]  extractedFiles= extractedDir.listFiles();
            if(extractedFiles != null) {
                for(File extractedFile : extractedFiles) {
                    moveFromExtracted(extractedFile,outputDir);
                }
            }
            // Delete the extracted  zip directory
            extractedDir.delete(); 
        }        
    }

    private static void moveFromExtracted(File file, File destDir) throws IOException {
        File destFile =  new File(destDir,file.getName());
        if (!file.renameTo(destFile)) {
            throw new IOException("Failed to move file after extraction from zip: " + file.getName());
        }
    }

    private static void processExcelFile(Connection con, File file)
            throws SQLException, FileNotFoundException, IOException {
        try (FileInputStream fis = new FileInputStream(file);
                Workbook workbook = WorkbookFactory.create(fis)) {

            Sheet sheet = workbook.getSheetAt(0);

            Row row = sheet.getRow(1);
            String filename = file.getName();

            // insertStatus(con, filename);
            System.out.println("Started Inserting Data into DB...");
            boolean success = true;
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row currentRow = sheet.getRow(i);
                String loggerID = currentRow.getCell(1).getStringCellValue();
                String testpointID = currentRow.getCell(2).getStringCellValue();
                double yearDouble = currentRow.getCell(3).getNumericCellValue();
                int year = (int) yearDouble;
                double monthDouble = currentRow.getCell(4).getNumericCellValue();
                int month = (int) monthDouble;
                String dateTime = currentRow.getCell(5).getStringCellValue();
                String onPotential = currentRow.getCell(7).getStringCellValue();
                double acVoltage = currentRow.getCell(8).getNumericCellValue();

                try {
                    insertIntoProcessedData(con, loggerID, testpointID, year, month, dateTime, onPotential, acVoltage);
                } catch (SQLException e) {
                    System.out.println("Error inserting data from file: " + filename);
                    e.printStackTrace();
                    success = false;
                    break; // Break out of loop if an error occurs
                }

            }
            String destinationFolder, destinationFolder_local;
            if (success) {
                System.out.println("Data inserted successfully from file: " + filename);
                destinationFolder = "F:\\Pipline\\processed";
                destinationFolder_local = "C:\\Users\\Venkat\\Downloads\\tempdocs\\test\\processed";
            } else {
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

    private static void insertIntoProcessedData(Connection con, String loggerId, String testpointID, int year,
            int month, String dateTime, String onPotential, double acVoltage) throws SQLException {
        String sql = "INSERT INTO U5CPDATALOGGERINFO (CDL_DATALOGGERID, CDL_TESTPOINT, CDL_YEAR, CDL_MONTH, CDL_DATE, CDL_ONPOTENTIAL, CDL_ACVOLTAGE) VALUES(?,?,?,?,?,?,?)";

        String new_sql = "INSERT INTO U5CPDATALOGGERINFO (CDL_SN, CDL_CPNPOTENTIAL,CDL_ONPOTENTIAL,CDL_ACVOLTAGE, CDL_ANODECURRENT,CDL_DCCURRENT,CDL_ACCURRENT,CDL_OFFPOTENTIAL, CDL_DATALOGGERID, CDL_TESTPOINT,CDL_DATE,CDL_MONTH,CDL_YEAR,CREATEDBY,CREATED,UPDATECOUNT) VALUES ((SELECT ISNULL(MAX(CDL_SN),0)+1 FROM U5CPDATALOGGERINFO), NULL, -1248.0 ,0.333, NULL,NULL,NULL,NULL, 'CPCU000785', 'TPTL4-77', CONVERT(DATETIME,'2023-12-30 23:52:00.0'),12,2023, 'R5', GETDATE(),0)";
        try (PreparedStatement pstmnt = con.prepareStatement(sql)) {
            pstmnt.setString(1, loggerId);
            pstmnt.setString(2, testpointID);
            pstmnt.setInt(3, year);
            pstmnt.setInt(4, month);
            pstmnt.setString(5, dateTime);
            pstmnt.setString(6, onPotential);
            pstmnt.setDouble(7, acVoltage);

            pstmnt.executeUpdate();
        }

    }

    private static void insertStatus(Connection con, String filename) throws SQLException {
        String sql = "INSERT INTO file_status (file_name, status) VALUES (?, 'Processed')";
        try (PreparedStatement pstmnt = con.prepareStatement(sql)) {
            pstmnt.setString(1, filename);
            pstmnt.executeUpdate();
        }
    }

}
