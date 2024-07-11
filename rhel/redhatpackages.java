package rhel;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.remote.RemoteWebDriver;
import org.openqa.selenium.support.ui.Select;

import java.io.*;
import java.net.URL;
import java.nio.channels.FileChannel;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;
import java.util.Properties;

public class redhatpackages {

    public static void main(String[] args) throws InterruptedException, IOException, EncryptedDocumentException, InvalidFormatException {
        // Read configuration from properties file
        Properties prop = new Properties();
        InputStream input = null;

        try {
            input = new FileInputStream("C:\\Users\\deben\\eclipse-workspace\\Selenium\\src\\rhel\\rhel_config.properties");
            prop.load(input);
        } catch (IOException ex) {
            ex.printStackTrace();
            return;
        } finally {
            if (input != null) {
                try {
                    input.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }

        // Get username, password, packageUrl, architecture, and output path from properties
        String username = prop.getProperty("username");
        String password = prop.getProperty("password");
        String packageUrl = prop.getProperty("packageUrl");
        String architecture = prop.getProperty("architecture");
        String outputpath = prop.getProperty("outputpath");
        String outputexcelfile = prop.getProperty("outputexcelfile");
        String existingexcelfile = prop.getProperty("existingexcelfile");
        String oldsheet = prop.getProperty("existingsheet");
        String newsheet = prop.getProperty("newsheet");

        // Check if the existing Excel file is locked
        File existingFile = new File(existingexcelfile);
        if (existingFile.exists() && isFileLocked(existingFile)) {
            closeOfficeApplication();
            // Wait a bit to ensure the file is released
            Thread.sleep(5000);
        }

        // Check if the output Excel file is locked
        File outputFile = new File(outputexcelfile);
        if (outputFile.exists() && isFileLocked(outputFile)) {
            closeOfficeApplication();
            // Wait a bit to ensure the file is released
            Thread.sleep(5000);
        }

        WebDriver driver = new ChromeDriver();
        driver.manage().window().maximize();
        driver.navigate().to("https://access.redhat.com/downloads/content/package-browser");
        driver.findElement(By.id("username-verification")).sendKeys(username);
        driver.findElement(By.id("login-show-step2")).click();
        Thread.sleep(2000);
        driver.findElement(By.id("password")).sendKeys(password);
        driver.findElement(By.id("rh-password-verification-submit-button")).click();
        System.out.println("Login Successfully");
        Thread.sleep(10000);
        WebElement consentBannerCloseButton = driver.findElement(By.id("truste-consent-button")); // Adjust the selector as necessary
        consentBannerCloseButton.click();

        driver.navigate().to(packageUrl);
        Thread.sleep(4000);
        WebElement versions = driver.findElement(By.xpath("//select[@id='evr' and @class='select-chosen linked']"));

        ((RemoteWebDriver) driver).executeScript("arguments[0].style.display='block';", versions);
        Select dropdown = new Select(versions);

        List<WebElement> allOptions = dropdown.getOptions();
        List<String> dropdownValues = new ArrayList<>();
        //int c = 0;
        for (WebElement option : allOptions) {
         //   if (c == 8) {
         //       break;
         //   }
            dropdownValues.add(option.getText());
        //    c = c + 1;
        }
        Workbook existingWorkbook = null;
        Sheet existingSheet = null;
        try {
            // Load existing workbook if it exists
            if (existingFile.exists()) {
                existingWorkbook = WorkbookFactory.create(existingFile);
                existingSheet = existingWorkbook.getSheet(oldsheet);
            }
        } catch (IOException ex) {
            ex.printStackTrace();
        }

        Workbook newWorkbook = new XSSFWorkbook();
        Sheet newSheet = newWorkbook.createSheet(newsheet);

        // Create header row for new sheet
        Row newHeaderRow = newSheet.createRow(0);
        Cell newHeaderCell0 = newHeaderRow.createCell(0);
        Cell newHeaderCell1 = newHeaderRow.createCell(1);
        Cell newHeaderCell2 = newHeaderRow.createCell(2);
        Cell newHeaderCell3 = newHeaderRow.createCell(3);
        Cell newHeaderCell4 = newHeaderRow.createCell(4);
        Cell newHeaderCell5 = newHeaderRow.createCell(5);

        newHeaderCell0.setCellValue("Package Name");
        newHeaderCell1.setCellValue("Package Architecture");
        newHeaderCell2.setCellValue("Version");
        newHeaderCell3.setCellValue("OS Type");
        newHeaderCell4.setCellValue("OS Version");
        newHeaderCell5.setCellValue("URL");

        int newRowCount = 1; // Start from row 1 since row 0 is the header row

        // Create the folder where the new files will be downloaded
        File newDownloadDir = new File(outputpath);
        if (!newDownloadDir.exists()) {
            if (newDownloadDir.mkdirs()) {
                System.out.println("Directory created successfully: " + newDownloadDir.getAbsolutePath());
            } else {
                System.out.println("Failed to create directory: " + newDownloadDir.getAbsolutePath());
                return;
            }
        }

        for (String value : dropdownValues) {
            versions = driver.findElement(By.xpath("//select[@id='evr' and @class='select-chosen linked']")); // Re-locate the dropdown element
            ((RemoteWebDriver) driver).executeScript("arguments[0].style.display='block';", versions);
            dropdown = new Select(versions);
            dropdown.selectByVisibleText(value);
            Thread.sleep(6000);

            List<WebElement> links = driver.findElements(By.tagName("a"));
            for (WebElement link : links) {
                String href = link.getAttribute("href");
                if (href != null && href.contains("access.cdn.redhat.com/content/origin/rpms")) {
                    String[] parts = extractParts(href, architecture);

                    // Check if URL exists in the existing Excel sheet
                    boolean urlExists = false;
                    if (existingSheet != null) {
                        for (int i = 1; i <= existingSheet.getLastRowNum(); i++) { // Start from row 1 since row 0 is the header row
                            Row row = existingSheet.getRow(i);
                            if (row != null) {
                                Cell cell = row.getCell(5); // Get the URL cell
                                if (cell != null) {
                                    String existingURL = cell.getStringCellValue();
                                    // Compare up to ".rpm?" to check if URL already exists
                                    if (existingURL != null && existingURL.contains(".rpm?")) {
                                        String existingURLPrefix = existingURL.substring(0, existingURL.indexOf(".rpm?") + 5); // Include ".rpm?"
                                        if (href.startsWith(existingURLPrefix)) {
                                            urlExists = true;
                                            break;
                                        }
                                    }
                                }
                            }
                        }
                    }

                    // Add URL to new Excel sheet and download if it doesn't already exist
                    if (!urlExists) {
                        Row newRow = newSheet.createRow(newRowCount++);
                        Cell newCell0 = newRow.createCell(0);
                        Cell newCell1 = newRow.createCell(1);
                        Cell newCell2 = newRow.createCell(2);
                        Cell newCell3 = newRow.createCell(3);
                        Cell newCell4 = newRow.createCell(4);
                        Cell newCell5 = newRow.createCell(5);

                        newCell0.setCellValue(parts[0]);
                        newCell1.setCellValue(parts[3]);
                        newCell2.setCellValue(parts[1]);
                        newCell3.setCellValue("redhat");
                        newCell4.setCellValue(parts[2]);
                        newCell5.setCellValue(href);

                        // Download the file
                        downloadFile(href, newDownloadDir.getAbsolutePath());
                        // Log after downloading each file
                        System.out.println("Downloaded file from URL: " + href);
                    } else {
                        System.out.println("URL already exists in " + existingexcelfile + " file , skipping: " + href);
                    }
                }
            }

            // Log after processing each version
            System.out.println("Completed processing version: " + value);
        }

        // Save the new Excel file with new download links
        try (FileOutputStream fileOut = new FileOutputStream(outputexcelfile)) {
            newWorkbook.write(fileOut);
        }

        // Close the workbooks
        if (existingWorkbook != null) {
            existingWorkbook.close();
        }
        newWorkbook.close();
        System.out.println("New Excel file created successfully.");
        // Close the WebDriver
        driver.quit();
        System.out.println("Process completed successfully.");
    }

    private static boolean isFileLocked(File file) {
        try (FileChannel channel = new RandomAccessFile(file, "rw").getChannel()) {
            java.nio.channels.FileLock lock = channel.tryLock();
            if (lock == null) {
                return true;
            }
            lock.release();
            return false;
        } catch (IOException e) {
            return true;
        }
    }

    private static void closeOfficeApplication() {
        try {
            // Execute PowerShell command to close all instances of Excel and WPS Office
            String[] commands = {
                "powershell.exe", "-Command",
                "Get-Process excel, et, wps | ForEach-Object { $_.CloseMainWindow() }"
            };
            ProcessBuilder processBuilder = new ProcessBuilder(commands);
            Process process = processBuilder.start();
            process.waitFor();
            System.out.println("Closed office applications (Excel, WPS Office).");
        } catch (IOException | InterruptedException e) {
            e.printStackTrace();
        }
    }

    private static String[] extractParts(String url, String architecture) {
        int rpmIndex = url.lastIndexOf('/') + 1;
        int queryIndex = url.lastIndexOf('?');
        String rpmPart = url.substring(rpmIndex, queryIndex == -1 ? url.length() : queryIndex);

        String[] parts = rpmPart.split("-");
        
//        // Extract "el9_4" part
//        int elIndexStart = url.indexOf("el");
//        int elIndexEnd = url.indexOf("/fd431d51");
//        String part3 = "";
//        if (elIndexStart != -1 && elIndexEnd != -1 && elIndexStart < elIndexEnd) {
//            part3 = url.substring(elIndexStart, elIndexEnd);
//        }
        
        String part3 = "";
        int elIndexStart = -1;
        if (url.contains("NetworkManager-libnm-devel")) {
            // Take the second occurrence of "el"
            elIndexStart = url.indexOf("el", url.indexOf("el") + 1);
        } else {
            // Take the first occurrence of "el"
            elIndexStart = url.indexOf("el");
        }
        if (elIndexStart != -1) {
            int elIndexEnd = url.indexOf("+", elIndexStart);
            if (elIndexEnd == -1) {
                elIndexEnd = url.indexOf("/", elIndexStart);
            }
            if (elIndexEnd != -1) {
                part3 = url.substring(elIndexStart, elIndexEnd);
            }
        }

        String part4 = "";
        if (url.contains(".src.rpm")) {
            part4 = architecture;
        } else {
            int archIndexStart = url.indexOf(architecture);
            int archIndexEnd = url.indexOf(".rpm");
            if (archIndexStart != -1 && archIndexEnd != -1 && archIndexStart < archIndexEnd) {
                part4 = url.substring(archIndexStart, archIndexEnd);
            }
        }

        // Extract version part like "2.4.5-8"
        String part2 = "";
        if (parts.length >= 2) {
            part2 = parts[parts.length - 2] + "-" + parts[parts.length - 1].substring(0, parts[parts.length - 1].indexOf('.'));
        }

        // Construct part1
        StringBuilder part1Builder = new StringBuilder();
        for (int i = 0; i < parts.length - 2; i++) {
            if (i > 0) {
                part1Builder.append("-");
            }
            part1Builder.append(parts[i]);
        }
        String part1 = part1Builder.toString();
        if (url.contains(".src.rpm")) {
            part1 += "-src";
        }
        return new String[]{part1, part2, part3, part4};
    }
    
    private static void downloadFile(String fileURL, String saveDir) throws IOException {
        URL url = new URL(fileURL);
        try (InputStream in = new BufferedInputStream(url.openStream());
             FileOutputStream fileOutputStream = new FileOutputStream(saveDir + "/" + Paths.get(url.getPath()).getFileName().toString())) {
            byte[] dataBuffer = new byte[1024];
            int bytesRead;
            while ((bytesRead = in.read(dataBuffer, 0, 1024)) != -1) {
                fileOutputStream.write(dataBuffer, 0, bytesRead);
            }
        }
    }
}
