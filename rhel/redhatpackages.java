package rhel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
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
import java.util.concurrent.TimeUnit;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class redhatpackages {

	 @SuppressWarnings("deprecation")
		public static void main(String[] args) throws InterruptedException, IOException, EncryptedDocumentException, InvalidFormatException {
	        // Read configuration from properties file
	        Properties prop = new Properties();
	        InputStream input = null;

	        try {
	            input = new FileInputStream("C:\\Users\\deben\\eclipse-workspace\\rhelpackages\\src\\redhat\\temporary.properties");
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

	        // Get common properties from config file
	        String username = prop.getProperty("username");
	        String password = prop.getProperty("password");
	        String packageUrls = prop.getProperty("packageUrls");
	        String baseOutputPath = prop.getProperty("outputpath");
	        String existingExcelFile = prop.getProperty("existingexcelfile");
	        String existingSheetName = prop.getProperty("existingsheet");

	        WebDriver driver = new ChromeDriver();
	        driver.manage().window().maximize();
	        driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
	        driver.navigate().to("https://access.redhat.com/downloads/content/package-browser");
	        driver.findElement(By.id("username-verification")).sendKeys(username);
	        driver.findElement(By.id("login-show-step2")).click();
	        Thread.sleep(2000);
	        driver.findElement(By.id("password")).sendKeys(password);
	        driver.findElement(By.id("rh-password-verification-submit-button")).click();
	        System.out.println("Login Successfully");
	        Thread.sleep(10000);
	        WebElement consentBannerCloseButton = driver.findElement(By.id("truste-consent-button"));
	        consentBannerCloseButton.click();

	        Workbook existingWorkbook = null;
	        Sheet existingSheet = null;
	        Sheet errorSheet = null;
	        File existingFile = new File(existingExcelFile);

	        // Check if the existing Excel file is locked
	        if (existingFile.exists() && isFileLocked(existingFile)) {
	            closeOfficeApplication();
	            // Wait a bit to ensure the file is released
	            Thread.sleep(5000);
	        }

	        try {
	            // Load existing workbook if it exists
	            if (existingFile.exists()) {
	                existingWorkbook = new XSSFWorkbook(new FileInputStream(existingFile));
	                existingSheet = existingWorkbook.getSheet(existingSheetName);
	                errorSheet = existingWorkbook.getSheet("No Data");
	            }

	            // If the existing sheet does not exist, create a new sheet
	            if (existingSheet == null) {
	                if (existingWorkbook == null) {
	                    existingWorkbook = new XSSFWorkbook(); // Create new workbook if it doesn't exist
	                }
	                existingSheet = existingWorkbook.createSheet(existingSheetName);
	                // Optionally create header row
	                Row headerRow = existingSheet.createRow(0);
	                headerRow.createCell(0).setCellValue("Package Name");
	                headerRow.createCell(1).setCellValue("Architecture");
	                headerRow.createCell(2).setCellValue("Version");
	                headerRow.createCell(3).setCellValue("OS Type");
	                headerRow.createCell(4).setCellValue("OS Version");
	                headerRow.createCell(5).setCellValue("URL");
	            }

	            // If the error sheet does not exist, create a new sheet
	            if (errorSheet == null) {
	                errorSheet = existingWorkbook.createSheet("No Data");
	                // Optionally create header row
	                Row errorHeaderRow = errorSheet.createRow(0);
	                errorHeaderRow.createCell(0).setCellValue("URL");
	            }
	        } catch (IOException ex) {
	            ex.printStackTrace();
	        }

	        int newRowCount = existingSheet.getLastRowNum() + 1; // Start from the first empty row
	        int errorRowCount = errorSheet.getLastRowNum() + 1; // Start from the first empty row in error sheet

	        // Split package URLs and process each one
	        String[] urlList = packageUrls.split(",");
	        String regexPackageName = ".*/content/([^/]+)/.*";
	        String regexArchitecture = ".*/([a-zA-Z0-9_]+)/[^/]+/package";
	        Pattern patternPackageName = Pattern.compile(regexPackageName);
	        Pattern patternArchitecture = Pattern.compile(regexArchitecture);

	        List<String> errorUrls = new ArrayList<>(); // List to store URLs with "We'll be back soon."

	        for (String packageUrl : urlList) {
	            // Extract package name and architecture
	            Matcher matcherPackageName = patternPackageName.matcher(packageUrl);
	            Matcher matcherArchitecture = patternArchitecture.matcher(packageUrl);

	            String packageName1 = "Unknown Package Name";
	            String architecture1 = "Unknown Architecture";

	            if (matcherPackageName.find()) {
	                packageName1 = matcherPackageName.group(1);
	            }

	            if (matcherArchitecture.find()) {
	                architecture1 = matcherArchitecture.group(1);
	            }

	            // Store the extracted values in variables
	            String extractedPackageName = packageName1;
	            String extractedArchitecture = architecture1;
	            String individualpaths = "/" + extractedPackageName + "/" + extractedArchitecture;

	            // Create the folder where the new files will be downloaded
	            File baseDownloadDir = new File(baseOutputPath + individualpaths);
	            if (!baseDownloadDir.exists()) {
	                if (baseDownloadDir.mkdirs()) {
	                    System.out.println("Base Directory created successfully: " + baseDownloadDir.getAbsolutePath());
	                } else {
	                    System.out.println("Failed to create base directory: " + baseDownloadDir.getAbsolutePath());
	                    return;
	                }
	            }

	            driver.navigate().to(packageUrl);
	            Thread.sleep(4000);
	            WebElement versions = driver.findElement(By.xpath("//select[@id='evr' and @class='select-chosen linked']"));

	            ((RemoteWebDriver) driver).executeScript("arguments[0].style.display='block';", versions);
	            Select dropdown = new Select(versions);

	            List<WebElement> allOptions = dropdown.getOptions();
	            List<String> dropdownValues = new ArrayList<>();
	            for (WebElement option : allOptions) {
	                dropdownValues.add(option.getText());
	            }

	            for (String value : dropdownValues) {
	            	Thread.sleep(4000);
	                versions = driver.findElement(By.xpath("//select[@id='evr' and @class='select-chosen linked']")); // Re-locate the dropdown element
	                ((RemoteWebDriver) driver).executeScript("arguments[0].style.display='block';", versions);
	                dropdown = new Select(versions);
	                dropdown.selectByVisibleText(value);
	                Thread.sleep(6000);

	                // Check for "We'll be back soon." message
	                if (driver.getPageSource().contains("We'll be back soon.")) {
	                    String errorUrl = driver.getCurrentUrl();
	                    errorUrls.add(errorUrl);
	                    System.out.println("Found 'We'll be back soon.' message for URL: " + errorUrl);
	                    driver.navigate().back();
	                    continue; // Move to the next version
	                }

	                for (int versionCount = 0; versionCount < 1; versionCount++) { // Process one version
	                    List<WebElement> links = driver.findElements(By.tagName("a"));
	                    for (WebElement link : links) {
	                        String href = link.getAttribute("href");
	                        if (href != null && href.contains("access.cdn.redhat.com/content/origin/rpms")) {
	                            String[] parts = extractParts(href, extractedArchitecture);

	                            // Check if URL exists in the existing Excel sheet
	                            boolean urlExists = false;
	                            if (existingSheet != null) {
	                                for (int i1 = 1; i1 <= existingSheet.getLastRowNum(); i1++) { // Start from row 1 since row 0 is the header row
	                                    Row row = existingSheet.getRow(i1);
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

	                            // Add URL to new sheet and download file if it doesn't exist in the existing sheet
	                            if (!urlExists) {
	                                Row newRow = existingSheet.createRow(newRowCount++);
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
	                                downloadFile(href, baseDownloadDir.getAbsolutePath());
	                                // Log after downloading each file
	                                System.out.println("Downloaded file from URL: " + href);
	                            } else {
	                                System.out.println("URL already exists in " + existingExcelFile + " file, skipping: " + href);
	                            }
	                        }
	                    }

	                    // Log after processing each version
	                    System.out.println("Completed processing version: " + value);
	                }
	            }
	        }


	        // Add URLs with "We'll be back soon." message to error sheet
	        for (String errorUrl : errorUrls) {
	            Row errorRow = errorSheet.createRow(errorRowCount++);
	            errorRow.createCell(0).setCellValue(errorUrl);
	        }

	        // Save the existing Excel file with new data
	        try (FileOutputStream fileOut = new FileOutputStream(existingExcelFile)) {
	            existingWorkbook.write(fileOut); // Write existing workbook with new data
	        }

	        // Close the workbook
	        existingWorkbook.close();
	        System.out.println("Data added successfully to the " + existingSheetName + " sheet in " + existingExcelFile);
	        // Close the WebDriver
	        driver.quit();
	        System.out.println("Process completed successfully.");
	    }

	    private static String[] extractParts(String url, String architecture) {
	        int rpmIndex = url.lastIndexOf('/') + 1;
	        int queryIndex = url.lastIndexOf('?');
	        String rpmPart = url.substring(rpmIndex, queryIndex == -1 ? url.length() : queryIndex);

	        String[] parts = rpmPart.split("-");



	        String part3 = "";
	        String part3_1 = "";
	        int elIndexStart = -1;
	        elIndexStart = url.indexOf(".el");     
	        if (elIndexStart == -1) {
	            elIndexStart = url.indexOf("+el");     
	        }
	        if (elIndexStart != -1) {
	            int elIndexEnd = url.indexOf("+", elIndexStart+1);
	            if (elIndexEnd == -1) {
	                elIndexEnd = url.indexOf("/", elIndexStart);
	            }
	            if (elIndexEnd != -1) {
	            	part3_1 = url.substring(elIndexStart, elIndexEnd);
	            }
	        }
	        
	        if (part3_1 != null && part3_1.length() > 0) {
	            part3=part3_1.substring(1);
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
	        
	// Extract version part 
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

	    private static boolean isFileLocked(File file) {
	        try (RandomAccessFile raf = new RandomAccessFile(file, "rw");
	             FileChannel channel = raf.getChannel()) {
	            return !channel.tryLock().isValid();
	        } catch (IOException e) {
	            return true; // Assume file is locked if an exception occurs
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
	}