package org.example;

import io.github.bonigarcia.wdm.WebDriverManager;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import java.io.FileOutputStream;
import java.io.IOException;

public class testing {
static WebDriver driver;

        public static void main(String[] args) {
        openBrowser("https://www.google.com");
    }

    public static void openBrowser(String browserName) {
        WebDriverManager.chromedriver().setup();
        driver = new ChromeDriver();
        driver.get(browserName);
    }

    public static void looking() {
        // Create a new workbook
        Workbook workbook = new XSSFWorkbook();
        // Create a new sheet
        Sheet sheet = workbook.createSheet("MySheet");

        // Create data
        String[] data = {"Name", "Age"};
        // Create a header row
        Row headerRow = sheet.createRow(0);
        for (int i = 0; i < data.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(data[i]);
        }

        // Sample data
        String[][] sampleData = {

                {"Lakshmi", "25"},
                {"Khaja", "30"},
                {"Akhil", "22"}
        };

        // Write sample data to the sheet
        for (int rowIndex = 1; rowIndex <= sampleData.length; rowIndex++) {
            Row row = sheet.createRow(rowIndex);
            for (int columnIndex = 0; columnIndex < data.length; columnIndex++) {
                Cell cell = row.createCell(columnIndex);
                cell.setCellValue(sampleData[rowIndex - 1][columnIndex]);
            }
        }

        try {
            // Save the workbook to a file
            FileOutputStream fileOut = new FileOutputStream("example.xlsx");
            workbook.write(fileOut);
            fileOut.close();
            System.out.println("Excel file created successfully.");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public void closeBrowser(){
        driver.quit();
    }


}
