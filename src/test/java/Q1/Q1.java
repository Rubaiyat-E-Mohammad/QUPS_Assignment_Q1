package Q1;

import io.github.bonigarcia.wdm.WebDriverManager;
import org.openqa.selenium.*;
import org.openqa.selenium.edge.EdgeDriver;


import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.Iterator;
import java.util.List;
import java.util.stream.Collectors;
import java.util.Comparator;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

public class Q1 {

    public static void main(String[] args) {
    	WebDriverManager.edgedriver().setup();
		WebDriver driver = new EdgeDriver();
        	
        try {
        	driver.get("https://www.google.com");
            driver.manage().window().maximize();

            // Get the current day
            String currentDay = getCurrentDay();

            // Load the Excel data
            FileInputStream fis = new FileInputStream("/Users/rubaiyatemohammad/skills/SQA/Selenium/Q1/src/test/java/Q1/Q1.xlsx");
            XSSFWorkbook wb = new XSSFWorkbook(fis);
            XSSFSheet sheet = wb.getSheet(currentDay);
            
            
            Iterator<Row> rowIterator = sheet.iterator();
            
            Row row = rowIterator.next();
            
            while (rowIterator.hasNext()) {
            row = rowIterator.next();

            String keyword = row.getCell(1).getStringCellValue(); // Assuming keyword is in the second column

            // Perform Google search for the keyword
            performGoogleSearch(driver, keyword);
            
            List<WebElement> suggestionElements = driver.findElements(By.xpath("//ul[@role='listbox']//li[@role='presentation']"));
            
            // Finding the longtest option
            String longestOption = "";
            for (WebElement suggestion : suggestionElements) {
                String temp = suggestion.getText();
                if(temp.length() > longestOption.length()) {
                	longestOption = temp;
                }
            }

            //Finding the shortest option
            String shortestOption = longestOption;
            for (WebElement suggestion : suggestionElements) {
                String temp = suggestion.getText();
                if(temp.length() < shortestOption.length()) {
                	shortestOption = temp;
                }
            }


            // Update the row in the Excel sheet
            row.createCell(2).setCellValue(longestOption); 
            row.createCell(3).setCellValue(shortestOption);
            }

            // Save the changes to the Excel file
            FileOutputStream fos = new FileOutputStream("/Users/rubaiyatemohammad/skills/SQA/Selenium/Q1/src/test/java/Q1/Q1.xlsx");
            wb.write(fos);
            fos.close();
            
        }catch(Exception e) {
        	e.printStackTrace();
        }finally {
        	driver.quit();
        }
    }

    private static String getCurrentDay() {
    	LocalDate currentDate = LocalDate.now();

        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("EEEE");

        String currentDay = currentDate.format(formatter);
        return currentDay;
    }
    
    private static void performGoogleSearch(WebDriver driver, String keyword) throws InterruptedException {
        WebElement searchBox = driver.findElement(By.name("q"));
        searchBox.clear(); // Clear the search box
        searchBox.sendKeys(keyword);
        Thread.sleep(2000);
    }
    

}


