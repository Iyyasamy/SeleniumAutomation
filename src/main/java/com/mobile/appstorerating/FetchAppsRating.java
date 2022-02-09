package com.mobile.appstorerating;

import io.github.bonigarcia.wdm.WebDriverManager;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeDriver;

import java.io.*;
import java.util.concurrent.TimeUnit;

public class FetchAppsRating {
    public WebDriver driver;

    public static void main(String args[]) throws IOException, FileNotFoundException {
        FetchAppsRating fetchAppsRating = new FetchAppsRating();
        fetchAppsRating.startDriver();
        String playStoreLink = "play.google.com";
        String appleStoreLink = "apps.apple.com";
        File file = new File("src/main/resources/AppsRating.xls");
        FileInputStream inputStream = new FileInputStream(file);
        HSSFWorkbook wb = new HSSFWorkbook(inputStream);
        HSSFSheet sheet = wb.getSheet("Automation");
        int rowCount = sheet.getLastRowNum() - sheet.getFirstRowNum();
        //iterate over each row to get app details & update the corresponding rating
        for (int i = 1; i <= rowCount; i++) {
            String appName = sheet.getRow(i).getCell(0).getStringCellValue().toLowerCase();
            fetchAppsRating.searchApp(appName);
            System.out.println("MobileApp name present in row " + i + " is: " + appName);
            sheet.getRow(i).createCell(1).setCellValue(fetchAppsRating.getAppRating(appName, playStoreLink));
            sheet.getRow(i).createCell(2).setCellValue(fetchAppsRating.getAppRating(appName, appleStoreLink));
            FileOutputStream outputStream = new FileOutputStream("src/main/resources/AppsRating.xls");
            wb.write(outputStream);
            outputStream.close();
        }
        fetchAppsRating.closeDriver();
    }

    public void startDriver() {
        WebDriverManager.chromedriver().setup();
        driver = new ChromeDriver();
        driver.navigate().to("https://www.google.com");
        driver.manage().window().maximize();
        driver.manage().timeouts().implicitlyWait(5, TimeUnit.SECONDS);
    }

    public void searchApp(String appName) {
        driver.findElement(By.xpath("//input[@name='q']")).clear();
        driver.findElement(By.xpath("//input[@name='q']")).sendKeys(appName);
        driver.findElement(By.xpath("//input[@name='q']")).sendKeys(Keys.ENTER);
    }

    // method to return the app rating from store
    public String getAppRating(String appName, String storeLink) {
        String rating = "N/A";
        int currentPage = 0, maxPage = 3;
        driver.findElement(By.xpath("//table[@role='presentation']//tr/td[2]")).click();
        do {
            try {
                rating = driver.findElement(By.xpath("//div[contains(@data-async-context,'query')]/div//a[contains(@href,'" + storeLink + "') and contains(@href,'" + appName + "')]//parent::div//parent::div//parent::div//span[contains(text(),'Rating: ')]")).getText();
                break;
            } catch (NoSuchElementException e) {
                driver.findElement(By.xpath("//span[text()='Next']")).click();
                currentPage++;
            }
        } while (currentPage < maxPage);
        return rating.replaceAll("Rating: ", "");
    }

    public void closeDriver() {
        driver.quit();
    }

}
