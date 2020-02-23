package com.Expekt;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;
import org.testng.internal.reflect.MethodMatcherException;
import java.io.FileInputStream;
import java.io.IOException;
import java.security.PublicKey;
import java.util.regex.Pattern;
import java.util.concurrent.TimeUnit;
import org.testng.annotations.*;
import static org.testng.Assert.*;
import org.openqa.selenium.*;
import org.openqa.selenium.By;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.WebDriver;
import org.testng.annotations.DataProvider;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.Test;
import org.testng.internal.reflect.MethodMatcherException;

public class test {
    private WebDriver driver;

    @BeforeMethod
    public void lounchBrowser() {
        System.setProperty("webdriver.chrome.driver", "C:\\Users\\Artur\\Downloads\\chromedriver.exe");
        WebDriver wd = new ChromeDriver();
        driver = new ChromeDriver();
        driver.get("https://en.expekt.com/register");
        driver.manage().window().maximize();
    }

    @Test(dataProvider = "Registration")
    public void login(String name, String sur, String addr, String code, String town, String mob, String user, String pass, String mail, String conf) {
        driver.get("https://en.expekt.com/register");
        driver.findElement(By.id("FirstName")).click();
        driver.findElement(By.id("FirstName")).sendKeys(name);
        driver.findElement(By.id("Surname")).click();
        driver.findElement(By.id("Surname")).sendKeys(sur);
        driver.findElement(By.id("Address")).click();
        driver.findElement(By.id("Address")).sendKeys(addr);
        driver.findElement(By.id("Postcode")).click();
        driver.findElement(By.id("Postcode")).sendKeys(code);
        driver.findElement(By.id("Town")).click();
        driver.findElement(By.id("Town")).sendKeys(town);
        driver.findElement(By.id("Mobile telephone number")).click();
        driver.findElement(By.id("Mobile telephone number")).sendKeys(mob);
        driver.findElement(By.id("Username")).click();
        driver.findElement(By.id("Username")).sendKeys(user);
        driver.findElement(By.id("Password")).click();
        driver.findElement(By.id("Password")).sendKeys(pass);
        driver.findElement(By.id("Email address")).click();
        driver.findElement(By.id("Email address")).sendKeys(mail);
        driver.findElement(By.id("Confirm email address")).click();
        driver.findElement(By.id("Confirm email address")).sendKeys(conf);
        driver.close();
    }

    @DataProvider(name = "Registration")

    public Object[][] getData() throws IOException {
        FileInputStream fileInputStream = new FileInputStream("exl\\123.xls");
        HSSFWorkbook workbook = new HSSFWorkbook(fileInputStream);
        HSSFSheet worksheet = workbook.getSheet("table");
        int rowCount = worksheet.getPhysicalNumberOfRows();
        Object[][] data = new Object[rowCount][10];

        for (int i = 0; i < rowCount; i++)
        {
            HSSFRow row = worksheet.getRow(i);

            HSSFCell name = row.getCell(0);
            if (name==null)
                data[i][0]="";
            else {
                name.setCellType(CellType.STRING);
                data[i][0] = name.getStringCellValue();
            }

            HSSFCell sur = row.getCell(1);
            if (sur==null)
                data[i][1]="";
            else {
                sur.setCellType(CellType.STRING);
                data[i][1] = sur.getStringCellValue();
            }

            HSSFCell addr = row.getCell(2);
            if (addr==null)
                data[i][2]="";
            else {addr.setCellType(CellType.STRING);
                data[i][2] = addr.getStringCellValue();
            }

            HSSFCell code = row.getCell(3);
            if (code==null)
                data[i][3]="";
            else {code.setCellType(CellType.STRING);
                data[i][3] = code.getStringCellValue();
            }

            HSSFCell town = row.getCell(4);
            if (town==null)
                data[i][4]="";
            else{town.setCellType(CellType.STRING);
            data[i][4] = town.getStringCellValue();
            }

            HSSFCell mob = row.getCell(5);
            if (mob==null)
                data[i][5]="";
            else{mob.setCellType(CellType.STRING);
            data[i][5] = mob.getStringCellValue();
            }

            HSSFCell user = row.getCell(6);
            if (user==null)
                data[i][6]="";
            else {user.setCellType(CellType.STRING);
                data[i][6] = user.getStringCellValue();
            }

            HSSFCell pass = row.getCell(7);
            if (pass==null)
                data[i][7]="";
            else{pass.setCellType(CellType.STRING);
            data[i][7] = pass.getStringCellValue();
            }

            HSSFCell mail = row.getCell(8);
            if (mail==null)
                data[i][8]="";
            else {
                mail.setCellType(CellType.STRING);
                data[i][8] = mail.getStringCellValue();
            }

            HSSFCell conf = row.getCell(9);
            if (conf==null)
                data[i][9]="";
            else {
                conf.setCellType(CellType.STRING);
                data[i][9] = conf.getStringCellValue();
            }


        }
        return data;


    }
}





