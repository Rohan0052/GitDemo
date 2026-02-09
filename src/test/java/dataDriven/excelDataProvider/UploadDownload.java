package dataDriven.excelDataProvider;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Duration;
//import java.util.ArrayList;
import java.util.Iterator;
import java.util.Scanner;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.Assert;

public class UploadDownload {

  public static void main(String[] args) throws IOException {

    try (Scanner sc = new Scanner(System.in)) {
      System.out.print("Enter Excel file path: ");
      String fileName = sc.nextLine();

      System.out.print("Enter fruit name: ");
      String fruitName = sc.nextLine();

      System.out.print("Enter column name: ");
      String text = sc.nextLine();

      System.out.print("Enter new value: ");
      String value = sc.nextLine();

      WebDriver driver = new ChromeDriver();
      driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(10));
      driver.manage().window().maximize();

      driver.get("https://rahulshettyacademy.com/upload-download-test/index.html");

      // Click download button
      driver.findElement(By.xpath("//div[@class='container']//button[@class='button']")).click();
      
      // Edit excel 
      
      int col = getColNumber(fileName, text);
      int row = getRowNumber(fileName, fruitName);

      Assert.assertTrue(updateCell(fileName, row, col, value));

      // Upload file
      WebElement upload = driver.findElement(By.xpath("//input[@type='file']"));
      upload.sendKeys(fileName);

      WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(15));

      // Wait for success message
      wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//div[contains(text(),'Updated Excel Data Successfully.')]")));

      String successMessage = driver.findElement(By.xpath("//div[contains(text(),'Updated Excel Data Successfully.')]")).getText();

      System.out.println(successMessage);

      // Assertion for success message
      Assert.assertEquals(successMessage, "Updated Excel Data Successfully.");

      // Wait for message to disappear
      wait.until(ExpectedConditions.invisibilityOfElementLocated(By.xpath("//div[contains(text(),'Updated Excel Data Successfully.')]")));

      // Get column id of Price
      String priceCol = driver.findElement(By.xpath("//div[text()='Price']")).getAttribute("data-column-id");

      // Get actual price for Apple
      String actualPrice = driver.findElement(By.xpath("//div[text()='" + fruitName + "']/parent::div/parent::div/div[@id='cell-" + priceCol + "-undefined']")).getText();

      System.out.println(actualPrice);

      // Assertion for price
      Assert.assertEquals(actualPrice, value);
      
      driver.close();

    }

  }

  private static boolean updateCell(String fileName, int row, int col, String value) throws IOException {
    // TODO Auto-generated method stub

    //		ArrayList<String> a = new ArrayList<String>();
    FileInputStream fis = new FileInputStream(fileName);
    XSSFWorkbook workbook = new XSSFWorkbook(fis);

    XSSFSheet sheet = workbook.getSheet("Sheet1");
    Row rowfield = sheet.getRow(row - 1);
    Cell cell = rowfield.getCell(col - 1);
    cell.setCellValue(value);

    FileOutputStream fos = new FileOutputStream(fileName);
    workbook.write(fos);
    workbook.close();
    fis.close();
    return true;

  }

  private static int getRowNumber(String fileName, String text) throws IOException {
    // TODO Auto-generated method stub

    //		ArrayList<String> a = new ArrayList<String>();
    FileInputStream fis = new FileInputStream(fileName);
    try (XSSFWorkbook workbook = new XSSFWorkbook(fis)) {
      XSSFSheet sheet = workbook.getSheet("Sheet1");

      Iterator < Row > rows = sheet.iterator();

      int k = 1;
      int rowIndex = -1;

      while (rows.hasNext()) {
        Row row = rows.next();
        Iterator < Cell > cells = row.cellIterator();

        while (cells.hasNext()) {

          Cell ce = cells.next();
          if (ce.getCellType() == CellType.STRING && ce.getStringCellValue().equalsIgnoreCase(text)) {

            rowIndex = k;
          }

        }

        k++;

      }

      return rowIndex;
    }
  }

  private static int getColNumber(String fileName, String colName) throws IOException {
    // TODO Auto-generated method stub

    //				ArrayList<String> a = new ArrayList<String>();
    FileInputStream fis = new FileInputStream(fileName);
    try (XSSFWorkbook workbook = new XSSFWorkbook(fis)) {
      XSSFSheet sheet = workbook.getSheet("Sheet1");

      Iterator < Row > rows = sheet.iterator();
      Row firstrow = rows.next();

      //			Iterator<Cell> firstcell = firstrow.cellIterator();

      Iterator < Cell > firstcell = firstrow.cellIterator();

      int k = 1;
      int column = 0;
      while (firstcell.hasNext()) {
        Cell value = firstcell.next();
        if (value.getStringCellValue().equalsIgnoreCase(colName)) {

          column = k;
        }

        k++;

      }
      System.out.println(column);
      return column;
    }

  }

}