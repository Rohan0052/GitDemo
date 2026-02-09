package dataDriven.excelDataProvider;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

public class dataProvide {
	
	@Test(dataProvider="driveTest")
	public void testCaseData(String greeting , String day , String id) {
		
		System.out.println(greeting+day+id);		
		
	}
	
	DataFormatter formatter = new DataFormatter();
	
	@DataProvider(name="driveTest")
	public Object[][] getData() throws IOException {
		
//		Object[][] data = { {"Hello","Sunday",1} , {"Namaste","Monday",2} , {"Hola","Tuesday",3} };
//		return data;
		
		FileInputStream fis = new FileInputStream("D:\\Selenium_practice\\ExcelDriven.xlsx");
		try (XSSFWorkbook wb = new XSSFWorkbook(fis)) {
			XSSFSheet sheet = wb.getSheetAt(0);
			int rowcount = sheet.getPhysicalNumberOfRows();
			XSSFRow row = sheet.getRow(0);
			int colcount = row.getLastCellNum();
			
			Object[][] data = new Object[rowcount-1][colcount];
			 
			for(int i=0 ; i<rowcount-1 ; i++){
				
				row = sheet.getRow(i+1);
				for(int j=0 ; j<colcount ; j++){
					
					XSSFCell cell = row.getCell(j);
					
					data[i][j] = formatter.formatCellValue(cell);				
					
				}
			}
			return data;
		}

		
	}

}
