package ReadExcelData;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;

import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadDataFromLoop1 {

	public static void main(String[] args) throws Exception {
		File src= new File("D:\\DXC Selenium Automation Class\\ApachePoiInSelenium\\TestDataOrangeHRM\\testdataorangeHRM.xlsx");
		FileInputStream fis=new FileInputStream(src);
		XSSFWorkbook wb=new XSSFWorkbook(fis);
		XSSFSheet sheet1=wb.getSheetAt(0);
		
		
		int rowcount =sheet1.getLastRowNum();
		System.out.println("Total Rows from Excel Sheet"+rowcount);
		
		for(int i=0;i<=rowcount;i++)
		{
			XSSFRichTextString data1=sheet1.getRow(i).getCell(0).getRichStringCellValue();
			System.out.println("Data from row is.."+i+"is"+data1);
			
			XSSFRichTextString data2=sheet1.getRow(i).getCell(1).getRichStringCellValue();
			System.out.println("Data from row is.."+i+"is"+data2);
			
		}
			wb.close();
	}

}
