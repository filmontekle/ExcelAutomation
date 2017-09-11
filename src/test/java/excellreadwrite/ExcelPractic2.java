package excellreadwrite;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelPractic2 {

	public static void main(String[] args) throws IOException {
		String filePath="./src/test/resources/TestData/AmazonSearchData.xlsx";
		FileInputStream input=new FileInputStream(filePath);
		XSSFWorkbook wb1=new XSSFWorkbook(input);		
		XSSFSheet sh1=wb1.getSheetAt(0);		
		XSSFRow rw=sh1.getRow(3);
		XSSFCell cell=rw.getCell(1);
		
		System.out.println(cell);
		
		int rowCount=sh1.getPhysicalNumberOfRows();
		System.out.println("number of rows:"+rowCount);
		
		for(int rows=1;rows<rowCount;rows++){
			String cells=sh1.getRow(rows).getCell(0).toString();
			String cell2=sh1.getRow(rows).getCell(1).toString();
			String cell3=sh1.getRow(rows).getCell(2).toString();
			
			System.out.println(cells+" "+cell2+" "+cell3);
		}
		
		

	
	}

}
