package excellreadwrite;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcellConditionalRead {
	
	public static void main(String[] args) throws IOException{
		String excellPath="/Users/filmontekle/Desktop/EmpData.xlsx";
		FileInputStream in=new FileInputStream(excellPath);
		XSSFWorkbook workbook=new XSSFWorkbook(in);	
		XSSFSheet worksheet=workbook.getSheet("TestData");
		
		int rowsCount=worksheet.getPhysicalNumberOfRows();
		System.out.println("Number of Rows in TestData:"+rowsCount);
		
		for(int rows=1; rows<rowsCount; rows++){
			String execute=worksheet.getRow(rows).getCell(0).toString();
			if(execute.equals("Y")){
				String searchItem=worksheet.getRow(rows).getCell(1).toString();
				System.out.println("Search for "+searchItem);
				
			}
			
			
			
			}
		
		in.close();
	
	}	
		
}
