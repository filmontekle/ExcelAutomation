package excellreadwrite;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcellConditionalWrite {
	
public static void main(String[] args) throws IOException{
		
		String excellPath="/Users/filmontekle/Desktop/EmpData.xlsx";		
		FileInputStream in=new FileInputStream(excellPath);		
		XSSFWorkbook workbook=new XSSFWorkbook(in);	
		XSSFSheet worksheet=workbook.getSheet("TestData");
		
		int rowNum=worksheet.getPhysicalNumberOfRows();
		System.out.println("rows count:"+rowNum);
		
		for(int rows=1; rows<rowNum; rows++){
			String row1=worksheet.getRow(rows).getCell(1).toString();
			if(row1.contains("Wooden Spoon")){
				XSSFCell cell=worksheet.getRow(rows).getCell(2);
				if(cell==null)
				cell=worksheet.getRow(rows).createCell(2);
				cell.setCellValue("Fail");
				break;
			}
			
		}
			FileOutputStream fileOutput=new FileOutputStream(excellPath);	
			workbook.write(fileOutput);
			
		fileOutput.close();
		workbook.close();
		in.close();
	}

}
