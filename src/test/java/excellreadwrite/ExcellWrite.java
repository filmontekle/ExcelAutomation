package excellreadwrite;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcellWrite {
	public static void main(String[] args) throws IOException{
		String excellPath="/Users/filmontekle/Desktop/EmpData.xlsx";
		FileInputStream in=new FileInputStream(excellPath);
		XSSFWorkbook workbook=new XSSFWorkbook(in);	
		XSSFSheet worksheet=workbook.getSheet("TestData");
		
		int rowsCount=worksheet.getPhysicalNumberOfRows();
		System.out.println("Number of Rows: "+rowsCount);
		
		in.close();
		
		XSSFCell cell=worksheet.getRow(1).getCell(2);
		
		if(cell==null){
			cell=worksheet.getRow(1).createCell(2);
		}
		cell.setCellValue("Pass");
		
		cell=worksheet.getRow(5).getCell(2);
		if(cell==null){
			cell=worksheet.getRow(5).createCell(2);
		}
		cell.setCellValue("Fail");
		

		cell=worksheet.getRow(2).getCell(2);
		if(cell==null){
			cell=worksheet.getRow(2).createCell(2);
		}
		cell.setCellValue("Fail");
		
		cell=worksheet.getRow(4).getCell(2);
		if(cell==null){
			cell=worksheet.getRow(4).createCell(2);
		}
		cell.setCellValue("Fail");
		
		cell=worksheet.getRow(3).getCell(2);
		if(cell==null){
			cell=worksheet.getRow(3).createCell(2);
		}
		cell.setCellValue("Pass");
		
		
		FileOutputStream out=new FileOutputStream(excellPath);
		workbook.write(out);
		
		
		
		
		
		
		
		in.close();
		out.close();
		
		
	
	
	
	
	
	
	}

}
