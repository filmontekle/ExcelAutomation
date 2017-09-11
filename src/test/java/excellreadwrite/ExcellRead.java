package excellreadwrite;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.util.SystemOutLogger;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcellRead {

	public static void main(String[] args) throws IOException{
		
		String excellPath="/Users/filmontekle/Desktop/EmpData.xlsx";
		//Open the file to read
		FileInputStream in=new FileInputStream(excellPath);
		//let the apache xssfWorkbook handle the data
		XSSFWorkbook workbook=new XSSFWorkbook(in);
		//Jump to the worksheet
		XSSFSheet worksheet=workbook.getSheet("sheet1");
		//Find out how may rows there are
		int rowsCount=worksheet.getPhysicalNumberOfRows();
		System.out.print("number of rows: "+rowsCount);
		//print first row and first cell data
		System.out.println();
		System.out.println("row 1: "+worksheet.getRow(0).getCell(0));
		System.out.println("Row 2 Cell 1 data:"+worksheet.getRow(1).getCell(0));
		System.out.println("rows 4 cell 2 data:"+worksheet.getRow(3).getCell(1));
		System.out.println("rows data:"+worksheet.getRow(4).getCell(1) + " In "+
				worksheet.getRow(4).getCell(2));
		System.out.println("data on row 2 cell 1:"+worksheet.getRow(1).getCell(0));
		System.out.println("last row second cell:"+worksheet.getRow(3).getCell(1));
		System.out.println("=========================");
		String CellValue=worksheet.getRow(3).getCell(1).toString();
		System.out.println("Data on Row 5: "+CellValue);
		
		//print all names+Department+employee ID
		for(int rowNum=1; rowNum<rowsCount; rowNum++){
			String name=worksheet.getRow(rowNum).getCell(1).toString();
			String dep=worksheet.getRow(rowNum).getCell(2).toString();
			String ID=worksheet.getRow(rowNum).getCell(0).toString();
			System.out.println(ID+" "+name+"==>"+dep);
					
			
		}
		
		
		in.close();
		
		
	}
}
