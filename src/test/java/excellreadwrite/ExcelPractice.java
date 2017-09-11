package excellreadwrite;

import java.io.FileInputStream;
import java.io.FileNotFoundException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class ExcelPractice {

	public static void main(String[] args) throws Exception {
		//.xls formats or older =//Workbook -->//Row --> //Cell
			
		String exelFilePath="./src/test/resources/TestData/Workbook2.xls";		
		FileInputStream fis=new FileInputStream(exelFilePath);
		//Workbook
		HSSFWorkbook wb1=new HSSFWorkbook(fis);
		//WorkSheet
		HSSFSheet sh1 = wb1.getSheet("Sheet1");
		//HSSFSheet sh1=wb1.getSheetAt(0);
		//Row
		HSSFRow rw = sh1.getRow(3);
		//Cell
		HSSFCell cell=rw.getCell(1);
		
		System.out.println(cell.toString());
		
		int rowNum=sh1.getPhysicalNumberOfRows();
		int colNum=sh1.getRow(0).getPhysicalNumberOfCells();
		
		for(int i=0;i<rowNum; i++){
			for(int j=0; j<colNum; j++){
				HSSFRow rw1 = sh1.getRow(i);
				//Cell
				HSSFCell cell1=rw.getCell(j);
				
				sh1.getRow(i).getCell(j);
				
				System.out.println(cell1);
				
			}
		}
	}

}
