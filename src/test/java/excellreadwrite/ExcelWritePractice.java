package excellreadwrite;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class ExcelWritePractice {

	public static void main(String[] args) throws IOException {
		//something changed in Git
		
		String exelFilePath="./src/test/resources/TestData/Workbook2.xls";		
		FileInputStream fis=new FileInputStream(exelFilePath);
		FileOutputStream fos=new FileOutputStream(exelFilePath);
		try{
		HSSFWorkbook wb1=new HSSFWorkbook(fis);
		HSSFSheet sh2=wb1.getSheetAt(0);
		HSSFRow rw = sh2.getRow(2);
		HSSFCell cell=rw.getCell(2);
		
		
		if(cell==null){
			rw.createCell(2);
			cell.setCellValue("Fail");
		}else{
			cell.setCellValue("Pass");
		}

		
		
		wb1.write(fos);
		}catch	(Exception e){
		
		}finally{	
		fis.close();
		fos.close();
		}
		
			

	}

}
