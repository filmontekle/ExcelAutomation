package excellreadwrite;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelPracticeUltimate {

	public static void main(String[] args) throws EncryptedDocumentException, InvalidFormatException, IOException {
	
		String exelFilePath="./src/test/resources/TestData/AmazonSearchData.xlsx";		
		FileInputStream fis=new FileInputStream(exelFilePath);
		
		Workbook wb=WorkbookFactory.create(fis);
		
		Sheet sh=wb.getSheetAt(0);
		Row row=sh.getRow(0);
		Cell cell=row.getCell(1);
		System.out.println("Cell:"+cell);
		
		int rowCount=sh.getPhysicalNumberOfRows();
		int colCount=row.getPhysicalNumberOfCells();
		
		for(int i=1; i<rowCount; i++){
			if(sh.getRow(i).getCell(0).toString().contains("Y")){			
			for(int j=0; i<colCount; j++){
		System.out.println(sh.getRow(i).getCell(j).toString()+"--");
	
			}
			System.out.println();
		}else{
			System.out.println("row number"+i+ " is skipped");
		}
		
		}
	}

}
