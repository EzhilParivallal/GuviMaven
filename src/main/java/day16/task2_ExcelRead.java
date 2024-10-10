package day16;

import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class task2_ExcelRead {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
     XSSFWorkbook book = new XSSFWorkbook("C:\\Users\\ezhil\\eclipse-workspace\\ExcelOperations\\src\\main\\java\\day16\\EmployeeDetails.xlsx");
     XSSFSheet sheet = book.getSheet("Sheet1");
     int rowcnt = sheet.getLastRowNum();
     int colcnt = sheet.getRow(0).getLastCellNum();
     for(int i=1;i<=rowcnt;i++) {
    	 XSSFRow row = sheet.getRow(i);
    	 for(int j=0;j<colcnt;j++) {
    		 XSSFCell cell=row.getCell(j);
    		 System.out.print(cell.getStringCellValue()+" ");
    	 }
    	 
    	 System.out.println();
     }
     book.close();
	}

}
