package day16;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class task1_ExcelWrite {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
      XSSFWorkbook book = new XSSFWorkbook();
      XSSFSheet sheet = book.createSheet("Sheet1");
      Object[][] data = {
    		  {"Name","Age","Emai"},
    		  {"John Doe",30,"john@test.com"},
    		  {"Jane Doe",28,"john@test.com"},
    		  {"Bob Smith",35,"jacky@example.com"},
    		  {"Swapnil",37,"swapnil@example.com"}
      };
      int rowcnt=0;
      for(Object[] row1: data) {
    	  XSSFRow row = sheet.createRow(rowcnt++);
    	  int colcnt=0;
    	  for(Object col: row1) {
    		  XSSFCell cell = row.createCell(colcnt++);
    		  if(col instanceof String) {
    			  cell.setCellValue((String)col);
    		  }
    		  else if(col instanceof Integer) {
    			  cell.setCellValue((Integer)col);
    		  }
    	  }
      }
      try {
		FileOutputStream output = new FileOutputStream("C:\\Users\\ezhil\\eclipse-workspace\\ExcelOperations\\src\\main\\java\\day16\\EmployeeDetails.xlsx");
	    book.write(output);
      } catch (FileNotFoundException e) {
		// TODO Auto-generated catch block
		e.printStackTrace();
	} catch (IOException e) {
		// TODO Auto-generated catch block
		e.printStackTrace();
	}
     book.close(); 
	}

}
