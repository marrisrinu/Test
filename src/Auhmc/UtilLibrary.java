package Auhmc;

import java.awt.HeadlessException;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;





public class UtilLibrary {
	int xlrows,xlcols;
	public String[][] readandwrite() throws Exception {

		// Read data from excel sheet
		FileInputStream fi = new FileInputStream("C:\\Users\\harini.b\\Desktop\\Enrollment\\pom.xlsx");
		Workbook wrkbook =WorkbookFactory.create(fi);
		Sheet wrksheet = wrkbook.getSheet("Sheet1");
		xlrows=wrksheet.getLastRowNum();
	    xlcols=wrksheet.getRow(0).getLastCellNum();
	    String localarray[][]=new String[xlrows][xlcols];
	    for( int i=0; i<xlrows; i++)
	    {
	    	Row r=wrksheet.getRow(i+1);
	    	int ccnt=r.getLastCellNum();
	    	for(int j=0; j<ccnt; j++)
	    	{
	    		Cell c=r.getCell(j);
	    		if(c.getCellType()==c.CELL_TYPE_STRING)
	    		{
	    			localarray[i][j]=c.getStringCellValue();
	    		}
	    		if(c.getCellType()==c.CELL_TYPE_NUMERIC)
	    		{
	    			int val=(int)c.getNumericCellValue();
	    			localarray[i][j]=""+val;
	        		}
	    		if(c.getCellType()==c.CELL_TYPE_BLANK){
	    			localarray[i][j]="";
	    		}
	    		}
	    	}
		return localarray;
	}
	
	public static void excelwrite(String status, int LastRow) throws Exception {
		  try {
		   FileInputStream file = new FileInputStream(new File(
		     "C:\\Users\\harini.b\\Desktop\\Enrollment\\testresult.xls"));

		    HSSFWorkbook workbook = new HSSFWorkbook(file);
		   HSSFSheet sheet = workbook.getSheetAt(0);

		    Row row = sheet.getRow(LastRow);

		    Cell cell2 = row.createCell(2); // Shift the cell value depending upon column size
		   cell2.setCellValue(status);
		   // System.out.println(status);
		   file.close();
		   FileOutputStream outFile = new FileOutputStream(new File(
		     "C:\\Users\\harini.b\\Desktop\\Enrollment\\testresult.xls"));
		   workbook.write(outFile);

		   }
		  catch (FileNotFoundException e) {
			   e.printStackTrace();
			  } catch (IOException e) {
			   e.printStackTrace();
			  } catch (HeadlessException e) {
			   e.printStackTrace();
			  }
			 }
}
