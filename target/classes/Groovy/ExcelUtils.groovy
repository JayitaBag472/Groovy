package Groovy;

import jxl.*
import jxl.write.*

public class ExcelUtils {
	

	 
	
	  public static void main(String[] args)
	  {
		//writeExcel()
		readExcel()
	  }
	   
	 /* static def writeExcel()
	  {
		WritableWorkbook workbook = Workbook.createWorkbook(new File("C:\\Users\\ejaybag\\Desktop\\Groovy\\src\\Groovy\\output.xls"))
		WritableSheet sheet = workbook.createSheet("Worksheet1", 0)
		 
		for(int i=0;i<3;i++)
		{
		  for(int j=0;j<5;j++)
		  {
			Label label = new Label(i, j, i+","+j);
			sheet.addCell(label);
		  }
		}
		 
		workbook.write()
		workbook.close()
		 
	  }*/
	   
	  static def readExcel()
	  {
		Workbook workbook1 = Workbook.getWorkbook(new File("C:\\Users\\ejaybag\\Desktop\\Groovy\\src\\Groovy\\InputData.xls"))
		Sheet sheet1 = workbook1.getSheet(0)
		Cell a1 = sheet1.getCell(0,28)
		//Cell b2 = sheet1.getCell(2,2)
		//Cell c2 = sheet1.getCell(2,1)
		 
		println a1.getContents();
		//println b2.getContents();
		//println c2.getContents();
		 
		workbook1.close()
	  }
	}

