package Utility;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReadUtil {
	
	public XSSFWorkbook wb;
	public XSSFSheet sh;
	public File src;
	public FileOutputStream fos;
	public ExcelReadUtil(String path) {
		try {
			src = new File(path);
			
				FileInputStream fis = new FileInputStream(src);
				wb = new XSSFWorkbook(fis);
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.getMessage();
		} 
		 try {
			fos= new FileOutputStream(src);
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}	
				
	}
	
	public String getData(String SheetName) {
		String data0 = null;
		sh = wb.getSheet(SheetName);
		int rownum = sh.getLastRowNum();
		
		for(int i = 0;i<=rownum;i++) {
			Row row  = sh.getRow(i);
			int colnum = row.getLastCellNum();
			
			for(int j=0;j<colnum;j++) {
				data0 = row.getCell(j).getStringCellValue();
				System.out.println(data0);
				
			}
		}
		return data0;
		
	
		
	}
	
	public void writeData(String SheetName) {
		sh = wb.getSheet(SheetName);
		int rownum = sh.getLastRowNum();
		
		for(int i = 0;i<=rownum;i++) {
			Row row  = sh.getRow(i);
			int colnum = row.getLastCellNum();
			for(int j =0;j<colnum;j++) {
				int colindex = row.getCell(j).getColumnIndex();
				System.out.println(colindex);
				if(colnum==colindex+1) {
				row.createCell(j+1).setCellValue("newDtaa");
				}
			}
	}
		
try {
	wb.write(fos);
	wb.close();
} catch (Exception e) {
	// TODO Auto-generated catch block
	e.printStackTrace();
}
}
}
