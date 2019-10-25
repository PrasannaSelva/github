package excelfileTest;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class testxl {

	public static void main(String[] args) throws Exception {
		// TODO Auto-generated method stub
		File f=new File("D:\\Selenium\\xl.xlsx");
		FileInputStream fis=new FileInputStream(f);
		Workbook wb=new XSSFWorkbook(fis);//interfce
		Sheet s=wb.getSheet("fj");
				//getSheetName("fj");//interfac
		int rowCount=s.getLastRowNum()-s.getFirstRowNum();
		for(int i=0;i<rowCount+1;i++)
		{
			Row row =s.getRow(i);
			for(int j=0;j<row.getLastCellNum();j++) {
				System.out.println(row.getCell(j).getStringCellValue());
			}
		}

	}

}
