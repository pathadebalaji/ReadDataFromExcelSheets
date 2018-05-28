package com.project.function;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadingExcelSheetData {

	public static void main(String[] args) throws IOException {

		FileInputStream fis = new FileInputStream("C:\\Selenium Learning\\ReadDataFromExcelSheets\\src\\com\\project\\excelSheet\\DataExcel.xlsx");
		
		XSSFWorkbook wb = new XSSFWorkbook(fis);

		XSSFSheet sh = wb.getSheet("Sheet1");
		
		System.out.println("Get total Number of Rows "+ sh.getPhysicalNumberOfRows());
		System.out.println("Get total Number of Colums "+ sh.getRow(0).getPhysicalNumberOfCells());
		
		for(int i=0; i<sh.getPhysicalNumberOfRows(); i++)
		{
			for(int j=0; j<sh.getRow(0).getPhysicalNumberOfCells(); j++)
			{
				System.out.print((sh.getRow(i).getCell(j).getStringCellValue())+"    ");
			}
			System.out.println();
		}
		
	}

}
