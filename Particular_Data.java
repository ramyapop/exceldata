package com.particulardata;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Particular_Data {////particulsar data
	public static void main(String[] args) throws IOException {
	
		File f = new File("C:\\Users\\ramya\\eclipse-workspace\\data_driven\\readdata\\data.xlsx");
		//converting excel into a file
	
		FileInputStream f1= new FileInputStream(f);
	//to read data from excel
	
		Workbook wb= new XSSFWorkbook(f1);//upcasting and opening workbook is a parent interface
		Sheet sheet = wb.getSheet("sheet1");
		Row row = sheet.getRow(1);
		Cell cell = row.getCell(2);//celltype la string or numeric
		CellType cellType = cell.getCellType();
		
		if (cellType.equals(cellType.STRING)) {//celltype--->cell(celltype might be string or numeric
			//if its string
			String st = cell.getStringCellValue();
			System.out.println(st);
			
		}
		
		else if (cellType.equals(cellType.NUMERIC)) {//double-->int-->string
			//if its numeric
			double d = cell.getNumericCellValue();//0.000 point ha eruka kudathu  so coverting to int
			int i=(int) d;//big to small narrowing  int--->string
			String value = String.valueOf(i);
			System.out.println(value);
		}
	wb.close();
	}

}
