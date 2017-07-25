package org.CSC.others;

import java.io.FileInputStream;

import java.io.IOException;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class MultipleSheets {

	public static void main(String[] args) throws IOException {
		FileInputStream file = new FileInputStream(
				"C:\\Users\\MaheshRam\\Desktop\\Selenium_files\\excelfiles\\datadrivenREGISTRATION.xlsx");
		XSSFWorkbook wb = new XSSFWorkbook(file);
		int sheetCount = wb.getNumberOfSheets();
		for (int i = 0; i <= sheetCount; i++) {
			XSSFSheet ws = wb.getSheet("Sheet" + i);
			int rowCount = ws.getLastRowNum();
			for (int j = 0; j <= rowCount; j++) {
				Row r = ws.getRow(j);
				int colCount = r.getLastCellNum();
				for (int k = 0; k <= colCount; k++) {
					System.out
							.print(r.getCell(k).getStringCellValue() + "    ");
				}
				System.out.println();
			}

		}

	}

}
