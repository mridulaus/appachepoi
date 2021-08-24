package com.demo.ExcelProject;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteDataIntoExcel {

	public static void main(String[] args) throws IOException {

		XSSFWorkbook workbook = new XSSFWorkbook();

		XSSFSheet spreadsheet = workbook.createSheet(" nailpaint Sheet ");

		XSSFRow row;

		Map<String, Object[]> nailinfo = new TreeMap<String, Object[]>();
		nailinfo.put("1", new Object[] { "Id", "brand", "colour", "price" });

		nailinfo.put("2", new Object[] { "1", "Oriflame", "lilac", "250" });

		nailinfo.put("3", new Object[] { "2", "Mabeline", "hotpink", "155" });

		nailinfo.put("4", new Object[] { "3", "Nyka", "purple", "200" });

		nailinfo.put("5", new Object[] { "4", "Bella", "iceyBlue", "176"});

	

		Set<String> keyid = nailinfo.keySet();
		int rowid = 0;

		for (String key : keyid) {
			row = spreadsheet.createRow(rowid++);
			Object[] objectArr = nailinfo.get(key);
			int cellid = 0;

			for (Object obj : objectArr) {
				Cell cell = row.createCell(cellid++);
				cell.setCellValue((String) obj);
			}
		}

		FileOutputStream out = new FileOutputStream(new File("C:\\Users\\user pc\\Documents\\nailpolishinfo.xlsx"));

		workbook.write(out);
		out.close();
		System.out.println("Writesheet.xlsx written successfully");


	}

}
