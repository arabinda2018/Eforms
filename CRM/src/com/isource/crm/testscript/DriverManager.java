package com.isource.crm.testscript;

import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DriverManager extends KeyWord {


static KeyWord kw;
	public static void main(String[] args) throws Exception {
		kw=new KeyWord();
		ArrayList a = new ArrayList();
		FileInputStream f = new FileInputStream("C:\\Users\\Arabinda Mohanty\\Desktop\\hubspot.xlsx");
		XSSFWorkbook wbks = new XSSFWorkbook(f);
		XSSFSheet s = wbks.getSheetAt(0);
		Iterator rowitr = s.iterator();
		while (rowitr.hasNext()) {
			Row row = (Row) rowitr.next();
			Iterator cellitr = row.cellIterator();
			while (cellitr.hasNext()) {
				Cell c = (Cell) cellitr.next();
				if (c.getCellTypeEnum() == CellType.STRING) {
					a.add(c.getStringCellValue());
				}
				if (c.getCellTypeEnum() == CellType.NUMERIC) {
					a.add(c.getNumericCellValue());
				}
				if (c.getCellTypeEnum() == CellType.BOOLEAN) {
					a.add(c.getBooleanCellValue());
				}

			}
		}
		for (int i = 0; i < a.size(); i++) {
			if (a.get(i).equals("openBrowser")) {
				String keyWordName = (String) a.get(i);
				String testData = (String) a.get(i + 1);
				String objectName = (String) a.get(i + 2);
				String runmode = (String) a.get(i + 3);
				System.out.println(keyWordName);
				System.out.println(testData);
				System.out.println(objectName);
				System.out.println(runmode);
				if (runmode.equals("yes")) {
					kw.openBrowser();
				}
			}
		}
	}
}
