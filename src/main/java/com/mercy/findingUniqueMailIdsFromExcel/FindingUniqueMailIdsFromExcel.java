package com.mercy.findingUniqueMailIdsFromExcel;

import java.io.File;
import java.io.IOException;
import java.util.Iterator;
import java.util.LinkedHashSet;

import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

public class FindingUniqueMailIdsFromExcel {
	public static void main(String[] args) throws BiffException, IOException {

		System.out.println("Program to findout unique mail ids from Excel");
		File file = new File("Attended_Students_Data.xls");

		// Getting the work book
		Workbook workbook = Workbook.getWorkbook(file);
		// Getting the particular sheet in the workbook
		Sheet sheet = workbook.getSheet(0);

		LinkedHashSet<String> set = new LinkedHashSet<>();

		for (int i=0; i<sheet.getRows(); i++) {
			String content = sheet.getCell(2, i).getContents();

			if (!content.trim().isEmpty()) {
				//				System.out.println("Contents : " + content);
				set.add(content);
			}
		}

		Iterator<String> it = set.iterator();
		
		while (it.hasNext()) {
			String email = it.next();
			if (!email.isEmpty()) {
				System.out.println(email);
			}
		}
	}
}
