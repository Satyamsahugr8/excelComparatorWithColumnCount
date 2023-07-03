package com.swing;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class MainCount {
	
	public static void main(String[] args) {
		
		try {
		FileInputStream file1Count = new FileInputStream(
					"C:\\Users\\SATYASAH\\OneDrive - Capgemini\\Desktop\\NewData\\Output\\ChildOutput_ComparedBy_Emp ID_Sheet1_file1.xlsx");
		
		XSSFWorkbook workBookCount = new XSSFWorkbook(file1Count);
		XSSFSheet sheetCount = workBookCount.getSheetAt(0);

		int totalNumberOfRowsInExcel1 = sheetCount.getLastRowNum();
		int totalNumberOfColumnInExcel1 = sheetCount.getRow(0).getLastCellNum();

		int columnIndex = 0;

		for (int i = 0; i < totalNumberOfColumnInExcel1; i++) {

			if ("status".equalsIgnoreCase(sheetCount.getRow(0).getCell(i).toString())) {
				columnIndex = i;
//				System.out.println(columnIndex);
			}
		}

		// Initialize counters
		int activeCount = 0;
		int inactiveCount = 0;

		// Iterate over the rows in the column
		for (Row row : sheetCount) {

			Cell cell = row.getCell(columnIndex);

			if (cell != null) {
				String cellValue = cell.getStringCellValue();

				// Assuming "Active" is considered active and "Inactive" is considered inactive
				if (cellValue.equalsIgnoreCase("Active")) {
					activeCount++;
				} else if (cellValue.equalsIgnoreCase("Inactive")) {
					inactiveCount++;
				}
			}
		}

		// creating new working and adding new rows for excel1
		XSSFWorkbook workBookOutput1 = new XSSFWorkbook();
		XSSFSheet sheetCreate1 = workBookOutput1.createSheet();
		XSSFRow rowCreated = null;

		for (int r = 0; r <= 1; r++) {
			rowCreated = sheetCreate1.createRow(r);

			for (int c = 0; c < 3; c++) {
				rowCreated.createCell(c);
			}
		}
//		sheetCreate1.createRow(1).createCell(2); //0,1,2

		// active
		sheetCreate1.getRow(0).getCell(0).setCellValue("status:");
		sheetCreate1.getRow(0).getCell(1).setCellValue("Active:");
		sheetCreate1.getRow(0).getCell(2).setCellValue(activeCount);

		//// Inactive
		sheetCreate1.getRow(1).getCell(1).setCellValue("Inactive:");
		sheetCreate1.getRow(1).getCell(2).setCellValue(inactiveCount);

//		System.out.println("ActiveCount:" + activeCount);
//		System.out.println("InactiveCount:" + inactiveCount);

		String target1Path = "C:\\Users\\SATYASAH\\OneDrive - Capgemini\\Desktop\\NewData\\Output\\Count.xlsx";

		FileOutputStream outputStream11 = new FileOutputStream(target1Path);
		workBookOutput1.write(outputStream11);
		workBookOutput1.close();
		
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
	}
}