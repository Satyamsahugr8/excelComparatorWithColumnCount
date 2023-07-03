package com.swing;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import javax.swing.JFrame;
import javax.swing.JOptionPane;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CountActiveInActive extends JFrame {
	
	
	@SuppressWarnings("unused")
	public int fetchExcel(String path1, String path2, int sheetNo1, int sheetNo2, int keyFile1, int keyFile2,
			String fileName1, String fileName2, String sheetName1, String sheetName2, String keyName1, String keyName2,
			String folderPath) {

		int counterMain = 0;

		try {

			FileInputStream file1 = new FileInputStream(path1);
			XSSFWorkbook workBook1 = new XSSFWorkbook(file1);
			XSSFSheet sheet1 = workBook1.getSheetAt(sheetNo1);

			FileInputStream file2 = new FileInputStream(path2);
			XSSFWorkbook workBook2 = new XSSFWorkbook(file2);
			XSSFSheet sheet2 = workBook2.getSheetAt(sheetNo2);

			// workBook1
			int totalNumberOfRowsInExcel1 = sheet1.getLastRowNum();
			int totalNumberOfColumnInExcel1 = sheet1.getRow(0).getLastCellNum();

			// workBook2
			int totalNumberOfRowsInExcel2 = sheet2.getLastRowNum();
			int totalNumberOfColumnInExcel2 = sheet2.getRow(0).getLastCellNum();
	
			// going to Excel1 key -> row = 1 to last
			for (int r = 1; r <= totalNumberOfRowsInExcel1; r++) {
				if (sheet1.getRow(r) == null) {
					continue;
				} else {

					if (sheet1.getRow(r).getCell(keyFile1) == null) {
						sheet1.removeRow(sheet1.getRow(r));
						continue;
					}

					// going to Excel2 key -> row = 1 to last
					for (int e = 1; e <= totalNumberOfRowsInExcel2; e++) {
						if (sheet2.getRow(e) == null) {
							continue;
						} else {
							if (sheet2.getRow(e).getCell(keyFile2) == null) {
								continue;
							}

							if ((sheet1.getRow(r).getCell(keyFile1).toString())
									.equals(sheet2.getRow(e).getCell(keyFile2).toString())) {

								XSSFRow rowOfSameKey1 = sheet1.getRow(r);
								sheet1.removeRow(rowOfSameKey1);
								break;
							}
						}
					}
				}
			} // for

//			System.out.println("-------------------------------------------------------------------------");

			String firstExcelPathCopy = path1;
			
			
			FileInputStream file1Copy = new FileInputStream(firstExcelPathCopy);
			XSSFWorkbook workBook1Copy = new XSSFWorkbook(file1Copy);
			XSSFSheet sheet1Copy = workBook1Copy.getSheetAt(sheetNo1);
			

			// going to Excel2 key -> row = 1 to last
			for (int rr = 1; rr <= totalNumberOfRowsInExcel2; rr++) {
				if (sheet2.getRow(rr) == null) {
					continue;
				} else {
					if (sheet2.getRow(rr).getCell(keyFile2) == null) {
						sheet2.removeRow(sheet2.getRow(rr));
						continue;
					}

					// going to Excel1 key -> row = 1 to last
					for (int e = 1; e <= totalNumberOfRowsInExcel1; e++) {
						if (sheet1Copy.getRow(e) == null) {
							continue;
						} else {
							if (sheet1Copy.getRow(e).getCell(keyFile1) == null) {
								continue;
							}

							if (sheet2.getRow(rr).getCell(keyFile2).toString().equals(sheet1Copy.getRow(e).getCell(keyFile1).toString())) {
								sheet2.removeRow(sheet2.getRow(rr));
								break;
							}
						}
					} // for
				}
			} // for
			workBook1Copy.close();
			// Upto here we have to two excel with some null or empty row
			// sheet1 and sheet2 as output only NO new sheet created

//-----------------------------------------------------------------------------------------------------------------			

			// counting null row in EXCEL 1
			int counter = 0;
			for (int rq = 0; rq <= totalNumberOfRowsInExcel1; rq++) {
				if (sheet1.getRow(rq) == null) {
					counter++;
				}
			}

//			System.out.println("totalNumberOfRows1:" + totalNumberOfRowsInExcel1);
//			System.out.println("counter:" + counter);

			// creating new working and adding new rows for excel1
			XSSFWorkbook workBookOutput1 = new XSSFWorkbook();
			XSSFSheet sheetCreate1 = workBookOutput1.createSheet();
			XSSFRow rowCreated = null;

			if (counter != 0) {

				int totalNumberOfRowsOfNewSheet = totalNumberOfRowsInExcel1 - counter;

//			System.out.println("totalNumberOfRowsOfNewSheet1:" + totalNumberOfRowsOfNewSheet);

				for (int rr = 0; rr <= totalNumberOfRowsOfNewSheet; rr++) {
					rowCreated = sheetCreate1.createRow(rr);

					for (int c = 0; c < totalNumberOfColumnInExcel1; c++) {
//					 XSSFCell cellCreated = rowCreated.createCell(c);
						rowCreated.createCell(c);
//						sheetCreate1.createRow(rr).createCell(c);
					}
				}

				for (int p = 0, u = 0; p <= totalNumberOfRowsInExcel1; p++) {
					if (sheet1.getRow(p) == null) {
						continue;
					} else {
						
						rowCreated = sheetCreate1.getRow(u);
						
						for (int d = 0; d < totalNumberOfColumnInExcel1; d++) {
							if (sheet1.getRow(p).getCell(d) == null) {
								continue;
							} else {
								if (sheet1.getRow(p).getCell(d).getCellType() == CellType.STRING) {
									rowCreated.getCell(d).setCellValue(sheet1.getRow(p).getCell(d).getStringCellValue());
								} else if (sheet1.getRow(p).getCell(d).getCellType() == CellType.NUMERIC) {
									rowCreated.getCell(d).setCellValue(sheet1.getRow(p).getCell(d).getNumericCellValue());
								} else if (sheet1.getRow(p).getCell(d).getCellType() == CellType.BOOLEAN) {
									rowCreated.getCell(d).setCellValue(sheet1.getRow(p).getCell(d).getBooleanCellValue());
								}
							}
						}
						u++;
					}
				}

//				System.out.println("NumberOfRow1:"+sheetCreate1.getLastRowNum());
//				if (sheetCreate1.getLastRowNum() > 0) {
				try {
					
					System.out.println("Unique Excel1 created");
					String target1Path = folderPath + "\\ChildOutput_ComparedBy_" + keyName1 + "_" + sheetName1 + "_"+ fileName1;
					
					FileOutputStream outputStream11 = new FileOutputStream(target1Path);
					workBookOutput1.write(outputStream11);
					
					
					
					workBookOutput1.close();
					
				} catch (FileNotFoundException ee) {
					counterMain++;
					JOptionPane.showMessageDialog(CountActiveInActive.this, "files 1 does'nt have unique data No Excel created",
							"Excel", JOptionPane.ERROR_MESSAGE);
				}

			} // if counter close

			else {

//				System.out.println("sheet1.getLastRowNum():"+sheet1.getLastRowNum());

//				if (sheet1.getLastRowNum() > 0) {
				try {
					System.out.println("Unique Excel1 created");
					String target1Path1 = folderPath + "\\ChildOutput_ComparedBy_" + keyName1 + "_" + sheetName1 + "_"
							+ fileName1;
					FileOutputStream outputStream1 = new FileOutputStream(target1Path1);
					workBook1.write(outputStream1);
					workBook1.close();
				} catch (FileNotFoundException ee) {
					counterMain++;
					JOptionPane.showMessageDialog(CountActiveInActive.this, "files 1 does'nt have unique data No Excel created",
							"Excel", JOptionPane.ERROR_MESSAGE);
				}
			}

			// counting null row in EXCEL 2
			int counter2 = 0;
			for (int r1 = 1; r1 <= totalNumberOfRowsInExcel2; r1++) {
				if (sheet2.getRow(r1) == null) {
					counter2++;
				}
			}

			// creating new working and adding new rows for excel2
			XSSFWorkbook workBookOutput2 = new XSSFWorkbook();
			XSSFSheet sheetCreate2 = workBookOutput2.createSheet();
			XSSFRow rowCreated2 = null;

			if (counter2 != 0) {

				int totalNumberOfRowsOfNewSheet2 = totalNumberOfRowsInExcel2 - counter2;

//			System.out.println("totalNumberOfRowsOfNewSheet2:" + totalNumberOfRowsOfNewSheet2);

//				rowCreated2 = sheetCreate2.createRow(r);
//				rowCreated2.createCell(c);
//				rowCreated2 = sheetCreate2.getRow(v);

//			XSSFCell cellCreated2 = null;

				for (int r2 = 0; r2 <= totalNumberOfRowsOfNewSheet2; r2++) {
					rowCreated2 = sheetCreate2.createRow(r2);
					for (int c = 0; c < totalNumberOfColumnInExcel2; c++) {
//				cellCreated2 = rowCreated2.createCell(c);
						rowCreated2.createCell(c);
					}
				}

				for (int p = 0, v = 0; p <= totalNumberOfRowsInExcel2; p++) {
					if (sheet2.getRow(p) == null) {
						continue;
					} else {
						rowCreated2 = sheetCreate2.getRow(v);
						for (int d = 0; d < totalNumberOfColumnInExcel2; d++) {
							if (sheet2.getRow(p).getCell(d) == null) {
								continue;
							} else {
//								rowCreated2.getCell(d).setCellValue(sheet2.getRow(p).getCell(d).toString());
								if (sheet2.getRow(p).getCell(d).getCellType() == CellType.STRING) {
									rowCreated2.getCell(d)
											.setCellValue(sheet2.getRow(p).getCell(d).getStringCellValue());
								} else if (sheet2.getRow(p).getCell(d).getCellType() == CellType.NUMERIC) {
									rowCreated2.getCell(d)
											.setCellValue(sheet2.getRow(p).getCell(d).getNumericCellValue());
								} else if (sheet2.getRow(p).getCell(d).getCellType() == CellType.BOOLEAN) {
									rowCreated2.getCell(d)
											.setCellValue(sheet2.getRow(p).getCell(d).getBooleanCellValue());
								}
							}
						}
					}
					v++;
				}

				// null row removed successfully
				// here we will have to two sheetCreate1 and sheetCreate2
				
				FileInputStream file1Count = new FileInputStream(folderPath + "\\ChildOutput_ComparedBy_" + keyName1 + "_" + sheetName1 + "_"+ fileName1);
				XSSFWorkbook workBookCount = new XSSFWorkbook(file1Count);
				XSSFSheet sheetCount = workBookCount.getSheetAt(sheetNo1);
				
				for (int i = 0; i < totalNumberOfColumnInExcel1; i++) {
					
					System.out.println(sheetCount.getRow(0).getCell(i).toString());
					
					if("status".equalsIgnoreCase(sheetCount.getRow(0).getCell(i).toString())) {
						int columnIndex = i;						
					}
					
				} 
				
	            // Initialize counters
	            int activeCount = 0;
	            int inactiveCount = 0;

	            // Iterate over the rows in the column
	            for (Row row : sheet1) {
	            	
	                int columnIndex = 0;
					Cell cell = row.getCell(columnIndex );

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
				
				
				
				
				
				
				
				

//				System.out.println("NumberOfRow2:"+sheetCreate2.getLastRowNum());
//				if (sheetCreate2.getLastRowNum() > 0) {
				try {
					System.out.println("Unique Excel2 created");
					String target1Path = folderPath + "\\ChildOutput_ComparedBy_" + keyName2 + "_" + sheetName2 + "_"
							+ fileName2;
					FileOutputStream outputStream22 = new FileOutputStream(target1Path);
					workBookOutput2.write(outputStream22);
					workBookOutput2.close();
				} catch (FileNotFoundException ee) {
					counterMain++;
					JOptionPane.showMessageDialog(CountActiveInActive.this, "files 2 does'nt have unique data No Excel created",
							"Excel", JOptionPane.ERROR_MESSAGE);
				}
//				} else {
//					counterMain++;
//					JOptionPane.showMessageDialog(Excel.this, "files 2 does'nt have unique data No Excel created",
//							"Excel", JOptionPane.ERROR_MESSAGE);
//				}
			} // if counter close
			else {
//				if (sheet2.getLastRowNum() > 0) {
				try {
					System.out.println("Unique Excel2 created");
					String target2Path2 = folderPath + "\\ChildOutput_ComparedBy_" + keyName2 + "_" + sheetName2 + "_"
							+ fileName2;
					FileOutputStream outputStream2 = new FileOutputStream(target2Path2);
					workBook2.write(outputStream2);
					workBook2.close();
				} catch (FileNotFoundException ee) {
					counterMain++;
					JOptionPane.showMessageDialog(CountActiveInActive.this, "files 2 does'nt have unique data No Excel created",
							"Excel", JOptionPane.ERROR_MESSAGE);
				}
//				} else {
//					counterMain++;
//					JOptionPane.showMessageDialog(Excel.this, "files 2 does'nt have unique data No Excel created",
//							"Excel", JOptionPane.ERROR_MESSAGE);
//				}
			}

			
			workBook1Copy.close();

//-----------------------------------------------------------------------------------------------------------

//			// upto unique data withOut Null row Completed
//			if (countDup1 != 0) {
//
//			} else {
//
//			}

			System.out.println("Unique......Done");

		} catch (Exception e) {
			e.printStackTrace();
		}

		return counterMain;
	} // end of fetch method
	
	
	
	
	
	
	
	
	public static void main(String[] args) throws EncryptedDocumentException, IOException {
		
		// Path to the existing Excel file
        String existingFilePath = "path/to/existing/file.xlsx";

        // Path to the new Excel file
        String newFilePath = "path/to/new/file.xlsx";

        	FileInputStream fis = new FileInputStream(new File(existingFilePath));
            Workbook workbook = WorkbookFactory.create(fis);
            FileOutputStream fos = new FileOutputStream(new File(newFilePath));
        
            // Get the first sheet of the workbook
            Sheet sheet = workbook.getSheetAt(0);

            // Get the column index (e.g., column A is 0, column B is 1, etc.)
            int columnIndex = 0; // Assuming the column is the first column (column A)

            // Initialize counters
            int activeCount = 0;
            int inactiveCount = 0;

            // Iterate over the rows in the column
            for (Row row : sheet) {
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
    }
}
