package votw;


import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;
import java.util.logging.Logger;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;



public class ExcelReader {

	private static Logger logger = Logger.getLogger(ExcelReader.class.getSimpleName());

	public static String[][] readExcelData(String filePath, String sheetName,
			String tableName) {
		String[][] testData = null;
		System.setProperty("sCurrentTestCase", sheetName);
		System.setProperty("sDataSheet", tableName);
		logger.info("Reading file, sheet, table: " + filePath + ", " + sheetName + ", " + tableName);
		try {
			HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(
					filePath));
			HSSFSheet sheet = workbook.getSheet(sheetName);
			HSSFCell[] boundaryCells = findCell(sheet, tableName);
			HSSFCell startCell = boundaryCells[0];
			HSSFCell endCell = boundaryCells[1];
			int startRow = startCell.getRowIndex() + 1;
			int endRow = endCell.getRowIndex();
			int startCol = startCell.getColumnIndex() + 1;
			int endCol = endCell.getColumnIndex() - 1;
			testData = new String[endRow - startRow + 1][endCol - startCol + 1];
			for (int i = startRow; i < endRow + 1; i++) {
				for (int j = startCol; j < endCol + 1; j++) {
					if (sheet.getRow(i).getCell(j).getCellType() == HSSFCell.CELL_TYPE_STRING) {
						testData[i - startRow][j - startCol] = sheet.getRow(i)
								.getCell(j).getStringCellValue();
					} else if (sheet.getRow(i).getCell(j).getCellType() == HSSFCell.CELL_TYPE_NUMERIC) {
						Double temp = sheet.getRow(i).getCell(j)
								.getNumericCellValue();
						testData[i - startRow][j - startCol] = String
								.valueOf(temp.intValue());
					}
				}
			}
		} catch (FileNotFoundException fnfe) {
			logger.severe("Could not read the Excel sheet");
			fnfe.printStackTrace();
		} catch (IOException ioe) {
			logger.severe("Could not read the Excel sheet");
			ioe.printStackTrace();
		}
		return testData;
	}

	public static HSSFCell[] findCell(HSSFSheet sheet, String text) {
		String pos = "start";
		HSSFCell[] cells = new HSSFCell[2];
		for (Row row : sheet) {
			for (Cell cell : row) {
				if (cell.getCellType() == HSSFCell.CELL_TYPE_STRING
						&& text.equals(cell.getStringCellValue())) {
					if (pos.equalsIgnoreCase("start")) {
						cells[0] = (HSSFCell) cell;
						// System.out.println("First Cell found at:"+cell.getRowIndex()+","+cell.getColumnIndex());
						pos = "end";
					} else {
						cells[1] = (HSSFCell) cell;
						// System.out.println("First Cell found at:"+cell.getRowIndex()+","+cell.getColumnIndex());
					}
				}
			}
		}
		return cells;
	}

	/*public String switchExcelFileForDevice(String fileName) {
		String deviceType = TestBed.deviceType;
		if (deviceType != null) {
			if (deviceType.equalsIgnoreCase("phone")
					|| deviceType.equalsIgnoreCase("tablet")) {
				String[] name = fileName.split(".xls");
				fileName = name[0] + "_" + deviceType + ".xls";

			}
		}
		System.out.println("Excel sheet name = " + fileName);
		return fileName;
	}*/

	/**
	 * CommonMethod(readExcelDataByRow) which reads the data from the excel
	 * sheet by Row Sunil
	 */
	public Map<String, Object> readExcelDataByRow(String filePath,
			String sheetName, int iRow) {
		Map<String, Object> data = new HashMap();

		try {
			HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(
					filePath));
			HSSFSheet sheet = workbook.getSheet(sheetName);
			int iCol = 1;
			while (sheet.getRow(iRow).getCell(iCol) != null) {
				if (sheet.getRow(iRow).getCell(iCol).getCellType() == HSSFCell.CELL_TYPE_STRING) {
					data.put(sheet.getRow(0).getCell(iCol).toString(), sheet
							.getRow(iRow).getCell(iCol).getStringCellValue()
							.trim());
				}
				if (sheet.getRow(iRow).getCell(iCol).getCellType() == HSSFCell.CELL_TYPE_NUMERIC) {
					data.put(sheet.getRow(0).getCell(iCol).toString(), sheet
							.getRow(iRow).getCell(iCol).getNumericCellValue());
				}
				iCol = iCol + 1;
			}

		} catch (FileNotFoundException e) {
			System.out.println("Could not read the Excel sheet");
			e.printStackTrace();
		} catch (IOException e) {
			System.out.println("Could not read the Excel sheet");
			e.printStackTrace();
		}
		return data;
	}

	/**
	 * CommonMethod(writeDataInExcel) which write data in the excel sheet by Row
	 * Sunil
	 */
	public Boolean writeDataInExcel(String sFilePath, String sSheetName,
			String sTableName, Map<String, String> mData) {
		Boolean bStatus = false;
		try {
			FileInputStream eFileInputStream = new FileInputStream(sFilePath);
			HSSFWorkbook workbook = new HSSFWorkbook(eFileInputStream);
			HSSFSheet sheet = workbook.getSheet(sSheetName);
			HSSFCell[] boundaryCells = findCell(sheet, sTableName);
			HSSFCell startCell = boundaryCells[0];
			HSSFCell endCell = boundaryCells[1];
			int startRow = startCell.getRowIndex();
			int endRow = endCell.getRowIndex();
			int startCol = startCell.getColumnIndex() + 1;
			int endCol = endCell.getColumnIndex() - 1;
			int iCurrentRow = 0;
			for (int i = startRow; i < endRow + 1; i++) {
				for (int j = startCol; j < endCol + 1; j++) {
					String sColoumnName = sheet.getRow(i).getCell(j)
							.getStringCellValue().trim();
					if (mData.containsKey(sColoumnName)) {
						try {
							sheet.getRow(i + 1)
									.getCell(j)
									.setCellValue(
											mData.get(sColoumnName).toString()
													.trim());
							bStatus = true;
						} catch (Exception e) {
							if (iCurrentRow != i + 1) {
								sheet.createRow(i + 1);
								iCurrentRow = i + 1;
							}
							sheet.getRow(i + 1).createCell(j);
							sheet.getRow(i + 1)
									.getCell(j)
									.setCellValue(
											mData.get(sColoumnName).toString()
													.trim());
						}
					}

					// Add the table name if it was removed
					if ((i == endRow - 1) && (j == endCol - 1)) {
						sheet.getRow(i).getCell(j).setCellValue(sTableName);
					}

				}

			}

			eFileInputStream.close();
			FileOutputStream output_file = new FileOutputStream(sFilePath);
			workbook.write(output_file);
			output_file.close();
		} catch (Exception e) {
			e.printStackTrace();
		}

		return bStatus;
	}
}
