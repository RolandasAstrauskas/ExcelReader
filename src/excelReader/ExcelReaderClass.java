package excelReader;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReaderClass {

	public ArrayList<WorkTimeItemArrayList> excelReader(File excelFile, int numberOfEmplyee, int month)
			throws IOException {

		ArrayList<WorkTimeItemArrayList> workTimeItemArrayList = new ArrayList<WorkTimeItemArrayList>();
		WorkTimeItemArrayList listOfNameAndTime;

		int plusRow = -4;

		FileInputStream fis = new FileInputStream(excelFile);

		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		XSSFSheet sheet = workbook.getSheetAt(month);

		for (int a = 0; a < numberOfEmplyee; a++) {

			List<WorkTimeItem> dayTimeList = new ArrayList<WorkTimeItem>();

			plusRow = plusRow + 4;

			
			XSSFRow rowName = sheet.getRow(11 + plusRow);

			XSSFCell cellName = rowName.getCell(1);

			String cellValueName = getCellValue(cellName);

			for (int i = 0; i < 1; i++) {

				
				XSSFRow row = sheet.getRow(10);

				for (int j = 0; j < 31; j++) {

					XSSFCell cell = row.getCell(j + 4);

					if (!isInteger(getCellValue(cell))) {

						break;

					} else {
						String cellValue = getCellValue(cell);

						int cellNumericValueMonthDay = Integer.parseInt(cellValue);

						dayTimeList.add(new WorkTimeItem(cellNumericValueMonthDay));

					}
				}
			}

			for (int i = 1; i < 2; i++) {
				
				XSSFRow row = sheet.getRow(i + 10 + plusRow);

				for (int j = 0; j < dayTimeList.size(); j++) {

					XSSFCell cell = row.getCell(j + 4);

					if (cell == null) {
						continue;

					} else {
						String cellValue = getCellValue(cell);

						WorkTimeItem workTimeItem = dayTimeList.get(j);

						dayTimeList.set(j, new WorkTimeItem(workTimeItem.getMonthDay(), cellValue));

					}
				}
			}

			for (int i = 2; i < 3; i++) {
				
				XSSFRow row = sheet.getRow(i + 10 + plusRow);

				for (int j = 0; j < dayTimeList.size(); j++) {

					XSSFCell cell = row.getCell(j + 4);

					if (cell == null) {
						continue;

					} else {
						String cellValue = getCellValue(cell);

						WorkTimeItem workTimeItem = dayTimeList.get(j);

						dayTimeList.set(j,
								new WorkTimeItem(workTimeItem.getMonthDay(), workTimeItem.getWorkStart(), cellValue));

					}
				}
			}

			for (int i = 3; i < 4; i++) {
				
				XSSFRow row = sheet.getRow(i + 10 + plusRow);

				for (int j = 0; j < dayTimeList.size(); j++) {

					XSSFCell cell = row.getCell(j + 4);

					if (cell == null) {
						continue;

					} else {
						String cellValue = getCellValue(cell);

						WorkTimeItem workTimeItem = dayTimeList.get(j);

						dayTimeList.set(j, new WorkTimeItem(workTimeItem.getMonthDay(), workTimeItem.getWorkStart(),
								workTimeItem.getWorkEnd(), cellValue));

					}
				}
			}

			for (int i = 4; i < 5; i++) {
				
				XSSFRow row = sheet.getRow(i + 10 + plusRow);

				for (int j = 0; j < dayTimeList.size(); j++) {

					XSSFCell cell = row.getCell(j + 4);

					if (cell == null) {
						continue;

					} else {
						String cellValue = getCellValue(cell);

						WorkTimeItem workTimeItem = dayTimeList.get(j);

						dayTimeList.set(j, new WorkTimeItem(workTimeItem.getMonthDay(), workTimeItem.getWorkStart(),
								workTimeItem.getWorkEnd(), workTimeItem.getLunchStart(), cellValue));

					}
				}
			}

			listOfNameAndTime = new WorkTimeItemArrayList((ArrayList<WorkTimeItem>) dayTimeList, cellValueName);

			workTimeItemArrayList.add(listOfNameAndTime);

		}

		workbook.close();
		fis.close();

		return workTimeItemArrayList;

	}

	private static String getCellValue(XSSFCell cell) {

		switch (cell.getCellType()) {
		case NUMERIC:
			return cell.getRawValue();

		case BOOLEAN:
			return String.valueOf(cell.getBooleanCellValue());

		case STRING:
			return cell.getStringCellValue();

		case FORMULA:
			return cell.getRawValue();

		default:
			return cell.getStringCellValue();
		}

	}

	private static boolean isInteger(String input) {
		try {
			Integer.parseInt(input);
			return true;
		} catch (Exception e) {
			return false;
		}
	}
}
