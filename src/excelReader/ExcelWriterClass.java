package excelReader;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelWriterClass {

	public void excelWrite(File excelFile, ArrayList<WorkTimeItem> workTimeList, int rowNumber, String employeeName,
			String writeOut) throws Exception {

		FileInputStream fis = new FileInputStream(excelFile);

		XSSFWorkbook wb = new XSSFWorkbook(fis);
		CellStyle backgroundStyleP = wb.createCellStyle();
		CellStyle backgroundStyleA = wb.createCellStyle();
		CellStyle backgroundStyleS = wb.createCellStyle();
		CellStyle backgroundStyleTable = wb.createCellStyle();

		backgroundStyleP.setFillForegroundColor(IndexedColors.GREEN.getIndex());
		backgroundStyleP.setFillPattern(FillPatternType.SOLID_FOREGROUND);

		backgroundStyleP.setBorderTop(BorderStyle.THIN);
		backgroundStyleP.setBorderBottom(BorderStyle.THIN);
		backgroundStyleP.setBorderLeft(BorderStyle.THIN);
		backgroundStyleP.setBorderRight(BorderStyle.THIN);

		backgroundStyleP.setLeftBorderColor(IndexedColors.BLACK.index);
		backgroundStyleP.setRightBorderColor(IndexedColors.BLACK.index);
		backgroundStyleP.setTopBorderColor(IndexedColors.BLACK.index);
		backgroundStyleP.setBottomBorderColor(IndexedColors.BLACK.index);

		backgroundStyleA.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
		backgroundStyleA.setFillPattern(FillPatternType.SOLID_FOREGROUND);

		backgroundStyleA.setBorderTop(BorderStyle.THIN);
		backgroundStyleA.setBorderBottom(BorderStyle.THIN);
		backgroundStyleA.setBorderLeft(BorderStyle.THIN);
		backgroundStyleA.setBorderRight(BorderStyle.THIN);

		backgroundStyleA.setLeftBorderColor(IndexedColors.BLACK.index);
		backgroundStyleA.setRightBorderColor(IndexedColors.BLACK.index);
		backgroundStyleA.setTopBorderColor(IndexedColors.BLACK.index);
		backgroundStyleA.setBottomBorderColor(IndexedColors.BLACK.index);

		backgroundStyleS.setFillForegroundColor(IndexedColors.RED.getIndex());
		backgroundStyleS.setFillPattern(FillPatternType.SOLID_FOREGROUND);

		backgroundStyleS.setBorderTop(BorderStyle.THIN);
		backgroundStyleS.setBorderBottom(BorderStyle.THIN);
		backgroundStyleS.setBorderLeft(BorderStyle.THIN);
		backgroundStyleS.setBorderRight(BorderStyle.THIN);

		backgroundStyleS.setLeftBorderColor(IndexedColors.BLACK.index);
		backgroundStyleS.setRightBorderColor(IndexedColors.BLACK.index);
		backgroundStyleS.setTopBorderColor(IndexedColors.BLACK.index);
		backgroundStyleS.setBottomBorderColor(IndexedColors.BLACK.index);

		backgroundStyleTable.setBorderTop(BorderStyle.THIN);
		backgroundStyleTable.setBorderBottom(BorderStyle.THIN);
		backgroundStyleTable.setBorderLeft(BorderStyle.THIN);
		backgroundStyleTable.setBorderRight(BorderStyle.THIN);

		backgroundStyleTable.setLeftBorderColor(IndexedColors.BLACK.index);
		backgroundStyleTable.setRightBorderColor(IndexedColors.BLACK.index);
		backgroundStyleTable.setTopBorderColor(IndexedColors.BLACK.index);
		backgroundStyleTable.setBottomBorderColor(IndexedColors.BLACK.index);

		XSSFSheet sh;
		XSSFRow row;

		int rowNumberNew = rowNumber + 1;

		for (int s = 0; s < workTimeList.size(); s++) {

			sh = wb.getSheetAt(s);
			row = sh.createRow(0);
			double plusHour = 7.5;
			for (int u = 0; u < 28; u++) {
				plusHour = plusHour + 0.5;
				Cell cell = row.createCell(8 + u);
				cell.setCellValue(plusHour);
				cell.setCellStyle(backgroundStyleTable);

			}

		}

		for (int i = 0; i < workTimeList.size(); i++) {

			WorkTimeItem workTimeItem = workTimeList.get(i);

			sh = wb.getSheetAt(i);
			row = sh.createRow(rowNumberNew);

			Cell cell = row.createCell(6);
			cell.setCellValue(employeeName);
			cell.setCellStyle(backgroundStyleTable);

			int start;
			int end;

			if (workTimeItem.getWorkStart().equals("P")) {
				start = 0;
				end = 0;

				for (int j = 8; j < 36; j++) {
					Cell cell2 = row.createCell(j);
					cell2.setCellStyle(backgroundStyleP);
					sh.setColumnWidth(j, 1500);

				}

			} else if (workTimeItem.getWorkStart().equals("A")) {
				start = 0;
				end = 0;

				for (int j = 8; j < 36; j++) {
					Cell cell2 = row.createCell(j);
					cell2.setCellStyle(backgroundStyleA);
					sh.setColumnWidth(j, 1500);

				}

			} else if (workTimeItem.getWorkStart().equals("S")) {

				start = 0;
				end = 0;

				for (int j = 8; j < 36; j++) {
					Cell cell2 = row.createCell(j);
					cell2.setCellStyle(backgroundStyleS);
					sh.setColumnWidth(j, 1500);
				}

			} else if (!isInteger(workTimeItem.getWorkStart())) {

				start = 0;
				end = 0;

			} else {

				start = Integer.parseInt(workTimeItem.getWorkStart());
				end = (Integer.parseInt(workTimeItem.getWorkEnd()));

			}

			int endNew = (end + end - 8);
			int startNew;

			if (start == 0) {
				startNew = 0;

			} else {
				startNew = (start - 8) + start;
			}

			for (int j = startNew; j < endNew; j++) {

				Cell cell2 = row.createCell(j);
				cell2.setCellStyle(backgroundStyleTable);
				cell2.setCellValue("V");
				sh.setColumnWidth(j, 1500);
			}
		}

		FileOutputStream out = new FileOutputStream(writeOut);
		wb.write(out);

		out.close();
		wb.close();

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
