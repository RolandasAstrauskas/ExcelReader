package excelReader;

import java.io.File;
import java.util.ArrayList;

import org.apache.poi.openxml4j.util.ZipSecureFile;

public class MainExcel {

	public static void main(String[] args) throws Exception {

		int numberOfEmply = 14;
		int month = 1;

		File excelFile = new File("Schedule");

		File excelFile2 = new File("ExcelFile");
		ZipSecureFile.setMinInflateRatio(-1.0d);

		String writeOut = "ExcelFileWhereToSave";

		ExcelReaderClass excelReader = new ExcelReaderClass();
		ExcelWriterClass excelWriterClass = new ExcelWriterClass();

		ArrayList<WorkTimeItemArrayList> workTimeItemArrayList2 = excelReader.excelReader(excelFile, numberOfEmply,
				month);

		for (int b = 0; b < numberOfEmply; b++) {

			WorkTimeItemArrayList workTimeItemArrayList = workTimeItemArrayList2.get(b);
			String employeeName = workTimeItemArrayList.getEmplyeeNameLastName();
			ArrayList<WorkTimeItem> workTimeItemList = workTimeItemArrayList.getWorkTimeItem();
			excelWriterClass.excelWrite(excelFile2, workTimeItemList, b, employeeName, writeOut);

		}

	}
}
