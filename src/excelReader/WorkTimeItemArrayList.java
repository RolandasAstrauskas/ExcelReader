package excelReader;

import java.util.ArrayList;

public class WorkTimeItemArrayList {

	private ArrayList<WorkTimeItem> workTimeItem;
	private String emplyeeNameLastName;
	

	public WorkTimeItemArrayList(ArrayList<WorkTimeItem> workTimeItem, String emplyeeNameLastName) {
		this.workTimeItem = workTimeItem;
		this.emplyeeNameLastName = emplyeeNameLastName;
	}

	public ArrayList<WorkTimeItem> getWorkTimeItem() {
		return workTimeItem;
	}

	public String getEmplyeeNameLastName() {
		return emplyeeNameLastName;
	}

	
	
}
