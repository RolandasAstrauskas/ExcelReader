package excelReader;

import java.util.ArrayList;

public class ListOfWorkTime {

	ArrayList<ArrayList<WorkTimeItemArrayList>> workTimeItemArrayList;

	public ListOfWorkTime(ArrayList<ArrayList<WorkTimeItemArrayList>> workTimeItemArrayList) {

		this.workTimeItemArrayList = workTimeItemArrayList;
	}

	public ArrayList<ArrayList<WorkTimeItemArrayList>> getWorkTimeItem() {
		return workTimeItemArrayList;
	}

}
