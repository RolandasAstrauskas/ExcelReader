package excelReader;

public class WorkTimeItem {

	private int monthDay;
	private String workStart;
	private String workEnd;
	private String lunchStart;
	private String lunchEnd;

	public WorkTimeItem(int monthDay, String workStart, String workEnd, String lunchStart, String lunchEnd) {
		this.monthDay = monthDay;
		this.workStart = workStart;
		this.workEnd = workEnd;
		this.lunchStart = lunchStart;
		this.lunchEnd = lunchEnd;
	}

	public WorkTimeItem(int monthDay, String workStart, String workEnd, String lunchStart) {
		this.monthDay = monthDay;
		this.workStart = workStart;
		this.workEnd = workEnd;
		this.lunchStart = lunchStart;
	}

	public WorkTimeItem(int monthDay, String workStart, String workEnd) {
		this.workEnd = workEnd;
		this.workStart = workStart;
		this.monthDay = monthDay;
	}

	public WorkTimeItem(int monthDay, String workStart) {
		this.workStart = workStart;
		this.monthDay = monthDay;
	}

	public WorkTimeItem(int monthDay) {
		this.monthDay = monthDay;
	}

	public WorkTimeItem() {
	}

	public int getMonthDay() {
		return monthDay;
	}

	public String getWorkStart() {
		return workStart;
	}

	public String getWorkEnd() {
		return workEnd;
	}

	public String getLunchStart() {
		return lunchStart;
	}

	public String getLunchEnd() {
		return lunchEnd;
	}

}
