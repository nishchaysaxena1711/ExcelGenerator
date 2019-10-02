package com.project.excelsheet;

import lombok.Getter;

@Getter
public class LedStatusBean {
	private String date;
	private String ledNumber;
	private String productType;
	private String productCategoryReference;
	private String coMakersName;
	private String serialNumber;
	private String driverMake;
	private String ledMake;
	private String switchingTest;
	private String lifeTest;
	private String startDate;
	private String startMonth;
	private String startYear;
	private String startHour;
	private String startMinute;
	private String endDate;
	private String endMonth;
	private String endYear;
	private String endHour;
	private String endMinute;
	private String setCyles;
	private String completedCycles;
	private String status;
	private String testedBy;
	private String approvedBy;

	public LedStatusBean(String[] metadata) {
		this.date = metadata[0];
		this.ledNumber = metadata[1];
		this.productType = metadata[2];
		this.productCategoryReference = metadata[3];
		this.coMakersName = metadata[4];
		this.serialNumber = metadata[5];
		this.driverMake = metadata[6];
		this.ledMake = metadata[7];
		this.switchingTest = metadata[8];
		this.lifeTest = metadata[9];
		this.startDate = metadata[10];
		this.startMonth = metadata[11];
		this.startYear = metadata[12];
		this.startHour = metadata[13];
		this.startMinute = metadata[14];
		this.endDate = metadata[15];
		this.endMonth = metadata[16];
		this.endYear = metadata[17];
		this.endHour = metadata[18];
		this.endMinute = metadata[19];
		this.setCyles = metadata[20];
		this.completedCycles = metadata[21];
		this.status = metadata[22];
		this.testedBy = metadata[23];
		this.approvedBy = metadata[24];
	}

	@Override
	public String toString() {
		return "Led properties [" +
				"date = " + date +
				", ledNumber = " + ledNumber +
				", productType = " + productType +
				", productCategoryReference = " + productCategoryReference + "\n" +
				", coMakersName = " + coMakersName +
				", serialNumber = " + serialNumber +
				", driverMake = " + driverMake +
				", ledMake = " + ledMake + "\n" +
				", switchingTest = " + switchingTest +
				", lifeTest = " + lifeTest +
				", startDate = " + startDate +
				", startMonth = " + startMonth + "\n" +
				", startYear = " + startYear +
				", startHour = " + startHour +
				", startMinute = " + startMinute +
				", endDate = " + endDate + "\n" +
				", endMonth = " + endMonth +
				", endYear = " + endYear +
				", endHour = " + endHour +
				", endMinute = " + endMinute + "\n" +
				", setCyles = " + setCyles +
				", completedCycles = " + completedCycles +
				", status = " + status +
				", testedBy = " + testedBy +
				", approvedBy = " + approvedBy +
				"]";
	}
}
