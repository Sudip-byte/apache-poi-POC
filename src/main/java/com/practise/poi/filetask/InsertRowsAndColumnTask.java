package com.practise.poi.filetask;

import com.spire.xls.ExcelVersion;
import com.spire.xls.Workbook;
import com.spire.xls.Worksheet;

import org.springframework.stereotype.Component;

import lombok.extern.slf4j.Slf4j;

@Component
@Slf4j
public class InsertRowsAndColumnTask {

	public String insertRowCol() {

		Workbook workbook = new Workbook();
		workbook.loadFromFile("Exams list.xlsx");

		Worksheet worksheet = workbook.getWorksheets().get(0);

		worksheet.insertRow(203);
		worksheet.insertColumn(2);

		worksheet.insertRow(207, 2);
		worksheet.insertColumn(5, 2);

		workbook.saveToFile("output/Exams list modified.xlsx", ExcelVersion.Version2013);

		return "Done";

	}

}
