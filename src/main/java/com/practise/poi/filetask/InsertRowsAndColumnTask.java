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
		workbook.loadFromFile("Country Data.xlsx");
		
		Worksheet worksheet = workbook.getWorksheets().get(0);
		
		 worksheet.insertRow(2);
		 worksheet.insertColumn(2);
		 
		 worksheet.insertRow(5, 2);
		 worksheet.insertColumn(5, 2);
		 
		 workbook.saveToFile("output/InsertRowsAndColumns.xslm", ExcelVersion.Version2013);
		 
		 return "Done";

	}

}
