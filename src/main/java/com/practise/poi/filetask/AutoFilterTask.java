package com.practise.poi.filetask;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.concurrent.atomic.AtomicInteger;

import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Component;

import com.spire.ms.System.Collections.ArrayList;

import lombok.extern.slf4j.Slf4j;

@Component
@Slf4j
public class AutoFilterTask {

	public void setAutoFilter() throws IOException {

		FileInputStream fis = new FileInputStream("Exams list.xlsx");

		Workbook workbook = new XSSFWorkbook(fis);
		Sheet sheet = workbook.getSheet("Exam List");

		sheet.setAutoFilter(new CellRangeAddress(0, 0, 0, 3));
		
		  DataValidationHelper dvHelper = sheet.getDataValidationHelper();
		  
		  CellRangeAddressList addressList = new CellRangeAddressList(3, 3, 1, 1);
		  DataValidationConstraint dvConstraint = dvHelper.createCustomConstraint("NOT(AND(D4=100,B4=\"XXX\"))");
		  DataValidation dataValidation = dvHelper.createValidation(dvConstraint,
		  addressList);
		  
		  dataValidation.setEmptyCellAllowed(true);
		  dataValidation.setShowErrorBox(true);
		  dataValidation.setErrorStyle(DataValidation.ErrorStyle.STOP); 
		  dataValidation.createErrorBox("Invalid Data", "Provide valid cell data");
		  
		  //dataValidation.setSuppressDropDownArrow(true);
		  sheet.addValidationData(dataValidation);
		 
		/*
		 * CellRangeAddress range = new CellRangeAddress(1, 314, 1, 1);
		 * 
		 * sheet.setArrayFormula("IF($D$2:$D$315=100,\"SUBJECT MISSING\",$B$2:$B$315)",
		 * range);
		 */

		FileOutputStream fileOut = new FileOutputStream("output/Exams-filtered.xlsx");
		workbook.write(fileOut);
		workbook.close();
		fileOut.close();

	}

}
