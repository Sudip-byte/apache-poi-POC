package com.practise.poi.filetask;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFName;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Component;

@Component
public class DropDownTask {

	public void dropDownTask() throws IOException {

		FileInputStream fis = new FileInputStream("DropFtr.xlsx");
		Workbook workbook = new XSSFWorkbook(fis);
		Sheet sheet = workbook.getSheet("Data");
		
		/*
		 * Name namedRange = workbook.createName(); namedRange.setNameName("CHOICES");
		 * namedRange.setRefersToFormula("$D$1:$E$1");
		 */
		
		XSSFName name = (XSSFName) workbook.createName();
		name.setNameName("CHOICES");
		name.setRefersToFormula("'Data'!$D$2:$D$3");
		XSSFName name1 = (XSSFName) workbook.createName();
		name1.setNameName("NO_CHOICES");
		name1.setRefersToFormula("'Data'!$E$2:$E$3");
		
		DataValidationHelper dvHelper = sheet.getDataValidationHelper();
		CellRangeAddressList addressList = new CellRangeAddressList(0, 0, 0, 0);
		DataValidationConstraint dvConstraint = dvHelper.createFormulaListConstraint("$D$1:$E$1");
		DataValidation dataValidation = dvHelper.createValidation(dvConstraint, addressList);
		dataValidation.setSuppressDropDownArrow(true);
		sheet.addValidationData(dataValidation);

		CellRangeAddressList addressList1 = new CellRangeAddressList(0, 0, 1, 1);
		DataValidationConstraint dvConstraint1 = dvHelper.createFormulaListConstraint("INDIRECT($A$1)");
		DataValidation dataValidation1 = dvHelper.createValidation(dvConstraint1, addressList1);
		dataValidation1.setSuppressDropDownArrow(true);
		sheet.addValidationData(dataValidation1);

		FileOutputStream fileOut = new FileOutputStream("output/DropFt.xlsx");
		try {
			workbook.write(fileOut);
			fileOut.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

}
