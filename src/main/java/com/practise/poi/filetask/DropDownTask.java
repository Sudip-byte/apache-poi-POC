package com.practise.poi.filetask;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFName;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Component;

import lombok.extern.slf4j.Slf4j;

@Component
@Slf4j
public class DropDownTask {

	public void dropDownTask() throws IOException {

		FileInputStream fis = new FileInputStream("DropFtr.xlsx");

		Workbook workbook = new XSSFWorkbook(fis);
		Sheet sheet = workbook.getSheet("Data");

		XSSFName name = (XSSFName) workbook.createName();
		name.setNameName("CHOICES");
		name.setRefersToFormula("'Data'!$D$2:$D$3");
		XSSFName name1 = (XSSFName) workbook.createName();
		name1.setNameName("NO_CHOICES");
		name1.setRefersToFormula("'Data'!$E$2:$E$3");
		XSSFName name2 = (XSSFName) workbook.createName();
		name2.setNameName("INC");
		name2.setRefersToFormula("'Data'!$G$2:$G$3");
		XSSFName name3 = (XSSFName) workbook.createName();
		name3.setNameName("COM");
		name3.setRefersToFormula("'Data'!$F$2:$F$2");

		DataValidationHelper dvHelper = sheet.getDataValidationHelper();
		
		CellRangeAddressList addressList = new CellRangeAddressList(0, 5, 0, 0);
		DataValidationConstraint dvConstraint = dvHelper.createFormulaListConstraint("$D$1:$E$1");
		DataValidation dataValidation = dvHelper.createValidation(dvConstraint, addressList);
		dataValidation.setSuppressDropDownArrow(true);
		sheet.addValidationData(dataValidation);

		CellRangeAddressList addressList1 = new CellRangeAddressList(0, 5, 1, 1);
		DataValidationConstraint dvConstraint1 = dvHelper.createFormulaListConstraint("INDIRECT($A1)");
		DataValidation dataValidation1 = dvHelper.createValidation(dvConstraint1, addressList1);
		dataValidation1.setSuppressDropDownArrow(true);
		sheet.addValidationData(dataValidation1);

		CellRangeAddressList addressList2 = new CellRangeAddressList(7, 12, 1, 1);
		DataValidationConstraint dvConstraint2 = dvHelper
				.createFormulaListConstraint("IF(A8=100,INDIRECT('Data'!$H$2),INDIRECT('Data'!$H$1))");
		DataValidation dataValidation2 = dvHelper.createValidation(dvConstraint2, addressList2);
		dataValidation2.setSuppressDropDownArrow(true);
		sheet.addValidationData(dataValidation2);
		
		// IF(C2="",Produce, INDIRECT("FakeRange"))
		
		/*
		 * sheet.forEach(row->{ double date = 0.0; int rowNum = row.getRowNum();
		 * 
		 * log.info("row number :: "+rowNum); if(rowNum!=0 &&
		 * row.getCell(16).getCellType().equals(CellType.NUMERIC)) {
		 * 
		 * date = row.getCell(16).getNumericCellValue(); } if(date!=0.0) {
		 * row.getCell(17).setCellValue(100); } });
		 */

		/*
		 * CellRangeAddress range = new CellRangeAddress(8, 9, 1, 1);
		 * 
		 * sheet.setArrayFormula("IF($A$9:$A$10=100,\"Completed\",$J$1:$J$2)", range);
		 */

		FileOutputStream fileOut = new FileOutputStream("output/DropFt.xlsx");
		try {
			workbook.write(fileOut);
			fileOut.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

}
