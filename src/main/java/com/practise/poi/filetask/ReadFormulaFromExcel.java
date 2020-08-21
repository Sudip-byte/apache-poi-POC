package com.practise.poi.filetask;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Component;

import lombok.extern.slf4j.Slf4j;

@Component
@Slf4j
public class ReadFormulaFromExcel {
	
	public String readFormula() throws IOException
	{
		return readExcelFormula("FormulaSheet.xlsx");
	}

	private String readExcelFormula(String fileName) throws IOException {
		
		FileInputStream fis = new FileInputStream(fileName);
		
		Workbook workbook = new XSSFWorkbook(fis);
		Sheet sheet = workbook.getSheetAt(0);
		
		sheet.forEach(row -> {
			row.forEach(cell-> {
				switch(cell.getCellType()){
            	case NUMERIC:
            		log.info("Numeric Value of cell :: "+cell.getNumericCellValue());
            		break;
            	case FORMULA:
            		log.info("Cell Formula="+cell.getCellFormula());
            		log.info("Cell Formula Result Type="+cell.getCachedFormulaResultType());
            		if(cell.getCachedFormulaResultType() == CellType.NUMERIC){
            			log.info("Formula Value="+cell.getNumericCellValue());
            		}
            	}
			});
		});
		
		workbook.close();
		
		return "FORMULA EXTRACTION COMPLETE";
	}

}
