package com.practise.poi.filetask;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.concurrent.atomic.AtomicReference;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Component;

import com.practise.poi.domain.Country;

import lombok.extern.slf4j.Slf4j;

@Component
@Slf4j
public class ReadExcelFileToList {

	public static List<Country> readExcelData(String fileName) {
		List<Country> countriesList = new ArrayList<>();

		try {
			// Create the input stream from the xlsx/xls file
			FileInputStream fis = new FileInputStream(fileName); // Accepting the file name here

			// Create Workbook instance for xlsx/xls file input stream
			Workbook workbook = null;
			if (fileName.toLowerCase().endsWith("xlsx")) {
				workbook = new XSSFWorkbook(fis);
			} else if (fileName.toLowerCase().endsWith("xls")) {
				workbook = new HSSFWorkbook(fis);
			}

			if (workbook != null) { // If workbook is not empty/null then proceed
				workbook.forEach(sheet -> sheet.forEach(row -> {// workbook contains many sheet , then for each sheet we
																// process further and each sheet will have many rows
																// then for each row we process further
					AtomicReference<String> name = new AtomicReference<>("");
					AtomicReference<String> shortCode = new AtomicReference<>("");
					row.forEach(cell -> { // each row will have many cells then for each cell we process further.

						switch (cell.getCellType()) {
						case STRING:
							if (shortCode.get().equalsIgnoreCase("")) {
								shortCode.set(cell.getStringCellValue());
								log.info("Short Code :: " + shortCode);
							} else if (name.get().equalsIgnoreCase("")) {
								// 2nd column
								name.set(cell.getStringCellValue());
								log.info("Name :: " + name);
							} else {
								// random data, leave it
								log.info("Random data::" + cell.getStringCellValue());
							}
							break;
						case NUMERIC:
							log.info("Random data::" + cell.getNumericCellValue());
						}

					});
					Country ct = Country.builder().countryCode(shortCode.get()).countryName(name.get()).build();
					countriesList.add(ct);
					log.info("Country Name :: Code -> " + ct.getCountryName() + " :: " + ct.getCountryCode());
				}));
			}

			// close file input stream
			fis.close();

		} catch (IOException e) {
			e.printStackTrace();
		}

		return countriesList;
	}

	public String readExcelData() {

		List<Country> list = readExcelData("Country Data.xlsx");
		log.info("Size " + list.size());
		if (list.isEmpty())
			return "Failure while reading excel";
		else
			return "Excel File reading was successfull";

	}

}
