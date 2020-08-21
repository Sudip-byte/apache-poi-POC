package com.practise.poi.filetask;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.Objects;
import java.util.concurrent.atomic.AtomicInteger;

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
public class WriteToExcelFile {

	public String executeWriteOperation() throws IOException {
		List<Country> list = ReadExcelFileToList.readExcelData("Country Data.xlsx");
		log.info("List :: " + list);
		writeCountryListToFile("Countries.xls", list);

		return "Write Successful";
	}

	private void writeCountryListToFile(String fileName, List<Country> list) throws IOException {

		Workbook workbook = null;

		if (fileName.endsWith("xlsx")) {
			workbook = new XSSFWorkbook();
		} else if (fileName.endsWith("xls")) {
			workbook = new HSSFWorkbook();
		}

		if (Objects.nonNull(workbook)) {

			Sheet sheet = workbook.createSheet("Countries");
			
			AtomicInteger rowIndex = new AtomicInteger(0);
			list.forEach(country -> {

				Row row = sheet.createRow(rowIndex.getAndIncrement());
				Cell cell0 = row.createCell(0);
				cell0.setCellValue(country.getCountryName());

				Cell cell1 = row.createCell(1);
				cell1.setCellValue(country.getCountryCode());

			});

			FileOutputStream fos = new FileOutputStream("Country-Modified.xlsx");
			workbook.write(fos);
			fos.close();

		}
	}

}
