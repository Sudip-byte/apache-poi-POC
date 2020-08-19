package com.practise.poi.controller;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import com.practise.poi.filetask.InsertRowsAndColumnTask;
import com.practise.poi.filetask.ReadExcelFileToList;

@RestController
public class ExcelOperationsController {
	
	@Autowired
	private ReadExcelFileToList readExcel;
	
	@Autowired
	private InsertRowsAndColumnTask insertRowCol;
	
	@GetMapping("/read")
	public String readExcel()
	{
		return readExcel.readExcelData();
	}
	
	@PostMapping("/insertRowCol")
	public String insertRowColInExcel()
	{
		return insertRowCol.insertRowCol();
	}

}
