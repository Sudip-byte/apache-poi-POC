package com.practise.poi.controller;

import java.io.IOException;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RestController;

import com.practise.poi.filetask.AutoFilterTask;
import com.practise.poi.filetask.DropDownTask;
import com.practise.poi.filetask.ImplementMacroCode;
import com.practise.poi.filetask.InsertRowsAndColumnTask;
import com.practise.poi.filetask.ReadExcelFileToList;
import com.practise.poi.filetask.ReadFormulaFromExcel;
import com.practise.poi.filetask.WriteToExcelFile;

@RestController
public class ExcelOperationsController {
	
	@Autowired
	private ReadExcelFileToList readExcel;
	
	@Autowired
	private InsertRowsAndColumnTask insertRowCol;
	
	@Autowired
	private WriteToExcelFile writeToExcel;
	
	@Autowired
	private ReadFormulaFromExcel readForumula;
	
	@Autowired
	private DropDownTask dropDownTask;
	
	@Autowired
	private AutoFilterTask filterTask;
	
	@Autowired
	private ImplementMacroCode macroCode;
	
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
	
	@GetMapping("/write")
	public String writeToExcel() throws IOException
	{
		return writeToExcel.executeWriteOperation();
	}
	
	@GetMapping("/readFormula")
	public String readFormula() throws IOException
	{
		return readForumula.readFormula();
	}
	
	@GetMapping("/generateDropdown")
	public void generateDropdown() throws IOException
	{
		
		dropDownTask.dropDownTask();
	}
	
	@GetMapping("/enableFilter")
	public void enableFilter() throws IOException
	{
		
		filterTask.setAutoFilter();
	}
	
	@GetMapping("/enableMacro")
	public void enableMacro() throws Exception
	{
		
		macroCode.addMacroCode();
	}

}
