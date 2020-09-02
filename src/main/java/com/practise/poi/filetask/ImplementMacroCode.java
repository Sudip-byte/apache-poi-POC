package com.practise.poi.filetask;

import org.springframework.stereotype.Component;
import com.aspose.cells.*;
import com.practise.poi.utility.Utils;

@Component
public class ImplementMacroCode {

	public void addMacroCode() throws Exception {
		// The path to the documents directory.
		//String dataDir = Utils.getSharedDataDir(ImplementMacroCode.class) + "BMO Project\\";

		// Load your source Excel file.
		Workbook wb = new Workbook("Demo.xlsm");
		
		
		Worksheet sheet = wb.getWorksheets().get("Pivot");
		
		
		// Add VBA Module
		int idx = wb.getVbaProject().getModules().add(sheet);

		// Access the VBA Module, set its name and codes
		VbaModule module = wb.getVbaProject().getModules().get(idx);
		module.setName("Macro1");

		module.setCodes("Sub Macro_Demo()\r\n" + 
				"'\r\n" + 
				"' Delete Macro\r\n" + 
				"'\r\n" + 
				"\r\n" + 
				"'\r\n" + 
				"    Range(\"A1\").Select\r\n" + 
				"    Range(Selection, Selection.End(xlToRight)).Select\r\n" + 
				"    Range(Selection, Selection.End(xlDown)).Select\r\n" + 
				"    pivotWS = ActiveSheet.Name\r\n" + 
				"    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _\r\n" + 
				"        \"Demo!R1C1:R35C7\", Version:=6).CreatePivotTable TableDestination:= _\r\n" + 
				"        pivotWS & \"!R3C1\", TableName:=\"PivotTable1\", DefaultVersion:=6\r\n" + 
				"    Sheets(pivotWS).Select\r\n" + 
				"    Cells(3, 1).Select\r\n" + 
				"    With ActiveSheet.PivotTables(\"PivotTable1\")\r\n" + 
				"        .ColumnGrand = True\r\n" + 
				"        .HasAutoFormat = True\r\n" + 
				"        .DisplayErrorString = False\r\n" + 
				"        .DisplayNullString = True\r\n" + 
				"        .EnableDrilldown = True\r\n" + 
				"        .ErrorString = \"\"\r\n" + 
				"        .MergeLabels = False\r\n" + 
				"        .NullString = \"\"\r\n" + 
				"        .PageFieldOrder = 2\r\n" + 
				"        .PageFieldWrapCount = 0\r\n" + 
				"        .PreserveFormatting = True\r\n" + 
				"        .RowGrand = True\r\n" + 
				"        .SaveData = True\r\n" + 
				"        .PrintTitles = False\r\n" + 
				"        .RepeatItemsOnEachPrintedPage = True\r\n" + 
				"        .TotalsAnnotation = False\r\n" + 
				"        .CompactRowIndent = 1\r\n" + 
				"        .InGridDropZones = False\r\n" + 
				"        .DisplayFieldCaptions = True\r\n" + 
				"        .DisplayMemberPropertyTooltips = False\r\n" + 
				"        .DisplayContextTooltips = True\r\n" + 
				"        .ShowDrillIndicators = True\r\n" + 
				"        .PrintDrillIndicators = False\r\n" + 
				"        .AllowMultipleFilters = False\r\n" + 
				"        .SortUsingCustomLists = True\r\n" + 
				"        .FieldListSortAscending = False\r\n" + 
				"        .ShowValuesRow = False\r\n" + 
				"        .CalculatedMembersInFilters = False\r\n" + 
				"        .RowAxisLayout xlCompactRow\r\n" + 
				"    End With\r\n" + 
				"    With ActiveSheet.PivotTables(\"PivotTable1\").PivotCache\r\n" + 
				"        .RefreshOnFileOpen = False\r\n" + 
				"        .MissingItemsLimit = xlMissingItemsDefault\r\n" + 
				"    End With\r\n" + 
				"    ActiveSheet.PivotTables(\"PivotTable1\").RepeatAllLabels xlRepeatLabels\r\n" + 
				"    ActiveSheet.PivotTables(\"PivotTable1\").AddDataField ActiveSheet.PivotTables( _\r\n" + 
				"        \"PivotTable1\").PivotFields(\"WTD_FTE HC\"), \"Sum of WTD_FTE HC\", xlSum\r\n" + 
				"    ActiveSheet.PivotTables(\"PivotTable1\").AddDataField ActiveSheet.PivotTables( _\r\n" + 
				"        \"PivotTable1\").PivotFields(\"WTD_Avail_Hrs\"), \"Sum of WTD_Avail_Hrs\", xlSum\r\n" + 
				"    ActiveSheet.PivotTables(\"PivotTable1\").AddDataField ActiveSheet.PivotTables( _\r\n" + 
				"        \"PivotTable1\").PivotFields(\"WTD_Productive_Hrs\"), \"Sum of WTD_Productive_Hrs\", _\r\n" + 
				"        xlSum\r\n" + 
				"    With ActiveSheet.PivotTables(\"PivotTable1\").PivotFields(\"Project_Name\")\r\n" + 
				"        .Orientation = xlRowField\r\n" + 
				"        .Position = 1\r\n" + 
				"    End With\r\n" + 
				"    With ActiveSheet.PivotTables(\"PivotTable1\").PivotFields(\"Project_DB_ID\")\r\n" + 
				"        .Orientation = xlRowField\r\n" + 
				"        .Position = 2\r\n" + 
				"    End With\r\n" + 
				"    With ActiveSheet.PivotTables(\"PivotTable1\").PivotFields(\"Account SECTOR\")\r\n" + 
				"        .Orientation = xlPageField\r\n" + 
				"        .Position = 1\r\n" + 
				"    End With\r\n" + 
				"    With ActiveSheet.PivotTables(\"PivotTable1\").PivotFields(\"Client_Name\")\r\n" + 
				"        .Orientation = xlPageField\r\n" + 
				"        .Position = 1\r\n" + 
				"    End With\r\n" + 
				"    ActiveSheet.PivotTables(\"PivotTable1\").PivotFields(\"Account SECTOR\"). _\r\n" + 
				"        CurrentPage = \"(All)\"\r\n" + 
				"    ActiveSheet.PivotTables(\"PivotTable1\").PivotFields(\"Account SECTOR\"). _\r\n" + 
				"        EnableMultiplePageItems = True\r\n" + 
				"    ActiveSheet.PivotTables(\"PivotTable1\").PivotFields(\"Client_Name\").CurrentPage _\r\n" + 
				"        = \"(All)\"\r\n" + 
				"    ActiveSheet.PivotTables(\"PivotTable1\").PivotFields(\"Client_Name\"). _\r\n" + 
				"        EnableMultiplePageItems = True\r\n" + 
				"    Range(\"A5\").Select\r\n" + 
				"    ActiveSheet.PivotTables(\"PivotTable1\").PivotFields(\"Project_Name\").Subtotals = _\r\n" + 
				"        Array(False, False, False, False, False, False, False, False, False, False, False, False)\r\n" + 
				"    ActiveSheet.PivotTables(\"PivotTable1\").PivotFields(\"Project_Name\").LayoutForm _\r\n" + 
				"        = xlTabular\r\n" + 
				"End Sub\r\n" + 
				"Private Sub Worksheet_Activate()\r\n" + 
				"    Macro_Demo\r\n" + 
				"End Sub\r\n" + 
				"\r\n" + 
				"Private Sub Worksheet_Deactivate()\r\n" + 
				"    ThisWorkbook.Worksheets(\"Pivot\").Cells.ClearContents\r\n" + 
				"End Sub\r\n" + 
				"");

		// Save the workbook
		wb.save("Output_Demo.xlsm", SaveFormat.XLSM);
		
	}

}
