/*
Copyright 2009 Yellow Snow 

Licensed under the Apache License, Version 2.0 (the "License"); you may not 
use this file except in compliance with the License. You may obtain a copy of 
the License at 

	http://www.apache.org/licenses/LICENSE-2.0 

Unless required by applicable law or agreed to in writing, software 
distributed under the License is distributed on an "AS IS" BASIS, WITHOUT 
WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied. See the 
License for the specific language governing permissions and limitations under 
the License. 
*/
import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.hssf.usermodel.HSSFSheet
import org.apache.poi.hssf.usermodel.HSSFCell
import org.apache.poi.ss.usermodel.WorkbookFactory
import org.apache.poi.hssf.util.CellReference


import java.io.IOException
import java.io.OutputStream
import java.math.BigDecimal
import java.util.Map
import java.text.DateFormat

public class SimpleXlsSlurper implements Iterable {
	static {
	}

	private def workbook
	private selection
	private sheets
	
	SimpleXlsSlurper(InputStream inputStream) {
		if (inputStream) {
			workbook = WorkbookFactory.create(inputStream)
			populate()
		} else {
			throw new IllegalArgumentException("Input Stream")
		}
	}
	SimpleXlsSlurper(String fileName) {
		if (fileName) {
			workbook = WorkbookFactory.create(new FileInputStream(fileName))
			populate()
		} else {
			throw new IllegalArgumentException("File Name")
		}
	}
	protected populate(){
		sheets = new ArrayList<HSSFSheet>(workbook.numberOfSheets)
		(0..(workbook.numberOfSheets-1)).each{sheets << workbook.getSheetAt(it)}
	} 
	Iterator iterator(){
		if (selection instanceof HSSFWorkbook) {
			return sheets.collect{it.sheetName}.iterator()
		} else if (selection instanceof HSSFSheet) {
			return selection.iterator()
		}
	}
	def row(Integer index){
		def sheet,row
		if (selection instanceof HSSFSheet){
			sheet = selection
		} else {
			if (!sheets.empty) {
				sheet = sheets[0]
			}
		}
		if (sheet) {
			selection = sheet.getRow(index)
		}
		return this
	}
	def sheet(Integer index){
		selection = sheets[index]
		return this
	}
	def sheet(String name){
		selection = workbook.getSheet(name)
		return this
	}
	def valueAt(String ref){
		def cellRef = new CellReference(ref)
		def rowNum = cellRef.row
		def cellNum = cellRef.col
		def aSheet,row,cell
		if (cellRef.sheetName) {
			sheet(cellRef.getSheetName())
			aSheet = selection
		} else if (selection instanceof HSSFSheet){
			aSheet = selection
		} else {
			if (!sheets.empty) {
				aSheet = sheets[0]
			}
		}
		if (aSheet){
			if (aSheet.physicalNumberOfRows > rowNum){
				row  = aSheet.getRow(rowNum)
			}
			cell = row?.getCell(cellNum)
		}
		return getCellValue(cell,cellRef)
	}
	private getCellValue(cell, cellRef){
		if (cell) {
			selection = null
			def result
			switch(cell.cellType) {
				case HSSFCell.CELL_TYPE_STRING:
					result = cell.richStringCellValue.string;
					break;
				case HSSFCell.CELL_TYPE_NUMERIC:
					if (DateUtil.isCellDateFormatted(cell)) {
					  result = cell.dateCellValue
					} else {
					  result = cell.numericCellValue
					}
					break;
				case HSSFCell.CELL_TYPE_BOOLEAN:
					result = cell.booleanCellValue
					break;
				case HSSFCell.CELL_TYPE_FORMULA:
					result = cell.cellFormula
					break;
				default:
					result = null
			}
			return result
		} else {
			throw new IllegalArgumentException("Bad cell reference [${cellRef}]")
		}
	}
}