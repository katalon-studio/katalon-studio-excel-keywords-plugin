package com.katalon.plugin.keyword.excel
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.DataFormat
import org.apache.poi.ss.usermodel.DataFormatter
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.ss.usermodel.WorkbookFactory
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook

import com.kms.katalon.core.annotation.Keyword
import com.kms.katalon.core.util.KeywordUtil

class ExcelWriteKeywords {

	@Keyword
	def createFileAndAddSheet(String filePath, String sheetName, List<List<Object>> rowsData) {

		XSSFWorkbook workbook = new XSSFWorkbook()

		addSheetAndWriteData(workbook, sheetName, rowsData)

		File file = new File(filePath)
		if (file.exists()) {
			file.delete()
		} else {
			KeywordUtil.logInfo('Creating a new file')
			file.createNewFile()
		}

		writeWorkbookToFile(workbook, file)
	}

	@Keyword
	def openFileAndAddSheet(String filePath, String sheetName, List<List<Object>> rowsData) {

		InputStream inputStream;

		try {

			KeywordUtil.logInfo('Opening file ' + filePath)

			inputStream = new FileInputStream(filePath)
			XSSFWorkbook workbook = WorkbookFactory.create(inputStream)

			addSheetAndWriteData(workbook, sheetName, rowsData)

			File file = new File(filePath)

			writeWorkbookToFile(workbook, file)
		} finally {

			inputStream.close()
		}
	}

	private writeWorkbookToFile(XSSFWorkbook workbook, File file) {
		KeywordUtil.logInfo('Writing to ' + file.getAbsolutePath())
		FileOutputStream outputStream;
		try {
			outputStream = new FileOutputStream(file)
			workbook.write(outputStream)
		} finally {
			outputStream.close()
		}
	}

	private void setCellStyle(XSSFWorkbook workbook, Cell cell, short dataFormat) {
		CellStyle cellStyle = workbook.createCellStyle()
		cellStyle.setDataFormat(dataFormat)
		cell.setCellStyle(cellStyle)
	}

	private void setCellStyle(XSSFWorkbook workbook, Cell cell, String dataFormat) {
		DataFormat format = workbook.createDataFormat();
		CellStyle cellStyle = workbook.createCellStyle()
		cellStyle.setDataFormat(format.getFormat(dataFormat))
		cell.setCellStyle(cellStyle)
	}

	private void addSheetAndWriteData(XSSFWorkbook workbook, String sheetName, List<List<Object>> rowsData) {
		KeywordUtil.logInfo('Creating sheet ' + sheetName)
		XSSFSheet sheet = workbook.createSheet(sheetName)

		KeywordUtil.logInfo('Adding rows')
		int rowNum = 0

		for (List<Object> rowData : rowsData) {
			addRow(workbook, sheet, rowNum, rowData)
			rowNum++
		}
	}

	private void addRow(XSSFWorkbook workbook, XSSFSheet sheet, int rowNum, List<Object> rowData) {
		Row row = sheet.createRow(rowNum)
		int colNum = 0;
		for (Object field : rowData) {
			addCell(workbook, row, colNum, field)
			colNum++
		}
	}

	private void addCell(XSSFWorkbook workbook, Row row, int colNum, Object field) {
		Cell cell = row.createCell(colNum)
		if (field instanceof Boolean) {
			setCellStyle(workbook, cell, (short) 0x0) // 0x0, 'General'
			cell.setCellValue((Boolean) field)
		} else if (field instanceof Integer || field instanceof Long) {
			setCellStyle(workbook, cell, '#')
			cell.setCellValue((Long) field)
		} else if (field instanceof Date) {
			setCellStyle(workbook, cell, (short) 0x16) // 0x16, 'm/d/yy h:mm'
			cell.setCellValue((Date) field)
		} else if (field instanceof Float || field instanceof Double || field instanceof BigDecimal) {
			setCellStyle(workbook, cell, (short) 0x4) // 0x4, '#,##0.00'
			cell.setCellValue((Double) field)
		} else if (field instanceof String) {
			setCellStyle(workbook, cell, (short) 0x31) // 0x31, 'text' - Alias for '@'
			cell.setCellValue((String) field)
		} else {
			setCellStyle(workbook, cell, (short) 0x0) // 0x0, 'General'
			cell.setCellValue(field.toString())
		}
	}
}