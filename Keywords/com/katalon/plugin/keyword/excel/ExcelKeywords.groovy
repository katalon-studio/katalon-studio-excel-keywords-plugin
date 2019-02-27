package com.katalon.plugin.keyword.excel
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook

import com.kms.katalon.core.annotation.Keyword
import com.kms.katalon.core.util.KeywordUtil

class ExcelKeywords {
	
	private void setCellStyle(XSSFWorkbook workbook, Cell cell, short dataFormat) {
		CellStyle cellStyle = workbook.createCellStyle()
		cellStyle.setDataFormat(dataFormat)
		cell.setCellStyle(cellStyle)
	}

	@Keyword
	def write(String filePath, String sheetName, Object[][] rowsData) {

		XSSFWorkbook workbook = new XSSFWorkbook()

		KeywordUtil.logInfo("Creating sheet " + sheetName)
		XSSFSheet sheet = workbook.createSheet(sheetName)

		KeywordUtil.logInfo("Adding rows")
		int rowNum = 0

		for (Object[] rowData : rowsData) {
			Row row = sheet.createRow(rowNum++)
			int colNum = 0;
			for (Object field : rowData) {
				Cell cell = row.createCell(colNum++)
				if (field instanceof Boolean) {
					setCellStyle(workbook, cell, (short) 0x0) // 0x0, "General"
					cell.setCellValue((Boolean) field)
				} else if (field instanceof Date) {
					setCellStyle(workbook, cell, (short) 0x16) // 0x16, "m/d/yy h:mm"
					cell.setCellValue((Date) field)
				} else if (field instanceof Float || field instanceof Double) {
					setCellStyle(workbook, cell, (short) 0x2) // 0x2, "0.00"
					cell.setCellValue((Double) field)
				} else if (field instanceof String) {
					setCellStyle(workbook, cell, (short) 0x31) // 0x31, "text" - Alias for "@"
					cell.setCellValue((String) field)
				}
			}
		}

		KeywordUtil.logInfo("Writing to " + filePath)
		File file = new File(filePath)
		if (!file.exists()) {
			file.createNewFile()
		}
		FileOutputStream outputStream = new FileOutputStream(file)
		workbook.write(outputStream)
		outputStream.close()
	}
}