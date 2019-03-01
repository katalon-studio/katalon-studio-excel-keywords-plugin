package com.katalon.plugin.keyword.excel
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.DataFormatter
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.ss.usermodel.WorkbookFactory

import com.kms.katalon.core.annotation.Keyword

class ExcelReadKeywords {

	@Keyword
	def readRow(String filePath, int sheetIndex, int rowIndex) {

		InputStream inputStream

		try {

			List<String> rowData = new ArrayList<String>()

			inputStream = new FileInputStream(filePath)
			Workbook workbook = WorkbookFactory.create(inputStream)

			Sheet sheet = workbook.getSheetAt(sheetIndex)

			DataFormatter dataFormatter = new DataFormatter()

			Row row = sheet.getRow(rowIndex)

			if (row == null) {
				return null
			}

			for (Cell cell : row) {
				String cellValue = dataFormatter.formatCellValue(cell)
				rowData.add(cellValue)
			}

			return rowData
		} finally {

			inputStream.close()
		}
	}
}