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

class ExcelReadKeywords {

	@Keyword
	def readRow(String filePath, int sheetIndex, int rowIndex) {

		List<String> rowData = new ArrayList<String>();

		File file = new File(filePath)
		Workbook workbook = WorkbookFactory.create(file);

		Sheet sheet = workbook.getSheetAt(sheetIndex);

		DataFormatter dataFormatter = new DataFormatter();

		Row row = sheet.getRow(rowIndex)

		if (row == null) {
			return null
		}

		for (Cell cell : row) {
			String cellValue = dataFormatter.formatCellValue(cell);
			rowData.add(cellValue)
		}

		return rowData;
	}
}