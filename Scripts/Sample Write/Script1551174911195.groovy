import com.katalon.plugin.keyword.excel.ExcelFormat
import com.kms.katalon.core.configuration.RunConfiguration as RunConfiguration

test(ExcelFormat.Excel97)

test(ExcelFormat.Excel2007)

def test(format) {
	
	String fileExtension = format.equals(ExcelFormat.Excel97) ? 'xls' : 'xlsx';
	
	String filePath = (((RunConfiguration.getProjectDir() + File.separator) + 'output') + File.separator) + 'excel.' + fileExtension
	
	String firstSheetName = 'My First Sheet'
	
	List<List<Object>> firstSheetData = [['Datatype', 'Example'], ['integer', 12345], ['float', 12345.12345], ['String', 'This is a string']
		, ['boolean', true], ['date', new Date()]]
	
	String secondSheetName = 'My Second Sheet'
	
	List<List<Object>> secondSheetData = [['Datatype', 'Example', 'Another example'], ['integer', 12345, 67890], ['float', 12345.12345, 67890.67890]
			, ['String', 'This is a string', 'This is another string'], ['boolean', true, false], ['date', new Date(), new Date()]
					, ['Datatype', 'Example', 'Another example'], ['integer', 12345, 67890], ['float', 12345.12345, 67890.67890]
							, ['String', 'This is a string', 'This is another string'], ['boolean', true, false], ['date', new Date(), new Date()]]
	
	CustomKeywords.'com.katalon.plugin.keyword.excel.ExcelWriteKeywords.createFileAndAddSheet'(format, filePath, firstSheetName, firstSheetData)
	
	CustomKeywords.'com.katalon.plugin.keyword.excel.ExcelWriteKeywords.openFileAndAddSheet'(filePath, secondSheetName, secondSheetData)
	
	def actualRow = CustomKeywords.'com.katalon.plugin.keyword.excel.ExcelReadKeywords.readRow'(filePath, 1, 1)
	
	List<Object> expectedRow = secondSheetData[1]
	for (int i = 0; i < expectedRow.size(); i++) {
		expectedRow[i] = expectedRow[i].toString()
	}
	
	assert expectedRow == actualRow
}