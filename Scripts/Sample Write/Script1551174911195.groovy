import org.junit.Assert

import com.kms.katalon.core.configuration.RunConfiguration as RunConfiguration

String filePath = (((RunConfiguration.getProjectDir() + File.separator) + 'output') + File.separator) + 'excel.xlsx'

String firstSheetName = 'My First Sheet'

Object[][] firstSheetData = [['Datatype', 'Example'], ['integer', 12345], ['float', 12345.12345], ['String', 'This is a string']
    , ['boolean', true], ['date', new Date()]]

CustomKeywords.'com.katalon.plugin.keyword.excel.ExcelKeywords.createFileAndAddSheet'(filePath, firstSheetName, firstSheetData)

String secondSheetName = 'My Second Sheet'

Object[][] secondSheetData = [['Datatype', 'Example', 'Another example'], ['integer', 12345, 67890], ['float', 12345.12345, 67890.67890]
	, ['String', 'This is a string', 'This is another string'], ['boolean', true, false], ['date', new Date(), new Date()]
	, ['Datatype', 'Example', 'Another example'], ['integer', 12345, 67890], ['float', 12345.12345, 67890.67890]
	, ['String', 'This is a string', 'This is another string'], ['boolean', true, false], ['date', new Date(), new Date()]]

CustomKeywords.'com.katalon.plugin.keyword.excel.ExcelKeywords.openFileAndAddSheet'(filePath, secondSheetName, secondSheetData)

def actualRow = CustomKeywords.'com.katalon.plugin.keyword.excel.ExcelKeywords.readRow'(filePath, 1, 1)

Object[] expectedRow = secondSheetData[1]
for (int i = 0; i < expectedRow.length; i++) {
	expectedRow[i] = expectedRow[i].toString()
}

Assert.assertArrayEquals(expectedRow, actualRow)