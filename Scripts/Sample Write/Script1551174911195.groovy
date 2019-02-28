import static com.kms.katalon.core.checkpoint.CheckpointFactory.findCheckpoint
import static com.kms.katalon.core.testcase.TestCaseFactory.findTestCase
import static com.kms.katalon.core.testdata.TestDataFactory.findTestData
import static com.kms.katalon.core.testobject.ObjectRepository.findTestObject
import com.kms.katalon.core.checkpoint.Checkpoint as Checkpoint
import com.kms.katalon.core.configuration.RunConfiguration as RunConfiguration
import com.kms.katalon.core.cucumber.keyword.CucumberBuiltinKeywords as CucumberKW
import com.kms.katalon.core.mobile.keyword.MobileBuiltInKeywords as Mobile
import com.kms.katalon.core.model.FailureHandling as FailureHandling
import com.kms.katalon.core.testcase.TestCase as TestCase
import com.kms.katalon.core.testdata.TestData as TestData
import com.kms.katalon.core.testobject.TestObject as TestObject
import com.kms.katalon.core.webservice.keyword.WSBuiltInKeywords as WS
import com.kms.katalon.core.webui.keyword.WebUiBuiltInKeywords as WebUI
import internal.GlobalVariable as GlobalVariable

String filePath = (((RunConfiguration.getProjectDir() + File.separator) + 'output') + File.separator) + 'excel.xlsx'

String sheetName = 'My First Sheet'

Object[][] data = [['Datatype', 'Example'], ['integer', 12345], ['float', 12345.12345], ['String', 'This is a string']
    , ['boolean', true], ['date', new Date()]]

CustomKeywords.'com.katalon.plugin.keyword.excel.ExcelKeywords.writeToNewFile'(filePath, sheetName, data)

String newSheetName = 'My Second Sheet'

CustomKeywords.'com.katalon.plugin.keyword.excel.ExcelKeywords.writeToNewSheet'(filePath, newSheetName, data)