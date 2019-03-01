import org.junit.Assert as Assert
import com.kms.katalon.core.configuration.RunConfiguration as RunConfiguration
import com.kms.katalon.core.webui.keyword.WebUiBuiltInKeywords as WebUI
import com.kms.katalon.core.mobile.keyword.MobileBuiltInKeywords as Mobile
import com.kms.katalon.core.cucumber.keyword.CucumberBuiltinKeywords as CucumberKW
import com.kms.katalon.core.webservice.keyword.WSBuiltInKeywords as WS
import static com.kms.katalon.core.testobject.ObjectRepository.findTestObject
import static com.kms.katalon.core.testdata.TestDataFactory.findTestData
import static com.kms.katalon.core.testcase.TestCaseFactory.findTestCase
import static com.kms.katalon.core.checkpoint.CheckpointFactory.findCheckpoint
import com.kms.katalon.core.model.FailureHandling as FailureHandling
import com.kms.katalon.core.testcase.TestCase as TestCase
import com.kms.katalon.core.testdata.TestData as TestData
import com.kms.katalon.core.testobject.TestObject as TestObject
import com.kms.katalon.core.checkpoint.Checkpoint as Checkpoint
import internal.GlobalVariable as GlobalVariable

String filePath = (((RunConfiguration.getProjectDir() + File.separator) + 'output') + File.separator) + 'excel.xlsx'

String firstSheetName = 'My First Sheet'

List<List> firstSheetData = [['Datatype', 'Example'], ['integer', 12345], ['float', 12345.12345], ['String', 'This is a string']
    , ['boolean', true], ['date', new Date()]]

CustomKeywords.'com.katalon.plugin.keyword.excel.ExcelWriteKeywords.createFileAndAddSheet'(filePath, firstSheetName, firstSheetData)

String secondSheetName = 'My Second Sheet'

List<List> secondSheetData = [['Datatype', 'Example', 'Another example'], ['integer', 12345, 67890], ['float', 12345.12345
        , 67890.67890], ['String', 'This is a string', 'This is another string'], ['boolean', true, false], ['date', new Date()
        , new Date()], ['Datatype', 'Example', 'Another example'], ['integer', 12345, 67890], ['float', 12345.12345, 67890.67890]
    , ['String', 'This is a string', 'This is another string'], ['boolean', true, false], ['date', new Date(), new Date()]]

CustomKeywords.'com.katalon.plugin.keyword.excel.ExcelWriteKeywords.openFileAndAddSheet'(filePath, secondSheetName, secondSheetData)

def actualRow = CustomKeywords.'com.katalon.plugin.keyword.excel.ExcelReadKeywords.readRow'(filePath, 1, 1)



List<List> expectedRow = secondSheetData[1]

for (int i = 0; i < expectedRow.size(); i++) {
    (expectedRow[i]) = expectedRow[i].toString()
}

assert expectedRow == actualRow

