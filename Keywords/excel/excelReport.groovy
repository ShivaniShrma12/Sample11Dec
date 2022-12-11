package excel

import static com.kms.katalon.core.checkpoint.CheckpointFactory.findCheckpoint
import static com.kms.katalon.core.testcase.TestCaseFactory.findTestCase
import static com.kms.katalon.core.testdata.TestDataFactory.findTestData
import static com.kms.katalon.core.testobject.ObjectRepository.findTestObject
import static com.kms.katalon.core.testobject.ObjectRepository.findWindowsObject

import com.kms.katalon.core.annotation.Keyword
import com.kms.katalon.core.checkpoint.Checkpoint
import com.kms.katalon.core.cucumber.keyword.CucumberBuiltinKeywords as CucumberKW
import com.kms.katalon.core.mobile.keyword.MobileBuiltInKeywords as Mobile
import com.kms.katalon.core.model.FailureHandling
import com.kms.katalon.core.testcase.TestCase
import com.kms.katalon.core.testdata.TestData
import com.kms.katalon.core.testobject.TestObject
import com.kms.katalon.core.webservice.keyword.WSBuiltInKeywords as WS
import com.kms.katalon.core.webui.keyword.WebUiBuiltInKeywords as WebUI
import com.kms.katalon.core.windows.keyword.WindowsBuiltinKeywords as Windows
import com.kms.katalon.keyword.excel.ExcelKeywords
import com.kms.katalon.core.configuration.RunConfiguration
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.ss.usermodel.Cell

import internal.GlobalVariable

public class excelReport {
	@Keyword
	public void excelWriteReport(String strTest, int TestCaseNum) {
		String filePath = RunConfiguration.getProjectDir() + '/Data Files/resultExcel.xlsx'
		ExcelKeywords excel = new ExcelKeywords()
		Workbook workbookXLSX = ExcelKeywords.getWorkbook(filePath)
		Sheet sheet1 = workbookXLSX.getSheet("Result")
		Row row = sheet1.getRow(TestCaseNum)
		int ColNum = row.getLastCellNum()
		println "value of i is " + TestCaseNum
		println "Total number of columns : " + ColNum
		println "Value in strTest is : " + strTest

		Cell cell = null
		if (cell==null) 
			cell = row.createCell(ColNum)
		
		cell.setCellValue(strTest)
		workbookXLSX.write(filePath)
		excel.saveWorkbook(filePath, workbookXLSX)
	}
}
