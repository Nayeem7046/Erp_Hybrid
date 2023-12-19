package driverFactory;

import org.openqa.selenium.WebDriver;
import org.testng.annotations.Test;

import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.ExtentTest;
import com.relevantcodes.extentreports.LogStatus;

import CommonFunctions.FunctionLibrary;
import utilities.ExcelFileUtil;
	
public class DriverScript {
	public WebDriver driver;
	String inputpath = "./FileInput/Controller.xlsx";
	String outputpath = "./FileOutput/Results.xlsx";
	String TestCases = "MasterTestClass";
	ExtentReports report;
	ExtentTest logger;
	public void starTest() throws Throwable
	{
		String Module_Status = "";
		//create object for ExcelfileUtil class
		ExcelFileUtil xl = new ExcelFileUtil(inputpath);
		//iterate all rows in testcases sheet
		for(int i=1;i<=xl.rowcount(TestCases);i++)
		{
			//read cell
			String execution_Status = xl.getCellData(TestCases, i, 2);
			if(execution_Status.equalsIgnoreCase("Y"))
			{
				//store all Module into variable
				String TCModule =xl.getCellData(TestCases, i, 1);
				report = new ExtentReports("./target/Reports/"+TCModule+FunctionLibrary.generateDate()+".html");
				logger = report.startTest(TCModule);
				//iterate all rows in TCModule
				for(int j=1;j<=xl.rowcount(TCModule);j++)
				{
					//read all cells from TCModule
					String Description = xl.getCellData(TCModule, j, 0);
					String Object_Type = xl.getCellData(TCModule, j, 1);
					String Locator_Type = xl.getCellData(TCModule, j, 2);
					String Locator_Value = xl.getCellData(TCModule, j, 3);
					String Test_Data = xl.getCellData(TCModule, j, 4);
					try {
						if(Object_Type.equalsIgnoreCase("startBrowser"))
						{
							driver = FunctionLibrary.startBrowser();
							logger.log(LogStatus.INFO, Description);
						}
						if(Object_Type.equalsIgnoreCase("openUrl"))
						{
							FunctionLibrary.openUrl(driver);
							logger.log(LogStatus.INFO, Description);
							
						}
						if(Object_Type.equalsIgnoreCase("waitForElement"))
						{
							FunctionLibrary.WaitForElement(driver, Locator_Type, Locator_Value, Test_Data);
							logger.log(LogStatus.INFO, Description);
						}
						if(Object_Type.equalsIgnoreCase("typeAction"))
						{
							FunctionLibrary.typeAction(driver, Locator_Type, Locator_Value, Test_Data);
							logger.log(LogStatus.INFO, Description);
						}
						if(Object_Type.equalsIgnoreCase("clickAction"))
						{
							FunctionLibrary.clickAction(driver, Locator_Type, Locator_Value);
							logger.log(LogStatus.INFO, Description);
						}
						if(Object_Type.equalsIgnoreCase("validateTitle"))
						{
							FunctionLibrary.ValidateTitle(driver, Test_Data);
							logger.log(LogStatus.INFO, Description);
						}
						if(Object_Type.equalsIgnoreCase("closeBrowser"))
						{
							FunctionLibrary.closeBrowser(driver);
							logger.log(LogStatus.INFO, Description);
						}
						if(Object_Type.equalsIgnoreCase("mouseClick"))
						{
							FunctionLibrary.mouseClick(driver);
							logger.log(LogStatus.INFO, Description);
						}
						if(Object_Type.equalsIgnoreCase("categoryTable"))
						{
							FunctionLibrary.categoryTable(driver, Test_Data);
							logger.log(LogStatus.INFO, Description);
						}
						if(Object_Type.equalsIgnoreCase("dropDownAction"))
						{
							FunctionLibrary.dropDownAction(driver, Locator_Type, Locator_Value, Test_Data);
							logger.log(LogStatus.INFO, Description);
						}
						if(Object_Type.equalsIgnoreCase("capturestock"))
						{
							FunctionLibrary.capturestock(driver, Locator_Type, Locator_Value);
							logger.log(LogStatus.INFO, Description);
						}
						if(Object_Type.equalsIgnoreCase("stockTable"))
						{
							FunctionLibrary.stockTable(driver);
							logger.log(LogStatus.INFO, Description);
						}
						if(Object_Type.equalsIgnoreCase("capturesupplier"))
						{
							FunctionLibrary.capturesupplier(driver, Locator_Type, Locator_Value);
							logger.log(LogStatus.INFO, Description);
						}
						if(Object_Type.equalsIgnoreCase("supplierTable"))
						{
							FunctionLibrary.supplierTable(driver);
							logger.log(LogStatus.INFO, Description);
						}
						//write as pass into TCModule
						xl.setCellData(TCModule, j, 5, "Pass", outputpath);
						logger.log(LogStatus.INFO, Description);
						Module_Status="True";
					}catch(Exception e)
					{
						//write as Fail into TCModule
						xl.setCellData(TCModule, j, 5, "Fail", outputpath);
						logger.log(LogStatus.FAIL, Description);
						Module_Status="False";
						System.out.println(e.getMessage());
					}
					if(Module_Status.equalsIgnoreCase("True"))
					{
						//write as pass into Testcases
						xl.setCellData(TestCases, i, 3, "Pass", outputpath);
					}
					if(Module_Status.equalsIgnoreCase("False"))
					{
						//write as Fail into Testcases
						xl.setCellData(TestCases, i, 3, "Fail", outputpath);
					}
					report.endTest(logger);
					report.flush();
				}
			}
			else
			{
				//write as blocked into Test cases sheet for flag N
				xl.setCellData(TestCases, i, 3, "Blocked", outputpath);
			}
		}
	}
}
