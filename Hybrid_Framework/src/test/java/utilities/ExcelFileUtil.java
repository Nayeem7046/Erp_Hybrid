package utilities;

import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelFileUtil {
	Workbook wb;
	//constructor for reading Excel path
	public ExcelFileUtil(String Excelpath) throws Throwable
	{
		FileInputStream fi = new FileInputStream(Excelpath);
		wb = WorkbookFactory.create(fi);
	}
	// method for counting no. of rows in a sheet
	public int rowcount(String sheetName)
	{
		return wb.getSheet(sheetName).getLastRowNum();
	}
	// method for reading cell data
	public String getCellData(String sheetName, int row, int column)
	{
		String data = " ";
		if(wb.getSheet(sheetName).getRow(row).getCell(column).getCellType()==CellType.NUMERIC)
		{
			int celldata = (int) wb.getSheet(sheetName).getRow(row).getCell(column).getNumericCellValue();
			data =String.valueOf(celldata);
		}
		else
		{
			data =wb.getSheet(sheetName).getRow(row).getCell(column).getStringCellValue();
		}
		return data;
	}
	//method for write status into new workbook
	public void setCellData(String sheetname,int row,int column,String status,String WriteExcelPath)throws Throwable
	{
		//get sheet from wb
		Sheet ws =wb.getSheet(sheetname);
		//getrow from sheet
		Row rownum=ws.getRow(row);
		//create cell
		Cell cell = rownum.createCell(column);
		//write status
		cell.setCellValue(status);
		if(status.equalsIgnoreCase("Pass"))
		{
			CellStyle style=wb.createCellStyle();
			Font font=wb.createFont();
			font.setColor(IndexedColors.GREEN.getIndex());
			font.setBold(true);
			style.setFont(font);
			rownum.getCell(column).setCellStyle(style);
		}
		else if(status.equalsIgnoreCase("Fail"))
		{
			CellStyle style=wb.createCellStyle();
			Font font=wb.createFont();
			font.setColor(IndexedColors.RED.getIndex());
			font.setBold(true);
			style.setFont(font);
			rownum.getCell(column).setCellStyle(style);
		}
		else if(status.equalsIgnoreCase("Blocked"))
		{
			CellStyle style=wb.createCellStyle();
			Font font=wb.createFont();
			font.setColor(IndexedColors.BLUE.getIndex());
			font.setBold(true);
			style.setFont(font);
			rownum.getCell(column).setCellStyle(style);
		}
		FileOutputStream fo = new FileOutputStream(WriteExcelPath);
		wb.write(fo);
	}

	/*
	public static void main(String[]args) throws Throwable
	{
		ExcelFileUtil xl=new ExcelFileUtil("C:/Users/nayee/Desktop/sample_data.xlsx");
		//count no. of rows in emp sheet
		int rc=xl.rowcount("Emp");
		System.out.println(rc);
		for(int i=1;i<=rc;i++)
		{
			String fname=xl.getCellData("Emp",i, 0);
			String mname=xl.getCellData("Emp",i, 1);
			String lname=xl.getCellData("Emp",i, 2);
			String eid=xl.getCellData("Emp",i, 3);
			System.out.println(fname+"  "+mname+"  "+lname+"  "+eid);
			//xl.setCellData("Emp", i, 4,"Pass","C:/Users/nayee/Desktop/Results.xlsx");
			//xl.setCellData("Emp", i, 4,"Fail","C:/Users/nayee/Desktop/Results.xlsx");
			xl.setCellData("Emp", i, 4,"Blocked","C:/Users/nayee/Desktop/Results.xlsx");
		}
	}
	 */
}