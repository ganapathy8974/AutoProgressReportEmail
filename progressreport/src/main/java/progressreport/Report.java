package progressreport;


import java.io.ByteArrayOutputStream;

import javax.activation.DataSource;
import javax.mail.util.ByteArrayDataSource;

import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Report {
	
	public DataSource getReport(String name,String roleNumber,String cls,int english,int tamil,int maths,int social,int science) throws Exception {
		
		
		//Excel Sheet and Workbook Creation. 
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet();	
		
		//new row creation and set cell data
		XSSFRow row =sheet.createRow(0);
		row.createCell(0).setCellValue("Name");
		sheet.addMergedRegion(new CellRangeAddress(0, 0, 1, 6));
		row.createCell(1).setCellValue(name);

		//new row creation and set cell data
		XSSFRow row1 =sheet.createRow(1);
		row1.createCell(0).setCellValue("Role Number");
		sheet.addMergedRegion(new CellRangeAddress(1, 1, 1, 6));
		row1.createCell(1).setCellValue(roleNumber);
				
		//new row creation and set cell data
		XSSFRow row2 =sheet.createRow(2);
		row2.createCell(0).setCellValue("Class");
		sheet.addMergedRegion(new CellRangeAddress(2, 2, 1, 6));
		row2.createCell(1).setCellValue(cls+" STD");
		
		//new row creation and set cell data
		XSSFRow row4 =sheet.createRow(4);
		row4.createCell(0).setCellValue("English");
		row4.createCell(1).setCellValue("Tamil");
		row4.createCell(2).setCellValue("Maths");
		row4.createCell(3).setCellValue("Social Science");
		row4.createCell(4).setCellValue("Science");
		row4.createCell(5).setCellValue("Totel");
		row4.createCell(6).setCellValue("Persentage");
		
		//new row creation and set cell data
		XSSFRow row5 =sheet.createRow(5);
		row5.createCell(0).setCellValue(english);
		row5.createCell(1).setCellValue(tamil);
		row5.createCell(2).setCellValue(maths);
		row5.createCell(3).setCellValue(social);
		row5.createCell(4).setCellValue(science);
		row5.createCell(5).setCellFormula("SUM(A6:E6)");
		row5.createCell(6).setCellFormula("(F6/500)*100");		
		
		//set column size auto here using for loop to implement all of columns 
		for(int i=0;i<row4.getPhysicalNumberOfCells();i++) {
			sheet.autoSizeColumn(i);
		}	
		
		//Store the Excel sheet in a byte array
		ByteArrayOutputStream byteReport = new ByteArrayOutputStream();
		workbook.write(byteReport);
		workbook.close();
		
		//Finally store the Byte Array into a Data Source and return it
		DataSource report = new ByteArrayDataSource(byteReport.toByteArray(), "application/vnd.ms-excel");		
		return report;
	}

}
