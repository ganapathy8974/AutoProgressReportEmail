import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReportEmail {

	public static void main(String[] args) throws IOException{
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet markReport = workbook.createSheet();
		markReport.addMergedRegion(new CellRangeAddress(0, 0, 1, 6));		
		XSSFRow rowhead = markReport.createRow(0);
		rowhead.createCell(0).setCellValue("Ganapathyramasubramanian");  
		rowhead.createCell(1).setCellValue("Ganu");		
		markReport.autoSizeColumn(0);
		FileOutputStream out = new FileOutputStream("C:\\Users\\Ganu\\Desktop\\ganutt.xlsx");
		workbook.write(out);
		out.close();
	}

}
