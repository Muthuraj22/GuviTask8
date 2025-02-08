package Task8;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class FileReadWriteTask {

	public static void main(String[] args) {
		
		//Defining file path
		String excelFilePath = "NewExcalWorkbook.xlsx";
		
		//create new workbook
		Workbook w = new XSSFWorkbook();
		
		//New Workbook
		Sheet Sheet1 = w.createSheet("DataSheet");
		
		//Writing data in sheet
		String[][] data = {
				{"Name", "Age", "Email"},
				{"John Doe", "30", "john@test.com"},
				{"Jane Doe", "28", "john@test.com"},
				{"Bob Smith", "35", "jacky@example.com"},
				{"Swapnil", "37", "swapnil@example.com"},
		};
		
		//Populate rows and cells
		int rowCount = 0;
		for(String[] rowData : data) {
			Row r = Sheet1.createRow(rowCount++);
			int colCount = 0;
			
			for(String cellData : rowData) {
				Cell cell = r.createCell(colCount++);
				cell.setCellValue(cellData);
			}
		}
		
		
		
		// Workbook to file
		try(FileOutputStream outputStream = new FileOutputStream(excelFilePath)){
			w.write(outputStream);
			w.close();
			System.out.println("Excel file created successfully");
		} catch (Exception e) {
			e.printStackTrace();
		}
		
		// Reading data from excel file
		try(FileInputStream inputStream = new FileInputStream(excelFilePath)){
			Workbook work = new XSSFWorkbook(inputStream);
			Sheet sheet = work.getSheetAt(0);
			for(Row row: sheet) {
				for(Cell cell : row) {
					System.out.println(cell.getStringCellValue()+ " ");
				}
			}
			System.out.println();
			work.close();
		}catch(IOException e) {
			e.printStackTrace();
		
	} 
	}
}
