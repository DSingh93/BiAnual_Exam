package utils;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class dataRead {

	static XSSFWorkbook workbook;
	static XSSFSheet sheet;

	public dataRead(String excelpath, String sheetname){

		try {
			workbook = new XSSFWorkbook(excelpath);
			sheet = workbook.getSheet(sheetname);

		}catch (Exception exp) {
			System.out.println(exp.getMessage());
			System.out.println(exp.getCause());
			exp.printStackTrace();
		}
	}


	public static void getCellData(int rowNum, int colNum){
		try {

			DataFormatter formatter = new DataFormatter();			
			Object value = formatter.formatCellValue(sheet.getRow(rowNum).getCell(colNum));

			System.out.println(value);

		} catch (Exception exp) {
			System.out.println(exp.getMessage());
			System.out.println(exp.getCause());
			exp.printStackTrace();
		}
	}


	public void getRowCount() {

		int rowCount = sheet.getPhysicalNumberOfRows();
		System.out.println("NO of Rows : "+rowCount);

	}

}
