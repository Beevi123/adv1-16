package hardcore;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Properties;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelFile {

	public  String getDataFromPropertyFile(String key) throws IOException {
		FileInputStream fis=new FileInputStream("./src/test/resources/commondata.properties");
		Properties p=new Properties();
		p.load(fis);
		return p.getProperty(key);
	}
		
		public String getDataExcel(String sheetName,int rowNum,int cellNum) throws EncryptedDocumentException, IOException {
			FileInputStream fis=new FileInputStream("C:\\Users\\faizh\\OneDrive\\InsertData.xlsx");
			Workbook wb=WorkbookFactory.create(fis);
			return wb.getSheet(sheetName).getRow(rowNum).getCell(cellNum).toString();
		}
			
			
	   public String getDataFromExcelIdf(String sheetName,int rowNum,int cellNum) throws EncryptedDocumentException, IOException {
		   FileInputStream fis= new FileInputStream("C:\\Users\\faizh\\OneDrive\\InsertData.xlsx");
		   Workbook wb= WorkbookFactory.create(fis);
			Sheet sheet = wb.getSheet(sheetName);
			Cell cell = sheet.getRow(rowNum).getCell(cellNum);
			DataFormatter data= new DataFormatter();
			return data.formatCellValue(cell);
			
			
			
	   }
			
		
	
		

	}


