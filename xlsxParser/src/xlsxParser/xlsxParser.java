package xlsxParser;

import java.io.File;  
import java.io.FileInputStream;  
import java.io.IOException; 
import java.lang.NumberFormatException;
import org.apache.poi.xssf.usermodel.XSSFSheet; 
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;  
import org.apache.poi.ss.usermodel.FormulaEvaluator;  
import org.apache.poi.ss.usermodel.Row; 
import java.util.Iterator;


public class xlsxParser {

	public static boolean isPrime(long number) {
		   if (number <= 1) {
		       return false;
		   }
		   for (int i = 2; i <= Math.sqrt(number); i++) {
		       if (number % i == 0) {
		           return false;
		       }
		   }
		   return true;
	}
	
	public static void main(String[] args) {
		// check if argument is given
		if(args.length == 0) {
			System.out.println("Neuveden parametr");
			return;
		}
		
		// parse file
		FileInputStream f;
		try {
			// parse given file
			f = new FileInputStream(new File(args[0]));
			XSSFWorkbook wb =  new XSSFWorkbook(f);
			XSSFSheet sheet = wb.getSheetAt(0);
			// iterate over rows
			Iterator<Row> itr = sheet.iterator();    //iterating over excel file  
			while (itr.hasNext()) {  
				Row row = itr.next();  
	
				// get only second column values
				Cell cell = row.getCell(1);
				if(cell == null)
					continue;
					
				// get only string values
				if(cell.getCellType() != Cell.CELL_TYPE_STRING)
					continue;
				
				// get string in cell
				String cellValue = cell.getStringCellValue();

				long cellNumberValue = 0;
				try {
					// convert string to number
					cellNumberValue = Long.parseLong(cellValue);
				} catch(NumberFormatException e) {
					// ignore string if error
					continue;
				}
				
				// check if number is smaller than 0
				if(cellNumberValue < 0) {
					// ignore number that is smaller than 0
					continue;
				}
				
				// check if number is prime
				if(isPrime(cellNumberValue)) {
					// print prime number
					System.out.println(cellNumberValue);
				}		
			}  		
		} catch(IOException e) {
			System.out.println("Soubor nenalezen");
		}
	}	
}
