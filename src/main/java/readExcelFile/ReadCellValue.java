package readExcelFile;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import org.apache.poi.ss.usermodel.Cell;  
import org.apache.poi.ss.usermodel.*;  
import org.apache.poi.ss.usermodel.Sheet;  
import org.apache.poi.ss.usermodel.Workbook;  

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadCellValue {
	
	public static void main(String[] args)   {  
	//obiectul clasei	
	ReadCellValue rc=new ReadCellValue();    
	//citeste valoarea de pe randul 14 coloana 4
	String dataCitita=rc.ReadCellData(14, 4);   
	System.out.println(dataCitita);  
	}  
	
	
	//metoda care citeste efectiv din excel 
	public String ReadCellData(int rand, int coloana)  {  
	String value=null;   
	Workbook wb=null;     
	
	try  {  
		//incarcam fisierul in java  
		FileInputStream fisier=new FileInputStream("Financial Sample.xlsx");  
		//construim un obiect de tip XSSWorkbook
		wb=new XSSFWorkbook(fisier); 
		
	}  catch(Exception e)  {  
		e.printStackTrace();  
	} 
	//citim sheetul pe baza de index
	Sheet sheet= wb.getSheetAt(0);  
	//reprezinta un rand din sheetul respectiv
	Row row=sheet.getRow(rand);  
	//reprezinta celula din coloana
	Cell cell=row.getCell(coloana);  
		//verific data type-ul celulei (putem avea mai multe data types nu doar astea 2 de mai jos)
    	if (cell.getCellType() == CellType.STRING) {
        	value=cell.getStringCellValue();     
} 
    	else if(cell.getCellType() == CellType.NUMERIC){
        	value=String.valueOf(cell.getNumericCellValue());    

    		}
    	return value;               

    	}
}
