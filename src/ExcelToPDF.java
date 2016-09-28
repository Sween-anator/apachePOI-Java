//Copyright Keagan Sweeney

import java.io.*;
import java.util.Iterator;
import java.util.Scanner;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

public class ExcelToPDF {
	public static void main(String[] args) throws IOException{
		System.out.println("This program will work only if items are in columns A, B and N.");
		System.out.println("Make sure that ALL files are in the same library.");
		System.out.println("Please enter excel sheet (Example- testexcel.xlsx): ");
		Scanner reader = new Scanner(System.in);
		String n = reader.next();
		
		File excel = new File(n);
		FileInputStream file = new FileInputStream(excel);
		XSSFWorkbook workbook = new XSSFWorkbook(file);
		XSSFSheet sheet = workbook.getSheetAt(0); 
		Iterator<Row> rowIterator = sheet.iterator();
		
		
		//Iterate through each row from first sheet
		for(int i = 0; i < sheet.getLastRowNum()+1; i ++){
			String value1;
			String value2 = "";
			String value3;
			
			String check1;
			String check2 = "";
			String against1;
			String against2 = "";
			
			Cell cell1 = sheet.getRow(i).getCell(0);
			Cell cell2 = sheet.getRow(i).getCell(1);
			Cell cell3 = sheet.getRow(i).getCell(13);
			
			check1 = cell1.toString();
			against1 = check1.substring(check1.length()-2,check1.length());
			if(cell2 != null){												
			check2 = cell2.toString();
			against2 = check2.substring(1, check2.length());
			}
			
			System.out.println(against2);
			
			if(against1.equals(".0")){
				value1 = cell1.toString();
				value1 = value1.substring(0,value1.length()-2);
			}else{
				value1 = cell1.toString();
			}
			
			if(against2.equals(".0")){
				value2 = cell2.toString();
				value2 = value2.substring(0,value2.length()-2);
			}else if(cell2 != null){
				value2 = cell2.toString();
			}
				value3 = cell3.toString();
			
			System.out.println(against1);
			System.out.println("PDF:"+value3);
			System.out.println("name:"+value1);
			if(value1.length() != 0){
			File oldFile = new File(value3);
			if(value2.length() != 0){
			File newFile = new File(value1 + "_" + value2);
			oldFile.renameTo(newFile);
			}else{
			File newFile = new File(value1);
			oldFile.renameTo(newFile);
			}
			}
		}
		System.out.println("Complete!");
	}
}
