import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class Main {
	
	public static final int STUDENTID = 0;
	public static final int LASTNAME = 1;
	public static final int FIRSTNAME = 2;
	public static final int UNITCODE = 3;
	public static final int UNITNAME = 4;
	public static final int MARK = 5;

	
	public static void main(String[] args) {
		try {
			// Variables setup
			ArrayList<ArrayList<String>> highestAchievers = new ArrayList<ArrayList<String>>();
			ArrayList<ArrayList<String>> temp = new ArrayList<ArrayList<String>>();
			ArrayList<Integer> mark = new ArrayList<Integer>();
			
			HSSFSheet sheet = readExcel("LCLARKSO2022.xls", 0);
			int idx = 1;
			HSSFRow row = sheet.getRow(idx);
			String unitCode = sheet.getRow(idx).getCell(UNITCODE).getStringCellValue();
			
			while (row != null) {
				System.out.println(idx);
				if (sheet.getRow(idx).getCell(UNITCODE).getStringCellValue().equals(unitCode)) {
					temp.add(new ArrayList<String>());
					temp.get(temp.size()-1).add( sheet.getRow(idx).getCell(STUDENTID).getStringCellValue() );
					temp.get(temp.size()-1).add( sheet.getRow(idx).getCell(LASTNAME).getStringCellValue() );
					
					if (sheet.getRow(idx).getCell(FIRSTNAME) != null) {
						temp.get(temp.size()-1).add( sheet.getRow(idx).getCell(FIRSTNAME).getStringCellValue() );
					} else {
						temp.get(temp.size()-1).add( null );
					}
					
					temp.get(temp.size()-1).add( sheet.getRow(idx).getCell(UNITCODE).getStringCellValue() );
					temp.get(temp.size()-1).add( sheet.getRow(idx).getCell(UNITNAME).getStringCellValue() );
					mark.add( (int) Math.round(sheet.getRow(idx).getCell(MARK).getNumericCellValue()) );
					idx++;
					row = sheet.getRow(idx);
				} else {
					addHighestAchievers(highestAchievers,temp, mark);
					temp.clear();
					mark.clear();
					unitCode = row.getCell(UNITCODE).getStringCellValue();
				}
			}
			addHighestAchievers(highestAchievers, temp, mark);
			
			//write to output Excel file
			HSSFWorkbook outputBook = writeWorkbook(highestAchievers, mark);
			writeExcel(outputBook);
			
		} catch (IOException e) {
			e.printStackTrace();
		}
		
	}
	
	public static HSSFSheet readExcel(String path, int sheetIdx) throws IOException {
		FileInputStream inputStream = new FileInputStream(path);
		
		HSSFWorkbook wb = new HSSFWorkbook(inputStream);
		
		HSSFSheet sheet = wb.getSheetAt(sheetIdx);
		
		return sheet;
	}
	
	public static void writeExcel(HSSFWorkbook outBook) throws IOException {
		FileOutputStream outputStream = new FileOutputStream("JavaBooks.xls");
        outBook.write(outputStream);
	}
	
	public static HSSFWorkbook writeWorkbook(ArrayList<ArrayList<String>> list, ArrayList<Integer> grade) {
		// Setup workbook
		HSSFWorkbook retVal = new HSSFWorkbook();
		HSSFSheet sheet = retVal.createSheet("First in Subject");
		// Create Title row
		HSSFRow title = sheet.createRow(0);
		title.createCell(0).setCellValue("Student ID");
		title.createCell(1).setCellValue("First Name");
		title.createCell(2).setCellValue("Last Name");
		title.createCell(3).setCellValue("Unit Code");
		title.createCell(4).setCellValue("Unit Name");
		// Fill all data rows
		int idx=1;
		for (ArrayList<String> studentEntry: list) {
			HSSFRow row = sheet.createRow(idx);
			row.createCell(0).setCellValue(studentEntry.get(STUDENTID));
			row.createCell(1).setCellValue(studentEntry.get(FIRSTNAME));
			row.createCell(2).setCellValue(studentEntry.get(LASTNAME));
			row.createCell(3).setCellValue(studentEntry.get(UNITCODE));
			row.createCell(4).setCellValue(studentEntry.get(UNITNAME));
			idx++;
		}
		return retVal;
	}
	
	public static void addHighestAchievers(ArrayList<ArrayList<String>> target, ArrayList<ArrayList<String>> temp, ArrayList<Integer> grade) {
		int topScore = grade.get(0);
		if( temp.size() < 10 ) return;
		if (topScore < 85) return; 
		// Add all high-achievers
		for(int i=0; i<temp.size(); i++) {
			// Add 1 high achiever data
			if (grade.get(i) == topScore) {
				target.add(temp.get(i));
			}
		}
	}
}
