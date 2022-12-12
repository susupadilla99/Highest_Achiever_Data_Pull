import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Comparator;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

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
			
			// Step 1: Setup
			HSSFSheet sheet = readExcel1("LCLARKSO2022.xls", 0);
			int idx = 1;
			HSSFRow row = sheet.getRow(idx);
			String unitCode = sheet.getRow(idx).getCell(UNITCODE).getStringCellValue();
			
			// Step 1: Generate list of all FirstInSubject
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
			
			// Step 1: Write FirstInSubject to output Excel Workbook
			HSSFWorkbook outputBook = writeWorkbook(highestAchievers, mark);
			
			
			// Step 2: Setup
			HSSFWorkbook db = readExcel2("HighAchieverDatabase.xls");
			idx = 0;
			ArrayList<ArrayList<String>> commendations = new ArrayList<ArrayList<String>>();
			
			// Step 2: Find all this-session commendations
			highestAchievers.sort( new StudentEntryComparator() );
			while (idx<highestAchievers.size()) {
				if (idx<highestAchievers.size() && highestAchievers.get(idx).get(0).equals(highestAchievers.get(idx+1).get(0))) {
					String id = highestAchievers.get(idx).get(STUDENTID);
					ArrayList<String> commd = new ArrayList<String>();
					commd.add(highestAchievers.get(idx).get(STUDENTID));
					commd.add(highestAchievers.get(idx).get(LASTNAME));
					commd.add(highestAchievers.get(idx).get(FIRSTNAME));
					String note = highestAchievers.get(idx).get(UNITCODE);
					idx++;
					while( idx< highestAchievers.size() && id.equals(highestAchievers.get(idx).get(STUDENTID)) ) {
						note += " - " + highestAchievers.get(idx).get(UNITCODE);
						idx++;
					}
					commd.add(note);
					commendations.add( commd );
				} else {
					idx++;
				}
			}
			
			System.out.println("----------------------");
			
			// Step 2: Filter out previously commended
			ArrayList<String> prevCommd = new ArrayList<String>();
			sheet = db.getSheetAt(1); // get sheet of all historic commended
			row = sheet.getRow(1);
			int prevCommdIdx = 1;
			// generate list of prev commended
			while(row != null) {
				prevCommd.add( ""+ (int) row.getCell(STUDENTID).getNumericCellValue() );
				prevCommdIdx++;
				row = sheet.getRow(prevCommdIdx);
			}
			
			System.out.println(prevCommd);
			
			printList(commendations);
			
			// filter our prev commended
			for(int i=0; i<commendations.size(); i++) {
				if ( contain(prevCommd,commendations.get(i).get(STUDENTID)) ) {
					System.out.println(commendations.get(i).get(FIRSTNAME) + " " + commendations.get(i).get(LASTNAME)+" removed.");
					commendations.remove(i);
				}
			}
			
			printList(commendations);
			
			
			writeExcel(outputBook);
			
		} catch (IOException e) {
			e.printStackTrace();
		}
		
	}
	
	public static HSSFSheet readExcel1(String path, int sheetIdx) throws IOException {
		FileInputStream inputStream = new FileInputStream(path);
		HSSFWorkbook wb = new HSSFWorkbook(inputStream);
		HSSFSheet sheet = wb.getSheetAt(sheetIdx);
		return sheet;
	}
	
	public static HSSFWorkbook readExcel2(String path) throws IOException {
		FileInputStream inputStream = new FileInputStream(path);
		HSSFWorkbook wb = new HSSFWorkbook(inputStream);
		return wb;
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
	
	public static boolean contain(ArrayList<String> list, String target) {
		for (String entry: list) {
			if (entry.equals(target)) return true;
		}
		return false;
	}
	
	private static void printList(ArrayList<ArrayList<String>> list) {
		System.out.println("-----------------------");
		for (ArrayList<String> entry: list) {
			System.out.println(entry);
		}
		System.out.println("-----------------------");
	}
	
}
