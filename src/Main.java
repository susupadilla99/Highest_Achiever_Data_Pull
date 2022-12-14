import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

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
			ArrayList<ArrayList<String>> firstInSubject = new ArrayList<ArrayList<String>>();
			ArrayList<Integer> mark = new ArrayList<Integer>();
			ArrayList<ArrayList<String>> temp = new ArrayList<ArrayList<String>>();
			ArrayList<Integer> tempMark = new ArrayList<Integer>();
			
			// Step 1: Setup
			HSSFSheet sheet = readExcel1("GeneratedFile.xls", 0);
			int idx = 1;
			HSSFRow row = sheet.getRow(idx);
			String unitCode = sheet.getRow(idx).getCell(UNITCODE).getStringCellValue();
			
			// Step 1: Generate list of all FirstInSubject
			while (row != null) {
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
					tempMark.add( (int) Math.round(sheet.getRow(idx).getCell(MARK).getNumericCellValue()) );
					idx++;
					row = sheet.getRow(idx);
				} else {
					addFirstInSubject(firstInSubject, mark, temp, tempMark);
					temp.clear();
					tempMark.clear();
					unitCode = row.getCell(UNITCODE).getStringCellValue();
				}
			}
			addFirstInSubject(firstInSubject, mark, temp, tempMark);
			
			// Step 1: Write FirstInSubject to output Excel Workbook
			HSSFWorkbook outputBook = writeWorkbook(firstInSubject, mark);
			
			
			// Step 2: Setup
			HSSFWorkbook db = readExcel2("HighAchieverDatabase.xls");
			idx = 0;
			ArrayList<ArrayList<String>> commendations = new ArrayList<ArrayList<String>>();
			ArrayList<ArrayList<String>> fisCopy = (ArrayList<ArrayList<String>>) firstInSubject.clone();
			
			// Step 2: Add all historical first in subjects
			sheet = db.getSheetAt(0); // get sheet of all historic first in subject
			row = sheet.getRow(1);
			int prevFirstIdx = 1;
			while (row!=null) {
				ArrayList<String> tmp = new ArrayList<String>();
				temp.add(new ArrayList<String>());
				tmp.add( ""+(int)row.getCell(STUDENTID).getNumericCellValue() );
				if (sheet.getRow(idx).getCell(1) != null) {
					tmp.add( row.getCell(1).getStringCellValue() );
				} else {
					tmp.add( null );
				}
				tmp.add( row.getCell(2).getStringCellValue() );
				tmp.add( row.getCell(UNITCODE).getStringCellValue() );
				tmp.add( row.getCell(UNITNAME).getStringCellValue() );
				firstInSubject.add(tmp);
				prevFirstIdx++;
				row = sheet.getRow(prevFirstIdx);
			}
			
			// Step 2: Find all available commendations
			firstInSubject.sort( new StudentEntryComparator() );
			while (idx<firstInSubject.size()) {
				if (idx<firstInSubject.size() && firstInSubject.get(idx).get(0).equals(firstInSubject.get(idx+1).get(0))) {
					String id = firstInSubject.get(idx).get(STUDENTID);
					ArrayList<String> commd = new ArrayList<String>();
					commd.add(firstInSubject.get(idx).get(STUDENTID));
					commd.add(firstInSubject.get(idx).get(LASTNAME));
					commd.add(firstInSubject.get(idx).get(FIRSTNAME));
					String note = firstInSubject.get(idx).get(UNITCODE);
					idx++;
					while( idx< firstInSubject.size() && id.equals(firstInSubject.get(idx).get(STUDENTID)) ) {
						note += " - " + firstInSubject.get(idx).get(UNITCODE);
						idx++;
					}
					commd.add(note);
					commendations.add( commd );
				} else {
					idx++;
				}
			}
			
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
			
			// filter our prev commended
			for(int i=0; i<commendations.size(); i++) {
				if ( prevCommd.contains(commendations.get(i).get(STUDENTID)) ) {
					commendations.remove(i);
					i--;
				}
			}
			
			// Step 3: Write data to relevant excel workbooks
			sheet = db.getSheetAt(0);
			String session = sheet.getRow(prevFirstIdx-1).getCell(6).getStringCellValue();
			if (session.charAt(1) == '1') { // Session 1
				session = "S2" + session.substring(2);
			} else { // Session 2
				session = "S1 " + ( Integer.parseInt(session.substring(3))+1 );
			}
			
			
			// write commendations to output book
			sheet = outputBook.createSheet("Commendations");
			HSSFRow title = sheet.createRow(0);
			title.createCell(0).setCellValue("StudentID");
			title.createCell(1).setCellValue("First Name");
			title.createCell(2).setCellValue("Last Name");
			title.createCell(3).setCellValue("Notes");
			
			idx = 1;
			for (ArrayList<String> entry: commendations) {
				row = sheet.createRow(idx);
				row.createCell(0).setCellValue(entry.get(STUDENTID));
				row.createCell(1).setCellValue(entry.get(FIRSTNAME));
				row.createCell(2).setCellValue(entry.get(LASTNAME));
				row.createCell(3).setCellValue(entry.get(3));
				idx++;
			}
			
			// write first in subject to database book
			sheet = db.getSheetAt(0);
			for (int i=0; i<fisCopy.size(); i++) {
				row = sheet.createRow(prevFirstIdx);
				row.createCell(0).setCellValue( Integer.parseInt(fisCopy.get(i).get(STUDENTID)) );
				row.createCell(1).setCellValue(fisCopy.get(i).get(FIRSTNAME));
				row.createCell(2).setCellValue(fisCopy.get(i).get(LASTNAME));
				row.createCell(3).setCellValue(fisCopy.get(i).get(UNITCODE));
				row.createCell(4).setCellValue(fisCopy.get(i).get(UNITNAME));
				row.createCell(5).setCellValue(mark.get(i));
				row.createCell(6).setCellValue(session);
				prevFirstIdx++;
			}
			
			// write commendations to database book
			sheet = db.getSheetAt(1);
			for (ArrayList<String> entry: commendations) {
				row = sheet.createRow(prevCommdIdx);
				row.createCell(0).setCellValue( Integer.parseInt(entry.get(STUDENTID)) );
				row.createCell(1).setCellValue(entry.get(FIRSTNAME));
				row.createCell(2).setCellValue(entry.get(LASTNAME));
				row.createCell(3).setCellValue(session);
				row.createCell(4).setCellValue(entry.get(3));
				prevCommdIdx++;
			}
			
			FileOutputStream outputStream = new FileOutputStream("HighAchiever "+session+".xls");
	        outputBook.write(outputStream);
	        
	        FileOutputStream dbStream = new FileOutputStream("HighAchieverDatabase.xls");
	        db.write(dbStream);
			
			
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
	
	// Write the initial workbook
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
		title.createCell(5).setCellValue("Mark");
		// Fill all data rows
		int idx=1;
		for (ArrayList<String> studentEntry: list) {
			HSSFRow row = sheet.createRow(idx);
			row.createCell(0).setCellValue(studentEntry.get(STUDENTID));
			row.createCell(1).setCellValue(studentEntry.get(FIRSTNAME));
			row.createCell(2).setCellValue(studentEntry.get(LASTNAME));
			row.createCell(3).setCellValue(studentEntry.get(UNITCODE));
			row.createCell(4).setCellValue(studentEntry.get(UNITNAME));
			row.createCell(5).setCellValue(grade.get(idx-1));
			idx++;
		}
		return retVal;
	}
	
	// Add students who achieved first in subject to "target" list. Also populate a "mark" list with students's grades
	public static void addFirstInSubject(ArrayList<ArrayList<String>> target, ArrayList<Integer> mark, ArrayList<ArrayList<String>> temp, ArrayList<Integer> tempMark) {
		int topScore = tempMark.get(0);
		if( temp.size() < 10 ) return;
		if (topScore < 85) return; 
		// Add all high-achievers
		for(int i=0; i<temp.size(); i++) {
			// Add 1 high achiever data
			if (tempMark.get(i) == topScore) {
				target.add(temp.get(i));
				mark.add(tempMark.get(i));
			}
		}
	}
	
	// Debugging function
	private static void printList(ArrayList<ArrayList<String>> list) {
		System.out.println("-----------------------");
		for (ArrayList<String> entry: list) {
			System.out.println(entry);
		}
		System.out.println("-----------------------");
	}
	
}
