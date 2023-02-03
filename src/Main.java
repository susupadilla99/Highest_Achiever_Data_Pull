import java.io.BufferedOutputStream;
import java.io.DataOutputStream;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.PrintWriter;
import java.util.ArrayList;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class Main {
	
	static final int STUDENTID = 0;
	static final int FIRSTNAME = 1;
	static final int LASTNAME = 2;
	static final int UNITCODE = 3;
	static final int UNITNAME = 4;
	static final int MARK = 5;
	static final int COURSECODE = 6;
	static final int COURSEVERSION = 6;
	static final int COURSEATTEMPT = 6;
	
	static final String FISCOMMENT = "FSEHIGHACH";
	static final String COMMDCOMMENT = "FSECOMM";

	static ArrayList<ArrayList<String>> firstInSubject = new ArrayList<ArrayList<String>>();
	static ArrayList<ArrayList<String>> commendations = new ArrayList<ArrayList<String>>();
	static ArrayList<Integer> mark = new ArrayList<Integer>();
	static int year = 0;
	static String session = "";
	
	
	public static void main(String[] args) {
		try {
			// Step 1: Setup
			HSSFWorkbook gFile = getWorkbook("GeneratedFile.xls");
			HSSFWorkbook dbFile = getWorkbook("HighAchieverDatabase.xls");
			HSSFWorkbook haFile = new HSSFWorkbook();
			HSSFWorkbook transcriptFile = new HSSFWorkbook();			
			
			HSSFSheet gSheet = gFile.getSheetAt(0);
			int gIdx = 2;
			HSSFRow gRow = gSheet.getRow(gIdx);
			String unitCode = gSheet.getRow(gIdx).getCell(6).getStringCellValue();

			// Grab some extra details
			year = (int) Math.round(gRow.getCell(4).getNumericCellValue());
			session = gRow.getCell(5).getStringCellValue();
			
			// Step 1: Generate list of all FirstInSubject
			ArrayList<ArrayList<String>> tempEntriesList = new ArrayList<ArrayList<String>>();
			ArrayList<Integer> tempMarkEntries = new ArrayList<Integer>();
			
			while (gRow.getCell(1)!=null) {
				if ( gSheet.getRow(gIdx).getCell(6).getStringCellValue().equals(unitCode) ) {
					//Extract Student Entry
					ArrayList<String> studentEntry = new ArrayList<String>();
					studentEntry.add( gRow.getCell(0).getStringCellValue() ); // Student ID
					// extra precaution cause First Name is nullable
					String stdFirstName = gRow.getCell(2)==null?"":gRow.getCell(2).getStringCellValue();
					studentEntry.add( stdFirstName ); // Student First Name
					studentEntry.add( gRow.getCell(1).getStringCellValue() ); // Student Last Name
					studentEntry.add( gRow.getCell(6).getStringCellValue() ); // Unit Code
					studentEntry.add( gRow.getCell(7).getStringCellValue() ); // Unit Name
					tempMarkEntries.add( (int) Math.round(gRow.getCell(18).getNumericCellValue()) ); // Mark
					studentEntry.add( gRow.getCell(13).getStringCellValue() ); // Course Code
					studentEntry.add( "" + (int) Math.round(gRow.getCell(15).getNumericCellValue()) ); // Course Version
					studentEntry.add( "" + (int) Math.round(gRow.getCell(16).getNumericCellValue()) ); // Course Attempt
					//Add Student Entry to Temporary List
					tempEntriesList.add(studentEntry);
					gIdx++;
					gRow = gSheet.getRow(gIdx);
				} else {
					addFirstInSubject(firstInSubject, mark, tempEntriesList, tempMarkEntries);
					tempEntriesList.clear();
					tempMarkEntries.clear();
					unitCode = gRow.getCell(6).getStringCellValue();
				}
			}
			addFirstInSubject(firstInSubject, mark, tempEntriesList, tempMarkEntries);
			
			// Step 1: Write FirstInSubject to output Excel Workbook
			writeFIS(haFile, firstInSubject, mark);
			
			
			// Step 2: Setup
			int dbIdx = 0;
			
			@SuppressWarnings("unchecked")
			ArrayList<ArrayList<String>> fisCopy = (ArrayList<ArrayList<String>>) firstInSubject.clone();
			
			// Step 2: Add all historical first in subjects
			HSSFSheet dbSheet = dbFile.getSheetAt(0); // get sheet of all historic first in subject
			HSSFRow dbRow = dbSheet.getRow(1);
			int prevFirstIdx = 1;
			while (dbRow!=null) {
				ArrayList<String> studentEntry = new ArrayList<String>();
				studentEntry.add( ""+(int)row.getCell(STUDENTID).getNumericCellValue() );
				if (sheet.getRow(idx).getCell(1) != null) {
					studentEntry.add( row.getCell(1).getStringCellValue() );
				} else {
					studentEntry.add( null );
				}
				studentEntry.add( row.getCell(2).getStringCellValue() );
				studentEntry.add( row.getCell(UNITCODE).getStringCellValue() );
				studentEntry.add( row.getCell(UNITNAME).getStringCellValue() );
				firstInSubject.add(studentEntry);
				prevFirstIdx++;
				row = sheet.getRow(prevFirstIdx);
			}
//			
//			// Step 2: Find all available commendations
//			firstInSubject.sort( new StudentEntryComparator() );
//			while (idx<firstInSubject.size()) {
//				if (idx<firstInSubject.size() && firstInSubject.get(idx).get(0).equals(firstInSubject.get(idx+1).get(0))) {
//					String id = firstInSubject.get(idx).get(STUDENTID);
//					ArrayList<String> commd = new ArrayList<String>();
//					commd.add(firstInSubject.get(idx).get(STUDENTID));
//					commd.add(firstInSubject.get(idx).get(LASTNAME));
//					commd.add(firstInSubject.get(idx).get(FIRSTNAME));
//					String note = firstInSubject.get(idx).get(UNITCODE);
//					idx++;
//					while( idx< firstInSubject.size() && id.equals(firstInSubject.get(idx).get(STUDENTID)) ) {
//						note += " - " + firstInSubject.get(idx).get(UNITCODE);
//						idx++;
//					}
//					commd.add(note);
//					commendations.add( commd );
//				} else {
//					idx++;
//				}
//			}
//			
//			// Step 2: Filter out previously commended
//			ArrayList<String> prevCommd = new ArrayList<String>();
//			sheet = db.getSheetAt(1); // get sheet of all historic commended
//			row = sheet.getRow(1);
//			int prevCommdIdx = 1;
//			// generate list of prev commended
//			while(row != null) {
//				prevCommd.add( ""+ (int) row.getCell(STUDENTID).getNumericCellValue() );
//				prevCommdIdx++;
//				row = sheet.getRow(prevCommdIdx);
//			}
//			
//			// filter our prev commended
//			for(int i=0; i<commendations.size(); i++) {
//				if ( prevCommd.contains(commendations.get(i).get(STUDENTID)) ) {
//					commendations.remove(i);
//					i--;
//				}
//			}
//			
//			// Step 3: Write data to relevant excel workbooks
//			sheet = db.getSheetAt(0);
//			String session = sheet.getRow(prevFirstIdx-1).getCell(6).getStringCellValue();
//			if (session.charAt(1) == '1') { // Session 1
//				session = "S2" + session.substring(2);
//			} else { // Session 2
//				session = "S1 " + ( Integer.parseInt(session.substring(3))+1 );
//			}
//			
//			
//			// write commendations to output book
//			sheet = outputBook.createSheet("Commendations");
//			HSSFRow title = sheet.createRow(0);
//			title.createCell(0).setCellValue("StudentID");
//			title.createCell(1).setCellValue("First Name");
//			title.createCell(2).setCellValue("Last Name");
//			title.createCell(3).setCellValue("Notes");
//			
//			idx = 1;
//			for (ArrayList<String> entry: commendations) {
//				row = sheet.createRow(idx);
//				row.createCell(0).setCellValue( Integer.parseInt(entry.get(STUDENTID)) );
//				row.createCell(1).setCellValue(entry.get(FIRSTNAME));
//				row.createCell(2).setCellValue(entry.get(LASTNAME));
//				row.createCell(3).setCellValue(entry.get(3));
//				idx++;
//			}
//			
//			// write first in subject to database book
//			sheet = db.getSheetAt(0);
//			for (int i=0; i<fisCopy.size(); i++) {
//				row = sheet.createRow(prevFirstIdx);
//				row.createCell(0).setCellValue( Integer.parseInt(fisCopy.get(i).get(STUDENTID)) );
//				row.createCell(1).setCellValue(fisCopy.get(i).get(FIRSTNAME));
//				row.createCell(2).setCellValue(fisCopy.get(i).get(LASTNAME));
//				row.createCell(3).setCellValue(fisCopy.get(i).get(UNITCODE));
//				row.createCell(4).setCellValue(fisCopy.get(i).get(UNITNAME));
//				row.createCell(5).setCellValue(mark.get(i));
//				row.createCell(6).setCellValue(session);
//				prevFirstIdx++;
//			}
//			
//			// write commendations to database book
//			sheet = db.getSheetAt(1);
//			for (ArrayList<String> entry: commendations) {
//				row = sheet.createRow(prevCommdIdx);
//				row.createCell(0).setCellValue( Integer.parseInt(entry.get(STUDENTID)) );
//				row.createCell(1).setCellValue(entry.get(FIRSTNAME));
//				row.createCell(2).setCellValue(entry.get(LASTNAME));
//				row.createCell(3).setCellValue(session);
//				row.createCell(4).setCellValue(entry.get(3));
//				prevCommdIdx++;
//			}
//			
//			FileOutputStream outputStream = new FileOutputStream("High Achiever "+session+".xls");
//	        outputBook.write(outputStream);
//	        
//	        FileOutputStream dbStream = new FileOutputStream("HighAchieverDatabase.xls");
//	        db.write(dbStream);
//			
//			
		} catch (IOException e) { // IOException Handling
			try {
				FileWriter fileWriter = new FileWriter("ERROR.txt");
			    PrintWriter printWriter = new PrintWriter(fileWriter);
				FileInputStream inputStream = new FileInputStream("HighAchieverDatabase.xlsx");
				printWriter.println("FILES IN THE WRONG FORMAT");
			    printWriter.println("The file \"HighAchieverDatabase\" is in the wrong format. Double check the following:");
			    printWriter.println("    -  The file \"HighestAchieverDatabase\" is in the .xls format instead of the .xlsx format ");
			    printWriter.close();
			} catch (IOException e1) {
				try {
					FileWriter fileWriter = new FileWriter("ERROR.txt");
				    PrintWriter printWriter = new PrintWriter(fileWriter);
					FileInputStream inputStream = new FileInputStream("GeneratedFile.xlsx");
					printWriter.println("FILES IN THE WRONG FORMAT");
				    printWriter.println("The file \"GeneratedFile\" is in the wrong format. Double check the following:");
				    printWriter.println("    -  The file \"GeneratedFile\" is in the .xls format instead of the .xlsx format ");
				    printWriter.close();
				} catch (IOException e2) {
					try {
						FileWriter fileWriter = new FileWriter("ERROR.txt");
					    PrintWriter printWriter = new PrintWriter(fileWriter);
					    printWriter.println("FILES CANNOT BE FOUND");
					    printWriter.println("An error occurred while trying to access excel files. Double check the following:");
					    printWriter.println("    -  Both files are renamed correctly as \"GeneratedFile\" and \"HighestAchieverDatabase\" ");
					    printWriter.println("    -  And you are not currently accessing any relevant Excel files");
					    printWriter.close();
					    e.printStackTrace();
					} catch (IOException e3) {
						System.exit(1);
					}
				}
			}
		} catch (NullPointerException e) { // Blank cell exception
			try {
				FileWriter fileWriter = new FileWriter("ERROR.txt");
			    PrintWriter printWriter = new PrintWriter(fileWriter);
			    printWriter.println("BLANK CELL DETECTED");
			    printWriter.println("A blank cell has been detected in the Excel files. Double check the following:");
			    printWriter.println("    -  No Cell in vital columns e.g Student ID, Unit Code, Mark are left as blank");
			    printWriter.println("    -  Check both GeneratedFile.xls and HighAchieverDatabase.xls");
			    printWriter.close();
			    e.printStackTrace();
			} catch (IOException e1) {
				System.exit(1);
			}
		} catch (Exception e) { // Unforeseen exception
			try {
			    FileWriter fileWriter = new FileWriter("ERROR.txt");
			    PrintWriter printWriter = new PrintWriter(fileWriter);
			    printWriter.println("UNEXPECTED ERROR");
			    printWriter.println("An unexpected error has occurred. Daniel did not have enough foresight to see this one coming.");
			    printWriter.println("At this point, you should call Daniel in and let him deal with his own mistake.");
			    printWriter.println("Alternatively, you can call in another computer science student and let them deal with Daniel's mess :)");
			    printWriter.println("Below is complete garbage that's meant for Daniel to look at. Daniel (or unfortunate computer science student), if you reading this, goodluck.");
			    printWriter.println("----------------------------------------------------------------------------------------------------------------------------------------------");
			    e.printStackTrace(printWriter);
			    e.printStackTrace();
			    printWriter.close();
			} catch (IOException e1) {
				System.exit(1);
			}
		}
		
	}
	
//	public static HSSFSheet readExcel1(String path, int sheetIdx) throws IOException {
//		FileInputStream inputStream = new FileInputStream(path);
//		HSSFWorkbook wb = new HSSFWorkbook(inputStream);
//		HSSFSheet sheet = wb.getSheetAt(sheetIdx);
//		return sheet;
//	}
//	
//	public static HSSFWorkbook readExcel2(String path) throws IOException {
//		FileInputStream inputStream = new FileInputStream(path);
//		HSSFWorkbook wb = new HSSFWorkbook(inputStream);
//		return wb;
//	}
	
	public static HSSFWorkbook getWorkbook(String path) throws IOException {
		FileInputStream inputStream = new FileInputStream(path);
		HSSFWorkbook workbook = new HSSFWorkbook(inputStream);
		return workbook;
	}
	
	// Write the initial workbook
	public static void writeFIS(HSSFWorkbook book, ArrayList<ArrayList<String>> list, ArrayList<Integer> grade) {
		// Setup workbook
		HSSFSheet sheet = book.createSheet("First in Subject");
		String currentSession = "";
		if (session.charAt(0) == 'F') {
			currentSession = "S1 "+year;
		} else if (session.charAt(0) =='S') {
			currentSession = "S2 "+year;
		}
		// Create Title row
		HSSFRow title = sheet.createRow(0);
		title.createCell(0).setCellValue("Student ID");
		title.createCell(1).setCellValue("First Name");
		title.createCell(2).setCellValue("Last Name");
		title.createCell(3).setCellValue("Unit Code");
		title.createCell(4).setCellValue("Unit Name");
		title.createCell(5).setCellValue("Mark");
		title.createCell(6).setCellValue("Session");
		// Fill all data rows
		int idx=1;
		for (ArrayList<String> studentEntry: list) {
			HSSFRow row = sheet.createRow(idx);
			row.createCell(0).setCellValue( Integer.parseInt(studentEntry.get(STUDENTID)) );
			row.createCell(1).setCellValue(studentEntry.get(FIRSTNAME));
			row.createCell(2).setCellValue(studentEntry.get(LASTNAME));
			row.createCell(3).setCellValue(studentEntry.get(UNITCODE));
			row.createCell(4).setCellValue(studentEntry.get(UNITNAME));
			row.createCell(5).setCellValue(grade.get(idx-1));
			row.createCell(6).setCellValue(currentSession);
			idx++;
		}
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
