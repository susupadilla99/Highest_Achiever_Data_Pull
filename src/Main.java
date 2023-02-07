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
	static String currentSession = "";
	
	
	public static void main(String[] args) {
		try {
			
			/** Step 1: Find all this-session first in subject students**/
			
			// setup
			HSSFWorkbook gFile = getWorkbook("GeneratedFile.xls");
			HSSFWorkbook dbFile = getWorkbook("HighAchieverDatabase.xls");
			HSSFWorkbook haFile = new HSSFWorkbook();
			HSSFWorkbook transcriptFISFile = new HSSFWorkbook();
			HSSFWorkbook transcriptCommdFile = new HSSFWorkbook();
			HSSFSheet gSheet = gFile.getSheetAt(0);
			int gIdx = 2;
			HSSFRow gRow = gSheet.getRow(gIdx);
			String unitCode = gSheet.getRow(gIdx).getCell(6).getStringCellValue();
			// setup year & session
			year = (int) Math.round(gRow.getCell(4).getNumericCellValue());
			session = gRow.getCell(5).getStringCellValue();
			if (session.charAt(0) == 'F') {
				currentSession = "S1 "+year;
			} else if (session.charAt(0) =='S') {
				currentSession = "S2 "+year;
			} else {
				currentSession = "S3 "+year;
			}
			
			// generate array list of all first in subject student
			ArrayList<ArrayList<String>> eachUnitsStudentList = new ArrayList<ArrayList<String>>();
			ArrayList<Integer> eachUnitsMarkList = new ArrayList<Integer>();
			while (gRow.getCell(1)!=null) {
				if ( gSheet.getRow(gIdx).getCell(6).getStringCellValue().equals(unitCode) ) {
					// extract student entry
					ArrayList<String> studentEntry = new ArrayList<String>();
					studentEntry.add( gRow.getCell(0).getStringCellValue() ); // Student ID
					// extra precaution cause First Name is nullable
					String stdFirstName = gRow.getCell(2)==null?"":gRow.getCell(2).getStringCellValue();
					studentEntry.add( stdFirstName ); // Student First Name
					studentEntry.add( gRow.getCell(1).getStringCellValue() ); // Student Last Name
					studentEntry.add( gRow.getCell(6).getStringCellValue() ); // Unit Code
					studentEntry.add( gRow.getCell(7).getStringCellValue() ); // Unit Name
					eachUnitsMarkList.add( (int) Math.round(gRow.getCell(18).getNumericCellValue()) ); // Mark
					studentEntry.add( gRow.getCell(13).getStringCellValue() ); // Course Code
					studentEntry.add( "" + (int) Math.round(gRow.getCell(15).getNumericCellValue()) ); // Course Version
					studentEntry.add( "" + (int) Math.round(gRow.getCell(16).getNumericCellValue()) ); // Course Attempt
					//add Student Entry to each-unit list
					eachUnitsStudentList.add(studentEntry);
					gIdx++;
					gRow = gSheet.getRow(gIdx);
				} else {
					addFirstInSubject(firstInSubject, mark, eachUnitsStudentList, eachUnitsMarkList);
					eachUnitsStudentList.clear();
					eachUnitsMarkList.clear();
					unitCode = gRow.getCell(6).getStringCellValue();
				}
			}
			addFirstInSubject(firstInSubject, mark, eachUnitsStudentList, eachUnitsMarkList);
	
			// write FirstInSubject to output Excel Workbook
			writeFIS(haFile, firstInSubject, mark);
			
			/** Step 2: Find all historical first in subject student from database & put them in array list **/
			
			// setup
			@SuppressWarnings("unchecked")
			ArrayList<ArrayList<String>> fisCopy = (ArrayList<ArrayList<String>>) firstInSubject.clone();
			HSSFSheet dbSheet = dbFile.getSheetAt(0); // get sheet of all historic first in subject
			HSSFRow dbRow = dbSheet.getRow(1);
			int dbIdx = 1;
			// generate list of all historical first in subject students
			while (dbRow!=null) {
				ArrayList<String> studentEntry = new ArrayList<String>();
				studentEntry.add( ""+(int)dbRow.getCell(0).getNumericCellValue() ); // Student ID
				// extra precaution because first name is nullable
				String stdFirstName = dbRow.getCell(1)==null?"":gRow.getCell(1).getStringCellValue();
				studentEntry.add( stdFirstName ); // Student First Name
				studentEntry.add( dbRow.getCell(2).getStringCellValue() ); // Student Last Name
				studentEntry.add( dbRow.getCell(3).getStringCellValue() ); // Unit Code
				studentEntry.add( dbRow.getCell(4).getStringCellValue() ); // Unit Name
				mark.add( (int)Math.round(dbRow.getCell(5).getNumericCellValue()) ); // Mark
				studentEntry.add( gRow.getCell(13).getStringCellValue() ); // Course Code
				studentEntry.add( "" + (int) Math.round(gRow.getCell(15).getNumericCellValue()) ); // Course Version
				studentEntry.add( "" + (int) Math.round(gRow.getCell(16).getNumericCellValue()) ); // Course Attempt
				// add student entry to first in subject list
				firstInSubject.add(studentEntry);
				dbIdx++;
				dbRow = dbSheet.getRow(dbIdx);
			}
			
			/** Step 3: Find all available commendations using the array list of ALL first in subject **/
			
			// sort first in subject list using student ID
			firstInSubject.sort( new StudentEntryComparator() );
			int idx = 0;
			while (idx<firstInSubject.size()) {
				// idx is valid & there are 2 consecutive first in subject entry
				if (idx<(firstInSubject.size()-1) && firstInSubject.get(idx).get(0).equals(firstInSubject.get(idx+1).get(0))) {
					String id = firstInSubject.get(idx).get(STUDENTID);
					ArrayList<String> commendationEntry = new ArrayList<String>();
					commendationEntry.add(firstInSubject.get(idx).get(STUDENTID)); // Student ID
					commendationEntry.add(firstInSubject.get(idx).get(FIRSTNAME)); // Student First Name
					commendationEntry.add(firstInSubject.get(idx).get(LASTNAME)); // Student Last Name
					String note = firstInSubject.get(idx).get(UNITCODE);
					idx++;
					// record all FIS units for this 1 student ID
					while( idx< firstInSubject.size() && id.equals(firstInSubject.get(idx).get(STUDENTID)) ) {
						note += " - " + firstInSubject.get(idx).get(UNITCODE);
						idx++;
					}
					commendationEntry.add(note); // Notes on units & mark of FIS
					commendationEntry.add(firstInSubject.get(idx).get(COURSECODE));
					commendationEntry.add(firstInSubject.get(idx).get(COURSEVERSION));
					commendationEntry.add(firstInSubject.get(idx).get(COURSEATTEMPT));
					commendations.add( commendationEntry );
				} else {
					idx++;
				}
			}

			/** Step 4: Filter out students who have been previously commended **/
			
			// setup
			ArrayList<String> prevCommd = new ArrayList<String>();
			HSSFSheet pcSheet = dbFile.getSheetAt(1); // get sheet of all historic commended
			HSSFRow pcRow = pcSheet.getRow(1);
			int pcIdx = 1;
			
			// generate list of previous commended student ID
			while(pcRow != null) {
				prevCommd.add( ""+ (int) pcRow.getCell(STUDENTID).getNumericCellValue() );
				pcIdx++;
				pcRow = pcSheet.getRow(pcIdx);
			}
			
			// filter our prev commended
			for(int i=0; i<commendations.size(); i++) {
				if ( prevCommd.contains(commendations.get(i).get(STUDENTID)) ) {
					commendations.remove(i);
					i--;
				}
			}

			/**	Step 5: Write data to relevant excel workbooks **/
			
			// write commendations to high achiever book
			HSSFSheet haSheet = haFile.createSheet("Commendations");
			HSSFRow title = haSheet.createRow(0);
			title.createCell(0).setCellValue("StudentID");
			title.createCell(1).setCellValue("First Name");
			title.createCell(2).setCellValue("Last Name");
			title.createCell(3).setCellValue("Notes");
			title.createCell(4).setCellValue("Session");
			
			int haIdx = 1;
			for (ArrayList<String> entry: commendations) {
				HSSFRow haRow = haSheet.createRow(haIdx);
				haRow.createCell(0).setCellValue( Integer.parseInt(entry.get(0)) );
				haRow.createCell(1).setCellValue(entry.get(1));
				haRow.createCell(2).setCellValue(entry.get(2));
				haRow.createCell(3).setCellValue(entry.get(3));
				haRow.createCell(4).setCellValue(currentSession);
				haIdx++;
			}
			
			// write first in subject to database book
			dbSheet = dbFile.getSheetAt(0);
			for (int i=0; i<fisCopy.size(); i++) {
				dbRow = dbSheet.createRow(dbIdx);
				dbRow.createCell(0).setCellValue( Integer.parseInt(fisCopy.get(i).get(STUDENTID)) );
				dbRow.createCell(1).setCellValue(fisCopy.get(i).get(FIRSTNAME));
				dbRow.createCell(2).setCellValue(fisCopy.get(i).get(LASTNAME));
				dbRow.createCell(3).setCellValue(fisCopy.get(i).get(UNITCODE));
				dbRow.createCell(4).setCellValue(fisCopy.get(i).get(UNITNAME));
				dbRow.createCell(5).setCellValue(mark.get(i));
				dbRow.createCell(6).setCellValue(currentSession);
				dbRow.createCell(7).setCellValue(fisCopy.get(i).get(COURSECODE));
				dbRow.createCell(8).setCellValue(fisCopy.get(i).get(COURSEVERSION));
				dbRow.createCell(9).setCellValue(fisCopy.get(i).get(COURSEATTEMPT));
				dbIdx++;
			}
			
			// write commendations to database book
			pcSheet = dbFile.getSheetAt(1);
			for (ArrayList<String> entry: commendations) {
				pcRow = pcSheet.createRow(pcIdx);
				pcRow.createCell(0).setCellValue( Integer.parseInt(entry.get(0)) ); // Student ID
				pcRow.createCell(1).setCellValue(entry.get(1)); // Student First Name
				pcRow.createCell(2).setCellValue(entry.get(2)); // Student Last Name
				pcRow.createCell(3).setCellValue(entry.get(3)); // Notes
				pcRow.createCell(4).setCellValue(currentSession); // Session
				pcRow.createCell(5).setCellValue(entry.get(4)); // Course Code
				pcRow.createCell(6).setCellValue(entry.get(5)); // Course Version
				pcRow.createCell(7).setCellValue(entry.get(6)); // Course Attempt
				pcIdx++;
			}
			
			// write to fis transcript file
			HSSFSheet fisSheet = transcriptFISFile.createSheet("in");
			title = fisSheet.createRow(0);
			title.createCell(0).setCellValue("stu_id");
			title.createCell(1).setCellValue("seq_no");
			title.createCell(2).setCellValue("cmt_cd");
			title.createCell(3).setCellValue("stu_cmt_effct_dt");
			title.createCell(4).setCellValue("stu_cmt_txt_1");
			title.createCell(4).setCellValue("spk_cd");
			title.createCell(4).setCellValue("spk_ver_no");
			title.createCell(4).setCellValue("ssp_att_no");
			title.createCell(4).setCellValue("avail_yr");
			title.createCell(4).setCellValue("sprd_cd");
			
			int fisIdx = 1;
			for (ArrayList<String> entry: fisCopy) {
				HSSFRow fisRow = fisSheet.createRow(fisIdx);
				fisRow.createCell(0).setCellValue( Integer.parseInt(entry.get(STUDENTID)) );
				fisRow.createCell(2).setCellValue( FISCOMMENT );
				fisRow.createCell(4).setCellValue( entry.get(UNITCODE) );
				fisRow.createCell(5).setCellValue( entry.get(COURSECODE) );
				fisRow.createCell(6).setCellValue( entry.get(COURSEVERSION) );
				fisRow.createCell(7).setCellValue( entry.get(COURSEATTEMPT) );
				fisRow.createCell(8).setCellValue( year );
				fisRow.createCell(9).setCellValue( session );
				fisIdx++;
			}
			
			// write to commendation transcript file
			HSSFSheet commdSheet = transcriptCommdFile.createSheet("in");
			title = commdSheet.createRow(0);
			title.createCell(0).setCellValue("stu_id");
			title.createCell(1).setCellValue("seq_no");
			title.createCell(2).setCellValue("cmt_cd");
			title.createCell(3).setCellValue("stu_cmt_effct_dt");
			title.createCell(4).setCellValue("stu_cmt_txt_1");
			title.createCell(4).setCellValue("spk_cd");
			title.createCell(4).setCellValue("spk_ver_no");
			title.createCell(4).setCellValue("ssp_att_no");
			title.createCell(4).setCellValue("avail_yr");
			title.createCell(4).setCellValue("sprd_cd");
			
			int commdIdx = 1;
			for (ArrayList<String> entry: commendations) {
				HSSFRow commdRow = commdSheet.createRow(commdIdx);
				commdRow.createCell(0).setCellValue( Integer.parseInt(entry.get(0)) );
				commdRow.createCell(2).setCellValue( COMMDCOMMENT );
				commdRow.createCell(5).setCellValue( entry.get(4) );
				commdRow.createCell(6).setCellValue( entry.get(5) );
				commdRow.createCell(7).setCellValue( entry.get(6) );
				commdRow.createCell(8).setCellValue( year );
				commdRow.createCell(9).setCellValue( session );
				commdIdx++;
			}
			
			
			FileOutputStream haStream = new FileOutputStream("High Achiever "+currentSession+".xls");
	        haFile.write(haStream);
	        
	        FileOutputStream dbStream = new FileOutputStream("HighAchieverDatabase.xls");
	        dbFile.write(dbStream);
			
	        FileOutputStream fisStream = new FileOutputStream("First in Subject Transcript Upload.xls");
	        transcriptFISFile.write(fisStream);

	        FileOutputStream commdStream = new FileOutputStream("Commendation Transcript Upload.xls");
	        transcriptCommdFile.write(commdStream);
	        
	        haFile.close();
	        dbFile.close();
	        transcriptFISFile.close();
	        transcriptCommdFile.close();
			
		} catch (IOException e) { // IOException Handling
			try {
				FileWriter fileWriter = new FileWriter("ERROR.txt");
			    PrintWriter printWriter = new PrintWriter(fileWriter);
				FileInputStream inputStream = new FileInputStream("HighAchieverDatabase.xlsx");
				printWriter.println("FILES IN THE WRONG FORMAT");
			    printWriter.println("The file \"HighAchieverDatabase\" is in the wrong format. Double check the following:");
			    printWriter.println("    -  The file \"HighestAchieverDatabase\" is in the .xls format instead of the .xlsx format ");
			    printWriter.close();
			    inputStream.close();
			} catch (IOException e1) {
				try {
					FileWriter fileWriter = new FileWriter("ERROR.txt");
				    PrintWriter printWriter = new PrintWriter(fileWriter);
					FileInputStream inputStream = new FileInputStream("GeneratedFile.xlsx");
					printWriter.println("FILES IN THE WRONG FORMAT");
				    printWriter.println("The file \"GeneratedFile\" is in the wrong format. Double check the following:");
				    printWriter.println("    -  The file \"GeneratedFile\" is in the .xls format instead of the .xlsx format ");
				    printWriter.close();
				    inputStream.close();
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
