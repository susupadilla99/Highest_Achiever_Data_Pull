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
			
			ArrayList<ArrayList<String>> highestAchievers = new ArrayList<ArrayList<String>>();
			ArrayList<ArrayList<String>> temp = new ArrayList<ArrayList<String>>();
			ArrayList<Integer> mark = new ArrayList<Integer>();
			
			HSSFWorkbook outputBook = new HSSFWorkbook();
			HSSFSheet out = outputBook.createSheet();
			int outIdx = 0;
			
			HSSFSheet sheet = readExcel("./src/rss/LCLARKSO2022.xls", 0);
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
					if( temp.size()>=10 ) {
						int topScore = mark.get(0);
						// Add all high-achievers
						for(int i=0; i<temp.size(); i++) {
							// Add 1 high achiever data
							if (mark.get(i) == topScore) {
								out.createRow(outIdx).createCell(0).setCellValue(temp.get(i).get(STUDENTID));
								out.getRow(outIdx).createCell(1).setCellValue(temp.get(i).get(FIRSTNAME));
								out.getRow(outIdx).createCell(2).setCellValue(temp.get(i).get(LASTNAME));
								out.getRow(outIdx).createCell(3).setCellValue(temp.get(i).get(UNITCODE));
								out.getRow(outIdx).createCell(4).setCellValue(temp.get(i).get(UNITNAME));
								out.getRow(outIdx).createCell(5).setCellValue(mark.get(i));
								outIdx++;
							}
						}
					}
					temp.clear();
					mark.clear();
					unitCode = row.getCell(UNITCODE).getStringCellValue();
				}
			}
			
			FileOutputStream outputStream = new FileOutputStream("JavaBooks.xls");
	        outputBook.write(outputStream);
			
			System.out.println(temp);
			
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
	
	public static void addHighestAchievers(ArrayList<ArrayList<String>> target, ArrayList<ArrayList<String>> temp, ArrayList<Integer> grade) {
		
	}
}
