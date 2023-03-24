
public class StudentEntry {
	// Mandatory - ish fields
	int studentID;
	String firstName;
	String lastName;
	String unitCode;
	String unitName;
	String courseCode;
	String courseName;
	int courseVersion;
	int courseAttempt;
	int mark;

	// Optional fields
	String notes;
	
	public StudentEntry(int id, String fName, String lName, String uCode, String uName, String cCode, String cName, int cVersion, int cAttempt, int m) {
		studentID = id;
		firstName = fName;
		lastName = lName;
		unitCode = uCode;
		unitName = uName;
		courseCode = cCode;
		courseName = cName;
		courseVersion = cVersion;
		courseAttempt = cAttempt;
		mark = m;
		notes = "";
	}

	public StudentEntry copy(){
		return new StudentEntry(studentID, firstName, lastName, unitCode, unitName, courseCode, courseName, courseVersion, courseAttempt, mark);
	}

}
