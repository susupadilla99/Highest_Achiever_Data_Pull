import java.util.Comparator;

public class StudentEntryComparator implements Comparator<StudentEntry> {
	@Override
	public int compare(StudentEntry o1, StudentEntry o2) {
		return o1.studentID - o2.studentID;
	}
}
