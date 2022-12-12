import java.util.ArrayList;
import java.util.Comparator;

public class StudentEntryComparator implements Comparator<ArrayList<String>> {
	@Override
	public int compare(ArrayList<String> o1, ArrayList<String> o2) {
		return o1.get(0).compareTo(o2.get(0));
	}
}
