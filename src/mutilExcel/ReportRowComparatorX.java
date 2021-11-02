package mutilExcel;

import java.util.Comparator;

public class ReportRowComparatorX implements Comparator<ReportRowX> {
    public int compare(ReportRowX o1, ReportRowX o2) {
        return o1.getSortKey().compareTo(o2.getSortKey());
    }
}
