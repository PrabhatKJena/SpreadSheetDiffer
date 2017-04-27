/**
 * Created by prajena on 4/13/17.
 */
public class DiffExcelRunner {
    public static void main(String[] args) throws Exception {
        DiffExcel diffExcel = new DiffExcel("/Users/prajena/Downloads/FT_test_data/left.xls",
                "/Users/prajena/Downloads/FT_test_data/right.xls",
                "RelatedActivity");
                //null);
        diffExcel.doDiff();
    }
}
