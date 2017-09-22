import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;

public class main {
    public static void main(String... args) throws IOException, InvalidFormatException, ParseException {
        SimpleDateFormat format = new SimpleDateFormat("dd.MM.yyyy");
        Date dateTake = format.parse("11.09.2017");
        ArrayList<BkExcel> list1 = new ArrayList<BkExcel>();
        ArrayList<BkExcel> list2 = new ArrayList<BkExcel>();
        ArrayList<BkExcel> list;
        int countOfFiles = DirectoryFileNames.GetFileNames().size();
        for (int i = 0; i < countOfFiles; i++) {
            list = Parser.parse("C:\\Users\\ApotinV\\Desktop\\от Жалгаса\\" + DirectoryFileNames.GetFileNames().get(i));
            list1 = BkExcel.check(list, dateTake);
            for (int j = 0; j < list1.size(); j++) {
                list2.add(list1.get(j));
            }
        }
        BkExcel.writeIntoExcel("C:\\Users\\ApotinV\\Desktop\\от Жалгаса\\aaaaa.xls",list2);







    }}
