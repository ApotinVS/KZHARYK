import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;

public class main {
    public static void main(String... args) throws IOException, InvalidFormatException {
        ArrayList<BkExcel> list = Parser.parse("C:\\Users\\ApotinV\\Desktop\\от Жалгаса\\16.xlsx");
        ArrayList<BkExcel> listout =  BkExcel.check(BkExcel.check(list));

        BkExcel.writeIntiExcel("C:\\Users\\ApotinV\\Desktop\\от Жалгаса\\11111111.xls",listout);
        //Connector.PushDB(list);



    }}
