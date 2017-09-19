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
        ArrayList<BkExcel> list = Parser.parse("C:\\Users\\ApotinV\\Desktop\\от Жалгаса\\16.xlsx");
        ArrayList<BkExcel> listout =  BkExcel.check(list, dateTake);



        BkExcel.writeIntoExcel("C:\\Users\\ApotinV\\Desktop\\от Жалгаса\\234.xls",listout);
        //Connector.PushDB(list);



    }}
