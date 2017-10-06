import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.ParseException;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;


public class Parser {

    public static ArrayList<BkExcel> parse(String path) throws IOException, InvalidFormatException, ParseException {
        BkExcel bkExcel ;
        ArrayList <BkExcel> listbkExcel = new ArrayList<BkExcel>();
        InputStream in = null;

        //HSSFWorkbook wb = null;
            try {
                in = new FileInputStream(path);
                // wb = new HSSFWorkbook(in);
            } catch (IOException e) {
                e.printStackTrace();
            }

            Workbook wb = WorkbookFactory.create(new File(path));
            Sheet sheet = wb.getSheetAt(0);
            XSSFRow row;
            Iterator<Row> rows = sheet.rowIterator();
            DataFormatter formatter = new DataFormatter();
            while (rows.hasNext()) {
                row = (XSSFRow) rows.next();
                bkExcel = new BkExcel(formatter.formatCellValue(row.getCell(0)),
                        formatter.formatCellValue(row.getCell(1)),
                        formatter.formatCellValue(row.getCell(2)),
                        formatter.formatCellValue(row.getCell(3)),
                        formatter.formatCellValue(row.getCell(4)),
                        formatter.formatCellValue(row.getCell(5)),
                        formatter.formatCellValue(row.getCell(6)),
                        formatter.formatCellValue(row.getCell(7)),
                        formatter.formatCellValue(row.getCell(8)),
                        formatter.formatCellValue(row.getCell(9)),
                        formatter.formatCellValue(row.getCell(10)),
                        formatter.formatCellValue(row.getCell(11)),
                        formatter.formatCellValue(row.getCell(12)),
                        formatter.formatCellValue(row.getCell(13)),
                        formatter.formatCellValue(row.getCell(14)),
                        formatter.formatCellValue(row.getCell(15)),
                        formatter.formatCellValue(row.getCell(16)),
                        formatter.formatCellValue(row.getCell(17)),
                        formatter.formatCellValue(row.getCell(18)),
                        formatter.formatCellValue(row.getCell(19)));
                listbkExcel.add(bkExcel);
            }

        return listbkExcel;
    }
}