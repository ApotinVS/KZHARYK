import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.Date;

public class BkExcel {

    public String number;
    public String plc;
    public String lSchet;
    public String KTP;
    public String konc;
    public String street;
    public String house;
    public String apartment;
    public String nSnFn;
    public String model;
    public String nCounter;
    public String energy;
    public String kWth;
    public String t1;
    public String t1v;
    public String t2;
    public String t2v;
    public String t3;
    public String t3v;
    public String summ;
    private Date dateNowT1;
    private Date dateNowT2;
    private Date dateNowT3;

    public static ArrayList<BkExcel> check (ArrayList<BkExcel> list/*, Date dateTake*/) throws ParseException {
        int sortNumber = 3;
        ArrayList<Integer> temp = new ArrayList<Integer>();
        ArrayList<BkExcel> listOut = new ArrayList<BkExcel>();
        ArrayList<BkExcel> listOut1 = new ArrayList<BkExcel>();

        for (int i = 2; i <list.size() ; i++) {
            if (!list.get(i).energy.equals("----") && !list.get(i).energy.equals("****") ){
                listOut.add(list.get(i));
            }
            else if ((!list.get(i).summ.equals("0,00") && !list.get(i).summ.equals("0.00") )/*&& !list.get(i).t1v.equals("****") && !list.get(i).t1.equals("****") && !list.get(i).energy.equals("----") &&
                !list.get(i).t2v.equals("****") && !list.get(i).t2.equals("****") /*&& !list.get(i).t3v.equals("****") && !list.get(i).t3.equals("****")*/ ){
            listOut.add(list.get(i));
            }
        }

        for (int i = 0; i < list.size(); i++) {
            if (list.get(i).summ.contains("."))
            list.get(i).summ = list.get(i).summ.substring(0,list.get(i).summ.indexOf("."));
            if (list.get(i).summ.contains(","))
            list.get(i).summ = list.get(i).summ.substring(0,list.get(i).summ.indexOf(","));

        }

        /*for (int i = 0; i <listOut.size() ; i++) {
            listOut.get(i).t1 = listOut.get(i).t1.substring(0, 5)+".2017";
            listOut.get(i).t2 = listOut.get(i).t2.substring(0, 5)+".2017";
            listOut.get(i).t3 = listOut.get(i).t3.substring(0, 5)+".2017";
            SimpleDateFormat format = new SimpleDateFormat("dd.MM.yyyy");
            listOut.get(i).dateNowT1 = format.parse(listOut.get(i).t1);
            listOut.get(i).dateNowT2 = format.parse(listOut.get(i).t2);
            listOut.get(i).dateNowT3 = format.parse(listOut.get(i).t3);

            if (dateTake.getTime() >= listOut.get(i).dateNowT1.getTime() && dateTake.getTime() >=
                    listOut.get(i).dateNowT2.getTime() && dateTake.getTime() >= listOut.get(i).dateNowT3.getTime()){
                long difference1 = dateTake.getTime() - listOut.get(i).dateNowT1.getTime();
                long difference2 = dateTake.getTime() - listOut.get(i).dateNowT2.getTime();
                long difference3 = dateTake.getTime() - listOut.get(i).dateNowT3.getTime();
                int days1 = (int) difference1 / (24 * 60 * 60 * 1000);
                int days2 = (int) difference2 / (24 * 60 * 60 * 1000);
                int days3 = (int) difference3 / (24 * 60 * 60 * 1000);
                if (days1 < sortNumber && days2 < sortNumber && days3 < sortNumber){
                    listOut1.add(listOut.get(i));
                }
            }
            else if (listOut.get(i).dateNowT1.getTime() > dateTake.getTime() || listOut.get(i).dateNowT2.getTime() >
                    dateTake.getTime() || listOut.get(i).dateNowT3.getTime() > dateTake.getTime()){
                long difference1 = dateTake.getTime() - listOut.get(i).dateNowT1.getTime();
                long difference2 = dateTake.getTime() - listOut.get(i).dateNowT2.getTime();
                long difference3 = dateTake.getTime() - listOut.get(i).dateNowT3.getTime();
                int days1 = (int) difference1 / (24 * 60 * 60 * 1000);
                int days2 = (int) difference2 / (24 * 60 * 60 * 1000);
                int days3 = (int) difference3 / (24 * 60 * 60 * 1000);
                if (days1 < sortNumber && days2 < sortNumber && days3 < sortNumber){
                    listOut1.add(listOut1.get(i));
                }
            }
        }
        listOut.clear();*/
       Collections.sort(listOut, COMPARE_BY_PLC);/*
        for (int i = 1; i < listOut1.size() ; i++) {
            if (listOut1.get(i).plc.equals(listOut1.get(i-1).plc)){
                temp.add(i);
                if (i == 1 &&  !listOut1.get(i).plc.equals(listOut1.get(i+1).plc)){
                    temp.add(i-1);
                }
                else if ( !listOut1.get(i).plc.equals(listOut1.get(i-2).plc) && i > 1){
                    temp.add(i-1);
                }
            }
        }*/

        Collections.sort(temp);

        /*for (int i = temp.size()-1; i >= 0; i--) {
            System.out.println(temp.get(i));
            listOut1.remove(listOut1.get(temp.get(i)));
        }*/


        return listOut;
    }

    public static final Comparator<BkExcel> COMPARE_BY_PLC = new Comparator<BkExcel>() {
        @Override
        public int compare(BkExcel bkExcel, BkExcel t1) {
           return Integer.parseInt(bkExcel.plc) - Integer.parseInt(t1.plc);
        }
    };

    public static void writeIntoExcel(String file, ArrayList<BkExcel> list){
        Workbook workbook = new HSSFWorkbook();
        Sheet sheet = workbook.createSheet();
        for (int i = 0; i < list.size(); i++) {
            Row row = sheet.createRow(i);
            Cell number = row.createCell(0);
            Cell plc = row.createCell(1);
            Cell lSchet = row.createCell(2);
            Cell KTP = row.createCell(3);
            Cell konc = row.createCell(4);
            Cell street = row.createCell(5);
            Cell house = row.createCell(6);
            Cell apartment = row.createCell(7);
            Cell nSnFn = row.createCell(8);
            Cell model = row.createCell(9);
            Cell nCounter = row.createCell(10);
            Cell energy = row.createCell(11);
            Cell kWth = row.createCell(12);
            Cell t1 = row.createCell(13);
            Cell t1v = row.createCell(14);
            Cell t2 = row.createCell(15);
            Cell t2v = row.createCell(16);
            Cell t3 = row.createCell(17);
            Cell t3v = row.createCell(18);
            Cell summ = row.createCell(19);

            number.setCellValue(list.get(i).number);
            plc.setCellValue(list.get(i).plc);
            lSchet.setCellValue(list.get(i).lSchet);
            KTP.setCellValue(list.get(i).KTP);
            konc.setCellValue(list.get(i).konc);
            street.setCellValue(list.get(i).street);
            house.setCellValue(list.get(i).house);
            apartment.setCellValue(list.get(i).apartment);
            nSnFn.setCellValue(list.get(i).nSnFn);
            model.setCellValue(list.get(i).model);
            nCounter.setCellValue(list.get(i).nCounter);
            energy.setCellValue(list.get(i).energy);
            kWth.setCellValue(list.get(i).kWth);
            t1.setCellValue(list.get(i).t1);
            t1v.setCellValue(list.get(i).t1v);
            t2.setCellValue(list.get(i).t2);
            t2v.setCellValue(list.get(i).t2v);
            t3.setCellValue(list.get(i).t3);
            t3v.setCellValue(list.get(i).t3v);
            summ.setCellValue(list.get(i).summ);
        }

        try (FileOutputStream out = new FileOutputStream(file)) {
            workbook.write(out);
        } catch (IOException e) {
            e.printStackTrace();
        }
        System.out.println("Excel файл успешно создан!");

    }

    @Override
    public String toString() {
        return "BkExcel{" +
                "plc='" + plc + '\'' +
                ", lSchet='" + lSchet + '\'' +
                ", KTP='" + KTP + '\'' +
                ", konc='" + konc + '\'' +
                ", street='" + street + '\'' +
                ", house='" + house + '\'' +
                ", apartment='" + apartment + '\'' +
                ", nSnFn='" + nSnFn + '\'' +
                ", model='" + model + '\'' +
                ", nCounter='" + nCounter + '\'' +
                ", energy='" + energy + '\'' +
                ", kWth='" + kWth + '\'' +
                ", t1='" + t1 + '\'' +
                ", t1v='" + t1v + '\'' +
                ", t2='" + t2 + '\'' +
                ", t2v='" + t2v + '\'' +
                ", t3='" + t3 + '\'' +
                ", t3v='" + t3v + '\'' +
                ", summ='" + summ + '\'' +
                '}';
    }

    public String getPlc() {
        return plc;
    }

    public void setPlc(String plc) {
        this.plc = plc;
    }

    public String getlSchet() {
        return lSchet;
    }

    public void setlSchet(String lSchet) {
        this.lSchet = lSchet;
    }

    public String getKTP() {
        return KTP;
    }

    public void setKTP(String KTP) {
        this.KTP = KTP;
    }

    public String getKonc() {
        return konc;
    }

    public void setKonc(String konc) {
        this.konc = konc;
    }

    public String getStreet() {
        return street;
    }

    public void setStreet(String street) {
        this.street = street;
    }

    public String getHouse() {
        return house;
    }

    public void setHouse(String house) {
        this.house = house;
    }

    public String getApartment() {
        return apartment;
    }

    public void setApartment(String apartment) {
        this.apartment = apartment;
    }

    public String getnSnFn() {
        return nSnFn;
    }

    public void setnSnFn(String nSnFn) {
        this.nSnFn = nSnFn;
    }

    public String getModel() {
        return model;
    }

    public void setModel(String model) {
        this.model = model;
    }

    public String getnCounter() {
        return nCounter;
    }

    public void setnCounter(String nCounter) {
        this.nCounter = nCounter;
    }

    public String getEnergy() {
        return energy;
    }

    public void setEnergy(String energy) {
        this.energy = energy;
    }

    public String getkWth() {
        return kWth;
    }

    public void setkWth(String kWth) {
        this.kWth = kWth;
    }

    public String getT1() {
        return t1;
    }

    public void setT1(String t1) {
        this.t1 = t1;
    }

    public String getT1v() {
        return t1v;
    }

    public void setT1v(String t1v) {
        this.t1v = t1v;
    }

    public String getT2() {
        return t2;
    }

    public void setT2(String t2) {
        this.t2 = t2;
    }

    public String getT2v() {
        return t2v;
    }

    public void setT2v(String t2v) {
        this.t2v = t2v;
    }

    public String getT3() {
        return t3;
    }

    public void setT3(String t3) {
        this.t3 = t3;
    }

    public String getT3v() {
        return t3v;
    }

    public void setT3v(String t3v) {
        this.t3v = t3v;
    }

    public String getSumm() {
        return summ;
    }

    public void setSumm(String summ) {
        this.summ = summ;
    }

    public BkExcel(String number, String plc, String lSchet, String KTP, String konc, String street, String house, String apartment, String nSnFn, String model, String nCounter, String energy, String kWth, String t1, String t1v, String t2, String t2v, String t3, String t3v, String summ) {

        this.number = number;
        this.plc = plc;
        this.lSchet = lSchet;
        this.KTP = KTP;
        this.konc = konc;
        this.street = street;
        this.house = house;
        this.apartment = apartment;
        this.nSnFn = nSnFn;
        this.model = model;
        this.nCounter = nCounter;
        this.energy = energy;
        this.kWth = kWth;
        this.t1 = t1;
        this.t1v = t1v;
        this.t2 = t2;
        this.t2v = t2v;
        this.t3 = t3;
        this.t3v = t3v;
        this.summ = summ;
    }
}
