import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

public class BkExcel {

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

    public static ArrayList<BkExcel> check (ArrayList<BkExcel> list){
        ArrayList<BkExcel> listOut = new ArrayList<BkExcel>();
        listOut.add(list.get(0));
        listOut.add(list.get(1));
        for (int i = 0; i <list.size() ; i++) {
            if (!list.get(i).t1v.equals("****") && !list.get(i).t1.equals("****") && list.get(i).energy.equals("****") &&
                    !list.get(i).t2v.equals("****") && !list.get(i).t2.equals("****") && !list.get(i).t3v.equals("****") && !list.get(i).t3.equals("****") ){
                listOut.add(list.get(i));

            }
        }
        return listOut;
    }
    public static void writeIntiExcel(String file, ArrayList<BkExcel> list){
        Workbook workbook = new HSSFWorkbook();
        Sheet sheet = workbook.createSheet();
        list = BkExcel.check(list);
        for (int i = 0; i < list.size(); i++) {

            Row row = sheet.createRow(i);
            Cell plc = row.createCell(0);
            Cell lSchet = row.createCell(1);
            Cell KTP = row.createCell(2);
            Cell konc = row.createCell(3);
            Cell street = row.createCell(4);
            Cell house = row.createCell(5);
            Cell apartment = row.createCell(6);
            Cell nSnFn = row.createCell(7);
            Cell model = row.createCell(8);
            Cell nCounter = row.createCell(9);
            Cell energy = row.createCell(10);
            Cell kWth = row.createCell(11);
            Cell t1 = row.createCell(12);
            Cell t1v = row.createCell(13);
            Cell t2 = row.createCell(14);
            Cell t2v = row.createCell(15);
            Cell t3 = row.createCell(16);
            Cell t3v = row.createCell(17);
            Cell summ = row.createCell(18);

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

    public BkExcel(String plc, String lSchet, String KTP, String konc, String street, String house, String apartment, String nSnFn, String model, String nCounter, String energy, String kWth, String t1, String t1v, String t2, String t2v, String t3, String t3v, String summ) {

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
