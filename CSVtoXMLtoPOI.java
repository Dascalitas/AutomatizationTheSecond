package com.dascalitas;

import com.thoughtworks.xstream.XStream;
import com.thoughtworks.xstream.annotations.XStreamAlias;
import com.thoughtworks.xstream.annotations.XStreamAsAttribute;
import com.thoughtworks.xstream.annotations.XStreamImplicit;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import java.io.*;
import java.net.URL;
import java.util.ArrayList;
import java.util.List;
import java.util.Scanner;

@XStreamAlias("ValCurs")
class ValCurs {

    @XStreamAlias("Date")
    @XStreamAsAttribute
    private String date;

    @XStreamAlias("name")
    @XStreamAsAttribute
    private String name;

    @XStreamImplicit(itemFieldName="Valute")
    private List<Valute> valutes;

    public String getDate() {
            return date;
        }

    public void setDate(String date) {
        this.date = date;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public List<Valute> getValutes() {
        return valutes;
    }

    public void setValutes(List<Valute> valutes) {
        this.valutes = valutes;
    }
}

@XStreamAlias("Valute")
class Valute {

    @XStreamAlias("NumCode")
    private String numCode;
    @XStreamAlias("CharCode")
    private String сharCode;
    @XStreamAlias("Nominal")
    private int nominal;
    @XStreamAlias("Name")
    private String name;
    @XStreamAlias("Value")
    private double value;

    @XStreamAlias("ID")
    @XStreamAsAttribute
    private String id;

    @Override
    public String toString() {
        return "Valute{" +
                "numCode='" + numCode + '\'' +
                ", сharCode='" + сharCode + '\'' +
                ", nominal=" + nominal +
                ", name='" + name + '\'' +
                ", value=" + value +
                ", id='" + id + '\'' +
                '}';
    }

    public String getNumCode() {
        return numCode;
    }

    public void setNumCode(String numCode) {
        this.numCode = numCode;
    }

    public String getСharCode() {
        return сharCode;
    }

    public void setСharCode(String сharCode) {
        this.сharCode = сharCode;
    }

    public int getNominal() {
        return nominal;
    }

    public void setNominal(int nominal) {
        this.nominal = nominal;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public double getValue() {
        return value;
    }

    public void setValue(double value) {
        this.value = value;
    }

    public String getId() {
        return id;
    }

    public void setId(String id) {
        this.id = id;
    }
}

public class CSVtoXMLtoPOI {
    public static void main(String[] args) throws FileNotFoundException {
        XStream xstream = new XStream();
        xstream.processAnnotations(ValCurs.class);
        xstream.processAnnotations(Valute.class);
        DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();
        DocumentBuilder dBuilder;
        Workbook wb = new HSSFWorkbook();


         Scanner scanner = new Scanner(new File("value.csv"));


        //Set the delimiter used in file
        scanner.useDelimiter(",");

        ArrayList date = new ArrayList();

        while (scanner.hasNext()) {
            date.add(scanner.next());
        }

        for (int i = 0; i < date.size(); i++) {
            System.out.println(date.get(i));
            try {
                URL site = new URL("https://bnm.md/en/official_exchange_rates?get_xml=1&date=" + date.get(i));

                ValCurs valCurs = (ValCurs) xstream.fromXML(site);
                Sheet sheet = wb.createSheet(valCurs.getDate() + " - " + valCurs.getName());

                Row row = sheet.createRow(0);
                row.createCell(0).setCellValue("ID");
                row.createCell(1).setCellValue("Name");
                row.createCell(2).setCellValue("Num Code");
                row.createCell(3).setCellValue("Value");
                row.createCell(4).setCellValue("Char code");
                row.createCell(5).setCellValue("Nominal");
                int cell = 0;
                for (Valute currentVal : valCurs.getValutes()) {
                    Row row2 = sheet.createRow(cell + 1);

                    row2.createCell(0).setCellValue(currentVal.getId());
                    row2.createCell(1).setCellValue(currentVal.getName());
                    row2.createCell(2).setCellValue(currentVal.getNumCode());
                    row2.createCell(3).setCellValue(currentVal.getValue());
                    row2.createCell(4).setCellValue(currentVal.getСharCode());
                    row2.createCell(5).setCellValue(currentVal.getNominal());
                    cell++;

                    System.out.println("done\n");
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
            try (OutputStream fileOut = new FileOutputStream("value.xls")) {
                wb.write(fileOut);
                System.out.println("finish");
            } catch (FileNotFoundException e) {
                e.printStackTrace();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }
