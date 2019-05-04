package com.dascalitas;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
import org.xml.sax.SAXException;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.xpath.XPath;
import javax.xml.xpath.XPathConstants;
import javax.xml.xpath.XPathExpressionException;
import javax.xml.xpath.XPathFactory;
import java.io.*;
import java.net.URL;
import java.net.URLConnection;
import java.util.ArrayList;
import java.util.Scanner;

public class CSVtoXMLtoPOI {
    public static void main(String[] args) {

        Scanner scanner = null;
        try {
            scanner = new Scanner(new File("value.csv"));
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }

        //Set the delimiter used in file
        scanner.useDelimiter(",");

        ArrayList date = new ArrayList();

        while (scanner.hasNext()) {
            date.add(scanner.next());
        }

        for (int i = 0; i < date.size(); i++) {

            Workbook wb = new HSSFWorkbook();

            try {
                URL site = new URL("https://bnm.md/en/official_exchange_rates?get_xml=1&date=13.02.2018");
                URLConnection connect = site.openConnection();
                DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();
                DocumentBuilder dBuilder;

                dBuilder = dbFactory.newDocumentBuilder();

                Document doc = dBuilder.parse(connect.getInputStream());
                doc.getDocumentElement().normalize();

                XPath xPath = XPathFactory.newInstance().newXPath();

                String expression = "//ValCurs[@Date = " + date.get(i) + "]";
                NodeList nodeList = (NodeList) xPath.compile(expression).evaluate(
                        doc, XPathConstants.NODESET);

                for (int j = 0; j < nodeList.getLength(); j++) {
                    Node nNode = nodeList.item(j);

                    if (nNode.getNodeType() == Node.ELEMENT_NODE) {
                        Element eElement = (Element) nNode;

                        Sheet sheet = wb.createSheet(eElement.getAttribute("Date") + " - " + eElement.getAttribute("name"));

                        Row row = sheet.createRow(0);
                        row.createCell(0).setCellValue("ID");
                        row.createCell(1).setCellValue("Name");
                        row.createCell(2).setCellValue("Num Code");
                        row.createCell(3).setCellValue("Value");
                        row.createCell(4).setCellValue("Char code");
                        row.createCell(5).setCellValue("Nominal");

                        NodeList moneyList = eElement.getElementsByTagName("Valute");

                        for (int val = 0; val < moneyList.getLength(); val++) {

                            Node valNode = moneyList.item(val);
                            Row row2 = sheet.createRow(val + 1);

                            if (valNode.getNodeType() == Node.ELEMENT_NODE) {
                                Element valElement = (Element) valNode;

                                row2.createCell(0).setCellValue(valElement.getAttribute("ID"));
                                row2.createCell(1).setCellValue(valElement
                                        .getElementsByTagName("Name")
                                        .item(0)
                                        .getTextContent());
                                row2.createCell(2).setCellValue(valElement
                                        .getElementsByTagName("NumCode")
                                        .item(0)
                                        .getTextContent());
                                row2.createCell(3).setCellValue(valElement
                                        .getElementsByTagName("Value")
                                        .item(0)
                                        .getTextContent());
                                row2.createCell(4).setCellValue(valElement
                                        .getElementsByTagName("CharCode")
                                        .item(0)
                                        .getTextContent());
                                row2.createCell(5).setCellValue(valElement
                                        .getElementsByTagName("Nominal")
                                        .item(0)
                                        .getTextContent());

                            }
                        }
                    }
                }
            } catch (ParserConfigurationException e) {
                e.printStackTrace();
            } catch (SAXException e) {
                e.printStackTrace();
            } catch (IOException e) {
                e.printStackTrace();
            } catch (XPathExpressionException e) {
                e.printStackTrace();
            }

            try (OutputStream fileOut = new FileOutputStream("value.xls")) {
                wb.write(fileOut);
            } catch (FileNotFoundException e) {
                e.printStackTrace();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }
}
