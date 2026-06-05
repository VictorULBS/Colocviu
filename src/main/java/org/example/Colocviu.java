package org.example;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.select.Elements;

import java.io.*;
import java.util.*;

public class Colocviu {

    public static List<String> fetchItems(String url, String itemClass){
        List<String> items = new ArrayList<>();
        try {
            Document doc = Jsoup.connect(url).get();

            Elements menuClassC = doc.select(itemClass);
            menuClassC.forEach( element -> items.add(element.text()));

        } catch (IOException e) {
            throw new RuntimeException(e);
        }
        System.out.println("Read success");
        return items;
    }

    public static void writeToXLS(List<String> items, String path){
        try {
            FileOutputStream out = new FileOutputStream(path + ".xlsx", false);
            Workbook workbook = new XSSFWorkbook();
            Sheet sheet = workbook.createSheet("output");
            int i = 0;
            for(String item : items) {
                Row row = sheet.createRow(i++);
                row.createCell(0).setCellValue(item);
            }
            workbook.write(out);
        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
        System.out.println("Write success");
    }

    public static void main(String[] args) {

        List<String> items = fetchItems("https://ulbsibiu.ro/", ".nav-link");
        writeToXLS(items, args[0]);
    }
}