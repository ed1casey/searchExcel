package com.bankapp;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.*;

public class ExcelComparator {

    public static void main(String[] args) {
        final long startTime = System.currentTimeMillis();

        try {
            Map<String, String> mainData = readMainFile("C:/Users/edcasey/Downloads/Telegram Desktop/возвраты_согласован.xlsx");
            Map<String, String> secondData = readSecondFile("C:/Users/edcasey/Downloads/Telegram Desktop/17.xls");

            compareData(mainData, secondData);

        } catch (IOException e) {
            e.printStackTrace();
        }

        final long elapsedTimeMillis = System.currentTimeMillis() - startTime;
        System.out.println(elapsedTimeMillis + "ms");
    }

    public static Map<String, String> readMainFile(String filePath) throws IOException {
        Map<String, String> mainData = new HashMap<>();
        FileInputStream fileInputStream = new FileInputStream(filePath);
        Workbook workbook = getWorkbook(fileInputStream, filePath);
        Sheet sheet = workbook.getSheetAt(0);

        // Пропускаем заголовок
        Iterator<Row> rowIterator = sheet.iterator();
        if (rowIterator.hasNext()) {
            rowIterator.next();
        }

        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            String zakupka = getCellValueAsString(row.getCell(0)); // закупка
            String partiya = getCellValueAsString(row.getCell(1)); // партия
            mainData.put(zakupka, partiya);
        }

        workbook.close();
        fileInputStream.close();
        return mainData;
    }

    public static Map<String, String> readSecondFile(String filePath) throws IOException {
        Map<String, String> secondData = new HashMap<>();
        FileInputStream fileInputStream = new FileInputStream(filePath);
        Workbook workbook = getWorkbook(fileInputStream, filePath);
        Sheet sheet = workbook.getSheetAt(0);

        Iterator<Row> rowIterator = sheet.iterator();
        if (rowIterator.hasNext()) {
            rowIterator.next();
        }

        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            String nomer = getCellValueAsString(row.getCell(0));
            String nomerZakupkiAksapta = getCellValueAsString(row.getCell(1));
            secondData.put(nomerZakupkiAksapta, nomer);
        }

        workbook.close();
        fileInputStream.close();
        return secondData;
    }

    public static Workbook getWorkbook(FileInputStream fileInputStream, String filePath) throws IOException {
        if (filePath.endsWith(".xlsx")) {
            return new XSSFWorkbook(fileInputStream);
        } else if (filePath.endsWith(".xls")) {
            return new HSSFWorkbook(fileInputStream);
        } else {
            throw new IllegalArgumentException("Неподдерживаемый формат файла: " + filePath);
        }
    }

    private static String getCellValueAsString(Cell cell) {
        if (cell == null) {
            return "";
        }
        DataFormatter formatter = new DataFormatter();
        return formatter.formatCellValue(cell).trim();
    }

    public static void compareData(Map<String, String> mainData, Map<String, String> secondData) {
        for (Map.Entry<String, String> entry : mainData.entrySet()) {
            String zakupka = entry.getKey();
            String partiyaMain = entry.getValue();

            String nomerFromSecondFile = secondData.get(zakupka);

            if (nomerFromSecondFile != null) {
                if (!partiyaMain.equals(nomerFromSecondFile)) {
                    System.out.println("Закупка " + zakupka + " не совпадает со вторым файлом. Партия в аксапте: " + partiyaMain + ", Номер партии из 1С: " + nomerFromSecondFile);
                }
            } else {
                System.out.println("Закупка " + zakupka + " отсутствует во втором файле.");
            }
        }
    }
}
