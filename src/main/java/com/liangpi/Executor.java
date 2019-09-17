package com.liangpi;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

public class Executor {
    private static String writePath;

    static {
        //可以配置
        writePath = "D:\\Java\\ReadAndExcel\\src\\main\\resources\\凉皮牛筋面" + WriteExcel.month + "月数据.xlsx";
    }

    public static void main(String[] args) throws IOException {
        LinkedHashMap<String, HashMap<String, HashMap<Integer, Double>>> allData = ReadExcelData();
        WriteData2Excel(allData);
    }

    private static void WriteData2Excel(LinkedHashMap<String, HashMap<String, HashMap<Integer, Double>>> allData) throws IOException {
        //写入数据

        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet("凉皮牛筋面月结数据");
        FileOutputStream fileOutputStream = new FileOutputStream(writePath);


        Set<Map.Entry<String, HashMap<String, HashMap<Integer, Double>>>> entries = allData.entrySet();
        //写入的时候控制上下左右
        int left = 0;
        int top = 0;
        int count = 0;
        for (Map.Entry<String, HashMap<String, HashMap<Integer, Double>>> entry : entries) {
            String name = entry.getKey();
            HashMap<String, HashMap<Integer, Double>> data = entry.getValue();
            WriteExcel.writeBlockCellRange(sheet, left, top, data, name);
            count++;
            if (count == 1) {
                left = left + 5;
            } else if (count == 2) {
                left += 4;
            } else if (count == 3) {
                left = left + 5;
            } else {
                count = 0;
                top = top + 27;
                left = 0;
            }
        }

        wb.write(fileOutputStream);
        fileOutputStream.close();
    }

    //读出数据
    private static LinkedHashMap<String, HashMap<String, HashMap<Integer, Double>>> ReadExcelData() throws IOException {
        ReadExcel readExcel = new ReadExcel();
        for (int i = WriteExcel.start; i <= WriteExcel.end; i++) {
            String path = Objects.requireNonNull(readExcel.getClass().getClassLoader().getResource(WriteExcel.month + "." + i + ".xlsx")).getPath();
            //更新日期
            readExcel.setDate(i);
            XSSFWorkbook workbook = new XSSFWorkbook(path);
            XSSFSheet sheetAt = workbook.getSheetAt(0);
            readExcel.readOneSheet(sheetAt);
            workbook.close();
        }
        return ReadExcel.getAllData();
    }
}
