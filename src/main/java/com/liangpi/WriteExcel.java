package com.liangpi;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;

import java.util.*;

public class WriteExcel {

//    private final static String rangeMonth = "（9.1-9.15）";
//    private final static String CommonTitle = "用货明细";
//    final static Integer start = 1;
//    final static Integer end = 15;
//    final static Integer month = 9;

    private final static String rangeMonth = "（8.16-8.31）";
    private final static String CommonTitle = "用货明细";
    final static Integer start = 16;
    final static Integer end = 31;
    final static Integer month = 8;



    public static void writeBlockCellRange(XSSFSheet sheet, int LeftOffset, int TopOffset, HashMap<String, HashMap<Integer, Double>> data, String name) {
        processTitle(sheet, LeftOffset, TopOffset, name);
        processBody(sheet, LeftOffset, TopOffset, data);
        processAllCount(sheet, LeftOffset, TopOffset, data, end, start, name);
    }

    //计算总量，且生成价格
    private static void processAllCount(XSSFSheet sheet, int leftOffset, int topOffset, HashMap<String, HashMap<Integer, Double>> data, Integer end, Integer begin, String name) {
        int start = topOffset + end - begin + 4;
        XSSFRow currentRow1 = sheet.getRow(start);
        if (currentRow1 == null) {
            currentRow1 = sheet.createRow(start);
        }

        XSSFRow currentRow3 = sheet.getRow(start + 2);
        if (currentRow3 == null) {
            currentRow3 = sheet.createRow(start + 2);
        }
        XSSFRow currentRow2 = sheet.getRow(start + 1);
        if (currentRow2 == null) {
            currentRow2 = sheet.createRow(start + 1);
        }

        HashMap<Integer, Double> piNum = data.get("皮");
        HashMap<Integer, Double> niuNum = data.get("牛");


        //合计/斤
        XSSFCell cell0 = currentRow1.createCell(leftOffset);
        XSSFCellStyle cellStyle = getCellStyle(sheet);
        XSSFWorkbook workbook = sheet.getWorkbook();
        XSSFFont font = workbook.createFont();
        font.setBold(true);
        font.setFontName("宋体");
        font.setColor(Font.COLOR_RED);
        cellStyle.setFont(font);
        cell0.setCellValue("合计/斤");
        cell0.setCellStyle(cellStyle);

        //pi
        double piSum = getSum(piNum);
        XSSFCell cell1 = currentRow1.createCell(leftOffset + 1);
        cell1.setCellValue(piSum);
        cell1.setCellStyle(cellStyle);

        //niu
        double niuSum = getSum(niuNum);
        XSSFCell cell2 = currentRow1.createCell(leftOffset + 2);
        cell2.setCellValue(niuSum);
        cell2.setCellStyle(cellStyle);

        XSSFCell cell3 = currentRow1.createCell(leftOffset + 3);
        cell3.setCellStyle(cellStyle);


        //获取单价
        HashMap<String, String> priceMap = Price.priceMap;
        String prices = priceMap.get(name);
        String[] split = prices.split(",");


        //每一样多少钱

        XSSFCell cell20 = currentRow2.createCell(0 + leftOffset);
        XSSFCell cell21 = currentRow2.createCell(1 + leftOffset);
        XSSFCell cell22 = currentRow2.createCell(2 + leftOffset);
        XSSFCell cell23 = currentRow2.createCell(3 + leftOffset);
        cell20.setCellStyle(cellStyle);
        cell21.setCellStyle(cellStyle);
        cell22.setCellStyle(cellStyle);
        cell23.setCellStyle(cellStyle);

        cell20.setCellValue("金额/元");
        cell21.setCellValue(Double.parseDouble(split[0]) * piSum);
        cell22.setCellValue(Double.parseDouble(split[1]) * niuSum);


        //金额总计

        XSSFCell cell30 = currentRow3.createCell(leftOffset + 0);
        XSSFCell cell31 = currentRow3.createCell(leftOffset + 1);
        XSSFCell cell32 = currentRow3.createCell(leftOffset + 2);
        XSSFCell cell33 = currentRow3.createCell(leftOffset + 3);

        cell30.setCellValue("金额总计");
        cell30.setCellStyle(cellStyle);

        //设置钱数

        double money = Double.parseDouble(split[0]) * piSum + Double.parseDouble(split[1]) * niuSum;
        cell31.setCellStyle(cellStyle);
        cell31.setCellValue(money);

        cell32.setCellStyle(cellStyle);
        cell33.setCellStyle(cellStyle);

    }

    private static double getSum(HashMap<Integer, Double> numMap) {
        double sum = 0;
        Set<Map.Entry<Integer, Double>> entries = numMap.entrySet();
        for (Map.Entry<Integer, Double> entry : entries) {
            sum += entry.getValue();
        }
        return sum;
    }

    private static void processBody(XSSFSheet sheet, int leftOffset, int topOffset, HashMap<String, HashMap<Integer, Double>> data) {


        //处理皮牛备注
        int topOffset2 = 2 + topOffset;//offset2是相对于秋林偏移→皮，牛，备注
        int leftOffset0 = 0 + leftOffset;//空格
        int leftOffset1 = 1 + leftOffset;//皮
        int leftOffset2 = 2 + leftOffset;//牛
        int leftOffset3 = 3 + leftOffset;//备注
        XSSFRow row2 = sheet.getRow(topOffset2);
        if (Objects.isNull(row2)) {
            //如果皮牛为空则生成一个row
            row2 = sheet.createRow(topOffset2);
        }

        XSSFCellStyle cellStyle = getCellStyle(sheet);

        //创建col并设置格式
        XSSFCell cell0 = row2.createCell(leftOffset0);
        cell0.setCellStyle(cellStyle);
        XSSFCell cell1 = row2.createCell(leftOffset1);
        cell1.setCellStyle(cellStyle);
        XSSFCell cell2 = row2.createCell(leftOffset2);
        cell2.setCellStyle(cellStyle);
        XSSFCell cell3 = row2.createCell(leftOffset3);
        cell3.setCellStyle(cellStyle);

        cell1.setCellValue("皮");
        cell2.setCellValue("牛");
        cell3.setCellValue("备注");


        //真正处理数据

        HashMap<Integer, Double> piMap = data.get("皮");
        HashMap<Integer, Double> niuMap = data.get("牛");

        setDate(sheet, start, end, leftOffset, topOffset);
        setPi(sheet, start, end, leftOffset, topOffset, piMap, niuMap);


    }

    private static void setPi(XSSFSheet sheet, Integer start, Integer end, int leftOffset, int topOffset, HashMap<Integer, Double> piMap, HashMap<Integer, Double> niuMap) {
        int topOffset3 = 3 + topOffset;//例如从8月16日开始，上下偏移就是3
        int leftOffset1 = 1 + leftOffset;
        int leftOffset2 = 2 + leftOffset;
        int leftOffset3 = 3 + leftOffset;
        XSSFCellStyle cellStyle = getCellStyle(sheet);
        for (int i = start; i <= end; i++) {
            XSSFRow currentRow = sheet.getRow(topOffset3);
            if (currentRow == null) {
                currentRow = sheet.createRow(topOffset3);
            }
            //设置pi的数量
            XSSFCell cell1 = currentRow.createCell(leftOffset1);
            cell1.setCellStyle(cellStyle);
            Double countPi = piMap.get(i);
            processCellValue(cell1, countPi);

            //设置niu的数量
            XSSFCell cell2 = currentRow.createCell(leftOffset2);
            cell2.setCellStyle(cellStyle);
            Double countNiu = niuMap.get(i);
            processCellValue(cell2, countNiu);

            XSSFCell cell3 = currentRow.createCell(leftOffset3);
            cell3.setCellStyle(cellStyle);

            topOffset3++;


        }

    }

    private static void processCellValue(XSSFCell cell1, Double count) {
        if (count == null || count < 0.0001) {
            cell1.setCellValue("——");
        } else {
            cell1.setCellValue(count);
        }
    }


    private static void setDate(XSSFSheet sheet, Integer start, Integer end, int leftOffset, int topOffset) {
        int topOffest3 = 3 + topOffset;
        XSSFCellStyle cellStyle = getCellStyle(sheet);
        //start为16，end为31
        for (int i = start; i <= end; i++) {
            XSSFRow currentRow = sheet.getRow(topOffest3);
            if (currentRow == null) {
                currentRow = sheet.createRow(topOffest3);
            }
            XSSFCell cell = currentRow.createCell(leftOffset + 0);
            cell.setCellStyle(cellStyle);
            cell.setCellValue(month + "月" + i + "日");
            topOffest3++;
        }

    }


    private static XSSFCellStyle getCellStyle(XSSFSheet sheet) {
        XSSFWorkbook workbook = sheet.getWorkbook();
        XSSFCellStyle cellStyle = workbook.createCellStyle();
        XSSFFont font = workbook.createFont();
        font.setFontName("宋体");
        font.setBold(true);
        cellStyle.setBorderTop(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBorderLeft(BorderStyle.THIN);

        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        cellStyle.setFont(font);

        return cellStyle;
    }

    private static void processTitle(XSSFSheet sheet, int LeftOffset, int TopOffset, String name) {
        CellRangeAddress cellAddresses = new CellRangeAddress(0 + TopOffset, 1 + TopOffset, 0 + LeftOffset, 3 + LeftOffset);
        XSSFCellStyle cellStyle = getCellStyle(sheet);

        sheet.addMergedRegion(cellAddresses);
        XSSFRow row0 = sheet.getRow(0 + TopOffset);
        if (row0 == null) {
            row0 = sheet.createRow(0 + TopOffset);
        }
        XSSFRow row1 = sheet.getRow(1 + TopOffset);
        if (row1 == null) {
            row1 = sheet.createRow(1 + TopOffset);
        }
        XSSFCell cell00 = row0.createCell(0 + LeftOffset);
        cell00.setCellValue(name + CommonTitle + rangeMonth);
        XSSFCell cell01 = row0.createCell(1 + LeftOffset);
        XSSFCell cell02 = row0.createCell(2 + LeftOffset);
        XSSFCell cell03 = row0.createCell(3 + LeftOffset);

        XSSFCell cel10 = row1.createCell(0 + LeftOffset);
        XSSFCell cel11 = row1.createCell(1 + LeftOffset);
        XSSFCell cel12 = row1.createCell(2 + LeftOffset);
        XSSFCell cel13 = row1.createCell(3 + LeftOffset);


        //设置title
        cell00.setCellStyle(cellStyle);
        cell01.setCellStyle(cellStyle);
        cell02.setCellStyle(cellStyle);
        cell03.setCellStyle(cellStyle);

        cel10.setCellStyle(cellStyle);
        cel11.setCellStyle(cellStyle);
        cel12.setCellStyle(cellStyle);
        cel13.setCellStyle(cellStyle);
    }


}
