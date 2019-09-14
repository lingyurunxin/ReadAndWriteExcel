package com.liangpi;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import java.util.*;


public class ReadExcel {
    //String为用户的名字，HashMap为客户的牛筋面和凉皮，字段只有这两个，每样都是一个月的数量
    //private static LinkedHashMap<String, HashMap<String, List<Double>>> allData = new LinkedHashMap<String, HashMap<String, List<Double>>>();

    //解释：String为客户名，HashMap存客户每个月牛和皮各多少斤，里层的HashMap存的是日期和多少斤
    private static LinkedHashMap<String, HashMap<String, HashMap<Integer,Double>>> allData = new LinkedHashMap<>();

    //偏移量可以配置，这个offset指的是例如包二相对于香榭丽col的偏移
    private int offset = 5;
    private int leftTop = 4;
    private int leftBottom = 30;
    private int rightTop = 4;
    private int rightBottom = 25;
    private static Integer date = 1;//每次读取的时候要更新


    //存储例如车这种奇怪的数据，String为客户的名字，map里面存的变量：开始row，结束row，开始column，结束column，凉皮和，牛筋面和，用map存放
    private static LinkedHashMap<String, HashMap<String, Object>> blockMap = new LinkedHashMap<String, HashMap<String, Object>>();

    //奇形怪状的Block不是随便装的，必须由指定的初始化
    private static ArrayList<String> specialBlockList = new ArrayList<String>(Arrays.asList("车"));

    public void readOneSheet(XSSFSheet sheet) {
        //处理例如车这种不规则客户，处理完之后，不规则客户的数据进入了allData里，并且可以获取不规则数据位置
        readBlockExcel(sheet);
        readExcel(sheet, leftTop, leftBottom, 0);
        readExcel(sheet, rightTop, rightBottom, offset);
    }

    /**
     * @param sheet  传入的sheet
     * @param top    例如香榭丽从第五行开始，这时候传入的是4
     * @param bottom 例如102是从31行结束，这时候传入的是30
     * @param offset 偏移，为了复用算法，例如包二，这时候相对香榭丽偏移则是6，香榭丽的偏移为0
     */
    public void readExcel(XSSFSheet sheet, int top, int bottom, int offset) {

        //处理完之后blockMap初始化完毕，不规则数据写入了block里
        for (int i = top; i <= bottom; i++) {
            XSSFRow row = sheet.getRow(i);

            XSSFCell nameCell = row.getCell(offset);
            String name = processCellName(nameCell);

            //如果blockMap不包含类似于车这种不规则客户，则不需要处理任何事情，大胆的读取
            if (!blockMap.containsKey(name)) {
                XSSFCell piCell = row.getCell(offset + 1);
                XSSFCell niuCell = row.getCell(offset + 2);
                double piNum = processCelltoDouble(piCell);
                double niuNum = processCelltoDouble(niuCell);
                processAllData(name, piNum, niuNum);
            } else {
                HashMap<String, Object> dataMap = blockMap.get(name);
                i = (Integer) dataMap.get("lastRow");
            }
        }

    }


    public void readBlockExcel(XSSFSheet sheet) {
        List<CellRangeAddress> mergedRegions = sheet.getMergedRegions();
        for (CellRangeAddress mergedRegion : mergedRegions) {

            int firstColumn = mergedRegion.getFirstColumn();
            int firstRow = mergedRegion.getFirstRow();
            int lastColumn = mergedRegion.getLastColumn();
            int lastRow = mergedRegion.getLastRow();
            String name = "";//例如车name
            XSSFRow row = sheet.getRow(firstRow);
            XSSFCell cell = row.getCell(firstColumn);
            name = processCellName(cell);
            if (specialBlockList.contains(name)) {
                HashMap<String, Object> map = new HashMap<String, Object>();
                map.put("firstColumn", firstColumn);
                map.put("firstRow", firstRow);
                map.put("lastColumn", lastColumn);
                map.put("lastRow", lastRow);
                blockMap.put(name, map);
                processBlockMap(sheet, name);
            }
        }
    }

    private String processCellName(XSSFCell cell) {
        String name = "";
        CellType cellType = cell.getCellType();
        if (CellType.NUMERIC.equals(cellType)) {
            name = cell.getRawValue();
        } else if (CellType.STRING.equals(cellType)) {
            name = cell.getStringCellValue();
        }
        return name;
    }

    //算出凉皮特殊块的凉皮和和牛筋面和,例如车name
    public void processBlockMap(XSSFSheet sheet, String name) {
        HashMap<String, Object> stringIntegerHashMap = blockMap.get(name);
        Integer firstRow = (Integer) stringIntegerHashMap.get("firstRow");
        Integer lastRow = (Integer) stringIntegerHashMap.get("lastRow");
        Integer firstColumn = (Integer) stringIntegerHashMap.get("firstColumn");

        double niuSum = 0;
        double piSum = 0;
        for (int i = firstRow; i <= lastRow; i++) {

            XSSFRow row = sheet.getRow(i);
            //这里的firstColumn相当于车的列，由于车为1列，所以其和lastColumn一样；
            XSSFCell piCell = row.getCell(firstColumn + 1);
            XSSFCell niuCell = row.getCell(firstColumn + 2);
            piSum += processCelltoDouble(piCell);
            niuSum += processCelltoDouble(niuCell);
        }
        HashMap<String, Object> data = blockMap.get(name);
        data.put("皮", piSum);
        data.put("牛", niuSum);
        processAllData(name, piSum, niuSum);
    }

    private void processAllData(String name, Double piSum, Double niuSum) {
        if (allData.containsKey(name)) {
            //String为皮，HashMap里装的是日期和数量
            HashMap<String, HashMap<Integer, Double>> dataMap = allData.get(name);
            HashMap<Integer, Double> piMap = dataMap.get("皮");
            HashMap<Integer, Double> niuMap = dataMap.get("牛");
            piMap.put(date,piSum);
            niuMap.put(date,niuSum);
        } else {
            HashMap<String, HashMap<Integer, Double>> dataMap = new HashMap<String, HashMap<Integer, Double>>();
            HashMap<Integer, Double> piMap = new HashMap<Integer, Double>();
            HashMap<Integer, Double> niuMap = new HashMap<Integer, Double>();
            piMap.put(date, piSum);
            niuMap.put(date,niuSum);
            dataMap.put("皮",piMap);
            dataMap.put("牛",niuMap);
            allData.put(name,dataMap);
        }
    }

    private double processCelltoDouble(XSSFCell cell) {
        CellType cellType = cell.getCellType();
        if (CellType.STRING.equals(cellType)) {
            //如果是String，很可能是——
            return 0;
        } else if (CellType.NUMERIC.equals(cellType)) {
            return cell.getNumericCellValue();
        } else if (CellType.BLANK.equals(cellType)) {
            return 0;
        } else {
            System.out.println("cellType:" + cellType + ";cellRow:" + cell.getRow().getRowNum() + ";cellCol:" + cell.getRowIndex());
            return 0;
        }
    }


    public static LinkedHashMap<String, HashMap<String, HashMap<Integer, Double>>> getAllData() {
        return allData;
    }

    public void setDate(Integer date) {
        this.date = date;
    }

    public Integer getDate() {
        return date;
    }

    public static void setAllData(LinkedHashMap<String, HashMap<String, HashMap<Integer, Double>>> allData) {
        ReadExcel.allData = allData;
    }



}
