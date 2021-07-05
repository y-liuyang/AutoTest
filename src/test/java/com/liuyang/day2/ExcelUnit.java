package com.liuyang.day2;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;

public class ExcelUnit {

    public static Workbook getWorkbook(String filePath){
        Workbook wk = null;
        try {
        if (filePath.endsWith("xls")) {
            File file = new File(filePath);
            InputStream io = new FileInputStream(file);
            wk = new HSSFWorkbook(io);
        }
        else if (filePath.endsWith("xlsx")){
            wk = new XSSFWorkbook(filePath);
        }
        }catch(IOException e){
            e.printStackTrace();
        }
        return  wk;

    }
    public static String getCellValue(Cell cell){
        String value = "";
        switch (cell.getCellType()){
            case STRING:
                value = String.valueOf(cell.getStringCellValue());
                return value;
            case NUMERIC:
                value = String.valueOf(cell.getNumericCellValue());
                return value;
        }
        return value;
    }
    @Test
    public void case23(){
        HashMap<String,String>[][] map = new HashMap[2][6];
        map[1][1].put("1","2");
        map[1][4].put("1","2");
        map[7][8].put("1","2");
        System.out.println(map);
    }

    public static String getCellValue(Sheet sheet,int row,int cell){
        Cell cell1 = sheet.getRow(row).getCell(cell);
        String value = ExcelUnit.getCellValue(cell1);
        return value;
    }
    public Object[][] testData(String filePath){
        ArrayList<String> list = new ArrayList<String>();

        Workbook wk = getWorkbook(filePath);
        Sheet sheet = wk.getSheetAt(0);
        //获取总行数
        int rows = sheet.getLastRowNum()+1;
        //列数
        int columns = sheet.getRow(0).getPhysicalNumberOfCells();
        HashMap<String,String>[][] map = new HashMap[1][1];
        for (int c = 0;c < columns;c++){
            String cellValue = getCellValue(sheet,0,c);
            list.add(cellValue);
        }
        for (int r = 1;r < rows;r++){
            for (int c = 0;c < columns;c++){
                String cellValue = getCellValue(sheet,r,c);

            }
        }
        return map;
    }

    @Test
    public void readFile(){
        String src = "C:\\Users\\Administrator\\Desktop\\1.xlsx";
        try {

            // 加载文件
            FileInputStream fis = new FileInputStream(src);

            // 加载workbook
            XSSFWorkbook wb = new XSSFWorkbook(fis);

            //加载sheet，这里我们只有一个sheet,默认是sheet1
            XSSFSheet sheet = wb.getSheetAt(0);


            int rows = sheet.getLastRowNum() + 1;
            //int columns = sheet.getColumns();
            //System.out.println("行数为" + rows + "   列数为" + columns);

        } catch (Exception e) {

            System.out.println(e.getMessage());

        }
    }
    }
