package com.liuyang.day2;

import java.awt.*;
import java.io.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

public class readFileTest {
    @Test
    public static void test() throws IOException {
        String excelFilePath = "C:\\Users\\Administrator\\Desktop\\data.xlsx";
        File file = new File(excelFilePath);    //根据上面的路径文件，创建一个文件对象
        FileInputStream inputStream = new FileInputStream(file);  //用于读取excel文件
        Workbook workBook = null;
        String fileExtensionName = excelFilePath.substring(excelFilePath.indexOf("."));  //截取文件名，获取 . 后面（包括.）的内容
        System.out.println(fileExtensionName);  //打印内容.xlsx
        if (fileExtensionName.equals(".xlsx")) {
            workBook = new XSSFWorkbook(inputStream);
        } else if (fileExtensionName.equals(".xls")) {
            workBook = new HSSFWorkbook(inputStream);
        }
        System.out.println(fileExtensionName);  //打印内容.xlsx
        Sheet sheet = workBook.getSheetAt(0);    //通过sheetName参数，生成Sheet对象
        System.out.println(sheet.getFirstRowNum());   //获取数据的第一行行号    0  //excel的行号和列号都是从0 开始
        System.out.println(sheet.getLastRowNum());   //获取数据的最后一行行号
        int rows = sheet.getLastRowNum() + 1;
        int columns = sheet.getRow(0).getPhysicalNumberOfCells();
        System.out.println("行数为" + rows + "   列数为" + columns);
        //List list = new ArrayList();
        //list.add("wer");
        /*
        Label label = null;
        sheet.getRow(0).createCell(2).setCellValue("Pass");
        sheet.getRow(1).createCell(2).setCellValue("Fail");
        sheet.getRow(2).createCell(2).setCellValue("N/A");
        sheet.getRow(3).createCell(2).setCellValue("Pass");
        System.out.println("已写完");
        //sheet.getRow(0).createCell(1).setCellValue("222w");
        //label = new Label("we",0);
        XSSFRow row1 = (XSSFRow) sheet.getRow (4);
        if (row1 == null) {
            row1 = (XSSFRow) sheet.createRow(4);
        }
        if (row1 != null) {
            row1.createCell(0).setCellValue ("rt");
        }


        FileOutputStream fout = new FileOutputStream(new File(excelFilePath));

        //覆盖写入内容
        workBook.write(fout);

        // 关闭文件
        fout.close();
        */

        HashMap<String,String> map = new HashMap<>();

        for (int i = 0; i < rows; i++) {
            System.out.println(i);
            for (int j = 0; j < columns; j++) {
                System.out.println(i+"  "+j);
                System.out.println(sheet.getRow(i).getCell(j).getStringCellValue());
                //map.put(sheet.getRow(i).getCell(j).getStringCellValue(),sheet.getRow(i).getCell(j+1).getStringCellValue());
            }
            for (int m =0;m<map.size();m++){
                System.out.println(map.get(m)+": ");
            }
        }



        /*
        System.out.println(Sheet.getRow(0).getCell(0).getStringCellValue());
        System.out.println(Sheet.getRow(0).getCell(1).getStringCellValue());
        System.out.println(Sheet.getRow(0).getCell(2).getStringCellValue());

        System.out.println("------------------------------");

        //System.out.println(Sheet.getRow(1).getCell(0).getStringCellValue());   //此处执行会发送错误，因为是数字类型，不是文字类型
        System.out.println(Sheet.getRow(1).getCell(1).getStringCellValue());
        System.out.println(Sheet.getRow(1).getCell(2).getStringCellValue());

        System.out.println("------------------------------");

        int totalRows = Sheet.getPhysicalNumberOfRows();  //获取总行数

        System.out.println(totalRows);  //返回总行数7

        System.out.println("------------------------------");

        int RowCells = Sheet.getRow(0).getPhysicalNumberOfCells(); //获取总列数

        System.out.println(RowCells);   //返回总列数3

        System.out.println("------------------------------");


        System.out.println(Sheet.getRow(2).getCell(0).toString());  //表格里是2，这里返回2.0

        System.out.println("------------------------------");

        System.out.println(Sheet.getRow(2).getCell(0).getNumericCellValue()); //表格里是2，这里返回2.0


        // getStringCellValue
        //java.lang.String getStringCellValue（）
        //以字符串形式获取单元格的值
        //对于数字单元格，我们抛出异常。对于空白单元格，我们返回一个空字符串。对于不是字符串公式的FormulaCells，我们抛出异常。
        //返回值：单元格的值作为字符串//

        System.out.println("------------------------------");

        //Sheet.getRow(2).getCell(0).setCellType(Cell.CELL_TYPE_STRING);  //设置单元格为string类型
        System.out.println(Sheet.getRow(2).getCell(0).getStringCellValue()); //返回值为2
    */


    }

}
