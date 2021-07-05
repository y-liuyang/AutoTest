package com.liuyang.day2;

import org.testng.annotations.Test;

import java.io.File;

public class MyFile {
    //判断文件是否存在
    public static boolean fileExists(String filePath){
        File file = new File(filePath);
        Boolean bo = file.exists();
        return bo;
    }
    //创建文件
    public static void creatFile(String filePath){
        if (fileExists(filePath)==false){
            File file = new File(filePath);
            file.mkdir();
        }
    }
        @Test
    public void case1(){
        String filePath = "C:\\1.txt";
        creatFile(filePath);
    }
}
