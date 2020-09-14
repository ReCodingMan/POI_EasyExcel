package com.cc;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.joda.time.DateTime;
import org.junit.jupiter.api.DynamicTest;
import org.junit.jupiter.api.Test;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

public class ExcelWriteTest {

    String PATH = "./";

    /**
     * 03版本
     */
    @Test
    public void testWrite03() throws IOException {
        //1、创建y一个工作簿 03
        Workbook workbook = new HSSFWorkbook();

        //2、创建一个工作表
        Sheet sheet = workbook.createSheet("橙子观众统计表");

        //3、创建一个行
        Row row1 = sheet.createRow(0);

        //4、创建一个单元格
        Cell cell1 = row1.createCell(0);
        cell1.setCellValue("今日新增观众");

        Cell cell2 = row1.createCell(1);
        cell2.setCellValue(666);

        //第二行
        Row row2 = sheet.createRow(1);
        Cell cell21 = row2.createCell(0);
        cell21.setCellValue("统计时间");

        Cell cell22 = row2.createCell(1);
        String time = new DateTime().toString("yyyy-MM-dd HH:mm:ss");
        cell22.setCellValue(time);

        //生成一张表（IO）,03版本使用 xls 结尾
        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "橙子统计表03.xls");

        workbook.write(fileOutputStream);

        //关闭流
        fileOutputStream.close();

        System.out.println("Excel生成关闭");
    }

    /**
     * 07版本
     */
    @Test
    public void testWrite07() throws IOException {
        //1、创建y一个工作簿 07
        Workbook workbook = new HSSFWorkbook();

        //2、创建一个工作表
        Sheet sheet = workbook.createSheet("橙子观众统计表");

        //3、创建一个行
        Row row1 = sheet.createRow(0);

        //4、创建一个单元格
        Cell cell1 = row1.createCell(0);
        cell1.setCellValue("今日新增观众");

        Cell cell2 = row1.createCell(1);
        cell2.setCellValue(666);

        //第二行
        Row row2 = sheet.createRow(1);
        Cell cell21 = row2.createCell(0);
        cell21.setCellValue("统计时间");

        Cell cell22 = row2.createCell(1);
        String time = new DateTime().toString("yyyy-MM-dd HH:mm:ss");
        cell22.setCellValue(time);

        //生成一张表（IO）,03版本使用 xlsx 结尾
        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "橙子统计表07.xlsx");

        workbook.write(fileOutputStream);

        //关闭流
        fileOutputStream.close();

        System.out.println("Excel生成关闭");
    }
}
