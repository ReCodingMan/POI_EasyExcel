package com.cc;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joda.time.DateTime;
import org.junit.jupiter.api.Test;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class ExcelReadTest {

    String PATH = "./";

    /**
     * 03版本读
     */
    @Test
    public void testRead03() throws IOException {

        //获取文件流
        FileInputStream fileInputStream = new FileInputStream(PATH + "橙子统计表03.xls");

        //1、创建一个工作簿 03，使用Excel能操作的这边都可以操作！
        Workbook workbook = new HSSFWorkbook(fileInputStream);

        //2、得到表
        Sheet sheet = workbook.getSheetAt(0);
        //3、得到行
        Row row = sheet.getRow(0);
        //4、得到列
        Cell cell = row.getCell(0);

        System.out.println(cell.getStringCellValue());
        fileInputStream.close();
    }

    /**
     * 07版本
     */
    @Test
    public void testRead07() throws IOException {

        //获取文件流
        FileInputStream fileInputStream = new FileInputStream(PATH + "橙子统计表07.xlsx");

        //1、创建y一个工作簿 07
        Workbook workbook = new XSSFWorkbook(fileInputStream);

        //2、得到表
        Sheet sheet = workbook.getSheetAt(0);
        //3、得到行
        Row row = sheet.getRow(0);
        //4、得到列
        Cell cell = row.getCell(1);

        System.out.println(cell.getNumericCellValue());
        fileInputStream.close();
    }
}
