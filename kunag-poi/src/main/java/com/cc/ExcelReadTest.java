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

    @Test
    public void testCellType() throws IOException{
        //获取文件流
        FileInputStream fileInputStream = new FileInputStream(PATH + "更新公式.xls");

        //1、创建一个工作簿 03，使用Excel能操作的这边都可以操作！
        Workbook workbook = new HSSFWorkbook(fileInputStream);

        //获取标题内容
        Sheet sheet = workbook.getSheetAt(0);
        Row rowTitle = sheet.getRow(0);
        if (rowTitle != null) {
            // 一定要掌握
            int cellCount = rowTitle.getPhysicalNumberOfCells();
            for (int cellNum = 0; cellNum < cellCount; cellNum++) {
                Cell cell = rowTitle.getCell(cellNum);
                if (cell != null) {
                    int cellType = cell.getCellType();
                    String cellValue = cell.getStringCellValue();
                    System.out.print(cellValue + " | ");
                }
            }
            System.out.println();
        }

        // 获取表中内容
        int rowCount = sheet.getPhysicalNumberOfRows();
        for (int rowNum = 0; rowNum < rowCount; rowNum++) {
            Row rowData = sheet.getRow(rowNum);
            if (rowData != null) {
                //读取列
                int cellCount = rowTitle.getPhysicalNumberOfCells();
                for (int cellNum = 0; cellNum < cellCount; cellNum++) {
                    System.out.print("[" + (rowNum+1) + "-" + (cellNum+1) + "]");

                    Cell cell = rowData.getCell(cellNum);
                    //匹配列的数据类型
                }
            }
        }
    }
}
