package com.cc;

import org.apache.poi.hssf.eventusermodel.HSSFUserException;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joda.time.DateTime;
import org.junit.jupiter.api.Test;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;

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

    /**
     * 工具类
     */
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
                    if (cell != null) {
                        int cellType = cell.getCellType();
                        String cellValue = "";

                        switch (cellType) {
                            case HSSFCell.CELL_TYPE_STRING://字符串
                                System.out.println("【String】");
                                cellValue = cell.getStringCellValue();
                                break;
                            case HSSFCell.CELL_TYPE_BOOLEAN://bool
                                System.out.println("【BOOLEAN】");
                                cellValue = String.valueOf(cell.getBooleanCellValue());
                                break;
                            case HSSFCell.CELL_TYPE_BLANK://null
                                System.out.println("【BLANK】");
                                break;
                            case HSSFCell.CELL_TYPE_NUMERIC://数字（日期，普通数字）
                                System.out.println("【NUMERIC】");
                                if (HSSFDateUtil.isCellDateFormatted(cell)) { //日期
                                    System.out.println("【日期】");
                                    Date date = cell.getDateCellValue();
                                    cellValue = new DateTime(date).toString("yyyy-MM-dd");
                                } else {
                                    System.out.println("【转换为字符串输出】");
                                    cell.setCellType(HSSFCell.CELL_TYPE_STRING);
                                    cellValue = cell.toString();
                                }
                                break;
                            case HSSFCell.CELL_TYPE_ERROR://bool
                                System.out.println("【数据类型错误】");
                                break;
                        }
                        System.out.println(cellValue);
                    }
                }
            }
        }
        fileInputStream.close();
    }

    /**
     * 公式（了解即可）
     */
    @Test
    public void testFormula() throws IOException {
        FileInputStream fileInputStream = new FileInputStream(PATH + "更新公式.xls");

        Workbook workbook = new HSSFWorkbook(fileInputStream);

        Sheet sheet = workbook.getSheetAt(0);

        Row row = sheet.getRow(4);
        Cell cell = row.getCell(0);

        //拿到计算公式 eval
        FormulaEvaluator formulaEvaluator = new HSSFFormulaEvaluator((HSSFWorkbook) workbook);

        //输出单元格内容
        int cellType = cell.getCellType();
        switch (cellType) {
            case Cell.CELL_TYPE_FORMULA:
                String formula = cell.getCellFormula();
                System.out.println(formula);

                //计算
                CellValue evaluate = formulaEvaluator.evaluate(cell);
                String cellValue = evaluate.formatAsString();
                System.out.println(cellValue);
                break;
        }
    }
}
