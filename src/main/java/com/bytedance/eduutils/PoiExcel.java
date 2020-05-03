package com.bytedance.eduutils;

import java.io.*;
import java.lang.reflect.Field;
import java.text.SimpleDateFormat;
import java.util.*;

import com.bytedance.eduutils.entity.User;
import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.multipart.commons.CommonsMultipartFile;

/**
 * excel读写工具类
 *
 * @author sun.kai
 * 2016年8月21日
 */
public class PoiExcel {

    public <T> ArrayList<List<? extends T>> readExcel(String path, Class clzz) {
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
        List<T> list = new LinkedList<T>();
        ArrayList<List<? extends T>> fileData = new ArrayList<>();
        File file = new File(path);
        FileInputStream fis = null;
        Workbook workBook = null;
        if (file.exists()) {
            try {
                fis = new FileInputStream(file);
                workBook = WorkbookFactory.create(fis);
                int numberOfSheets = workBook.getNumberOfSheets();
                for (int s = 0; s < numberOfSheets; s++) { // sheet工作表
                    Sheet sheetAt = workBook.getSheetAt(s);
//                  String sheetName = sheetAt.getSheetName(); //获取工作表名称
                    int rowsOfSheet = sheetAt.getPhysicalNumberOfRows(); // 获取当前Sheet的总列数
                    System.out.println("当前表格的总行数:" + rowsOfSheet);
                    for (int r = 0; r < rowsOfSheet; r++) { // 总行
                        Row row = sheetAt.getRow(r);
                        if (row == null) {
                            continue;
                        } else {
                            int rowNum = row.getRowNum();
//                            System.out.println("当前行:" + rowNum);
                            int numberOfCells = row.getPhysicalNumberOfCells();
                            ArrayList<Object> list1 = new ArrayList<>();
                            for (int c = 0; c < numberOfCells; c++) { // 总列(格)
                                Cell cell = row.getCell(c);
                                if (cell == null) {
                                    continue;
                                } else {
                                    int cellType = cell.getCellType();
                                    switch (cellType) {
                                        case Cell.CELL_TYPE_STRING: // 代表文本
                                            String stringCellValue = cell.getStringCellValue();
                                            list1.add(stringCellValue);
//                                            System.out.print(stringCellValue + "\t");
                                            break;
                                        case Cell.CELL_TYPE_NUMERIC: // 数字||日期
                                            boolean cellDateFormatted = DateUtil.isCellDateFormatted(cell);
                                            if (cellDateFormatted) {
                                                Date dateCellValue = cell.getDateCellValue();
                                                System.out.print(sdf.format(dateCellValue) + "\t");
                                            } else {
                                                double numericCellValue = cell.getNumericCellValue();
//                                                System.out.print(numericCellValue + "\t");
                                                list1.add(numericCellValue);
                                            }
                                            break;
                                        case Cell.CELL_TYPE_BLANK: // 空白格
                                            String stringCellBlankValue = cell.getStringCellValue();
//                                            System.out.print(stringCellBlankValue + "\t");
                                            list1.add(stringCellBlankValue);
                                            break;
                                        case Cell.CELL_TYPE_BOOLEAN: // 布尔型
                                            boolean booleanCellValue = cell.getBooleanCellValue();
                                            list1.add(booleanCellValue);
//                                            System.out.print(booleanCellValue + "\t");
                                            break;

                                        case Cell.CELL_TYPE_ERROR: // 错误
                                            byte errorCellValue = cell.getErrorCellValue();
                                            System.out.print(errorCellValue + "\t");
                                            break;
                                        case Cell.CELL_TYPE_FORMULA: // 公式
                                            int cachedFormulaResultType = cell.getCachedFormulaResultType();
                                            System.out.print(cachedFormulaResultType + "\t");
                                            break;
                                    }
                                }
                            }
                            fileData.add((List<? extends T>) list1);
                            System.out.println(" \t ");
                        }
//                        System.out.println("");
                    }
                }
                if (fis != null) {
                    fis.close();
                }
            } catch (Exception e) {
                // TODO Auto-generated catch block
                e.printStackTrace();
            }
        } else {
            System.out.println("文件不存在!");
        }
        return fileData;
    }
//    public static <T> Map<Integer, List<? extends T>> readExcel(String path, Class clzz) {
//        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
//        List<T> list = new LinkedList<T>();
//        Map<Integer, List<? extends T>> map = new HashMap<Integer, List<? extends T>>();
//        File file = new File(path);
//        FileInputStream fis = null;
//        Workbook workBook = null;
//        if (file.exists()) {
//            try {
//                fis = new FileInputStream(file);
//                workBook = WorkbookFactory.create(fis);
//                int numberOfSheets = workBook.getNumberOfSheets();
//                for (int s = 0; s < numberOfSheets; s++) { // sheet工作表
//                    Sheet sheetAt = workBook.getSheetAt(s);
////                  String sheetName = sheetAt.getSheetName(); //获取工作表名称
//                    int rowsOfSheet = sheetAt.getPhysicalNumberOfRows(); // 获取当前Sheet的总列数
//                    System.out.println("当前表格的总行数:" + rowsOfSheet);
//                    for (int r = 0; r < rowsOfSheet; r++) { // 总行
//                        Row row = sheetAt.getRow(r);
//                        if (row == null) {
//                            continue;
//                        } else {
//                            int rowNum = row.getRowNum();
//                            System.out.println("当前行:" + rowNum);
//                            int numberOfCells = row.getPhysicalNumberOfCells();
//                            ArrayList<Object> list1 = new ArrayList<>();
//                            for (int c = 0; c < numberOfCells; c++) { // 总列(格)
//                                Cell cell = row.getCell(c);
//                                if (cell == null) {
//                                    continue;
//                                } else {
//                                    int cellType = cell.getCellType();
//                                    switch (cellType) {
//                                        case Cell.CELL_TYPE_STRING: // 代表文本
//                                            String stringCellValue = cell.getStringCellValue();
//                                            list1.add(stringCellValue);
////                                            System.out.print(stringCellValue + "\t");
//                                            break;
//                                        case Cell.CELL_TYPE_NUMERIC: // 数字||日期
//                                            boolean cellDateFormatted = DateUtil.isCellDateFormatted(cell);
//                                            if (cellDateFormatted) {
//                                                Date dateCellValue = cell.getDateCellValue();
//                                                System.out.print(sdf.format(dateCellValue) + "\t");
//                                            } else {
//                                                double numericCellValue = cell.getNumericCellValue();
////                                                System.out.print(numericCellValue + "\t");
//                                                list1.add(numericCellValue);
//                                            }
//                                            break;
//                                        case Cell.CELL_TYPE_BLANK: // 空白格
//                                            String stringCellBlankValue = cell.getStringCellValue();
////                                            System.out.print(stringCellBlankValue + "\t");
//                                            list1.add(stringCellBlankValue);
//                                            break;
//                                        case Cell.CELL_TYPE_BOOLEAN: // 布尔型
//                                            boolean booleanCellValue = cell.getBooleanCellValue();
//                                            list1.add(booleanCellValue);
////                                            System.out.print(booleanCellValue + "\t");
//                                            break;
//
//                                        case Cell.CELL_TYPE_ERROR: // 错误
//                                            byte errorCellValue = cell.getErrorCellValue();
//                                            System.out.print(errorCellValue + "\t");
//                                            break;
//                                        case Cell.CELL_TYPE_FORMULA: // 公式
//                                            int cachedFormulaResultType = cell.getCachedFormulaResultType();
//                                            System.out.print(cachedFormulaResultType + "\t");
//                                            break;
//                                    }
//                                }
//                                map.put(rowNum, (List<? extends T>) list1);
//                            }
//                            System.out.println(" \t ");
//                        }
//                        System.out.println("");
//                    }
//                }
//                if (fis != null) {
//                    fis.close();
//                }
//            } catch (Exception e) {
//                // TODO Auto-generated catch block
//                e.printStackTrace();
//            }
//        } else {
//            System.out.println("文件不存在!");
//        }
//        return map;
//    }


    @SuppressWarnings("resource")
    public <T> void writeExcel(String path, List<T> list, Class<T> clzz) throws IOException {
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
        Field[] declaredFields = clzz.getDeclaredFields();
        File file = new File(path);

        Workbook workbook = null;
        try {
            if (file.exists()) {
//                fos = new FileOutputStream(file);
                String suffix = getSuffix(path);
                if (suffix.equalsIgnoreCase("XLSX")) {
                    workbook = new XSSFWorkbook();
                } else if (suffix.equalsIgnoreCase("XLS")) {
                    workbook = new HSSFWorkbook();
                } else {
                    throw new Exception("当前文件不是excel文件");
                }
                Sheet sheet = workbook.createSheet(); // 生成工作表
                Row row = sheet.createRow(0);
                for (int m = 0; m < declaredFields.length; m++) { // 设置title
                    Field field = declaredFields[m];
                    field.setAccessible(true);
                    Cell cell = row.createCell(m);
                    String name = field.getName();
                    cell.setCellType(Cell.CELL_TYPE_STRING);
                    cell.setCellValue(name);
                }
                for (int i = 0; i < list.size(); i++) { // 数据
                    T instance = list.get(i);
                    row = sheet.createRow(i + 1);
                    for (int j = 0; j < declaredFields.length; j++) {
                        Field field = declaredFields[j];
                        Object value = field.get(instance);
                        String fieldTypeName = field.getGenericType().getTypeName();
                        field.setAccessible(true);
                        Cell cell = row.createCell(j);
                        switch (fieldTypeName) {// content
                            case "long":
                                double d = Double.valueOf(value.toString());
                                cell.setCellValue(d);
                                break;
                            case "float":
                                CellStyle floatStyle = workbook.createCellStyle();
                                short format = workbook.createDataFormat().getFormat(".00");// 保留2位精度
                                floatStyle.setDataFormat(format);
                                double d1 = Double.parseDouble(String.valueOf(value));
                                cell.setCellStyle(floatStyle);
                                cell.setCellValue(d1);
                                break;
                            case "int":
                                double d2 = Double.parseDouble(String.valueOf(value));
                                cell.setCellValue(d2);
                                break;
                            case "java.lang.Integer":
                                double d3 = Double.parseDouble(String.valueOf(value));
                                cell.setCellValue(d3);
                                break;
                            case "java.lang.String":
                                String s = value.toString();
                                cell.setCellValue(s);
                                break;
                            case "java.util.Date":
                                CellStyle dateStyle = workbook.createCellStyle();
                                short df = workbook.createDataFormat().getFormat("yyyy-mm-dd");
                                dateStyle.setDataFormat(df);
                                cell.setCellStyle(dateStyle);
                                String format2 = sdf.format(value);
                                Date date = sdf.parse(format2);
                                cell.setCellValue(date);
                                break;
                        }
                    }
                    FileOutputStream fos = new FileOutputStream(file);
                    workbook.write(fos);
                    fos.close();
                }
            } else {
                if (file.createNewFile()) {
                    writeExcel(path, list, clzz);
                } else {
                    System.out.println("创建Excel表格失败!");
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static String getSuffix(String path) {
        String substring = path.substring(path.lastIndexOf(".") + 1);
        return substring;
    }

}