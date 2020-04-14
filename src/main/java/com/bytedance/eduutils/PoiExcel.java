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

    public static <T> ArrayList<List<? extends T>> readExcel(String path, Class clzz) {
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
    public static <T> void writeExcel(String path, List<T> list, Class<T> clzz) throws IOException {
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

//            if (fos != null) {
////                fos.close();
//            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static String getSuffix(String path) {
        String substring = path.substring(path.lastIndexOf(".") + 1);
        return substring;
    }

//    private static Logger logger = Logger.getLogger(PoiExcel.class);
//    private final static String xls = "xls";
//    private final static String xlsx = "xlsx";
//
//    /**
//     * 读入excel文件，解析后返回
//     *
//     * @param file
//     * @throws IOException
//     */
//    public static List<String[]> readExcel(MultipartFile file) throws IOException {
//        //检查文件
//        checkFile(file);
//        //获得Workbook工作薄对象
//        Workbook workbook = getWorkBook(file);
//        //创建返回对象，把每行中的值作为一个数组，所有行作为一个集合返回
//        List<String[]> list = new ArrayList<String[]>();
//        if (workbook != null) {
//            for (int sheetNum = 0; sheetNum < workbook.getNumberOfSheets(); sheetNum++) {
//                //获得当前sheet工作表
//                Sheet sheet = workbook.getSheetAt(sheetNum);
//                if (sheet == null) {
//                    continue;
//                }
//                //获得当前sheet的开始行
//                int firstRowNum = sheet.getFirstRowNum();
//                //获得当前sheet的结束行
//                int lastRowNum = sheet.getLastRowNum();
//                //循环除了第一行的所有行
//                for (int rowNum = firstRowNum + 1; rowNum <= lastRowNum; rowNum++) {
//                    //获得当前行
//                    Row row = sheet.getRow(rowNum);
//                    if (row == null) {
//                        continue;
//                    }
//                    //获得当前行的开始列
//                    int firstCellNum = row.getFirstCellNum();
//                    //获得当前行的列数
//                    int lastCellNum = row.getPhysicalNumberOfCells();
//                    String[] cells = new String[row.getPhysicalNumberOfCells()];
//                    //循环当前行
//                    for (int cellNum = firstCellNum; cellNum < lastCellNum; cellNum++) {
//                        Cell cell = row.getCell(cellNum);
//                        cells[cellNum] = getCellValue(cell);
//                    }
//                    list.add(cells);
//                }
//            }
//            workbook.close();
//        }
//        return list;
//    }
//
//    public static void checkFile(MultipartFile file) throws IOException {
//        //判断文件是否存在
//        if (null == file) {
//            logger.error("文件不存在！");
//            throw new FileNotFoundException("文件不存在！");
//        }
//        //获得文件名
//        String fileName = file.getOriginalFilename();
//        //判断文件是否是excel文件
//        if (!fileName.endsWith(xls) && !fileName.endsWith(xlsx)) {
//            logger.error(fileName + "不是excel文件");
//            throw new IOException(fileName + "不是excel文件");
//        }
//    }
//
//    public static Workbook getWorkBook(MultipartFile file) {
//        //获得文件名
//        String fileName = file.getOriginalFilename();
//        //创建Workbook工作薄对象，表示整个excel
//        Workbook workbook = null;
//        try {
//            //获取excel文件的io流
//            InputStream is = file.getInputStream();
//            //根据文件后缀名不同(xls和xlsx)获得不同的Workbook实现类对象
//            if (fileName.endsWith(xls)) {
//                //2003
//                workbook = new HSSFWorkbook(is);
//            } else if (fileName.endsWith(xlsx)) {
//                //2007
//                workbook = new XSSFWorkbook(is);
//            }
//        } catch (IOException e) {
//            logger.info(e.getMessage());
//        }
//        return workbook;
//    }
//
//    public static String getCellValue(Cell cell) {
//        String cellValue = "";
//        if (cell == null) {
//            return cellValue;
//        }
//        //把数字当成String来读，避免出现1读成1.0的情况
//        if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
//            cell.setCellType(Cell.CELL_TYPE_STRING);
//        }
//        //判断数据的类型
//        switch (cell.getCellType()) {
//            case Cell.CELL_TYPE_NUMERIC: //数字
//                cellValue = String.valueOf(cell.getNumericCellValue());
//                break;
//            case Cell.CELL_TYPE_STRING: //字符串
//                cellValue = String.valueOf(cell.getStringCellValue());
//                break;
//            case Cell.CELL_TYPE_BOOLEAN: //Boolean
//                cellValue = String.valueOf(cell.getBooleanCellValue());
//                break;
//            case Cell.CELL_TYPE_FORMULA: //公式
//                cellValue = String.valueOf(cell.getCellFormula());
//                break;
//            case Cell.CELL_TYPE_BLANK: //空值
//                cellValue = "";
//                break;
//            case Cell.CELL_TYPE_ERROR: //故障
//                cellValue = "非法字符";
//                break;
//            default:
//                cellValue = "未知类型";
//                break;
//        }
//        return cellValue;
//    }
//

}