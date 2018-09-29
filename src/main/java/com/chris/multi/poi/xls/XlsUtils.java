package com.chris.multi.poi.xls;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;

import java.io.*;
import java.lang.reflect.Field;
import java.sql.Timestamp;
import java.time.Instant;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Set;

/**
 * Created by Chris Chen
 * 2018/09/17
 * Explain: Java操作Excel表格的工具
 */

public class XlsUtils {
    private static final int CHAR_WIDTH = 512;

    /**
     * 创建一个工作簿，并且添加一张表，写入数据
     *
     * @param xlsWorkSheetInfo
     * @param saveFileName
     * @param <T>
     * @return
     */
    public static <T> Boolean exportToXls(XlsWorkSheetInfo<T> xlsWorkSheetInfo, String saveFileName) {
        //创建一个工作簿
        HSSFWorkbook workbook = new HSSFWorkbook();
        //向工作簿天机按工作表
        addToXls(xlsWorkSheetInfo, workbook, null);
        return saveXlsFile(workbook, saveFileName);
    }

    /**
     * 导出到输出流OutputStream
     * 不包含设置回调
     *
     * @param xlsWorkSheetInfo
     * @param os
     * @param <T>
     * @return
     */
    public static <T> Boolean exportToXlsOutputStream(XlsWorkSheetInfo<T> xlsWorkSheetInfo, OutputStream os) {
        return exportToXlsOutputStream(xlsWorkSheetInfo, os, null);
    }

    /**
     * 导出到输出流OutputStream
     * 包含设置回调
     *
     * @param xlsWorkSheetInfo
     * @param os
     * @param <T>
     * @return
     */
    public static <T> Boolean exportToXlsOutputStream(XlsWorkSheetInfo<T> xlsWorkSheetInfo, OutputStream os, XlsSetupAdapter xlsSetupAdapter) {
        //创建一个工作簿
        HSSFWorkbook workbook = new HSSFWorkbook();
        if (xlsSetupAdapter != null) {
            xlsSetupAdapter.workBookSetup(workbook);
        }
        //向工作簿添加工作表
        addToXls(xlsWorkSheetInfo, workbook, xlsSetupAdapter);
        try {
            workbook.write(os);
            return true;
        } catch (IOException e) {
            e.printStackTrace();
        }
        return false;
    }

    /**
     * 把包含工作表信息的集合写入到xls表格
     *
     * @param xlsWorkSheetInfoSet
     * @param saveFileName
     * @return
     */
    public static Boolean exportToXls(Set<XlsWorkSheetInfo> xlsWorkSheetInfoSet, String saveFileName) {
        //创建一个工作簿
        HSSFWorkbook workbook = new HSSFWorkbook();
        //循环向工作簿添加工作表
        for (XlsWorkSheetInfo xlsWorkSheetInfo : xlsWorkSheetInfoSet) {
            addToXls(xlsWorkSheetInfo, workbook, null);
        }
        return saveXlsFile(workbook, saveFileName);
    }

    /**
     * 导出到输出流OutputStream
     *
     * @param xlsWorkSheetInfoList
     * @param os
     * @return
     */
    public static Boolean exportToXlsOutputStream(List<XlsWorkSheetInfo> xlsWorkSheetInfoList, OutputStream os) {
        //创建一个工作簿
        HSSFWorkbook workbook = new HSSFWorkbook();
        //循环向工作簿添加工作表
        for (XlsWorkSheetInfo xlsWorkSheetInfo : xlsWorkSheetInfoList) {
            addToXls(xlsWorkSheetInfo, workbook, null);
        }
        //xlsWorkSheetInfoList.stream().forEach(workbookInfo -> addToXls(workbookInfo, workbook));//有中文乱码
        try {
            workbook.write(os);
            return true;
        } catch (IOException e) {
            e.printStackTrace();
        }
        return false;
    }

    /**
     * 向工作簿添加一张表
     *
     * @param xlsWorkSheetInfo
     * @param workbook
     * @return
     */
    public static <T> Boolean addToXls(XlsWorkSheetInfo<T> xlsWorkSheetInfo, HSSFWorkbook workbook, XlsSetupAdapter xlsSetupAdapter) {
        Class<T> clazz = xlsWorkSheetInfo.getClazz();
        Field[] fields = clazz.getDeclaredFields();
        try {
            //创建工作表
            Sheet sheet = workbook.createSheet(buildSheetName(xlsWorkSheetInfo));
            //预先美化工作表
            if (xlsSetupAdapter != null) {
                xlsSetupAdapter.workSheetSetup(sheet);
            }
            //遍历字段填充表头
            Row headRow = sheet.createRow(0);
            int headColIndex = 0;
            for (Field field : fields) {
                //获取字段注解
                String colName = getXlsColumnName(field);
                //设置列宽，先设置自动
                int colWidth = getXlsColumnWidth(field);
                if (colWidth >= 0) {
                    sheet.setColumnWidth(headColIndex, colWidth * CHAR_WIDTH);
                } else {
                    sheet.setColumnWidth(headColIndex, (colName.length() + 2) * CHAR_WIDTH);
                }
                //给标题栏设置一个背景色
                HSSFCellStyle cellStyle = workbook.createCellStyle();
                cellStyle.setFillForegroundColor(IndexedColors.BLUE.index);
                Cell cell = headRow.createCell(headColIndex);
                cell.setCellStyle(cellStyle);

                //写入标题
                cell.setCellValue(colName);
                headColIndex++;
            }
            //数据
            List<T> dataList = xlsWorkSheetInfo.getDataList();
            if (dataList != null) {
                for (int rowIndex = 1, len = dataList.size(); rowIndex <= len; rowIndex++) {
                    Object obj = dataList.get(rowIndex - 1);
                    Row dataRow = sheet.createRow(rowIndex);
                    //运用反射，把各个字段的值填充到xls表格
                    int colIndex = 0;
                    for (Field field : fields) {
                        field.setAccessible(true);
                        Object value = field.get(obj);
                        if (value == null) {
                            colIndex++;
                            continue;//空值不写
                        }
                        if ((value instanceof Integer) ||
                                (value instanceof Long) ||
                                (value instanceof Double) ||
                                (value instanceof Float)) {
                            dataRow.createCell(colIndex++).setCellValue(Integer.parseInt(String.valueOf(value)));
                        } else {
                            dataRow.createCell(colIndex++).setCellValue(String.valueOf(value));
                        }
                        field.setAccessible(false);
                    }
                }
            }
            return true;
        } catch (IllegalAccessException e) {
            e.printStackTrace();
        }
        return false;
    }

    /**
     * 设置表的列宽
     *
     * @param field
     * @return
     */
    private static int getXlsColumnWidth(Field field) {
        XlsColumn xlsColumns = field.getAnnotation(XlsColumn.class);
        int width = -1;//列宽
        if (xlsColumns != null) {
            //如果注解不为空且width有值，将width作为列宽
            width = xlsColumns.width();
        }
        //如果最终返回值为-1，将不娶设置列宽，由Excel自己设置
        return width;
    }

    /**
     * 获取字段列名
     *
     * @param field
     * @return
     */
    private static String getXlsColumnName(Field field) {
        XlsColumn xlsColumns = field.getAnnotation(XlsColumn.class);
        String colName = null;//列名
        if (xlsColumns != null) {
            //如果注解不为空且value有值，将value作为列名
            colName = xlsColumns.value();
        }
        //如果cloName仍为空，则以字段名为列名
        if (colName == null || "".equals(colName)) {
            field.setAccessible(true);
            colName = field.getName();
            field.setAccessible(false);
        }
        return colName;
    }

    /**
     * 该字段对应的列是否规定必填
     *
     * @param field
     * @return
     */
    private static boolean isRequired(Field field) {
        XlsColumn xlsColumns = field.getAnnotation(XlsColumn.class);
        boolean required = false;//默认并未规定必填
        if (xlsColumns != null) {
            //如果注解不为空且width有值，将width作为列宽
            required = xlsColumns.required();
        }
        return required;
    }

    /**
     * 创建工作表的名称
     *
     * @param xlsWorkSheetInfo
     * @param <T>
     * @return
     */
    private static <T> String buildSheetName(XlsWorkSheetInfo<T> xlsWorkSheetInfo) {
        int pageIndex = xlsWorkSheetInfo.getPageIndex();
        String title = xlsWorkSheetInfo.getTitle();
        return pageIndex == -1 ? title : title + "(" + pageIndex + ")";
    }

    /**
     * 向工作簿添加多张表
     *
     * 此方法依赖poi 4.0
     * 因为同事Petter使用3.14进行处理，加之目前此方法暂时不用，故此暂时弃用
     * 待后经过协调在进行处理
     * 出错的代码注释掉
     *
     * @param workSheetInfoSet
     * @param workbook
     * @return
     */
    /*
    public static Boolean addToXls(Set<XlsWorkSheetInfo> workSheetInfoSet, HSSFWorkbook workbook) {
        for (XlsWorkSheetInfo workSheetInfo : workSheetInfoSet) {
            addToXls(workSheetInfo, workbook);
        }
        try {
            //workbook.write();
            return true;
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                workbook.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        return false;
    }
    */

    /**
     * 从文件读取xls表格内容
     * 指定表索引号
     *
     * @param xlsFileName
     * @param sheetIndex
     * @param clazz
     * @param <T>
     * @return
     */
    public static <T> List<T> readFromXlsFile(String xlsFileName, int sheetIndex, Class<T> clazz) {
        File file = new File(xlsFileName);
        if (!file.exists()) {
            return null;
        }
        InputStream is = null;
        try {
            is = new FileInputStream(file);
            return readFromInputStream(is, sheetIndex, clazz);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } finally {
            try {
                is.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        return null;
    }

    /**
     * 从文件读取xls表格内容
     * 匹配所有表
     *
     * @param xlsFileName
     * @param clazz
     * @param <T>
     * @return
     */
    public static <T> List<T> readFromXlsFile(String xlsFileName, Class<T> clazz) {
        File file = new File(xlsFileName);
        if (!file.exists()) {
            return null;
        }
        InputStream is = null;
        try {
            is = new FileInputStream(file);
            return readFromInputStream(is, clazz);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } finally {
            try {
                is.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        return null;
    }

    /**
     * 从输入流中读取表格内容
     * 指定表索引号
     *
     * @param is
     * @param sheetIndex
     * @param clazz
     * @param <T>
     * @return
     */
    public static <T> List<T> readFromInputStream(InputStream is, int sheetIndex, Class<T> clazz) {
        try {
            HSSFWorkbook workbook = new HSSFWorkbook(is);
            return readFromWorkbook(workbook, sheetIndex, clazz);
        } catch (IOException e) {
            e.printStackTrace();
        }
        return null;
    }

    /**
     * 从输入流中读取表格内容
     * 匹配所有表
     *
     * @param is
     * @param clazz
     * @param <T>
     * @return
     */
    public static <T> List<T> readFromInputStream(InputStream is, Class<T> clazz) {
        try {
            HSSFWorkbook workbook = new HSSFWorkbook(is);
            return readFromWorkbook(workbook, clazz);
        } catch (IOException e) {
            e.printStackTrace();
        }
        return null;
    }

    /**
     * 从工作簿读取指定数据匹配的数据
     * 系统遍历工作簿，将类型匹配的工作表的数据全部读取出来
     *
     * @param workbook
     * @param clazz
     * @param <T>
     * @return
     */
    public static <T> List<T> readFromWorkbook(HSSFWorkbook workbook, Class<T> clazz) {
        List<T> dataList = new ArrayList<>();
        //获取工作簿中表的总数
        int count = workbook.getNumberOfSheets();
        for (int sheetIndex = 0; sheetIndex < count; sheetIndex++) {
            List<T> list = readFromWorkbook(workbook, sheetIndex, clazz);
            if (list != null) {
                dataList.addAll(list);
            }
        }
        return dataList;
    }

    /**
     * 从工作簿读取指定数据匹配的数据
     * 系统会自动根据字段名匹配数据，并将匹配到的数据全部读出来
     *
     * @param workbook
     * @param sheetIndex
     * @param clazz
     * @param <T>
     * @return
     */
    public static <T> List<T> readFromWorkbook(HSSFWorkbook workbook, int sheetIndex, Class<T> clazz) {
        //获得工作表
        HSSFSheet sheet = workbook.getSheetAt(sheetIndex);
        //检验结构是否匹配
        HSSFRow headRow = sheet.getRow(0);
        Field[] fields = clazz.getDeclaredFields();
        //类型不匹配则退出
        if (!matchClass(fields, headRow)) {
            return null;
        }
        //System.out.println("Found match sheets.");
        // 遍历取出数据
        List<T> objList = new ArrayList<>();
        int length = fields.length;
        int maxLines = getMaxLines(clazz);
        for (int rowIndex = 1; rowIndex < maxLines; rowIndex++) {
            HSSFRow dataRow = sheet.getRow(rowIndex);
            //找到空行 这里要求有效数据中间不能出现空行
            if (dataRow == null) {
                System.out.println("found " + (rowIndex - 1) + " records.");
                break;
            }
            T obj = getInstance(clazz);
            if (obj == null) {
                continue;
            }
            //遍历字段赋值
            for (int colIndex = 0; colIndex < length; colIndex++) {
                Field field = fields[colIndex];
                HSSFCell cell = dataRow.getCell(colIndex);
                field.setAccessible(true);
                setValueFromCell(obj, field, cell);
                field.setAccessible(false);
            }
            objList.add(obj);
        }
        return objList;
    }

    /**
     * 获取设定的每张表最大数据行数
     *
     * @param clazz
     * @param <T>
     * @return
     */
    private static <T> int getMaxLines(Class<T> clazz) {
        int maxLines = 65534;
        XlsSheet xlsSheet = clazz.getAnnotation(XlsSheet.class);
        if (xlsSheet == null) {
            return maxLines;
        }
        int ml = xlsSheet.maxLines();
        if (ml > 0 && ml < 65534) {
            return ml;
        }
        return maxLines;
    }

    /**
     * 从一个xls单元格给一个对象的字段赋值
     *
     * @param obj
     * @param field
     * @param cell
     * @param <T>
     */
    private static <T> void setValueFromCell(T obj, Field field, HSSFCell cell) {
        String typeName = field.getType().getName();
        if (cell == null) {
            return;
        }
        try {

            if (int.class.getName().equals(typeName) || Integer.class.getName().equals(typeName)) {
                field.set(obj, (int) cell.getNumericCellValue());
                return;
            }
            if (short.class.getName().equals(typeName) || Short.class.getName().equals(typeName)) {
                field.set(obj, (short) cell.getNumericCellValue());
                return;
            }
            if (long.class.getName().equals(typeName) || Long.class.getName().equals(typeName)) {
                field.set(obj, (long) cell.getNumericCellValue());
                return;
            }
            if (float.class.getName().equals(typeName) || Float.class.getName().equals(typeName)) {
                field.set(obj, (float) cell.getNumericCellValue());
                return;
            }
            if (double.class.getName().equals(typeName) || Double.class.getName().equals(typeName)) {
                field.set(obj, cell.getNumericCellValue());
                return;
            }

            if (boolean.class.getName().equals(typeName) || Boolean.class.getName().equals(typeName)) {
                field.set(obj, cell.getBooleanCellValue());
                return;
            }

            if (Date.class.getName().equals(typeName)) {
                field.set(obj, cell.getDateCellValue());
                return;
            }
            if (Timestamp.class.getName().equals(typeName)) {
                field.set(obj, new Timestamp(cell.getDateCellValue().getTime()));
                return;
            }
            if (Instant.class.getName().equals(typeName)) {
                field.set(obj, Instant.ofEpochMilli(cell.getDateCellValue().getTime()));
                return;
            }
            //如果上面都不匹配，就全部按照字符串进行读取
            cell.setCellType(CellType.STRING);//强转为字符串类型 poi 4.0
            field.set(obj, cell.getStringCellValue());
            return;
        } catch (IllegalAccessException e) {
            //e.printStackTrace();
        }

    }

    /**
     * 给一个类创建一个实例
     *
     * @param clazz
     * @param <T>
     * @return
     */
    private static <T> T getInstance(Class<T> clazz) {
        try {
            return clazz.newInstance();
        } catch (InstantiationException e) {
            e.printStackTrace();
        } catch (IllegalAccessException e) {
            e.printStackTrace();
        }
        return null;
    }

    /**
     * 根据一个类的字段列表和xls数据表的表头行判断导入数据类型是否匹配
     *
     * @param fields
     * @param headRow
     * @return
     */
    private static boolean matchClass(Field[] fields, HSSFRow headRow) {
        int length = fields.length;
        for (int i = 0; i < length; i++) {
            String colName = null;
            Field field = fields[i];
            XlsColumn xlsColumn = field.getAnnotation(XlsColumn.class);
            if (xlsColumn != null) {
                colName = xlsColumn.value();
            }
            if (colName == null || "".equals(colName)) {
                colName = field.getName();
            }
            if (!headRow.getCell(i).getStringCellValue().equalsIgnoreCase(colName)) {
                return false;
            }
        }
        return true;
    }

    /**
     * 根据给定的类的集合，从工作簿中读取匹配的数据，并返回一个对象集合
     *
     * @param workbook
     * @param clazz
     * @return
     */
    public static Set<List<Object>> readFromWorkbook(Workbook workbook, Set<Class<?>> clazz) {
        return null;
    }

    /**
     * 保存xls文件
     *
     * @param workbook
     * @param saveFileName
     * @return
     */
    private static Boolean saveXlsFile(HSSFWorkbook workbook, String saveFileName) {
        //创建输出流
        File file = new File(saveFileName);
        if (!file.getParentFile().exists()) {
            file.getParentFile().mkdirs();
        }
        OutputStream os = null;
        try {
            os = new FileOutputStream(file);
            workbook.write(os);
            return true;
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            if (os != null) {
                try {
                    os.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }

        return false;
    }
}
