package com.chris.multi.poi.xls;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.*;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Set;

/**
 * Created by Chris Chen
 * 2018/09/17
 * Explain: Java操作Excel表格的工具
 */

public class PoiUtils {
    /**
     * 创建一个工作簿，并且添加一张表，写入数据
     *
     * @param workSheetInfo
     * @param saveFileName
     * @param <T>
     * @return
     */
    public static <T> Boolean exportToXls(WorkSheetInfo<T> workSheetInfo, String saveFileName) {
        //创建一个工作簿
        HSSFWorkbook workbook = new HSSFWorkbook();
        //向工作簿天机按工作表
        addToXls(workSheetInfo, workbook);
        return saveXlsFile(workbook, saveFileName);
    }

    /**
     * 导出到输出流OutputStream
     *
     * @param workSheetInfo
     * @param os
     * @param <T>
     * @return
     */
    public static <T> Boolean exportToXlsOutputStream(WorkSheetInfo<T> workSheetInfo, OutputStream os) {
        //创建一个工作簿
        HSSFWorkbook workbook = new HSSFWorkbook();
        //向工作簿天机按工作表
        addToXls(workSheetInfo, workbook);
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
     * @param workSheetInfoSet
     * @param saveFileName
     * @return
     */
    public static Boolean exportToXls(Set<WorkSheetInfo> workSheetInfoSet, String saveFileName) {
        //创建一个工作簿
        HSSFWorkbook workbook = new HSSFWorkbook();
        //循环向工作簿添加工作表
        for (WorkSheetInfo workSheetInfo : workSheetInfoSet) {
            addToXls(workSheetInfo, workbook);
        }
        return saveXlsFile(workbook, saveFileName);
    }

    /**
     * 导出到输出流OutputStream
     *
     * @param workSheetInfoSet
     * @param os
     * @return
     */
    public static Boolean exportToXlsOutputStream(Set<WorkSheetInfo> workSheetInfoSet, OutputStream os) {
        //创建一个工作簿
        HSSFWorkbook workbook = new HSSFWorkbook();
        //循环向工作簿添加工作表
        for (WorkSheetInfo workSheetInfo : workSheetInfoSet) {
            addToXls(workSheetInfo, workbook);
        }
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
     * @param workSheetInfo
     * @param workbook
     * @return
     */
    public static <T> Boolean addToXls(WorkSheetInfo<T> workSheetInfo, HSSFWorkbook workbook) {
        Class<T> clazz = workSheetInfo.getClazz();
        Field[] fields = clazz.getDeclaredFields();
        try {
            //创建工作表
            Sheet sheet = workbook.createSheet(workSheetInfo.getTitle());
            //遍历字段填充表头
            Row headRow = sheet.createRow(0);
            int headColIndex = 0;
            for (Field field : fields) {
                //获取字段注解
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
                headRow.createCell(headColIndex++).setCellValue(colName);
            }
            //数据
            List<T> dataList = workSheetInfo.getDataList();
            for (int rowIndex = 1, len = dataList.size(); rowIndex <= len; rowIndex++) {
                Object obj = dataList.get(rowIndex - 1);
                Row dataRow = sheet.createRow(rowIndex);
                //运用反射，把各个字段的值填充到xls表格
                int colIndex = 0;
                for (Field field : fields) {
                    field.setAccessible(true);
                    Object value = field.get(obj);
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
            return true;
        } catch (IllegalAccessException e) {
            e.printStackTrace();
        }
        return false;
    }

    /**
     * 向工作簿添加多张表
     *
     * @param workSheetInfoSet
     * @param workbook
     * @return
     */
    public static Boolean addToXls(Set<WorkSheetInfo> workSheetInfoSet, HSSFWorkbook workbook) {
        for (WorkSheetInfo workSheetInfo : workSheetInfoSet) {
            addToXls(workSheetInfo, workbook);
        }
        try {
            workbook.write();
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

    /**
     * 从文件读取xls表格内容
     *
     * @param xlsFileName
     * @param sheetIndex
     * @param clazz
     * @param <T>
     * @return
     */
    public static <T> List<T> readFromXls(String xlsFileName, int sheetIndex, Class<T> clazz) {
        File file = new File(xlsFileName);
        if (!file.exists()) {
            return null;
        }
        InputStream is = null;
        try {
            is = new FileInputStream(file);
            return readFromXls(is, sheetIndex, clazz);
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
     *
     * @param is
     * @param sheetIndex
     * @param clazz
     * @param <T>
     * @return
     */
    public static <T> List<T> readFromXls(InputStream is, int sheetIndex, Class<T> clazz) {
        try {
            HSSFWorkbook workbook = new HSSFWorkbook(is);
            return readFromXls(workbook, sheetIndex, clazz);
        } catch (IOException e) {
            e.printStackTrace();
        }
        return null;
    }

    /**
     * 从工作簿读取指定数据匹配的数据
     * 不用指定哪张表，系统会自动根据字段名匹配数据，并将匹配到的数据全部读出来
     *
     * @param workbook
     * @param sheetIndex
     * @param clazz
     * @param <T>
     * @return
     */
    public static <T> List<T> readFromXls(HSSFWorkbook workbook, int sheetIndex, Class<T> clazz) {
        //获得工作表
        HSSFSheet sheet = workbook.getSheetAt(sheetIndex);
        //检验结构是否匹配
        HSSFRow headRow = sheet.getRow(0);
        Field[] fields = clazz.getDeclaredFields();
        //类型不匹配则退出
        if (!matchClass(fields, headRow)) {
            System.out.println("类型不匹配");
            return null;
        }
        System.out.println("找到合适的表");
        // 遍历取出数据
        List<T> objList = new ArrayList<>();
        int length = fields.length;
        for (int rowIndex = 1; rowIndex < 65535; rowIndex++) {
            HSSFRow dataRow = sheet.getRow(rowIndex);
            //找到空行 这里要求有效数据中间不能出现空行
            if (dataRow == null) {
                System.out.println("找到" + (rowIndex - 1) + "条数据");
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
     * 从一个xls单元格给一个对象的字段赋值
     *
     * @param obj
     * @param field
     * @param cell
     * @param <T>
     */
    private static <T> void setValueFromCell(T obj, Field field, HSSFCell cell) {
        String typeName = field.getType().getName();
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
            //如果上面都不匹配，就全部按照字符串进行读取
            field.set(obj, cell.getStringCellValue());
            return;
        } catch (IllegalAccessException e) {
            e.printStackTrace();
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
    public static Set<List<Object>> readFromXls(Workbook workbook, Set<Class<?>> clazz) {
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
