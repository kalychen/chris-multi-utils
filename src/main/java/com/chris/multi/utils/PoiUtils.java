package com.chris.multi.utils;

import com.chris.multi.model.WorkSheetInfo;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.*;
import java.lang.reflect.Field;
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
                field.setAccessible(true);
                String fieldname = field.getName();
                field.setAccessible(false);
                headRow.createCell(headColIndex++).setCellValue(fieldname);
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
            return null;
        }
        // 遍历取出数据 todo 如何确定有效数据的最大行
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
            if (!headRow.getCell(i).getStringCellValue().equalsIgnoreCase(fields[i].getName())) {
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
