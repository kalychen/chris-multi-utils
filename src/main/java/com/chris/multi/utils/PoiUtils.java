package com.chris.multi.utils;

import com.chris.multi.model.WorkSheetInfo;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.util.List;
import java.util.Set;

/**
 * Created by Chris Chen
 * 2018/09/17
 * Explain: Java操作Excel表格的工具
 */

public class PoiUtils {
    public static <T> Boolean exportToXls(WorkSheetInfo<T> workSheetInfo) {
        String saveFileName = "G:/temp/chris-test-02.xls";
        Class<T> clazz = workSheetInfo.getClazz();
        Field[] fields = clazz.getDeclaredFields();
        try {
            //创建一个工作簿
            HSSFWorkbook workbook = new HSSFWorkbook();
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
//                dataRow.createCell(0).setCellValue(user.id);
//                dataRow.createCell(1).setCellValue(user.name);
//                dataRow.createCell(2).setCellValue(user.age);
//                dataRow.createCell(3).setCellValue(user.address);
                //运用反射，把各个字段的值填充到xls表格
                int colIndex = 0;
                for (Field field : fields) {
                    field.setAccessible(true);
                    Object value = field.get(obj);
                    field.setAccessible(false);
                    dataRow.createCell(colIndex++).setCellValue(String.valueOf(value));
                }
            }
            //创建输出流
            OutputStream os = new FileOutputStream(new File(saveFileName));
            workbook.write(os);
            workbook.close();
            return true;

        } catch (IOException e) {
            e.printStackTrace();
        } catch (IllegalAccessException e) {
            e.printStackTrace();
        }
        return false;
    }

    /**
     * 把包含工作表信息的集合写入到xls表格
     *
     * @param workSheetInfoSet
     * @return
     */
    public static Boolean exportToXls(Set<WorkSheetInfo> workSheetInfoSet) {
        return false;
    }

    /**
     * 向工作簿添加一张表
     *
     * @param workSheetInfo
     * @param workbook
     * @return
     */
    public static Boolean addToXls(WorkSheetInfo workSheetInfo, Workbook workbook) {
        return false;
    }

    /**
     * 向工作簿添加多张表
     *
     * @param workSheetInfoSet
     * @param workbook
     * @return
     */
    public static Boolean addToXls(Set<WorkSheetInfo> workSheetInfoSet, Workbook workbook) {
        return false;
    }

    /**
     * 从工作簿读取指定数据匹配的数据
     * 不用指定哪张表，系统会自动根据字段名匹配数据，并将匹配到的数据全部读出来
     *
     * @param workbook
     * @param clazz
     * @param <T>
     * @return
     */
    public static <T> List<T> readFromXls(Workbook workbook, Class<T> clazz) {
        return null;
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
}
