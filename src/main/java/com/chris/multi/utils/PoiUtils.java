package com.chris.multi.utils;

import com.chris.multi.model.WorkSheetInfo;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.List;
import java.util.Set;

/**
 * Created by Chris Chen
 * 2018/09/17
 * Explain: Java操作Excel表格的工具
 */

public class PoiUtils {
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
     * @param workbook
     * @param clazz
     * @return
     */
    public static Set<List<Object>> readFromXls(Workbook workbook, Set<Class<?>> clazz) {
        return null;
    }
}
