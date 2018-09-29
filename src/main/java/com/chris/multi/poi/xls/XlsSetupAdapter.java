package com.chris.multi.poi.xls;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;

/**
 * Created by Chris Chen
 * 2018/09/27
 * Explain: 电子表工作簿设置
 */

public interface XlsSetupAdapter {
    /**
     * 设置工作簿
     *
     * @param workbook
     */
    void workBookSetup(HSSFWorkbook workbook);

    /**
     * 设置工作表
     *
     * @param workSheet
     */
    void workSheetSetup(Sheet workSheet);
}
