package com.chris.multi.poi.xls;

import java.util.List;

/**
 * Created by Chris Chen
 * 2018/09/17
 * Explain: 工作表数据信息
 */
public class WorkSheetInfo<T> {
    private String title;//工作表名称
    private int pageIndex = -1;//分页的页码 不设置则为-1
    private String author;//作者
    private Long time;//操作时间
    private List<T> dataList;//工作表中每行的数据
    private Class<T> clazz;

    private static final int MAXLINES = 65534;

    private WorkSheetInfo() {

    }

    public static <T> WorkSheetInfo get(Class<T> clazz) {
        WorkSheetInfo<T> workSheetInfo = new WorkSheetInfo<>();
        workSheetInfo.clazz = clazz;
        return workSheetInfo;
    }

    public String getTitle() {
        return title;
    }

    public WorkSheetInfo setTitle(String title) {
        this.title = title;
        return this;
    }

    public int getPageIndex() {
        return pageIndex;
    }

    public WorkSheetInfo setPageIndex(int pageIndex) {
        this.pageIndex = pageIndex;
        return this;
    }

    public String getAuthor() {
        return author;
    }

    public WorkSheetInfo setAuthor(String author) {
        this.author = author;
        return this;
    }

    public Long getTime() {
        return time;
    }

    public WorkSheetInfo setTime(Long time) {
        this.time = time;
        return this;
    }

    public List<T> getDataList() {
        return dataList;
    }

    public WorkSheetInfo setDataList(List<T> dataList) {
        if (dataList.size() > getMaxLines()) {
            throw new RuntimeException("One page can not contains data than " + MAXLINES + "lines.");
        }
        this.dataList = dataList;
        return this;
    }

    public Class<T> getClazz() {
        return clazz;
    }

    public WorkSheetInfo setClazz(Class<T> clazz) {
        this.clazz = clazz;
        return this;
    }

    /**
     * 获取设定的最大行数
     *
     * @return
     */
    private int getMaxLines() {
        XlsSheet xlsSheet = getClazz().getAnnotation(XlsSheet.class);
        if (xlsSheet == null) {
            return MAXLINES;
        }
        int ml = xlsSheet.maxLines();
        if (ml > 0 && ml < 65534) {
            return ml;
        }
        return MAXLINES;
    }
}
