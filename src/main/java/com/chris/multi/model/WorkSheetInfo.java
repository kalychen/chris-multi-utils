package com.chris.multi.model;

import java.util.List;

/**
 * Created by Chris Chen
 * 2018/09/17
 * Explain: 工作表数据信息
 */
public class WorkSheetInfo<T> {
    private String title;//工作表名称
    private String author;//作者
    private Long time;//操作时间
    private List<T> dataList;//工作表中每行的数据

    public WorkSheetInfo() {
    }

    public WorkSheetInfo(String title, String author, Long time, List<T> dataList) {
        this.title = title;
        this.author = author;
        this.time = time;
        this.dataList = dataList;
    }

    public String getTitle() {
        return title;
    }

    public void setTitle(String title) {
        this.title = title;
    }

    public String getAuthor() {
        return author;
    }

    public void setAuthor(String author) {
        this.author = author;
    }

    public Long getTime() {
        return time;
    }

    public void setTime(Long time) {
        this.time = time;
    }

    public List<T> getDataList() {
        return dataList;
    }

    public void setDataList(List<T> dataList) {
        this.dataList = dataList;
    }
}
