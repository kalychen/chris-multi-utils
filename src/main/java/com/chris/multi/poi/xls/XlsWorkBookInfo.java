package com.chris.multi.poi.xls;

import java.sql.Timestamp;
import java.util.List;

/**
 * Created by Chris Chen
 * 2018/09/27
 * Explain: 工作簿数据信息
 */

public class XlsWorkBookInfo {
    private String name;//工作簿名称
    private String author;//作者
    private String lastModifier;//上一个修改者
    private Timestamp createTime;
    private Timestamp lastAccessTime;

    private XlsSetupAdapter xlsSetupAdapter;//细节设置适配器

    private List<XlsWorkSheetInfo> xlsWorkSheetInfoList;

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getAuthor() {
        return author;
    }

    public void setAuthor(String author) {
        this.author = author;
    }

    public String getLastModifier() {
        return lastModifier;
    }

    public void setLastModifier(String lastModifier) {
        this.lastModifier = lastModifier;
    }

    public Timestamp getCreateTime() {
        return createTime;
    }

    public void setCreateTime(Timestamp createTime) {
        this.createTime = createTime;
    }

    public Timestamp getLastAccessTime() {
        return lastAccessTime;
    }

    public void setLastAccessTime(Timestamp lastAccessTime) {
        this.lastAccessTime = lastAccessTime;
    }

    public List<XlsWorkSheetInfo> getXlsWorkSheetInfoList() {
        return xlsWorkSheetInfoList;
    }

    public void setXlsWorkSheetInfoList(List<XlsWorkSheetInfo> xlsWorkSheetInfoList) {
        this.xlsWorkSheetInfoList = xlsWorkSheetInfoList;
    }

    public XlsSetupAdapter getXlsSetupAdapter() {
        return xlsSetupAdapter;
    }

    public void setXlsSetupAdapter(XlsSetupAdapter xlsSetupAdapter) {
        this.xlsSetupAdapter = xlsSetupAdapter;
    }
}
