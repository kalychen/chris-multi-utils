package com.chris.multi.model;

import com.chris.multi.poi.xls.XlsColumn;
import com.chris.multi.poi.xls.XlsSheet;

/**
 * Created by Chris Chen
 * 2018/09/17
 * Explain:
 */
@XlsSheet(value = "学生表")
public class StuModel {
    @XlsColumn("编号")
    private int id;
    @XlsColumn("姓名")
    private String name;
    @XlsColumn("年级")
    private String grade;
    @XlsColumn("班号")
    private String schoolClass;
    @XlsColumn("英语科成绩")
    private int englishScore;

    public StuModel() {
    }

    public StuModel(int id, String name, String grade, String schoolClass, int englishScore) {
        this.id = id;
        this.name = name;
        this.grade = grade;
        this.schoolClass = schoolClass;
        this.englishScore = englishScore;
    }

    public int getId() {
        return id;
    }

    public void setId(int id) {
        this.id = id;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getGrade() {
        return grade;
    }

    public void setGrade(String grade) {
        this.grade = grade;
    }

    public String getSchoolClass() {
        return schoolClass;
    }

    public void setSchoolClass(String schoolClass) {
        this.schoolClass = schoolClass;
    }

    public int getEnglishScore() {
        return englishScore;
    }

    public void setEnglishScore(int englishScore) {
        this.englishScore = englishScore;
    }
}
