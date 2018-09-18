package com.chris.multi.sample;

import com.chris.multi.poi.xls.PoiUtils;
import com.chris.multi.poi.xls.WorkSheetInfo;
import com.chris.multi.sample.model.StuModel;
import com.chris.multi.sample.model.UserModel;

import java.io.*;
import java.util.*;

/**
 * Created by Chris Chen
 * 2018/09/17
 * Explain:
 */

public class MainTest {
    private static final String saveFileName = "G:/temp1/chris-test-02.xls";

    public static void main(String[] args) {
        outTemplate();
    }

    //输出表模板
    private static void outTemplate() {
        exportMultiSheet();//只要数据为空，生成的就是只有表头的表模板
    }

    private static void importXls() {
        List<StuModel> stuModels = PoiUtils.readFromXlsFile(saveFileName, StuModel.class);
        for (StuModel stu : stuModels) {
            System.out.println(stu.getId() + "-->" + stu.getName() + "-->" + stu.getGrade() + "-->" + stu.getSchoolClass() + "-->" + stu.getEnglishScore());
        }
    }

    private static void exportMultiSheet() {
        List<WorkSheetInfo> workSheetInfoList = new ArrayList<>();
        workSheetInfoList.add(getStuInfo(1));
        workSheetInfoList.add(getStuInfo(2));
        workSheetInfoList.add(getUserInfo());
        OutputStream os = null;
        File file = new File(saveFileName);
        if (!file.getParentFile().exists()) {
            file.getParentFile().mkdirs();
        }
        try {

            os = new FileOutputStream(file);
            PoiUtils.exportToXlsOutputStream(workSheetInfoList, os);
            os.flush();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                os.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }

    }

    private static WorkSheetInfo<StuModel> getStuInfo(int pageIndex) {

        List<StuModel> stuList = new ArrayList<>();
        for (int i = 1; i <= 10; i++) {
            stuList.add(new StuModel(i, "学生 page " + pageIndex +"->"+ i, "三年级", "五班", new Random().nextInt(100)));
        }
        WorkSheetInfo<StuModel> workSheetInfo = WorkSheetInfo.get(StuModel.class)
                .setTitle("学生表")
                .setPageIndex(pageIndex)
                .setTime(System.currentTimeMillis());
//                .setDataList(stuList);

        return workSheetInfo;
    }

    private static WorkSheetInfo<UserModel> getUserInfo() {
        List<UserModel> userList = new ArrayList<>();
        for (int i = 1; i <= 100; i++) {
            userList.add(new UserModel(i, "name " + i, new Random().nextInt(100), "addr " + i));
        }
        WorkSheetInfo<UserModel> workSheetInfo = WorkSheetInfo.get(UserModel.class)
                .setTitle("用户表")
                .setTime(System.currentTimeMillis());
//                .setDataList(userList);

        return workSheetInfo;
    }
}
