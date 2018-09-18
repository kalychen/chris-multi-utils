package com.chris.multi.sample;

import com.chris.multi.sample.model.StuModel;
import com.chris.multi.sample.model.UserModel;
import com.chris.multi.poi.xls.WorkSheetInfo;
import com.chris.multi.poi.xls.PoiUtils;

import java.io.*;
import java.util.*;

/**
 * Created by Chris Chen
 * 2018/09/17
 * Explain:
 */

public class MainTest {
    private static final String saveFileName = "G:/temp1/chris-test-06.xls";

    public static void main(String[] args) {
        outTemplate();
    }

    //输出表模板
    private static void outTemplate() {
        test3();//只要数据为空，生成的就是只有表头的表模板
    }

    private static void test4() {
        List<StuModel> stuModels = PoiUtils.readFromXls(saveFileName, 0, StuModel.class);
        for (StuModel stu : stuModels) {
            System.out.println(stu.getId() + "-->" + stu.getName() + "-->" + stu.getGrade() + "-->" + stu.getSchoolClass() + "-->" + stu.getEnglishScore());
        }
    }

    private static void test3() {
        Set<WorkSheetInfo> workSheetInfoSet = new HashSet<>();
        workSheetInfoSet.add(getStuInfo());
        workSheetInfoSet.add(getUserInfo());
        OutputStream os = null;
        File file = new File(saveFileName);
        if (!file.getParentFile().exists()) {
            file.getParentFile().mkdirs();
        }
        try {

            os = new FileOutputStream(file);
            PoiUtils.exportToXlsOutputStream(workSheetInfoSet, os);
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

    private static WorkSheetInfo<StuModel> getStuInfo() {

        List<StuModel> stuList = new ArrayList<>();
        for (int i = 1; i <= 100; i++) {
            stuList.add(new StuModel(i, "学生 " + i, "三年级", "五班", new Random().nextInt(100)));
        }
        WorkSheetInfo<StuModel> workSheetInfo = new WorkSheetInfo<>(StuModel.class);
        workSheetInfo.setTitle("学生表");
        workSheetInfo.setAuthor("Chris Chen");
        workSheetInfo.setTime(System.currentTimeMillis());
        //workSheetInfo.setDataList(stuList);

        return workSheetInfo;
    }

    private static WorkSheetInfo<UserModel> getUserInfo() {
        List<UserModel> userList = new ArrayList<>();
        for (int i = 1; i <= 100; i++) {
            userList.add(new UserModel(i, "name " + i, new Random().nextInt(100), "addr " + i));
        }
        WorkSheetInfo<UserModel> workSheetInfo = new WorkSheetInfo<>(UserModel.class);
        workSheetInfo.setTitle("用户表2");
        workSheetInfo.setAuthor("Chris Chen");
        workSheetInfo.setTime(System.currentTimeMillis());
        //workSheetInfo.setDataList(userList);

        return workSheetInfo;
    }
}