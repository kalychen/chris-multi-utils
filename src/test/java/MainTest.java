import com.chris.multi.model.StuModel;
import com.chris.multi.model.UserModel;
import com.chris.multi.model.WorkSheetInfo;
import com.chris.multi.utils.PoiUtils;

import java.util.ArrayList;
import java.util.List;
import java.util.Random;

/**
 * Created by Chris Chen
 * 2018/09/17
 * Explain:
 */

public class MainTest {
    public static void main(String[] args) {
        test2();
    }

    private static void test2() {
        List<StuModel> stuList = new ArrayList<>();
        for (int i = 1; i <= 100; i++) {
            stuList.add(new StuModel(i,"学生 "+i,"三年级","五班",new Random().nextInt(100)));
        }
        WorkSheetInfo<StuModel> workSheetInfo = new WorkSheetInfo<>(StuModel.class);
        workSheetInfo.setTitle("学生表");
        workSheetInfo.setAuthor("Chris Chen");
        workSheetInfo.setTime(System.currentTimeMillis());
        workSheetInfo.setDataList(stuList);

        PoiUtils.exportToXls(workSheetInfo);
    }

    private static void test1() {
        List<UserModel> userList = new ArrayList<>();
        for (int i = 1; i <= 100; i++) {
            userList.add(new UserModel(i, "name " + i, new Random().nextInt(100), "addr " + i));
        }
        WorkSheetInfo<UserModel> workSheetInfo = new WorkSheetInfo<>(UserModel.class);
        workSheetInfo.setTitle("用户表2");
        workSheetInfo.setAuthor("Chris Chen");
        workSheetInfo.setTime(System.currentTimeMillis());
        workSheetInfo.setDataList(userList);

        PoiUtils.exportToXls(workSheetInfo);
    }
}
