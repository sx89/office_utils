package com.bytedance.eduutils.core;

import com.bytedance.eduutils.PoiExcel;
import com.bytedance.eduutils.entity.User;
import org.apache.log4j.Logger;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class Core {
    /**
     * 读取excel文件中的用户信息，保存在数据库中
     *
     * @param excelFile
     */
    private static org.apache.log4j.Logger logger = Logger.getLogger(PoiExcel.class);

    public static void main(String[] args) throws IOException {
        PoiExcel poiExcel = new PoiExcel();
        ArrayList<List<?>> file1 = poiExcel.readExcel("file/1.XLSX", User.class);

        List<User> users = new ArrayList<User>();
        for (int i = 1; i < file1.size(); i++) {
            List<?> temp = file1.get(i);
            Object o = temp.get(2);
            double o1 = (double) o;
            int i1 = (int) o1;
            users.add(new User("hello" + (String) temp.get(0), (String) temp.get(1), i1));
        }
        poiExcel.writeExcel("file/2.XLSX", users, User.class);
    }

}
