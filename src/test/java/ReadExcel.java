import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joda.time.DateTime;
import org.junit.Test;
import utils.WriteExcel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.*;

public class ReadExcel {
    @Test
    public void test(){

        String path = "E:\\Wechat_Files\\WeChat Files\\wxid_drci3s9roofq22\\FileStorage\\File\\2022-09\\上下班打卡_日报_20220801-20220831.xlsx";
        utils.ReadExcel readExcel= new utils.ReadExcel();
        Map<Integer, List<Object>> map = readExcel.ReadExcelUtil(path, 4);
        for (int i = 0; i < map.size(); i++) {
            List<Object> list = map.get(i);
            for (int j = 0; j < list.size(); j++) {
                String st = (String) list.get(j);
                if (st.contains("已离职")){
                    map.remove(i);
                }
            }
        }
        Map<Integer, List<Object>> newMap = readExcel.NewMap(map);

        List<String> names = new ArrayList<>();

        for (int i = 0; i < newMap.size(); i++) {
            List<Object> list = newMap.get(i);
            names.add(String.valueOf(list.get(1)));
        }
        names = readExcel.duplicateRemoval(names);
        System.out.println(names);
        Map<String,Map<String,String>> nameMap = new HashMap<>();
        int count = 0;
        Integer integer = null;
        while (count<names.size()){
            Map<String, String> hashMap = new HashMap<>();
            for (int i = 0; i < newMap.size(); i++) {
                List<Object> list = newMap.get(i);
                if (integer == null){
                    integer = readExcel.queryCoordinates("正常", list);
                }
                if (String.valueOf(list.get(1)).contains(names.get(count))){
                    String date = String.valueOf(list.get(0));
                    String stuts = String.valueOf(list.get(integer));
                    hashMap.put(date,stuts);
                }
            }
            nameMap.put(names.get(count),hashMap);
            count++;
        }
        System.out.println(nameMap);
        WriteExcel writeExcel = new WriteExcel();
        try {
            writeExcel.writeExcel("D:\\","a",nameMap);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

}
