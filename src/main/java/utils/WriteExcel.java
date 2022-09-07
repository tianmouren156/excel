package utils;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

public class WriteExcel {
    public WriteExcel() {
    }

    public boolean writeExcel(String path, String name,Map<String,Map<String,String>> map) throws IOException {
        Workbook workbook=new HSSFWorkbook();
        Sheet sheet = workbook.createSheet();

        Set<String> set = new TreeSet<>();
        for (String s : map.keySet()) {
            for (String s1 : map.get(s).keySet()) {
                set.add(s1);
            }
        }
        List<String> list = new ArrayList<>(set);
        ArrayList<String> keyName = new ArrayList<>(map.keySet());
        System.out.println(keyName);
        List<List<String>> data = new ArrayList<>();
        List<String> mi = new ArrayList<>();
        for (String s : list) {
            mi.add(s);
        }
        data.add(mi);
        for (String s : keyName) {
            List<String> data1 = new ArrayList<>();
            data1.add(s);
            Map<String, String> stringMap = map.get(s);
            for (String s1 : list) {
                String s2 = stringMap.get(s1);
            data1.add(s2);
            }
            data.add(data1);
        }
        System.out.println(data);
        int rowNum = 0;
        for (List<String> datum : data) {
            int cellNum = 0;
            if (rowNum==0)cellNum=1;
            Row row = sheet.createRow(rowNum++);
            for (String s : datum) {

                Cell cell = row.createCell(cellNum++);
                cell.setCellValue(s);
            }

        }
        FileOutputStream fos= new FileOutputStream(path+"/test01.xls");
        workbook.write(fos);
        fos.close();
        return false;
    }
}
