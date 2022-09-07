package utils;



import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joda.time.DateTime;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.*;

public class ReadExcel {
    public ReadExcel() {
    }

    public Map<Integer, List<Object>> ReadExcelUtil(String path, Integer count){
        Map<Integer, List<Object>> map = new HashMap<>();
        List<Object> list = null;
        Integer key = 0;
        try {
            XSSFWorkbook workbook= new XSSFWorkbook(new FileInputStream(path));
            XSSFSheet sheetAt = workbook.getSheetAt(0);
            int lastRowNum = sheetAt.getLastRowNum();
            for (int i = count; i < lastRowNum; i++) {
                XSSFRow row = sheetAt.getRow(i);
                int lastCellNum = row.getLastCellNum();
                list = new ArrayList<Object>();
                for (int j = 0; j < lastCellNum-1; j++) {
                    XSSFCell cell = row.getCell(j);
                    CellType type = cell.getCellType();
                    Object value = "";
                    switch (type){

                        case STRING:
                            value = cell.getStringCellValue();
                            break;
                        case NUMERIC:
                            if (HSSFDateUtil.isCellDateFormatted(cell)){
                                Date dateCellValue = cell.getDateCellValue();
                                value = new DateTime(dateCellValue).toString("yyyy/mm/dd");
                            }else {
                                cell.setCellType(CellType.STRING);
                                value = cell.toString();
                            }
                            break;
                    }
                    list.add(value);
                }
                map.put(key++, list);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
        return map;
    }

    public Map<Integer, List<Object>> NewMap(Map<Integer, List<Object>> map){
        Map<Integer, List<Object>> hashMap = new HashMap<>();
        int key = 0;
        for (int i = 0; i < map.size(); i++) {
            if (map.get(i)!=null){
                hashMap.put(key++,map.get(i));
            }
        }
        return hashMap;
    }

    public List<String> duplicateRemoval(List<String> list){
        return new ArrayList<>(new TreeSet<String>(list));
    }

    public Integer queryCoordinates(String val,List<Object> list){
        Integer count = null;
        for (int i = 0; i < list.size(); i++) {
            if (String.valueOf(list.get(i)).contains(val)){
                count = i;
            return count;
            }
        }
        return count;
    }
}
