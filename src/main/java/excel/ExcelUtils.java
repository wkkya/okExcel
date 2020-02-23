package excel;

import annotation.ImportExcel;
import com.alibaba.fastjson.JSONObject;
import com.sun.istack.internal.NotNull;
import com.sun.istack.internal.Nullable;
import org.apache.commons.codec.binary.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.util.StringUtil;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class ExcelUtils {
    private static Logger logger = LoggerFactory.getLogger(ExcelUtils.class);

    public static List<Object> readExcel(@NotNull String path, @NotNull Class<?> clazz){
        return readExcel(path, 1, 2, clazz);
    }

    public static List<Object> readExcel(@NotNull String path, int titleLine, int start, @NotNull Class<?> clazz){

        InputStream inputStream = null;
        try {
            inputStream = new FileInputStream(new File(path));
        } catch (FileNotFoundException e) {
            logger.error(path + "文件未找到", e);
        }

        HSSFWorkbook hssfWorkbook = null;
        try {
            hssfWorkbook = new HSSFWorkbook(inputStream);
        } catch (IOException e) {
            logger.error("构建HssfWorkbook失效", e);
        }

        List<Map<String, Object>> list = new ArrayList<Map<String, Object>>();
        HSSFSheet sheet = hssfWorkbook.getSheetAt(0);
        HSSFRow titleRow = sheet.getRow(titleLine - 1);
        //通过数值限定读取范围，避免无标题多数据行多次不必要io
        short firstCellNum = titleRow.getFirstCellNum();
        short lastCellNum = titleRow.getLastCellNum();
        for(int i=start-1; i<=sheet.getLastRowNum(); i++){
            Map<String, Object> map = new HashMap<String, Object>();
            HSSFRow row = sheet.getRow(i);
            for (int j = firstCellNum; j<lastCellNum; j++){
                HSSFCell titleCell = titleRow.getCell(j);
                HSSFCell cell = row.getCell(j);
                map.put(titleCell.getStringCellValue(), getCellValue(cell));
            }
            list.add(map);
        }

        List<Object> result = new ArrayList<Object>();
        //获取所有带注解的属性
        for (Map<String, Object> map : list){
            Object o = null;
            try {
                o = clazz.newInstance();
                Field[] declaredFields = clazz.getDeclaredFields();
                for(Field field:declaredFields){
                    if (field.isAnnotationPresent(ImportExcel.class)){
                        ImportExcel annotation = field.getAnnotation(ImportExcel.class);
                        String value = annotation.value();
                        Object o1 = map.get(value);
                        field.setAccessible(true);
                        field.set(o, o1);
                    }
                }
            } catch (InstantiationException e) {
                logger.error("反射生成实体异常", e);
            } catch (IllegalAccessException e) {
                logger.error("反射生成实体异常", e);
            }
            result.add(o);
        }
        logger.info(JSONObject.toJSONString(result));
        return result;
    }

    private static Object getCellValue(HSSFCell cell){
        Object defValue = "";
        int cellType = cell.getCellType();
        switch (cellType){
            case HSSFCell.CELL_TYPE_NUMERIC:
                defValue = cell.getNumericCellValue();
                break;
            case HSSFCell.CELL_TYPE_STRING:
                defValue = cell.getStringCellValue();
                break;
            case HSSFCell.CELL_TYPE_BOOLEAN:
                defValue = cell.getBooleanCellValue();
                break;
            case HSSFCell.CELL_TYPE_FORMULA:
                defValue = cell.getCellFormula();
                break;
            default:
                defValue = "";
                break;
        }
        return defValue;
    }

}
