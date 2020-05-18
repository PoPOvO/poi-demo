package org.xli.excel;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.security.InvalidParameterException;
import java.util.*;

/**
 * @author xli
 * @Description 操作Excel的基本抽象类，如果需要将具体的POJO和Excle映射可以继承该类
 * @Date 创建于 2019/1/24 9:26
 */
public abstract class AbstractExcelMappingPOJO<T> {
    /**
     * 使用者导入时对表格处理的类，因为表格的格式可能不同，例如存在标题，不存在标题等。需要使用时根据情况处理
     *
     * @param sheet
     * @param sheetList
     */
    abstract public void processImportSheet(Sheet sheet, List<T> sheetList);

    /**
     * 使用者导出时对表格处理的类，如设置样式，标题，表头等
     *
     * @param workbook
     * @param sheet
     * @param sheetList
     */
    abstract public void processExportSheet(Workbook workbook, Sheet sheet, List<T> sheetList);

    private boolean isXLS(String name) {
        return name.length() > 4 && name.endsWith(".xls");
    }

    private boolean isXLSX(String name) {
        return name.length() > 5 && name.endsWith(".xlsx");
    }

    /**
     * 导入Excel到List的模板方法
     *
     * @param file
     * @return
     */
    public Map<String, List<T>> importExcel(File file) {
        if (file == null || !file.exists()) {
            throw new InvalidParameterException("文件不存在");
        }

        try (FileInputStream fis = new FileInputStream(file);) {
            //工作簿对象
            Workbook workbook = null;

            //判断Excel的版本，分别使用不同的类处理
            if (isXLS(file.getName())) {
                workbook = new HSSFWorkbook(fis);
            } else if (isXLSX(file.getName())) {
                workbook = new XSSFWorkbook(fis);
            } else {
                throw new InvalidParameterException("文件格式不正确");
            }

            Map<String, List<T>> res = new LinkedHashMap<>();
            Iterator<Sheet> iterator = workbook.sheetIterator();

            //迭代表格
            while (iterator.hasNext()) {
                //表格对象
                Sheet sheet = iterator.next();
                List<T> sheetList = new ArrayList<>();
                res.put(sheet.getSheetName(), sheetList);
                //使用者对表格进行处理
                processImportSheet(sheet, sheetList);
            }

            return res;
        } catch (IOException e) {
            e.printStackTrace();
        }

        return null;
    }

    /**
     * 导出List数据到Excel的模板方法
     *
     * @param map
     * @param target
     */
    public void exportExcel(Map<String, List<T>> map, File target) {
        if (map == null || map.size() == 0 || target == null || (!isXLS(target.getName()) && !isXLSX(target.getName()))) {
            throw new InvalidParameterException("参数不合法");
        }

        boolean isXLS = isXLS(target.getName()) ? true : false;
        try (FileOutputStream fos = new FileOutputStream(target);) {
            //根据文件格式创建工作簿对象
            Workbook workbook = isXLS ? new HSSFWorkbook() : new XSSFWorkbook();

            //生成表格
            for (String name : map.keySet()) {
                Sheet sheet = workbook.createSheet(name);
                //使用者处理表格和类之间的对应关系
                processExportSheet(workbook, sheet, map.get(name));
            }

            //将工作簿数据写入到输出流
            workbook.write(fos);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
