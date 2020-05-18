package org.xli.excel;

import org.apache.poi.ss.usermodel.*;
import org.xli.excel.entity.WinningStudent;

import java.io.File;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

/**
 * @author xli
 * @Description
 * @Date 创建于 2019/1/23 19:36
 */
public class ExcelMappingWinningStudent extends AbstractExcelMappingPOJO<WinningStudent> {
    private static final int DATA_START = 1;

    @Override
    public void processImportSheet(Sheet sheet, List<WinningStudent> sheetList) {
        Iterator<Row> rowIterator = sheet.iterator();

        //迭代表格的行
        int i = 0;
        while (rowIterator.hasNext()) {
            //行对象
            Row row = rowIterator.next();

            //跳过表头并获取单元格的值
            if (row.getRowNum() >= DATA_START) {
                WinningStudent o = new WinningStudent();
                o.setSerialNum((int) row.getCell(i++).getNumericCellValue());
                o.setItem(row.getCell(i++).getStringCellValue());
                o.setName(row.getCell(i++).getStringCellValue());
                o.setClassName(row.getCell(i++).getStringCellValue());
                sheetList.add(o);
                i = 0;
            }
        }
    }

    @Override
    public void processExportSheet(Workbook workbook, Sheet sheet, List<WinningStudent> sheetList) {
        //创建样式
        CellStyle cellStyle = newDefaultCellStyle(workbook);

        //设置标题
        int i = 0;
        Row head = sheet.createRow(i++);
        head.createCell(0).setCellValue("序号");
        head.createCell(1).setCellValue("项目");
        head.createCell(2).setCellValue("姓名");
        head.createCell(3).setCellValue("班级");
        //为表头设置样式
        for (Cell c: head) {
            c.setCellStyle(cellStyle);
        }

        for (; i <= sheetList.size(); i++) {
            Row row = sheet.createRow(i);
            row.setRowStyle(cellStyle);
            WinningStudent winningStudent = sheetList.get(i-1);

            row.createCell(0).setCellValue(winningStudent.getSerialNum());
            row.createCell(1).setCellValue(winningStudent.getItem());
            row.createCell(2).setCellValue(winningStudent.getName());
            row.createCell(3).setCellValue(winningStudent.getClassName());

            //为每个单元格设置样式
            for (Cell c: row) {
                c.setCellStyle(cellStyle);
            }
        }
    }

    //样式
    public CellStyle newDefaultCellStyle(Workbook workbook) {
        //设置字体
        Font font = workbook.createFont();
        font.setFontName("宋体");
        font.setFontHeightInPoints((short) 12);

        //设置样式
        CellStyle cellStyle = workbook.createCellStyle();
        //水平居中
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        //垂直居中
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        //将字体应用到样式
        cellStyle.setFont(font);

        return cellStyle;
    }

    public static void main(String[] args) {
        ExcelMappingWinningStudent excelToWinningStudent = new ExcelMappingWinningStudent();
        Map<String, List<WinningStudent>> map = excelToWinningStudent.importExcel(
                new File("D:\\xli\\Work\\IDEA-Storage\\poi-demo\\src\\main\\resources\\a.xls"));

        for (String key : map.keySet()) {
            for (WinningStudent winningStudent: map.get(key)) {
                System.out.println(winningStudent);
            }
        }

        excelToWinningStudent.exportExcel(map,
                new File("D:\\xli\\Work\\IDEA-Storage\\poi-demo\\src\\main\\resources\\a(copy).xls"));
    }
}
