package org.xli.excel;

import org.apache.poi.ss.usermodel.*;
import org.xli.excel.entity.StudentBaseInfo;

import java.io.File;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

/**
 * @author xli
 * @Description
 * @Date 创建于 2019/1/23 19:36
 */
public class ExcelMappingStudentBaseInfo extends AbstractExcelMappingPOJO<StudentBaseInfo> {
    private static final int DATA_START = 1;

    @Override
    public void processImportSheet(Sheet sheet, List<StudentBaseInfo> sheetList) {
        Iterator<Row> rowIterator = sheet.iterator();

        //迭代表格的行
        int i = 0;
        while (rowIterator.hasNext()) {
            //行对象
            Row row = rowIterator.next();

            //跳过表头并获取单元格的值
            if (row.getRowNum() >= DATA_START) {
                StudentBaseInfo o = new StudentBaseInfo();
                o.setSerialNum(Integer.parseInt(row.getCell(i++).getStringCellValue()));
                o.setId(String.valueOf(row.getCell(i++).getNumericCellValue()));
                o.setName(row.getCell(i++).getStringCellValue());
                o.setClassName(row.getCell(i++).getStringCellValue());
                sheetList.add(o);
                i = 0;
            }
        }
    }

    @Override
    public void processExportSheet(Workbook workbook, Sheet sheet, List<StudentBaseInfo> sheetList) {
        //创建样式
        CellStyle cellStyle = newDefaultCellStyle(workbook);

        //设置表头
        int i = 0;
        Row head = sheet.createRow(i++);
        head.createCell(0).setCellValue("班内序号");
        head.createCell(1).setCellValue("学号");
        head.createCell(2).setCellValue("姓名");
        head.createCell(3).setCellValue("班级");
        //为表头设置样式
        for (Cell c: head) {
            c.setCellStyle(cellStyle);
        }

        for (; i <= sheetList.size(); i++) {
            Row row = sheet.createRow(i);
            row.setRowStyle(cellStyle);
            StudentBaseInfo studentBaseInfo = sheetList.get(i-1);

            row.createCell(0).setCellValue(studentBaseInfo.getSerialNum());
            row.createCell(1).setCellValue(studentBaseInfo.getId());
            row.createCell(2).setCellValue(studentBaseInfo.getName());
            row.createCell(3).setCellValue(studentBaseInfo.getClassName());

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
        font.setFontHeightInPoints((short) 16);

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
        ExcelMappingStudentBaseInfo excelToWinningStudent = new ExcelMappingStudentBaseInfo();
        Map<String, List<StudentBaseInfo>> map = excelToWinningStudent.importExcel(
                new File("D:\\xli\\Work\\IDEA-Storage\\poi-demo\\src\\main\\resources\\b.xlsx"));

        for (String key : map.keySet()) {
            for (StudentBaseInfo studentBaseInfo: map.get(key)) {
                System.out.println(studentBaseInfo);
            }
        }

        excelToWinningStudent.exportExcel(map,
                new File("D:\\xli\\Work\\IDEA-Storage\\poi-demo\\src\\main\\resources\\b(copy).xlsx"));
    }
}
