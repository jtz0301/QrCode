package com.jz.util;

import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import com.jz.entity.Admin;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.math.BigDecimal;
import java.net.URLEncoder;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.LinkedHashMap;
import java.util.Map;

/**
 * Created by Cheung on 2017/12/19.
 * <p>
 * Apache POI操作Excel对象
 * HSSF：操作Excel 2007之前版本(.xls)格式,生成的EXCEL不经过压缩直接导出
 * XSSF：操作Excel 2007及之后版本(.xlsx)格式,内存占用高于HSSF
 * SXSSF:从POI3.8 beta3开始支持,基于XSSF,低内存占用,专门处理大数据量(建议)。
 * <p>
 * 注意:
 * 值得注意的是SXSSFWorkbook只能写(导出)不能读(导入)
 * <p>
 * 说明:
 * .xls格式的excel(最大行数65536行,最大列数256列)
 * .xlsx格式的excel(最大行数1048576行,最大列数16384列)
 */
public class ExcelUtil {

    public static final String DEFAULT_DATE_PATTERN = "yyyy-MM-dd";// 默认日期格式
    public static final int DEFAULT_COLUMN_WIDTH = 17;// 默认列宽

    public static void exportExcel(HttpServletResponse response, LinkedHashMap<String, String> headMap, String fileName, JSONArray dataArray) throws Exception {
        // 告诉浏览器用什么软件可以打开此文件
        response.setContentType("application/vnd.ms-excel");
        // 下载文件的默认名称
        response.setHeader("Content-Disposition", "attachment;filename=" + URLEncoder.encode(fileName, "utf-8"));
        exportExcel(headMap, dataArray, response.getOutputStream());
    }

    /**
     * 导出Excel(.xlsx)格式
     *
     * @param headMap   表格头信息集合
     * @param dataArray 数据数组
     * @param os        文件输出流
     */
    public static void exportExcel(LinkedHashMap<String, String> headMap, JSONArray dataArray, OutputStream os) {
        String datePattern = DEFAULT_DATE_PATTERN;
        int minBytes = DEFAULT_COLUMN_WIDTH;

        /**
         * 声明一个工作薄
         */
        SXSSFWorkbook workbook = new SXSSFWorkbook();// 大于1000行时会把之前的行写入硬盘
        workbook.setCompressTempFiles(true);

        // head样式
        CellStyle headerStyle = workbook.createCellStyle();
        
        headerStyle.setAlignment(HorizontalAlignment.CENTER);// 水平居中
        headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);// 垂直居中
        headerStyle.setFillForegroundColor(IndexedColors.PALE_BLUE.getIndex());// 设置颜色
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);// 前景色纯色填充
        headerStyle.setBorderTop(BorderStyle.THIN);
        headerStyle.setBorderRight(BorderStyle.THIN);
        headerStyle.setBorderBottom(BorderStyle.THIN);
        headerStyle.setBorderLeft(BorderStyle.THIN);
        Font headerFont = workbook.createFont();
        headerFont.setFontHeightInPoints((short) 12);
        headerFont.setBold(true);
        headerStyle.setFont(headerFont);

        // 单元格样式
        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        cellStyle.setBorderTop(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBorderLeft(BorderStyle.THIN);
        Font cellFont = workbook.createFont();
        cellFont.setBold(false);
        cellStyle.setFont(cellFont);

        /**
         * 生成一个(带名称)表格
         */
        SXSSFSheet sheet = (SXSSFSheet) workbook.createSheet();
        //设置页名称
        workbook.setSheetName(0, "sheet");
        sheet.createFreezePane(0, 1, 1, 1);// (单独)冻结前三行

        /**
         * 生成head相关信息+设置每列宽度
         */
        int[] colWidthArr = new int[headMap.size()];// 列宽数组
        String[] headKeyArr = new String[headMap.size()];// headKey数组
        String[] headValArr = new String[headMap.size()];// headVal数组
        int i = 0;
        for (Map.Entry<String, String> entry : headMap.entrySet()) {
            headKeyArr[i] = entry.getKey();
            headValArr[i] = entry.getValue();
            int bytes = headKeyArr[i].getBytes().length;
            colWidthArr[i] = bytes < minBytes ? minBytes : bytes;
            i++;
        }


        /**
         * 遍历数据集合，产生Excel行数据
         */
        int rowIndex = 0;
        for (Object obj : dataArray) {
            // 生成title+head信息
            if (rowIndex == 0) {
                SXSSFRow headerRow = (SXSSFRow) sheet.createRow(0);// head行
                for (int j = 0; j < headValArr.length; j++) {
                    headerRow.createCell(j).setCellValue(headValArr[j]);
                    headerRow.getCell(j).setCellStyle(headerStyle);
                }
                rowIndex = 1;
            }

            JSONObject jo = (JSONObject) JSONObject.toJSON(obj);
            // 生成数据
            SXSSFRow dataRow = (SXSSFRow) sheet.createRow(rowIndex);// 创建行
            for (int k = 0; k < headKeyArr.length; k++) {
                SXSSFCell cell = (SXSSFCell) dataRow.createCell(k);// 创建单元格
                Object o = jo.get(headKeyArr[k]);
                String cellValue = "";

                if (o == null) {
                    cellValue = "";
                } else if (o instanceof Date) {
                    cellValue = new SimpleDateFormat(datePattern).format(o);
                } else if (o instanceof Float || o instanceof Double) {
                    cellValue = new BigDecimal(o.toString()).setScale(2, BigDecimal.ROUND_HALF_UP).toString();
                } else {
                    cellValue = o.toString();
                }
                cell.setCellValue(cellValue);
                colWidthArr[k] = colWidthArr[k] < cellValue.length() ? cellValue.length() : colWidthArr[k];
                sheet.setColumnWidth(k, (colWidthArr[k]+3) * 256);// 设置列宽
                cell.setCellStyle(cellStyle);
            }
            rowIndex++;
        }

        try {
            workbook.write(os);
            workbook.dispose();// 释放workbook所占用的所有windows资源
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                if (os != null) {
//                    os.flush();// 刷新此输出流并强制将所有缓冲的输出字节写出
                    os.close();// 关闭流
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }


    public static void main(String[] args) throws IOException {
        // 模拟10W条数据
        int count = 1000;
        JSONArray studentArray = new JSONArray();
        for (int i = 0; i < count; i++) {
            //AdminUser s = new AdminUser();
        	Admin s=new Admin();
            s.setName("POI" + i);
            studentArray.add(s);
        }

        LinkedHashMap<String, String> headMap = new LinkedHashMap<>();
        headMap.put("name", "姓名");
//        headMap.put("age", "年龄");
//        headMap.put("birthday", "生日");
//        headMap.put("height", "身高");
//        headMap.put("weight", "体重");
//        headMap.put("sex", "性别");
//        titleList.add(titleMap);
//        titleList.add(headMap);


        File file = new File("D://ExcelExportDemo/");
        if (!file.exists()) file.mkdir();// 创建该文件夹目录
        OutputStream os = null;
        Date date = new Date();
        try {
            // .xlsx格式
            os = new FileOutputStream(file.getAbsolutePath() + "/" + date.getTime() + ".xlsx");
            System.out.println("正在导出xlsx...");
            ExcelUtil.exportExcel(headMap, studentArray, os);
            System.out.println("导出完成...共" + count + "条数据,用时" + (System.nanoTime()- date.getTime()) + "ms");
            System.out.println("文件路径：" + file.getAbsolutePath() + "/" + date.getTime() + ".xlsx");
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            os.close();
        }

    }
}
