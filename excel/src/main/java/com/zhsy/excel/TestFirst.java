package com.zhsy.excel;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Random;

/**
 * Created by Administrator on 16-5-25.
 */
public class TestFirst {
    public static void actual() throws IOException {
        int columns = 7;
        //row1-column1-6,row1-column7
        String row05 = "營業人使用二聯式收銀機統一發票明細表", row06 = "印表日：2016/05/13";
        //row2
        String row10 = "中國民國", row11 = "105年1月份", row13 = "稅籍編號";
        //row3
        String row20 = "統一編號", row21 = "23343680", row23 = "營業人名稱", row24 = "古吉系統科技股份有限公司";
        //row4
        String row30 = "開立日期", row31 = "開立發票起訖號碼", row33 = "應稅發票金額", row34 = "免稅銷售額", row36 = "誤開作廢發票號碼";

        //footer
        String col0 = null;

        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet("sheet");
        //style
        HSSFCellStyle style = workbook.createCellStyle();
        style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        style.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);

        //row1
        HSSFRow row = sheet.createRow(0);
        HSSFCell cell = row.createCell(0);
        cell.setCellValue(row05);
        cell.setCellStyle(style);
        cell = row.createCell(6);
        cell.setCellValue(row06);
        cell.setCellStyle(style);
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 5));
        //row2
        row = sheet.createRow(1);
        cell = row.createCell(0);
        cell.setCellValue(row10);
        cell = row.createCell(1);
        cell.setCellValue(row11);
        cell = row.createCell(2);
        cell.setCellValue(row13);
        sheet.addMergedRegion(new CellRangeAddress(1, 1, 2, 3));
        //row3
        row = sheet.createRow(2);
        cell = row.createCell(0);
        cell.setCellValue(row20);
        cell = row.createCell(1);
        cell.setCellValue(row21);
        cell = row.createCell(2);
        cell.setCellValue(row23);
        sheet.addMergedRegion(new CellRangeAddress(2, 2, 2, 3));
        cell = row.createCell(4);
        cell.setCellValue(row24);
        //row4
        row = sheet.createRow(3);
        cell = row.createCell(0);
        cell.setCellValue(row30);
        cell = row.createCell(1);
        cell.setCellValue(row31);
        cell = row.createCell(2);
        cell.setCellValue(row33);
        sheet.addMergedRegion(new CellRangeAddress(3, 3, 2, 3));
        cell = row.createCell(4);
        cell.setCellValue(row34);
        cell = row.createCell(5);
        cell.setCellValue(row36);
        sheet.addMergedRegion(new CellRangeAddress(3, 3, 5, 6));

        //
        sheet.setAutobreaks(true);
        //output
        OutputStream out = new FileOutputStream("e:/test/1.xls");
        workbook.write(out);
        out.close();
        workbook.close();
    }

    public static void main(String[] args) throws IOException {
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet("sheet");

        Random random = new Random();
        for (int i = 0; i < 5; i++) {
            for (int j = 0; j < 3; j++) {
                HSSFRow row = sheet.createRow(i);
                HSSFCell cell = row.createCell(j);
                cell.getCellStyle().getDataFormat();
                if (j == 0) {
                    cell.setCellValue(4 + "");
                } else {
                    cell.setCellValue(new Date());
                }
            }
        }
//        HSSFDateUtil.isCellDateFormatted();

        //output
        OutputStream out = new FileOutputStream("e:/test/1.xls");
        workbook.write(out);
        out.close();
        workbook.close();
    }

    public static void base() throws Exception {
        String[] title = {"id", "name", "age"};

        List<String[]> list = new ArrayList<>();
        for (int i = 0; i < 10; i++) {
            list.add(new String[]{"no:" + i, "name" + i, 7 * i + ""});
        }

        OutputStream out = new FileOutputStream("e:/test/1.xls");

        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet("sheet");
        HSSFCellStyle style = workbook.createCellStyle();

        /*style begin*/
        style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        style.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
        // sheet.setMargin(HSSFSheet.BottomMargin, 0.7);
        // sheet.setMargin(HSSFSheet.LeftMargin, 0.3);
        // sheet.setMargin(HSSFSheet.RightMargin, 0.3);
        // sheet.setMargin(HSSFSheet.TopMargin, 0.7);
        /*style end*/

        HSSFRow row = sheet.createRow(0);// head
        for (int i = 0; i < title.length; i++) {
            HSSFCell cell = row.createCell(i);
            cell.setCellValue(title[i]);
            cell.setCellStyle(style);
        }

        for (int i = 0; i < list.size(); i++) {// body
            row = sheet.createRow(i + 1);
            for (int j = 0; j < title.length; j++) {
                HSSFCell cell = row.createCell(j);
                cell.setCellValue(list.get(i)[j]);
                cell.setCellStyle(style);
            }
        }
//        sheet.addMergedRegion(new CellRangeAddress(0, 4, 1, 2));//merge:row row col col

        workbook.write(out);
        out.close();
        workbook.close();
    }
}
