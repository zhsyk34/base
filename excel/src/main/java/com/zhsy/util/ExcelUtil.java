package com.zhsy.util;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * Created by Administrator on 16-5-25.
 */
public class ExcelUtil {

    public static HSSFRow getRow(HSSFSheet sheet, int rowIndex) {
        return sheet == null ? null : sheet.getLastRowNum() < rowIndex ? null : sheet.getRow(rowIndex);
    }

    public static HSSFCell getCell(HSSFRow row, int colIndex) {
        return row == null ? null : row.getLastCellNum() < colIndex ? null : row.getCell(colIndex);
    }

    public static HSSFCell getCell(HSSFSheet sheet, int rowIndex, int colIndex) {
        HSSFRow row = getRow(sheet, rowIndex);
        return row == null ? null : getCell(row, colIndex);
    }

    public static void save(HSSFSheet sheet, int rowIndex, int colIndex, Object value) {
        if (sheet == null) {
            return;
        }
        HSSFRow row = getRow(sheet, rowIndex);
        if (row == null) {
            row = sheet.createRow(rowIndex);
        }

        HSSFCell cell = getCell(row, colIndex);
        if (cell == null) {
            cell = row.createCell(colIndex);
        }
        setValue(cell, value);
    }

    //set alignment h&v center in all cell default
    public static void setAlignment(HSSFSheet sheet) {
        HSSFCellStyle style = sheet.getWorkbook().createCellStyle();
        style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        style.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);

        int rowMax = sheet.getLastRowNum();
        System.out.println(rowMax);
        for (int i = 0; i <= rowMax; i++) {
            HSSFRow row = sheet.getRow(i);
            int colMax = row.getLastCellNum();
            System.out.println(colMax);
            for (int j = 0; j < colMax; j++) {
                row.getCell(j).setCellStyle(style);
            }
        }
    }

    public static void setSize(HSSFSheet sheet, int width, int height) {//TODO
        final int DPI = 20;
        width = width * DPI;
        height = height * DPI;
        int rowMax = sheet.getLastRowNum();
        for (int i = 0; i < rowMax; i++) {
            sheet.getRow(i).setHeight((short) height);
        }

        int colMax = getLastColNum(sheet);
        for (int i = 0; i < colMax; i++) {
            sheet.setColumnWidth(i, width);
        }
    }

    public static int getLastColNum(HSSFSheet sheet) {
        int rowMax = sheet.getLastRowNum();
        int colMax = 0;
        for (int i = 0; i < rowMax; i++) {
            colMax = Math.max(sheet.getRow(i).getLastCellNum(), colMax);
        }
        System.out.printf("row:%d,col:%d", rowMax, colMax);
        return colMax;
    }

    public static boolean merge(HSSFSheet sheet, int firstRow, int lastRow, int firstCol, int lastCol, Object value) {
        if (sheet == null) {
            return false;
        }
        CellRangeAddress region = new CellRangeAddress(firstRow, lastRow, firstCol, lastCol);
        sheet.addMergedRegion(region);
        save(sheet, firstRow, firstCol, value);
        return true;
    }

    public static void setValue(HSSFCell cell, Object value) {//TODO
        if (value != null) {
            if (value instanceof String) {
                cell.setCellValue((String) value);
            }
            if (value instanceof Double) {
                cell.setCellValue((Double) value);
            }
        }
    }

    public static boolean merge(HSSFSheet sheet, int firstRow, int lastRow, int firstCol, int lastCol) {
        return merge(sheet, firstRow, lastRow, firstCol, lastCol, null);
    }

    public static boolean export(HSSFWorkbook workbook, String path) {
        FileOutputStream output = null;
        try {
            new File(path).createNewFile();
            output = new FileOutputStream(path);
            output.flush();
            workbook.write(output);
            return true;
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            if (output != null) {
                try {
                    output.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
        return false;
    }
}
