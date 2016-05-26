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
    /*
    getLastRowNum:index(start with 0)
    getLastCellNum:no(start with 1)
    */

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

    public static HSSFCell save(HSSFSheet sheet, int rowIndex, int colIndex, Object value) {
        if (sheet == null) {
            return null;
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
        return cell;
    }

    /*
    EXCEL列高度的单位是磅,Apache POI的行高度单位是缇(twip)
    http://fireinwind.iteye.com/blog/2168655
    1cm = 567twip
    */
    public static void setSize(HSSFSheet sheet, double width, double height) {//TODO
        final int CM2TWIP = 567;
        width = width * CM2TWIP;
        height = height * CM2TWIP;
        int rowMax = getRowSize(sheet), colMax = getColSize(sheet);

        if (height > 0) {
            for (int i = 0; i < rowMax; i++) {
                HSSFRow row = getRow(sheet, i);
                if (row != null) {
                    row.setHeight((short) height);
                }
            }
        }

        for (int i = 0; i < colMax; i++) {
            if (width > 0) {
                sheet.setColumnWidth(i, (short) width);
            } else {
                sheet.autoSizeColumn(i, true);
            }
        }
    }

    public static int getRowSize(HSSFSheet sheet) {
        return sheet.getLastRowNum() + 1;
    }

    public static int getColSize(HSSFSheet sheet) {
        int rowMax = getRowSize(sheet);
        int colMax = 0;
        for (int i = 0; i < rowMax; i++) {
            HSSFRow row = sheet.getRow(i);
            if (row != null) {
                colMax = Math.max(row.getLastCellNum(), colMax);
            }
        }
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

    public static boolean merge(HSSFSheet sheet, int firstRow, int lastRow, int firstCol, int lastCol) {
        return merge(sheet, firstRow, lastRow, firstCol, lastCol, null);
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

    public static void setBaseCss(HSSFSheet sheet) {
        HSSFCellStyle style = sheet.getWorkbook().createCellStyle();
        style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        style.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
        // style.setWrapText(true);
        style.setBorderTop((short) 1);
        style.setBorderRight((short) 1);
        style.setBorderBottom((short) 1);
        style.setBorderLeft((short) 1);

        setCss(sheet, style);
    }

    public static void setBorder(HSSFSheet sheet) {
        HSSFCellStyle style = sheet.getWorkbook().createCellStyle();
        style.setBorderTop((short) 1);
        style.setBorderRight((short) 1);
        style.setBorderBottom((short) 1);
        style.setBorderLeft((short) 1);

        setCss(sheet, style);
    }

    //set alignment h&v center in all cell default
    public static void setAlignment(HSSFSheet sheet) {
        HSSFCellStyle style = sheet.getWorkbook().createCellStyle();
        style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        style.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);

        setCss(sheet, style);
    }

    public static void setCss(HSSFSheet sheet, HSSFCellStyle style) {
        int rowMax = getRowSize(sheet), colMax = getColSize(sheet);
        // System.out.println(rowMax + "-" + colMax);
        for (int i = 0; i < rowMax; i++) {
            for (int j = 0; j < colMax; j++) {
                HSSFCell cell = getCell(sheet, i, j);
                if (cell == null) {
                    cell = save(sheet, i, j, null);
                }
                cell.setCellStyle(style);
            }
        }
    }
}
