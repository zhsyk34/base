package com.zhsy;

import com.zhsy.util.ExcelUtil;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.IOException;

/**
 * Created by Administrator on 16-5-25.
 */
public class Update {

    public static void main(String[] args) throws IOException {
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet("test new");

        for (int row = 0; row < 5; row++) {
            for (int col = 0; col < 4; col++) {
                if (row == 0 & col > 1) {
                    continue;
                }
                ExcelUtil.save(sheet, row, col, (row + 1) + "-" + (col + 1));
            }
        }

        ExcelUtil.setAlignment(sheet);
//        ExcelUtil.setSize(sheet, 20, 20);//TODO
        ExcelUtil.merge(sheet, 0, 0, 3, 11, "merge-data");

        ExcelUtil.export(workbook, "C:\\Documents and Settings\\Administrator\\桌面\\1.xls");
        workbook.close();
    }
}
