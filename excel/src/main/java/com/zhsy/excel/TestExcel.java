package com.zhsy.excel;

import com.zhsy.util.ExcelUtil;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.IOException;

/**
 * Created by Administrator on 16-5-25.
 */
public class TestExcel {

    public static void main(String[] args) throws IOException {
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet("test new");

        /*header*/
        String head00 = "營業人使用二聯式收銀機統一發票明細表", head06 = "印表日：2016/05/13";
        String head10 = "中國民國", head11 = "105年1月份", head13 = "稅籍編號";
        String head20 = "統一編號", head21 = "23343680", head23 = "營業人名稱", head24 = "古吉系統科技股份有限公司";
        String head30 = "開立日期", head31 = "開立發票起訖號碼", head32 = "應稅發票金額", head34 = "免稅銷售額", head35 = "誤開作廢發票號碼";

        // row1
        ExcelUtil.save(sheet, 0, 0, head00);
        ExcelUtil.merge(sheet, 0, 0, 0, 5);//
        ExcelUtil.save(sheet, 0, 6, head06);
        // row2
        ExcelUtil.save(sheet, 1, 0, head10);
        ExcelUtil.save(sheet, 1, 1, head11);
        ExcelUtil.save(sheet, 1, 3, head13);
        // row3
        ExcelUtil.save(sheet, 2, 0, head20);
        ExcelUtil.save(sheet, 2, 1, head21);
        ExcelUtil.save(sheet, 2, 3, head23);
        ExcelUtil.save(sheet, 2, 4, head24);
        // row4
        ExcelUtil.save(sheet, 3, 0, head30);
        ExcelUtil.save(sheet, 3, 1, head31);
        ExcelUtil.save(sheet, 3, 2, head32);
        ExcelUtil.merge(sheet, 3, 3, 2, 3);//
        ExcelUtil.save(sheet, 3, 4, head34);
        ExcelUtil.save(sheet, 3, 5, head35);
        ExcelUtil.merge(sheet, 3, 3, 5, 6);//

        //footer
        String foot00 = "申報單位", foot02 = "作癈分數";
        String foot12 = "空白發票起訖號碼";
        String foot22 = "銷售額及稅額計算";
        String foot32 = "項目", foot33 = "應稅", foot34 = "免/零稅銷售額";
        String foot43 = "發票總金額\n(1)", foot44 = "銷售額\n(2)=(1)X100/105", foot45 = "稅額\n(3)=(2)X5%";
        String foot61 = "（請蓋統一發票專用章）", foot62 = "本表合計";
        String foot71 = "申報日期：    年    月    日", foot72 = "本月總計";

        int begin = 8, index = begin;
        //row1
        ExcelUtil.save(sheet, index, 0, foot00);
        ExcelUtil.save(sheet, index, 1, "");
        ExcelUtil.save(sheet, index, 2, foot02);
        //row2
        index++;
        ExcelUtil.save(sheet, index, 2, foot12);
        //row3
        index++;
        ExcelUtil.save(sheet, index, 2, foot22);
        //row4
        index++;
        ExcelUtil.save(sheet, index, 2, foot32);
        ExcelUtil.save(sheet, index, 3, foot33);
        ExcelUtil.save(sheet, index, 6, foot34);
        //row5
        index++;
        ExcelUtil.save(sheet, index, 3, foot43);
        ExcelUtil.save(sheet, index, 4, foot44);
        ExcelUtil.save(sheet, index, 5, foot45);
        //row7
        index += 2;
        ExcelUtil.save(sheet, index, 1, foot61);
        ExcelUtil.save(sheet, index, 2, foot62);
        //row8
        index++;
        ExcelUtil.save(sheet, index, 1, foot71);
        ExcelUtil.save(sheet, index, 2, foot72);
        //merge
        ExcelUtil.merge(sheet, begin, index, 0, 0);
        ExcelUtil.merge(sheet, begin, begin + 5, 1, 1);
        ExcelUtil.merge(sheet, begin, begin, 2, 3);
        ExcelUtil.merge(sheet, begin, begin, 4, 6);
        //
        ExcelUtil.merge(sheet, begin + 1, begin + 1, 2, 3);
        ExcelUtil.merge(sheet, begin + 1, begin + 1, 4, 6);
        //
        ExcelUtil.merge(sheet, begin + 2, begin + 2, 2, 5);
        //
        ExcelUtil.merge(sheet, begin + 3, begin + 5, 2, 2);
        ExcelUtil.merge(sheet, begin + 3, begin + 3, 3, 5);
        ExcelUtil.merge(sheet, begin + 3, begin + 5, 6, 6);

        //style
        ExcelUtil.setBaseCss(sheet);
        ExcelUtil.setSize(sheet, 10, 0.9);//TODO

        ExcelUtil.export(workbook, "C:\\Documents and Settings\\Administrator\\桌面\\1.xls");
        workbook.close();
    }
}
