package com.dxy.demo;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.usermodel.charts.LineChartData;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

/**
 * @author Created by daixiyang on 2018/5/9
 */
public class QuickGuide {

    public static void main(String[] args) throws IOException, InvalidFormatException {

        String baseDir = "/Users/dxy/Desktop/xls2xlsx/";

        //Source File Path
        String xlsFilePath = baseDir + "chart.xls";

        InputStream inputStream = new FileInputStream(xlsFilePath);
        Workbook wb = new HSSFWorkbook(inputStream);

        Sheet sheet = wb.getSheetAt(1);
        sheet.getDrawingPatriarch();
    }

}
