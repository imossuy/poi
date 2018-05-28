package com.dxy.demo;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFShape;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

/**
 * Created by daixiyang on 2018/5/28
 */
public class ReadFile {

    public static void main(String[] args) throws IOException {
        String baseDir = "/Users/dxy/Desktop/xls2xlsx/";
        String xlsFilePath = baseDir + "size.xls";
        String xlsxFilePath = baseDir + "size1.xlsx";

        HSSFWorkbook hssfWorkbook = new HSSFWorkbook(new FileInputStream(xlsFilePath));
        XSSFWorkbook xssfWorkbook  = new XSSFWorkbook(new FileInputStream(xlsxFilePath));

        HSSFSheet hssfSheet = hssfWorkbook.getSheetAt(0);
        XSSFSheet xssfSheet = xssfWorkbook.getSheetAt(0);

        short hssfFontHeight = hssfWorkbook.getFontAt((short) 0).

        int hssfColumnWidth = hssfSheet.getColumnWidth(0);
        float hssfColumnWidthPixel = hssfSheet.getColumnWidthInPixels(0);
        int xssfColumnWidth = xssfSheet.getColumnWidth(0);
        float xssfColumnWidthPixel = xssfSheet.getColumnWidthInPixels(0);



        Row hssfRow = hssfSheet.getRow(0);
        Row xssfRow = xssfSheet.getRow(0);

        int hssfRowHeight = hssfRow.getHeight();
        float hssfRowHeightPoint = hssfRow.getHeightInPoints();
        int xssfRowHeight = xssfRow.getHeight();
        float xssfRowHeightPoint = xssfRow.getHeightInPoints();

        HSSFClientAnchor hssfClientAnchor = (HSSFClientAnchor) hssfSheet.getDrawingPatriarch().getChildren().get(0).getAnchor();
        XSSFClientAnchor xssfClientAnchor = (XSSFClientAnchor) xssfSheet.getDrawingPatriarch().getShapes().get(0).getAnchor();

        System.out.println("-----hssf------");
        System.out.println(hssfColumnWidth);
        System.out.println(hssfColumnWidthPixel);
        System.out.println(hssfRowHeight);
        System.out.println(hssfRowHeightPoint);
        System.out.println(hssfClientAnchor.getCol1());
        System.out.println(hssfClientAnchor.getRow1());
        System.out.println(hssfClientAnchor.getCol2());
        System.out.println(hssfClientAnchor.getRow2());
        System.out.println(hssfClientAnchor.getDx1());
        System.out.println(hssfClientAnchor.getDy1());
        System.out.println(hssfClientAnchor.getDx2());
        System.out.println(hssfClientAnchor.getDy2());
        System.out.println(hssfClientAnchor.getAnchorType());
        System.out.println("---------------");

        System.out.println("-----xssf------");
        System.out.println(xssfColumnWidth);
        System.out.println(xssfColumnWidthPixel);
        System.out.println(xssfRowHeight);
        System.out.println(xssfRowHeightPoint);
        System.out.println(xssfClientAnchor.getCol1());
        System.out.println(xssfClientAnchor.getRow1());
        System.out.println(xssfClientAnchor.getCol2());
        System.out.println(xssfClientAnchor.getRow2());
        System.out.println(xssfClientAnchor.getDx1());
        System.out.println(xssfClientAnchor.getDy1());
        System.out.println(xssfClientAnchor.getDx2());
        System.out.println(xssfClientAnchor.getDy2());
        System.out.println(xssfClientAnchor.getAnchorType());

        float a = hssfClientAnchor.getDx2()/1024.0f*hssfColumnWidthPixel*9525;
        System.out.println(a);

        System.out.println((float) xssfClientAnchor.getDx2()/a);

        hssfSheet.setColumnWidth();

    }
}
