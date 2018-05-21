package com.dxy.demo;

import org.apache.poi.hssf.model.InternalSheet;
import org.apache.poi.hssf.record.AutoFilterInfoRecord;
import org.apache.poi.hssf.record.ColumnInfoRecord;
import org.apache.poi.hssf.record.NameRecord;
import org.apache.poi.hssf.record.RecordBase;
import org.apache.poi.hssf.record.aggregates.ColumnInfoRecordsAggregate;
import org.apache.poi.hssf.usermodel.HSSFAnchor;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFComment;
import org.apache.poi.hssf.usermodel.HSSFEvaluationWorkbook;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.usermodel.HSSFPatriarch;
import org.apache.poi.hssf.usermodel.HSSFPicture;
import org.apache.poi.hssf.usermodel.HSSFPictureData;
import org.apache.poi.hssf.usermodel.HSSFPolygon;
import org.apache.poi.hssf.usermodel.HSSFShape;
import org.apache.poi.hssf.usermodel.HSSFShapeGroup;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFSimpleShape;
import org.apache.poi.hssf.usermodel.HSSFTextbox;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.PictureData;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.SimpleShape;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.PaneInformation;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFChildAnchor;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFHyperlink;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFShape;
import org.apache.poi.xssf.usermodel.XSSFShapeGroup;
import org.apache.poi.xssf.usermodel.XSSFSimpleShape;
import org.apache.poi.xssf.usermodel.XSSFTextBox;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import static com.sun.org.apache.xml.internal.security.keys.keyresolver.KeyResolver.iterator;

/**
 * @author daixiyang
 * @created 2018/5/8
 */
public class FileConversionXlsToXlsx {

    public static void main(String[] args) throws FileNotFoundException {
        String baseDir = "/Users/dxy/Desktop/xls2xlsx/";

        String xlsFilePath = baseDir + "chart.xls";
        String xlsxFilePath = convertXls2Xlsx(xlsFilePath);
    }

    private static String convertXls2Xlsx(String xlsFilePath) {
        Map cellStyleMap = new HashMap();
        String xlsxFilePath = null;
        Workbook workbookIn = null;
        File xlsxFile = null;
        Workbook workbookOut = null;
        OutputStream out = null;
        String xlsx = ".xlsx";
        try {
            InputStream inputStream = new FileInputStream(xlsFilePath);
            xlsxFilePath = xlsFilePath.substring(0, xlsFilePath.lastIndexOf('.')) + xlsx;
            workbookIn = new HSSFWorkbook(inputStream);
            xlsxFile = new File(xlsxFilePath);
            if (xlsxFile.exists()) {
                xlsxFile.delete();
            }
            workbookOut = new XSSFWorkbook();
            int sheetCnt = workbookIn.getNumberOfSheets();

            CreationHelper factory = workbookOut.getCreationHelper();




            for (int i = 0; i < sheetCnt; i++) {
                Sheet sheetIn = workbookIn.getSheetAt(i);
                Sheet sheetOut = workbookOut.createSheet(sheetIn.getSheetName());

                copySheetProperties(i,sheetIn,sheetOut,workbookIn,workbookOut);


                HSSFPatriarch drawingIn = (HSSFPatriarch) sheetIn.getDrawingPatriarch();
                XSSFDrawing drawingOut = (XSSFDrawing) sheetOut.createDrawingPatriarch();

                Iterator<Row> rowIt = sheetIn.rowIterator();
                while (rowIt.hasNext()) {
                    Row rowIn = rowIt.next();
                    Row rowOut = sheetOut.createRow(rowIn.getRowNum());
                    copyRowProperties(rowOut, rowIn,sheetIn,sheetOut,drawingIn,drawingOut,cellStyleMap);
                }

                if(drawingIn != null){
                    for (HSSFShape hssfShape : drawingIn) {
                        if(hssfShape instanceof HSSFPicture){
                            //图片
                            HSSFPicture picture = (HSSFPicture) hssfShape;
                            HSSFClientAnchor hssfClientAnchor = picture.getClientAnchor();
                            XSSFClientAnchor xssfClientAnchor = drawingOut.createAnchor(
                                    hssfClientAnchor.getDx1(),
                                    hssfClientAnchor.getDy1(),
                                    hssfClientAnchor.getDx2(),
                                    hssfClientAnchor.getDy2(),
                                    hssfClientAnchor.getCol1(),
                                    hssfClientAnchor.getRow1(),
                                    hssfClientAnchor.getCol2(),
                                    hssfClientAnchor.getRow2());
                            HSSFPictureData hssfPictureData = picture.getPictureData();
                            int pictureIndex = workbookOut.addPicture(
                                    hssfPictureData.getData(),hssfPictureData.getFormat());
                            drawingOut.createPicture(xssfClientAnchor,pictureIndex);

                        }else if(hssfShape instanceof HSSFPolygon){
                            HSSFPolygon hssfPolygon = (HSSFPolygon) hssfShape;
                            HSSFAnchor hssfAnchor = hssfPolygon.getAnchor();
                            XSSFClientAnchor xssfClientAnchor = drawingOut.createAnchor(
                                    hssfAnchor.getDx1(),
                                    hssfAnchor.getDy1(),
                                    hssfAnchor.getDx2(),
                                    hssfAnchor.getDy2(),
                                    0,0,0,0);

                            drawingOut.createSimpleShape(xssfClientAnchor);
                        }else if(hssfShape instanceof HSSFComment){
                            HSSFComment hssfComment = (HSSFComment) hssfShape;
                            ClientAnchor clientAnchor = hssfComment.getClientAnchor();
                            XSSFClientAnchor xssfClientAnchor = drawingOut.createAnchor(
                                    clientAnchor.getDx1(),
                                    clientAnchor.getDy1(),
                                    clientAnchor.getDx2(),
                                    clientAnchor.getDy2(),
                                    clientAnchor.getCol1(),
                                    clientAnchor.getRow1(),
                                    clientAnchor.getCol2(),
                                    clientAnchor.getRow2());

//                            XSSFComment xssfComment = drawingOut.createCellComment(xssfClientAnchor);
//                            xssfComment.setAuthor(hssfComment.getAuthor());
//                            xssfComment.setString(hssfComment.getString().getString());
//                            xssfComment.setAddress(hssfComment.getAddress());

                        }else if(hssfShape instanceof HSSFTextbox){
                            HSSFTextbox hssfTextbox = (HSSFTextbox) hssfShape;
                            HSSFAnchor hssfAnchor = hssfTextbox.getAnchor();

                            XSSFClientAnchor xssfClientAnchor = drawingOut.createAnchor(
                                    hssfAnchor.getDx1(),
                                    hssfAnchor.getDy1(),
                                    hssfAnchor.getDx2(),
                                    hssfAnchor.getDy2(),
                                    0,
                                    0,
                                    0,
                                    0);

                            XSSFTextBox xssfTextBox = drawingOut.createTextbox(xssfClientAnchor);
                            hssfTextbox2XssfTextbox(factory, hssfTextbox, xssfTextBox);

                        }else if(hssfShape instanceof HSSFShapeGroup){
                            HSSFShapeGroup hssfShapeGroup = (HSSFShapeGroup) hssfShape;

                            XSSFShapeGroup xssfShapeGroup = drawingOut.createGroup((XSSFClientAnchor) factory.createClientAnchor());
                            xssfShapeGroup.setCoordinates(
                                    hssfShapeGroup.getX1(),
                                    hssfShapeGroup.getY1(),
                                    hssfShapeGroup.getX2(),
                                    hssfShapeGroup.getY2());
                            int fillColor = hssfShapeGroup.getFillColor();
//                            xssfShapeGroup.setFillColor(
//                                    fillColor & 0x000000ff,
//                                    (fillColor & 0x0000ff00) >> 8,
//                                    (fillColor & 0x00ff0000) >> 8);
//                            xssfShapeGroup.setLineStyle(hssfShapeGroup.getLineStyle());
//                            int lineStyleColor = hssfShapeGroup.getLineStyleColor();
//                            xssfShapeGroup.setLineStyleColor(
//                                    lineStyleColor & 0x000000ff,
//                                    (lineStyleColor & 0x0000ff00) >> 8,
//                                    (lineStyleColor & 0x00ff0000) >> 8);
//
//                            xssfShapeGroup.setLineWidth(hssfShapeGroup.getLineWidth());
//                            xssfShapeGroup.setNoFill(hssfShapeGroup.isNoFill());

                            for(HSSFShape childShape : hssfShapeGroup.getChildren()){
                                if(childShape instanceof HSSFTextbox){
                                    HSSFTextbox hssfTextbox = (HSSFTextbox) childShape;
                                    HSSFAnchor hssfAnchor = hssfTextbox.getAnchor();

//                                    XSSFTextBox xssfTextBox = xssfShapeGroup.createTextbox(
//                                            new XSSFChildAnchor(
//                                                    hssfAnchor.getDx1(),
//                                                    hssfAnchor.getDy1(),
//                                                    hssfAnchor.getDx2(),
//                                                    hssfAnchor.getDy2()));
//                                    hssfTextbox2XssfTextbox(factory, hssfTextbox, xssfTextBox);

                                }
                            }
                        }else if(HSSFSimpleShape.class.isInstance(hssfShape)){

                            HSSFSimpleShape hssfSimpleShape = (HSSFSimpleShape) hssfShape;
                            HSSFClientAnchor hssfAnchor = (HSSFClientAnchor) hssfSimpleShape.getAnchor();
                            XSSFSimpleShape xssfSimpleShape = drawingOut.createSimpleShape(
                                    drawingOut.createAnchor(
                                            hssfAnchor.getDx1(),
                                            hssfAnchor.getDy1(),
                                            hssfAnchor.getDx2(),
                                            hssfAnchor.getDy2(),
                                            hssfAnchor.getCol1(),
                                            hssfAnchor.getRow1(),
                                            hssfAnchor.getCol2(),
                                            hssfAnchor.getRow2()));
//                            xssfSimpleShape.setBottomInset(hssfSimpleShape.);
//                            xssfSimpleShape.setLeftInset();
//                            xssfSimpleShape.setRightInset();
                            xssfSimpleShape.setShapeType(5);

                            try {
//                                String str = hssfSimpleShape.getString().getString();
//                                xssfSimpleShape.setText((XSSFRichTextString) factory.createRichTextString(str));
                            }catch (Exception e){
                                e.printStackTrace();
                            }
//                            xssfSimpleShape.setTextAutofit();
//                            xssfSimpleShape.setTextDirection();
//                            xssfSimpleShape.setTextHorizontalOverflow();
//                            xssfSimpleShape.setTextVerticalOverflow();
//                            xssfSimpleShape.setTopInset();˙
//                            xssfSimpleShape.setVerticalAlignment();
//                            xssfSimpleShape.setWordWrap();
//                            int fillColor = hssfSimpleShape.getFillColor();
//                            xssfSimpleShape.setFillColor(
//                                    fillColor & 0x000000ff,
//                                    (fillColor & 0x0000ff00) >> 8,
//                                    (fillColor & 0x00ff0000) >> 8);
//                            xssfSimpleShape.setLineStyle(hssfSimpleShape.getLineStyle());
//                            int lineStyleColor = hssfSimpleShape.getLineStyleColor();
//                            xssfSimpleShape.setLineStyleColor(
//                                    fillColor & 0x000000ff,
//                                    (fillColor & 0x0000ff00) >> 8,
//                                    (fillColor & 0x00ff0000) >> 8);
//                            xssfSimpleShape.setLineWidth(hssfSimpleShape.getLineWidth());
//                            xssfSimpleShape.setNoFill(hssfSimpleShape.isNoFill());



                        }
                    }

                }

//                copySheetPropertiesAfter(i,sheetIn,sheetOut,workbookIn,workbookOut);

            }

            copyWorkbookProperties(workbookIn,workbookOut);

            out = new BufferedOutputStream(new FileOutputStream(xlsxFile));
            workbookOut.write(out);
        } catch (Exception ex) {
            ex.printStackTrace();
            xlsxFilePath = null;
        } finally {
            try {
                if (workbookOut != null) {
                    workbookOut.close();
                }
                if (workbookIn != null) {
                    workbookIn.close();
                }
                if (out != null) {
                    out.close();
                }
            } catch (Exception ex) {
                ex.printStackTrace();
            }
        }
        return xlsxFilePath;
    }



    private static List<ColumnInfoRecord> getColumnInfoRecordList(ColumnInfoRecordsAggregate columnInfoRecordsAggregate) {
        try {
            Field field = ColumnInfoRecordsAggregate.class.getDeclaredField("records");
            field.setAccessible(true);
            return (List<ColumnInfoRecord>) field.get(columnInfoRecordsAggregate);

        } catch (NoSuchFieldException e) {
//
            e.printStackTrace();
        } catch (IllegalAccessException e) {
//
            e.printStackTrace();
        }
        return null;
    }

    private static void copyWorkbookProperties(Workbook workbookIn, Workbook workbookOut) {
        workbookOut.setMissingCellPolicy(workbookIn.getMissingCellPolicy());

//        workbookOut.setHidden(workbookIn.isHidden());  //not implement
        workbookOut.setActiveSheet(workbookIn.getActiveSheetIndex());
        workbookOut.setFirstVisibleTab(workbookIn.getFirstVisibleTab());
        workbookOut.setForceFormulaRecalculation(workbookIn.getForceFormulaRecalculation());
        workbookOut.setMissingCellPolicy(workbookIn.getMissingCellPolicy());
    }

    private static void copySheetProperties(int index, Sheet sheetIn, Sheet sheetOut, Workbook workbookIn, Workbook workbookOut) {
//        workbookOut.setPrintArea();
//        workbookOut.setSelectedTab();
//        workbookOut.setSheetHidden(index,workbookIn.isSheetHidden(index));

        HSSFEvaluationWorkbook hssfEvaluationWorkbook = HSSFEvaluationWorkbook.create((HSSFWorkbook) workbookIn);
//        hssfEvaluationWorkbook.getNameXPtg(NameRecord.BUILTIN_FILTER_DB,null);


        workbookOut.setSheetVisibility(index,workbookIn.getSheetVisibility(index));

        InternalSheet internalSheet = getInternalSheet(sheetIn);
        List<RecordBase> recordBaseList = internalSheet.getRecords();
        for(RecordBase recordBase : recordBaseList){
            if(recordBase instanceof ColumnInfoRecordsAggregate){
                ColumnInfoRecordsAggregate columnInfoRecordsAggregate = (ColumnInfoRecordsAggregate) recordBase;
                List<ColumnInfoRecord> columnInfoRecordList = getColumnInfoRecordList(columnInfoRecordsAggregate);
                for(ColumnInfoRecord columnInfoRecord : columnInfoRecordList){
                    int colfirst = columnInfoRecord.getFirstColumn();
                    int collast = columnInfoRecord.getLastColumn();
                    if(columnInfoRecord.getCollapsed()){
                        for(int i=colfirst; i<=collast; i++){
                            sheetOut.setColumnGroupCollapsed(i,true);
                        }
                    }
                    if(columnInfoRecord.getHidden()){
                        for(int i=colfirst; i<=collast; i++){
                            sheetOut.setColumnHidden(i,true);
                        }
                    }
                    sheetOut.setColumnWidth(colfirst,columnInfoRecord.getColumnWidth());
                }
            }else if(AutoFilterInfoRecord.class.isInstance(recordBase)){
                AutoFilterInfoRecord autoFilterInfoRecord = (AutoFilterInfoRecord) recordBase;
//                sheetOut.setAutoFilter(CellRangeAddress);
//                sheetIn.setAutoFilter()

            }else if(NameRecord.class.isInstance(recordBase)){
                NameRecord nameRecord = (NameRecord) recordBase;
//                nameRecord.
            }

        }


//        sheetOut.setDefaultColumnStyle();
//        sheetOut.setDefaultColumnWidth();

        sheetOut.setActiveCell(sheetIn.getActiveCell());
//        sheetOut.setArrayFormula();
        sheetOut.setAutobreaks(sheetIn.getAutobreaks());

        for(int i : sheetIn.getColumnBreaks()){
            sheetOut.setColumnBreak(i);
        }

        sheetOut.setDefaultRowHeight(sheetIn.getDefaultRowHeight());
        sheetOut.setDefaultRowHeightInPoints(sheetIn.getDefaultRowHeightInPoints());
//        sheetOut.setDisplayFormulas();
        sheetOut.setDisplayGridlines(sheetIn.isDisplayGridlines());
        sheetOut.setDisplayGuts(sheetIn.getDisplayGuts());
//        sheetOut.setDisplayRowColHeadings();
//        sheetOut.setDisplayZeros();
        sheetOut.setFitToPage(sheetIn.getFitToPage());
        sheetOut.setForceFormulaRecalculation(sheetIn.getForceFormulaRecalculation());
        sheetOut.setHorizontallyCenter(sheetIn.getHorizontallyCenter());
        sheetOut.setMargin(Sheet.LeftMargin,sheetIn.getMargin(Sheet.LeftMargin));
        sheetOut.setMargin(Sheet.RightMargin,sheetIn.getMargin(Sheet.RightMargin));
        sheetOut.setMargin(Sheet.BottomMargin,sheetIn.getMargin(Sheet.BottomMargin));
        sheetOut.setMargin(Sheet.TopMargin,sheetIn.getMargin(Sheet.TopMargin));
        sheetOut.setMargin(Sheet.HeaderMargin,sheetIn.getMargin(Sheet.HeaderMargin));
        sheetOut.setMargin(Sheet.FooterMargin,sheetIn.getMargin(Sheet.FooterMargin));
//        sheetOut.setPrintGridlines();
//        sheetOut.setPrintRowAndColumnHeadings();
        sheetOut.setRepeatingColumns(sheetIn.getRepeatingColumns());
        sheetOut.setRepeatingRows(sheetIn.getRepeatingRows());

        for(int i : sheetIn.getRowBreaks()){
            sheetOut.setColumnBreak(i);
        }
//        sheetOut.setRowGroupCollapsed();
        sheetOut.setRowSumsBelow(sheetIn.getRowSumsBelow());
        sheetOut.setRowSumsRight(sheetIn.getRowSumsRight());
//        sheetOut.setRightToLeft();
        sheetOut.setSelected(sheetIn.isSelected());
        sheetOut.setVerticallyCenter(sheetIn.getVerticallyCenter());
//        sheetOut.setZoom();

        for(CellRangeAddress cellRangeAddress : sheetIn.getMergedRegions()){
            sheetOut.addMergedRegion(cellRangeAddress);
        }

//        sheetOut.addMergedRegionUnsafe();

//        for(DataValidation dataValidation : sheetIn.getDataValidations()){
//            sheetOut.addValidationData(dataValidation);
//        }

        PaneInformation paneInformation = sheetIn.getPaneInformation();
        if(paneInformation != null){
            if(paneInformation.isFreezePane()){
                sheetOut.createFreezePane(
                        paneInformation.getVerticalSplitPosition(),
                        paneInformation.getHorizontalSplitPosition(),
                        paneInformation.getHorizontalSplitTopRow(),
                        paneInformation.getVerticalSplitLeftColumn());
            }else {
                sheetOut.createSplitPane(
                        paneInformation.getVerticalSplitPosition(),
                        paneInformation.getHorizontalSplitPosition(),
                        paneInformation.getVerticalSplitLeftColumn(),
                        paneInformation.getHorizontalSplitTopRow(),
                        paneInformation.getActivePane());
            }
        }


    }

    private static InternalSheet getInternalSheet(Sheet sheetIn) {

        try {
            Method method = HSSFSheet.class.getDeclaredMethod("getSheet");
            method.setAccessible(true);
            return (InternalSheet) method.invoke(sheetIn);
        } catch (NoSuchMethodException e) {
//
            e.printStackTrace();
        } catch (IllegalAccessException e) {
//
            e.printStackTrace();
        } catch (InvocationTargetException e) {
//
            e.printStackTrace();
        }
        return null;
    }

    private static void hssfTextbox2XssfTextbox(CreationHelper factory, HSSFTextbox hssfTextbox, XSSFTextBox xssfTextBox) {
        xssfTextBox.setBottomInset(hssfTextbox.getMarginBottom());
        int fillColor = hssfTextbox.getFillColor();
//        xssfTextBox.setFillColor(
//                fillColor & 0x000000ff,
//                (fillColor & 0x0000ff00) >> 8,
//                (fillColor & 0x00ff0000) >> 8);
//        xssfTextBox.setLeftInset(hssfTextbox.getMarginLeft());
//        xssfTextBox.setLineStyle(hssfTextbox.getLineStyle());
//        int lineStyleColor = hssfTextbox.getLineStyleColor();
//        xssfTextBox.setLineStyleColor(
//                lineStyleColor & 0x000000ff,
//                (lineStyleColor & 0x0000ff00) >> 8,
//                (lineStyleColor & 0x00ff0000) >> 8);
        xssfTextBox.setLineWidth(hssfTextbox.getLineWidth());
        xssfTextBox.setNoFill(hssfTextbox.isNoFill());
        xssfTextBox.setRightInset(hssfTextbox.getMarginRight());
        xssfTextBox.setShapeType(202);
        xssfTextBox.setText((XSSFRichTextString) factory.createRichTextString(hssfTextbox.getString().getString()));
//                            xssfTextBox.setTextAutofit();
//                            xssfTextBox.setTextDirection();
//                            xssfTextBox.setTextHorizontalOverflow();
//                            xssfTextBox.setTextVerticalOverflow();
        xssfTextBox.setTopInset(hssfTextbox.getMarginTop());
//                            xssfTextBox.setVerticalAlignment();
//                            xssfTextBox.setWordWrap();
    }

    private static void copyRowProperties(Row rowOut, Row rowIn, Sheet sheetIn, Sheet sheetOut, HSSFPatriarch drawingIn, XSSFDrawing drawingOut, Map cellStyleMap) {
        rowOut.setRowNum(rowIn.getRowNum());
        rowOut.setHeight(rowIn.getHeight());
        rowOut.setHeightInPoints(rowIn.getHeightInPoints());
        rowOut.setZeroHeight(rowIn.getZeroHeight());
//        rowOut.setRowStyle();

//        sheetOut.setRowGroupCollapsed(rowIn.getRowNum(),sheetIn.setRowGroupCollapsed(););

        Iterator<Cell> cellIt = rowIn.cellIterator();
        while (cellIt.hasNext()) {
            Cell cellIn = cellIt.next();
            Cell cellOut = rowOut.createCell(cellIn.getColumnIndex(), cellIn.getCellTypeEnum());


            Comment comment = cellIn.getCellComment();
            if(comment != null){
                HSSFComment hssfComment = (HSSFComment) comment;
                ClientAnchor clientAnchor = hssfComment.getClientAnchor();
                XSSFClientAnchor xssfClientAnchor = drawingOut.createAnchor(
                        clientAnchor.getDx1(),
                        clientAnchor.getDy1(),
                        clientAnchor.getDx2(),
                        clientAnchor.getDy2(),
                        clientAnchor.getCol1(),
                        clientAnchor.getRow1(),
                        clientAnchor.getCol2(),
                        clientAnchor.getRow2());

                ClientAnchor clientAnchor1 = rowOut.getSheet().getWorkbook().getCreationHelper().createClientAnchor();

                XSSFComment xssfComment = drawingOut.createCellComment(clientAnchor1);
                xssfComment.setAuthor(hssfComment.getAuthor());
                xssfComment.setString(hssfComment.getString().getString());
                xssfComment.setAddress(hssfComment.getAddress());
//                cellOut.setCellComment(xssfComment);
            }


            rowOut.getSheet().setColumnWidth(cellOut.getColumnIndex(),
                    rowIn.getSheet().getColumnWidth(cellIn.getColumnIndex()));
            copyCellProperties(cellOut, cellIn, cellStyleMap,sheetIn,sheetOut);
        }

    }

    private static void copyCellProperties(Cell cellOut, Cell cellIn, Map cellStyleMap, Sheet sheetIn, Sheet sheetOut) {


//        cellOut.setCellComment();
//        cellOut.setCellValue();
//        cellOut.setCellStyle();
//        cellOut.setCellFormula(cellIn.getCellFormula());
//        cellOut.setAsActiveCell();
//        cellOut.setCellErrorValue(cellIn.getErrorCellValue());
        cellOut.setCellType(cellIn.getCellTypeEnum());
        if(cellIn.getHyperlink() != null){
            cellOut.setHyperlink(new XSSFHyperlink(cellIn.getHyperlink()));
        }


        Workbook wbOut = cellOut.getSheet().getWorkbook();
        HSSFPalette hssfPalette = ((HSSFWorkbook) cellIn.getSheet().getWorkbook()).getCustomPalette();
        switch (cellIn.getCellType()) {
            case Cell.CELL_TYPE_BLANK:
                break;

            case Cell.CELL_TYPE_BOOLEAN:
                cellOut.setCellValue(cellIn.getBooleanCellValue());
                break;

            case Cell.CELL_TYPE_ERROR:
                cellOut.setCellValue(cellIn.getErrorCellValue());
                break;

            case Cell.CELL_TYPE_FORMULA:
                if(cellIn.getCellFormula() != null){
                    cellOut.setCellFormula(cellIn.getCellFormula());
                }
                break;

            case Cell.CELL_TYPE_NUMERIC:
                cellOut.setCellValue(cellIn.getNumericCellValue());
                break;

            case Cell.CELL_TYPE_STRING:
                cellOut.setCellValue(cellIn.getStringCellValue());
                break;
            default:
                break;
        }
        HSSFCellStyle styleIn = (HSSFCellStyle) cellIn.getCellStyle();
        XSSFCellStyle styleOut = null;


        if (cellStyleMap.get(styleIn.getIndex()) != null) {
            styleOut = (XSSFCellStyle) cellStyleMap.get(styleIn.getIndex());
        } else {
            styleOut = (XSSFCellStyle) wbOut.createCellStyle();
            copyCellStyleProperties(styleIn,styleOut,sheetIn,sheetOut);
//            styleOut.setAlignment(styleIn.getAlignment());
            DataFormat format = wbOut.createDataFormat();
            styleOut.setDataFormat(format.getFormat(styleIn.getDataFormatString()));
            HSSFColor forgroundColor = styleIn.getFillForegroundColorColor();
            if (forgroundColor != null) {
                short[] foregroundColorValues = forgroundColor.getTriplet();
                styleOut.setFillForegroundColor(new XSSFColor(new java.awt.Color(foregroundColorValues[0],
                        foregroundColorValues[1], foregroundColorValues[2])));
                styleOut.setFillPattern(styleIn.getFillPattern());
            }
            styleOut.setFillPattern(styleIn.getFillPattern());
            styleOut.setBorderBottom(styleIn.getBorderBottom());
            styleOut.setBorderLeft(styleIn.getBorderLeft());
            styleOut.setBorderRight(styleIn.getBorderRight());
            styleOut.setBorderTop(styleIn.getBorderTop());
            HSSFColor bottom = hssfPalette.getColor(styleIn.getBottomBorderColor());
            if (bottom != null) {
                short[] bottomColorArray = bottom.getTriplet();
                styleOut.setBottomBorderColor(new XSSFColor(new java.awt.Color(bottomColorArray[0],
                        bottomColorArray[1], bottomColorArray[2])));
            }
            HSSFColor top = hssfPalette.getColor(styleIn.getTopBorderColor());
            if (top != null) {
                short[] topColorArray = top.getTriplet();
                styleOut.setTopBorderColor(new XSSFColor(new java.awt.Color(topColorArray[0], topColorArray[1],
                        topColorArray[2])));
            }
            HSSFColor left = hssfPalette.getColor(styleIn.getLeftBorderColor());
            if (left != null) {
                short[] leftColorArray = left.getTriplet();
                styleOut.setLeftBorderColor(new XSSFColor(new java.awt.Color(leftColorArray[0], leftColorArray[1],
                        leftColorArray[2])));
            }
            HSSFColor right = hssfPalette.getColor(styleIn.getRightBorderColor());
            if (right != null) {
                short[] rightColorArray = right.getTriplet();
                styleOut.setRightBorderColor(new XSSFColor(new java.awt.Color(rightColorArray[0], rightColorArray[1],
                        rightColorArray[2])));
            }
            HSSFColor background = hssfPalette.getColor(styleIn.getFillBackgroundColor());
            if (background != null) {
                short[] backgroundColorArray = background.getTriplet();
                styleOut.setFillBackgroundColor(new XSSFColor(new java.awt.Color(backgroundColorArray[0], backgroundColorArray[1],
                        backgroundColorArray[2])));
            }
            HSSFColor foreground = hssfPalette.getColor(styleIn.getFillForegroundColor());
            if (foreground != null) {
                short[] foregroundColorArray = foreground.getTriplet();
                styleOut.setFillForegroundColor(new XSSFColor(new java.awt.Color(foregroundColorArray[0], foregroundColorArray[1],
                        foregroundColorArray[2])));
            }
//            styleOut.setVerticalAlignment(styleIn.getVerticalAlignment());
//            styleOut.setHidden(styleIn.getHidden());
//            styleOut.setIndention(styleIn.getIndention());
//            styleOut.setLocked(styleIn.getLocked());
//            styleOut.setRotation(styleIn.getRotation());
//            styleOut.setShrinkToFit(styleIn.getShrinkToFit());
//            styleOut.setVerticalAlignment(styleIn.getVerticalAlignment());
//            styleOut.setWrapText(styleIn.getWrapText());
//            cellOut.setCellComment(cellIn.getCellComment());
            cellStyleMap.put(styleIn.getIndex(), styleOut);
        }
        cellOut.setCellStyle(styleOut);
    }

    private static void copyCellStyleProperties(HSSFCellStyle styleIn, XSSFCellStyle styleOut, Sheet sheetIn, Sheet sheetOut) {
        styleOut.setFillPattern(styleIn.getFillPatternEnum());
        styleOut.setVerticalAlignment(styleIn.getVerticalAlignmentEnum());
        styleOut.setAlignment(styleIn.getAlignmentEnum());
        styleOut.setBorderBottom(styleIn.getBorderBottomEnum());
        styleOut.setBorderLeft(styleIn.getBorderLeftEnum());
        styleOut.setBorderRight(styleIn.getBorderRightEnum());
        styleOut.setBorderTop(styleIn.getBorderTopEnum());
//        styleOut.setBottomBorderColor(styleIn.getBottomBorderColor());
//        styleOut.setDataFormat(styleIn.getDataFormat());
//        styleOut.setFillForegroundColor(styleIn.getFillForegroundColor());
        styleOut.setHidden(styleIn.getHidden());
        styleOut.setIndention(styleIn.getIndention());
//        styleOut.setLeftBorderColor(styleIn.getLeftBorderColor());
        styleOut.setLocked(styleIn.getLocked());
//        styleOut.setRightBorderColor(styleIn.getRightBorderColor());
        styleOut.setRotation(styleIn.getRotation());
        styleOut.setShrinkToFit(styleIn.getShrinkToFit());
//        styleOut.setTopBorderColor(styleIn.getTopBorderColor());
        styleOut.setWrapText(styleIn.getWrapText());
//        styleOut.setBorderColor();
//        styleOut.setFillBackgroundColor(styleIn.getFillBackgroundColor());
        styleOut.setQuotePrefixed(styleIn.getQuotePrefixed());

        HSSFFont hssfFont = styleIn.getFont(sheetIn.getWorkbook());
        XSSFFont xssfFont = (XSSFFont) sheetOut.getWorkbook().createFont();
        xssfFont.setBold(hssfFont.getBold());
        xssfFont.setFontHeightInPoints(hssfFont.getFontHeightInPoints());
        xssfFont.setFontHeight(hssfFont.getFontHeight());
        xssfFont.setFontName(hssfFont.getFontName());
        xssfFont.setColor(hssfFont.getColor());
        xssfFont.setItalic(hssfFont.getItalic());
        xssfFont.setUnderline(hssfFont.getUnderline());
        xssfFont.setCharSet(hssfFont.getCharSet());
//        xssfFont.setFamily(hssfFont);
//        xssfFont.setScheme(hssfFont.);
        xssfFont.setStrikeout(hssfFont.getStrikeout());
//        xssfFont.setThemeColor(hssfFont.g);
//        xssfFont.setThemesTable(hssfFont.get);
        xssfFont.setTypeOffset(hssfFont.getTypeOffset());

        styleOut.setFont(xssfFont);
    }
}
