package lib;

import lib.model.Attribute;
import lib.model.Column;
import lib.model.Excel;
import lib.model.ExcelPoint;
import lib.model.Sheet;
import lib.model.Row;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.xssf.usermodel.*;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

public class ApachePoiExcelBuilder {
    private String path;
    private XSSFWorkbook workbook = new XSSFWorkbook();
    private XSSFCellStyle solidStyle = null;
    private XSSFCellStyle normalStyle = null;
    private XSSFCellStyle numberStyle = null;
    private XSSFCellStyle percentCellStyle = null;
    private XSSFCellStyle headerDefaultStyle = null;

    private Excel excel;

    private static final String HEADER_FONT_NAME = "Arial";
    private static final String CELL_FONT_NAME = "Arial";

    public void save() throws IOException {
        FileOutputStream fileOutputStream = new FileOutputStream(path);
        workbook.write(fileOutputStream);
        fileOutputStream.close();
    }

    public ApachePoiExcelBuilder path(String path) {
        this.path = path;
        return this;
    }

    public ApachePoiExcelBuilder excel(Excel excel) {
        this.excel = excel;
        return this;
    }

    public ApachePoiExcelBuilder generate() {

        for (Sheet sheet : excel.sheets()) {
            XSSFSheet poiSheet = createPoiSheet(sheet);
            Integer lastRowIndex = sheet.startPoint().getRowIndex();
            Integer columnIndexOffset = sheet.startPoint().getColumnIndex();

            List<Row> rows = sheet.rowsWithHeaderIfAvailable();

            for (Row row : rows) {
                addRow(row.columns(), poiSheet, columnIndexOffset, lastRowIndex++);
            }

            Integer lastColumn = ExcelUtil.getInstance().getMaxColumnIndex(rows);

            if (sheet.autoSizeColumn()) {
                for (Integer i = columnIndexOffset; i < lastColumn + columnIndexOffset; i++) {
                    poiSheet.autoSizeColumn(i, true);
                }
            }
        }

        return this;
    }

    private void addRow(List<Column> columns, XSSFSheet poiSheet, Integer columnIndexOffset, Integer rowIndex) {
        XSSFRow row = poiSheet.createRow(rowIndex);
        XSSFCell cell;

        for (Column column : columns) {

            Integer columnIndex = column.index() + columnIndexOffset;

            if (column.colSpan() != null) {

                for (int index = columnIndex; index < (columnIndex + column.colSpan()); index++) {
                    Cell dummyCell = row.createCell(index);
                    dummyCell.setCellStyle(createCellStyle(column.style()));
                }

                poiSheet.addMergedRegion(new CellRangeAddress(
                        rowIndex
                        , rowIndex
                        , columnIndex
                        , column.colSpan() + columnIndex - 1));
            }

            cell = row.createCell(columnIndex);

            if (column.value() != null) {
                if (column.type() != null
                        && column.type().getType() == Cell.CELL_TYPE_NUMERIC) {
                    cell.setCellValue(Double.parseDouble(column.value()));
                } else {
                    cell.setCellValue(column.value());
                }
            }

            if (column.formula() != null) {
                cell.setCellFormula(column.formula());
            }

            if (column.imagePath() != null) {
                drawPicture(poiSheet, row, column.imagePath(), ExcelPoint.newPoint(columnIndex, rowIndex));
            }

            cell.setCellStyle(createCellStyle(column.style()));

            CellUtil.setAlignment(cell, workbook, column.alignmentIndex());
        }

    }

    private void drawPicture(XSSFSheet sheet, org.apache.poi.ss.usermodel.Row row, String path, ExcelPoint excelPoint) {
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        try {
            BufferedImage image = ImageIO.read(getClass().getResource(path));
            ImageIO.write(image, "png", baos);
        } catch (IOException e) {
            e.printStackTrace();
        }

        byte[] bytes = baos.toByteArray();
        Integer pictureIdx = workbook.addPicture(bytes, Workbook.PICTURE_TYPE_PNG);

        // Create the drawing patriarch.  This is the top level container for all shapes.
        Drawing drawing = sheet.createDrawingPatriarch();

        //add a picture shape
        ClientAnchor anchor = workbook.getCreationHelper().createClientAnchor();
        anchor.setCol1(excelPoint.getColumnIndex());
        anchor.setRow1(excelPoint.getRowIndex());

        Cell cell = row.createCell(excelPoint.getColumnIndex() + 1);
        CellUtil.setAlignment(cell, workbook, CellStyle.ALIGN_CENTER);
        Picture pict = drawing.createPicture(anchor, pictureIdx);

        //auto-size picture relative to its top-left corner
        pict.resize();
    }

    private XSSFCellStyle createCellStyle(Attribute style) {
        switch (style) {
            case STYLE_NUMBER:
                return createNumberCellStyle();
            case STYLE_DEFAULT:
                return createNormalCellStyle();
            case STYLE_PERCENT:
                return createPercentCellStyle();
            case STYLE_SOLID:
                return createSolidCellStyle();
            case STYLE_SOLID_NUMBER:
                break;
            case STYLE_SOLID_PERCENT:
                break;
            case STYLE_HEADER:
                return createHeaderDefaultStyle();
            case STYLE_NONE:
                return null;
        }

        return workbook.createCellStyle();
    }

    public XSSFCellStyle createSolidCellStyle() {

        if (solidStyle != null)
            return solidStyle;

        Font font = workbook.createFont();
        font.setFontName(CELL_FONT_NAME);
        font.setBold(true);

        solidStyle = workbook.createCellStyle();
        solidStyle.setFillForegroundColor(IndexedColors.LIGHT_BLUE.getIndex());
        solidStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
        solidStyle.setBorderBottom(CellStyle.BORDER_THIN);
        solidStyle.setBorderTop(CellStyle.BORDER_THIN);
        solidStyle.setBorderLeft(CellStyle.BORDER_THIN);
        solidStyle.setBorderRight(CellStyle.BORDER_THIN);
        solidStyle.setFont(font);

        return solidStyle;
    }

    public XSSFCellStyle createPercentCellStyle() {
        if (percentCellStyle != null)
            return percentCellStyle;

        percentCellStyle = workbook.createCellStyle();
        percentCellStyle.setDataFormat(getDateFormat().getFormat("#,##0.00"));
        percentCellStyle.setBorderBottom(CellStyle.BORDER_THIN);
        percentCellStyle.setBorderTop(CellStyle.BORDER_THIN);
        percentCellStyle.setBorderLeft(CellStyle.BORDER_THIN);
        percentCellStyle.setBorderRight(CellStyle.BORDER_THIN);

        return percentCellStyle;
    }


    public XSSFCellStyle createNumberCellStyle() {
        if (numberStyle != null)
            return numberStyle;

        numberStyle = workbook.createCellStyle();
        numberStyle.setDataFormat(getDateFormat().getFormat("#,###0"));
        numberStyle.setBorderBottom(CellStyle.BORDER_THIN);
        numberStyle.setBorderTop(CellStyle.BORDER_THIN);
        numberStyle.setBorderLeft(CellStyle.BORDER_THIN);
        numberStyle.setBorderRight(CellStyle.BORDER_THIN);

        return numberStyle;
    }

    public XSSFCellStyle createHeaderDefaultStyle() {
        if (headerDefaultStyle != null) {
            return headerDefaultStyle;
        }

        Font font = workbook.createFont();
        font.setFontName(HEADER_FONT_NAME);
        font.setBold(true);

        headerDefaultStyle = workbook.createCellStyle();
        headerDefaultStyle.setFont(font);
        return headerDefaultStyle;

    }

    public XSSFDataFormat getDateFormat() {
        return workbook.createDataFormat();
    }

    public XSSFCellStyle createNormalCellStyle() {
        if (normalStyle != null)
            return normalStyle;

        Font font = workbook.createFont();
        font.setFontName(CELL_FONT_NAME);

        normalStyle = workbook.createCellStyle();
        normalStyle.setBorderBottom(CellStyle.BORDER_THIN);
        normalStyle.setBorderTop(CellStyle.BORDER_THIN);
        normalStyle.setBorderLeft(CellStyle.BORDER_THIN);
        normalStyle.setBorderRight(CellStyle.BORDER_THIN);
        normalStyle.setFont(font);

        return normalStyle;
    }

    private XSSFSheet createPoiSheet(Sheet sheet) {
        XSSFSheet poiSheet;
        poiSheet = workbook.createSheet(sheet.name());
        poiSheet.getCTWorksheet()
                .getSheetViews()
                .getSheetViewArray(0)
                .setRightToLeft(sheet.direction().equals(Attribute.DIRECTION_RTL));

        poiSheet.setDisplayGridlines(false);

        return poiSheet;
    }

    public static ApachePoiExcelBuilder getNewInstance() {
        return new ApachePoiExcelBuilder();
    }

}
