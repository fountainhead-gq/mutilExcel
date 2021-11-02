package mutilExcel;

import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.regex.Pattern;
import org.apache.poi.hssf.usermodel.HSSFFooter;
import org.apache.poi.hssf.usermodel.HSSFHeader;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Footer;
import org.apache.poi.ss.usermodel.Header;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

public class ReportDefinition {
    private String sheetName = "sheet1";

    private String reportName = "";

    private String operatorName = "";

    private List<ReportParameter> parametes = new ArrayList<>();

    private List<ReportRow> headers = new ArrayList<>();

    private List<ReportRow> rows = new ArrayList<>();

    private SXSSFWorkbook workbook;

    private Sheet sheet;

    private boolean prepared = false;

    private CellStyle headerStyle;

    private CellStyle propertyStyle;

    private CellStyle cellStyle;

    private CellStyle stringCellStyle;

    private CellStyle longStringCellStyle;

    private CellStyle numberCellStyle;

    private CellStyle dateCellStyle;

    private CellStyle currencyStyle;

    private CellStyle percentageStyle;

    private CellStyle wrapStyle;

    public static String HEADER_STYLE = "headerStyle";

    public static String CELL_STYLE = "cellStyle";

    public static String STRING_STYLE = "stringStyle";

    public static String LONG_STRING_STYLE = "longStringStyle";

    public static String NUMBER_STYLE = "numberStyle";

    public static String DATE_STYLE = "dateStyle";

    public static String PROPERTY_STYLE = "propertyStyle";

    public static String CURRENCY_STYLE = "currencyStyle";

    public static String PERCENTAGE_STYLE = "percentageStyle";

    private HashMap<String, CellStyle> styles = new HashMap<>();

    private boolean printProperty = true;

    private int headerRowIndex = 0;

    private int dataStartRowIndex = 0;

    private int dataEndRowIndex = 0;

    private int dataStartColIndex = 0;

    private int dataEndColIndex = 0;

    private int titleStartRowIndex = 0;

    private int titleEndRowIndex = 0;

    private int titleStartColIndex = 0;

    private int titleEndColIndex = 0;

    private int defaultColumnWidth = 3000;

    private int minColumnWidth = 238;

    private boolean printable = false;

    private SimpleDateFormat fm = new SimpleDateFormat("yyyy-MM-dd HH:mm:dd");

    public CellStyle getBoldStyle(Cell cell) {
        CellStyle cellStyle = this.workbook.createCellStyle();
        cellStyle.cloneStyleFrom(cell.getCellStyle());
        Font font = this.workbook.createFont();
        font.setBoldweight((short)700);
        cellStyle.setFont(font);
        return cellStyle;
    }

    public void setCellBold(int rowIndex, int colIndex) {
        if (this.sheet.getRow(rowIndex) != null) {
            Row row = this.sheet.getRow(rowIndex);
            if (row.getCell(colIndex) != null) {
                Cell cell = this.sheet.getRow(rowIndex).getCell(colIndex);
                cell.setCellStyle(getBoldStyle(cell));
            }
        }
    }

    public CellStyle generateCellStyle() {
        CellStyle style = this.workbook.createCellStyle();
        return style;
    }

    public void setCellStyle(int rowIndex, int colIndex, CellStyle style) {
        if (this.sheet.getRow(rowIndex) != null) {
            Row row = this.sheet.getRow(rowIndex);
            if (row.getCell(colIndex) != null) {
                Cell cell = row.getCell(colIndex);
                cell.setCellStyle(style);
            } else {
                System.out.println("Cell is null");
            }
        } else {
            System.out.println("row is null...");
        }
    }

    public void setBottomLine(int rowIndex, int colIndex, short linewidth) {
        if (this.sheet.getRow(rowIndex) != null) {
            Row row = this.sheet.getRow(rowIndex);
            if (row.getCell(colIndex) != null) {
                Cell cell = this.sheet.getRow(rowIndex).getCell(colIndex);
                CellStyle cellStyle = this.workbook.createCellStyle();
                cellStyle.cloneStyleFrom(cell.getCellStyle());
                cellStyle.setBorderBottom(linewidth);
                cell.setCellStyle(getBoldStyle(cell));
            }
        }
    }

    public CellStyle getNoBorderStyle(Cell cell) {
        CellStyle cellStyle = this.workbook.createCellStyle();
        cellStyle.cloneStyleFrom(cell.getCellStyle());
        cellStyle.setBorderLeft((short)0);
        cellStyle.setBorderRight((short)0);
        cellStyle.setBorderTop((short)0);
        cellStyle.setBorderBottom((short)0);
        return cellStyle;
    }

    public void setBorder(Cell cell, short width) {
        CellStyle cellStyle = cell.getCellStyle();
        cellStyle.setBorderLeft(width);
        cellStyle.setBorderRight(width);
        cellStyle.setBorderTop(width);
        cellStyle.setBorderBottom(width);
    }

    public void setBorder(int rowIndex, int colIndex, short width) {
        if (this.sheet.getRow(rowIndex) != null) {
            Row row = this.sheet.getRow(rowIndex);
            if (row.getCell(colIndex) != null) {
                Cell cell = this.sheet.getRow(rowIndex).getCell(colIndex);
                setBorder(cell, width);
            }
        }
    }

    public ReportDefinition() {
        this.workbook = new SXSSFWorkbook();
        this.sheet = this.workbook.createSheet(this.sheetName);
    }

    private void prepareStyles() {
        this.headerStyle = this.workbook.createCellStyle();
        if (!this.printable) {
            this.headerStyle.setBorderLeft((short)1);
            this.headerStyle.setBorderRight((short)1);
            this.headerStyle.setBorderTop((short)1);
            this.headerStyle.setBorderBottom((short)1);
        }
        this.headerStyle.setFillBackgroundColor((short)49);
        this.headerStyle.setFillForegroundColor((short)12);
        Font headerFont = this.workbook.createFont();
        headerFont.setBoldweight((short)700);
        headerFont.setFontHeightInPoints((short)10);
        headerFont.setFontName("Aria");
        this.headerStyle.setFont(headerFont);
        this.headerStyle.setVerticalAlignment((short)1);
        this.headerStyle.setAlignment((short)2);
        this.headerStyle.setWrapText(true);
        Font dataFont = this.workbook.createFont();
        dataFont.setFontHeightInPoints((short)10);
        dataFont.setFontName("Aria");
        this.cellStyle = this.workbook.createCellStyle();
        if (!this.printable) {
            this.cellStyle.setBorderLeft((short)1);
            this.cellStyle.setBorderRight((short)1);
            this.cellStyle.setBorderTop((short)1);
            this.cellStyle.setBorderBottom((short)1);
        }
        this.cellStyle.setVerticalAlignment((short)1);
        this.cellStyle.setAlignment((short)2);
        this.cellStyle.setFont(dataFont);
        this.propertyStyle = this.workbook.createCellStyle();
        if (!this.printable) {
            this.propertyStyle.setBorderLeft((short)1);
            this.propertyStyle.setBorderRight((short)1);
            this.propertyStyle.setBorderTop((short)1);
            this.propertyStyle.setBorderBottom((short)1);
        }
        this.propertyStyle.setFont(headerFont);
        this.propertyStyle.setVerticalAlignment((short)1);
        this.propertyStyle.setAlignment((short)1);
        this.stringCellStyle = this.workbook.createCellStyle();
        if (!this.printable) {
            this.stringCellStyle.setBorderLeft((short)1);
            this.stringCellStyle.setBorderRight((short)1);
            this.stringCellStyle.setBorderTop((short)1);
            this.stringCellStyle.setBorderBottom((short)1);
        }
        this.stringCellStyle.setVerticalAlignment((short)1);
        this.stringCellStyle.setAlignment((short)1);
        this.stringCellStyle.setFont(dataFont);
        this.longStringCellStyle = this.workbook.createCellStyle();
        if (!this.printable) {
            this.longStringCellStyle.setBorderLeft((short)1);
            this.longStringCellStyle.setBorderRight((short)1);
            this.longStringCellStyle.setBorderTop((short)1);
            this.longStringCellStyle.setBorderBottom((short)1);
        }
        this.longStringCellStyle.setVerticalAlignment((short)1);
        this.longStringCellStyle.setAlignment((short)1);
        this.longStringCellStyle.setWrapText(true);
        this.longStringCellStyle.setFont(dataFont);
        this.numberCellStyle = this.workbook.createCellStyle();
        if (!this.printable) {
            this.numberCellStyle.setBorderLeft((short)1);
            this.numberCellStyle.setBorderRight((short)1);
            this.numberCellStyle.setBorderTop((short)1);
            this.numberCellStyle.setBorderBottom((short)1);
        }
        this.numberCellStyle.setVerticalAlignment((short)1);
        this.numberCellStyle.setAlignment((short)3);
        this.numberCellStyle.setFont(dataFont);
        DataFormat format = this.workbook.createDataFormat();
        this.dateCellStyle = this.workbook.createCellStyle();
        this.dateCellStyle.setDataFormat(format.getFormat("yyyy-MM-dd"));
        if (!this.printable) {
            this.dateCellStyle.setBorderLeft((short)1);
            this.dateCellStyle.setBorderRight((short)1);
            this.dateCellStyle.setBorderTop((short)1);
            this.dateCellStyle.setBorderBottom((short)1);
        }
        this.dateCellStyle.setVerticalAlignment((short)1);
        this.dateCellStyle.setAlignment((short)3);
        this.dateCellStyle.setFont(dataFont);
        this.currencyStyle = this.workbook.createCellStyle();
        DataFormat currencyformat = this.workbook.createDataFormat();
        this.currencyStyle.setDataFormat(currencyformat.getFormat("#,##0.00"));
        if (!this.printable) {
            this.currencyStyle.setBorderLeft((short)1);
            this.currencyStyle.setBorderRight((short)1);
            this.currencyStyle.setBorderTop((short)1);
            this.currencyStyle.setBorderBottom((short)1);
        }
        this.currencyStyle.setVerticalAlignment((short)1);
        this.currencyStyle.setAlignment((short)3);
        this.currencyStyle.setFont(dataFont);
        this.wrapStyle = this.workbook.createCellStyle();
        this.wrapStyle.setWrapText(true);
        this.wrapStyle.setFont(dataFont);
        this.percentageStyle = this.workbook.createCellStyle();
        DataFormat percentageFormat = this.workbook.createDataFormat();
        this.percentageStyle.setDataFormat(percentageFormat.getFormat("0.00%"));
        if (!this.printable) {
            this.percentageStyle.setBorderLeft((short)1);
            this.percentageStyle.setBorderRight((short)1);
            this.percentageStyle.setBorderTop((short)1);
            this.percentageStyle.setBorderBottom((short)1);
        }
        this.percentageStyle.setVerticalAlignment((short)1);
        this.percentageStyle.setAlignment((short)3);
        this.percentageStyle.setFont(dataFont);
        this.styles.put(HEADER_STYLE, this.headerStyle);
        this.styles.put(PROPERTY_STYLE, this.propertyStyle);
        this.styles.put(STRING_STYLE, this.stringCellStyle);
        this.styles.put(LONG_STRING_STYLE, this.longStringCellStyle);
        this.styles.put(CELL_STYLE, this.cellStyle);
        this.styles.put(NUMBER_STYLE, this.numberCellStyle);
        this.styles.put(DATE_STYLE, this.dateCellStyle);
        this.styles.put(CURRENCY_STYLE, this.currencyStyle);
        this.styles.put(PERCENTAGE_STYLE, this.percentageStyle);
    }

    public void addRow(ReportRow row) {
        row.setRowNumber(5 + this.parametes.size() + 1 + this.headers.size() + this.rows.size());
        this.rows.add(row);
    }

    public void addHeader(String header) {
        if (this.headers.size() == 0)
            this.headers.add(new ReportRow());
        ((ReportRow)this.headers.get(0)).addCell(header);
    }

    public void addHeaderWidth(String header, int width) {
        if (this.headers.size() == 0)
            this.headers.add(new ReportRow());
        ((ReportRow)this.headers.get(0)).addCellWidth(header, width);
    }

    public void addHeaderRow(ReportRow row) {
        this.headers.add(row);
    }

    public void addParameter(String propertyName, String propertyValue) {
        this.parametes.add(new ReportParameter(propertyName, propertyValue));
    }

    public void mergeCells(int firstRow, int lastRow, int firstCol, int lastCol) {
        CellRangeAddress address = new CellRangeAddress(firstRow, lastRow, firstCol, lastCol);
        this.workbook.getSheet(this.sheetName).addMergedRegion(address);
    }

    public String getCellValue(int rowIndex, int colIndex) {
        String value = "";
        Sheet sheet = this.workbook.getSheet(this.sheetName);
        if (sheet.getRow(rowIndex) == null)
            return null;
        Row row = sheet.getRow(rowIndex);
        if (row.getCell(colIndex) == null)
            return null;
        Cell cell = row.getCell(colIndex);
        switch (cell.getCellType()) {
            case 0:
                value = String.valueOf(cell.getNumericCellValue());
                break;
            case 1:
                value = cell.getStringCellValue();
                break;
        }
        return value;
    }

    public void mergeSameCellsInRow(int theRow, int firstCol, int lastCol) {
        if (!this.prepared)
            prepareExcel();
        int fromCol = firstCol;
        int toCol = firstCol + 1;
        for (int i = firstCol + 1; i <= lastCol; i++) {
            toCol = i;
            String v1 = getCellValue(theRow, fromCol);
            String v2 = getCellValue(theRow, toCol);
            if (v1 == null || v2 == null) {
                toCol--;
            } else if (!v1.equals(v2)) {
                if (fromCol + 1 != toCol)
                    mergeCells(theRow, theRow, fromCol, toCol - 1);
                fromCol = toCol;
            }
        }
        mergeCells(theRow, theRow, fromCol, toCol);
    }

    public void mergeSameCellsInColumn(int theCol, int firstRow, int lastRow) {
        if (!this.prepared)
            prepareExcel();
        int fromRow = firstRow;
        int toRow = firstRow + 1;
        for (int i = firstRow + 1; i <= lastRow; i++) {
            toRow = i;
            String v1 = getCellValue(fromRow, theCol);
            String v2 = getCellValue(toRow, theCol);
            if (v1 == null || v2 == null) {
                toRow--;
            } else if (!v1.equals(v2)) {
                if (fromRow + 1 != toRow)
                    mergeCells(fromRow, toRow - 1, theCol, theCol);
                fromRow = toRow;
            }
        }
        mergeCells(fromRow, toRow, theCol, theCol);
    }

    public int getFirstHeaderRow() {
        return this.parametes.size();
    }

    public int getLastHeaderRow() {
        return this.parametes.size() + this.headers.size() - 1;
    }

    public int getFirstContentRow() {
        return this.parametes.size() + this.headers.size();
    }

    public int getLastContentRow() {
        return this.parametes.size() + this.headers.size() + this.rows.size() - 1;
    }

    public int getLastColumn() {
        if (this.headers.size() == 0)
            return 0;
        return ((ReportRow)this.headers.get(0)).size() - 1;
    }

    private void insertPropertyCells(int rownumber, String paramName, String paraValue, boolean merge) {
        Row row = this.sheet.createRow(rownumber);
        Cell cellPropertyName = row.createCell(0);
        cellPropertyName.setCellStyle(this.propertyStyle);
        cellPropertyName.setCellType(1);
        cellPropertyName.setCellValue(paramName);
        int columneWidth1 = this.sheet.getColumnWidth(0);
        int length1 = (paramName.getBytes()).length * 256;
        if (length1 < columneWidth1) {
            this.sheet.setColumnWidth(0, columneWidth1);
        } else {
            this.sheet.setColumnWidth(0, (int)(length1 * 1.1D));
        }
        Cell cellPropertyValue = row.createCell(1);
        cellPropertyValue.setCellStyle(this.stringCellStyle);
        cellPropertyValue.setCellType(1);
        cellPropertyValue.setCellValue(paraValue);
        int columneWidth2 = this.sheet.getColumnWidth(1);
        int length2 = (paraValue.getBytes()).length * 256;
        if (length2 < columneWidth2) {
            this.sheet.setColumnWidth(1, columneWidth2);
        } else {
            this.sheet.setColumnWidth(1, (int)(length2 * 1.1D));
        }
        if (merge)
            mergeCells(rownumber, rownumber, 0, 1);
    }

    public void prepareExcel() {
        if (!this.prepared) {
            prepareStyles();
            int rowNumber = 0;
            if (this.printProperty) {
                insertPropertyCells(rowNumber++, "报表信息", "", true);
                insertPropertyCells(rowNumber++, "报表名称", getReportName(), false);
                insertPropertyCells(rowNumber++, "制表人", getOperatorName(), false);
                insertPropertyCells(rowNumber++, "制表时间", this.fm.format(new Date()), false);
                insertPropertyCells(rowNumber++, "报表参数", "", true);
                for (ReportParameter parameter : this.parametes)
                    insertPropertyCells(rowNumber++, parameter.getPropertyName(), parameter.getPropertyValue(), false);
                rowNumber++;
            }
            this.dataStartRowIndex = rowNumber;
            this.titleStartRowIndex = rowNumber;
            for (int h = 0; h < this.headers.size(); h++) {
                Row row = this.sheet.createRow(rowNumber++);
                for (int i = 0; i < ((ReportRow)this.headers.get(h)).size(); i++) {
                    Cell cellHeader = row.createCell(i);
                    cellHeader.setCellStyle(this.headerStyle);
                    cellHeader.setCellType(1);
                    cellHeader.setCellValue(((ReportCell)((ReportRow)this.headers.get(h)).getCells().get(i)).getCellContent());
                    this.titleEndColIndex = i;
                }
            }
            this.titleEndRowIndex = rowNumber - 1;
            for (Iterator<ReportRow> r = (Iterator<ReportRow>)this.rows.iterator(); ((Iterator)r).hasNext(); ) {
                ReportRow reportRow = ((Iterator<ReportRow>)r).next();
                Row row = this.sheet.createRow(rowNumber++);
                for (int i = 0; i < reportRow.getCells().size(); i++) {
                    Cell cell = row.createCell(i);
                    ReportCell rc = reportRow.getCells().get(i);
                    String cellValue = rc.getCellContent();
                    Integer cellFormat = rc.getCellFormat();
                    cell.setCellType(1);
                    cell.setCellStyle(this.styles.get(rc.getCellStyleName()));
                    if (cellFormat.intValue() == -1) {
                        cell.setCellFormula(rc.getCellContent());
                        cell.setCellType(2);
                    }
                    if (cellFormat.intValue() == 2 && isDate(cellValue))
                        try {
                            SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
                            cell.setCellValue(sdf.parse(cellValue));
                        } catch (Exception ex) {
                            cell.setCellType(1);
                            cell.setCellValue(cellValue);
                        }
                    if (cellFormat.intValue() == 1)
                        try {
                            if (isInteger(cellValue)) {
                                if (cellValue.length() < 11)
                                    cell.setCellValue(Integer.parseInt(cellValue));
                                cell.setCellType(0);
                            } else if (isDecimal(cellValue)) {
                                cell.setCellValue(Double.parseDouble(cellValue));
                                cell.setCellType(0);
                            }
                        } catch (Exception ex) {
                            cell.setCellValue(cellValue);
                            cell.setCellType(1);
                        }
                    if (cellFormat.intValue() == 0) {
                        cell.setCellValue(cellValue);
                        cell.setCellType(1);
                    }
                    if (cellFormat.intValue() == 3) {
                        cell.setCellType(1);
                        cell.setCellValue(rc.getCellContent());
                    }
                    rc.isBold();
                    rc.isNoBorder();
                }
                row.setHeightInPoints(row.getHeightInPoints() * 1.2F);
            }
            this.dataEndRowIndex = rowNumber - 1;
            this.dataStartColIndex = 0;
            if (this.headers.size() > 0)
                for (int i = 0; i < ((ReportRow)this.headers.get(0)).size(); i++)
                    this.dataEndColIndex = i;
            if (this.rows.size() > 0)
                for (int i = 0; i < ((ReportRow)this.rows.get(0)).size(); i++) {
                    ReportCell rc = ((ReportRow)this.rows.get(0)).getCells().get(i);
                    this.dataEndColIndex = i;
                }
            if (this.printable)
                addSignFooter(rowNumber);
        }
        this.prepared = true;
    }

    public void prepareExcelEng() {
        if (!this.prepared) {
            prepareStyles();
            int rowNumber = 0;
            if (this.printProperty) {
                insertPropertyCells(rowNumber++, "Report information", "", true);
                insertPropertyCells(rowNumber++, "Report name", getReportName(), false);
                insertPropertyCells(rowNumber++, "Tabulator", getOperatorName(), false);
                insertPropertyCells(rowNumber++, "Created Time", this.fm.format(new Date()), false);
                insertPropertyCells(rowNumber++, "Report Parameters", "", true);
                for (ReportParameter parameter : this.parametes)
                    insertPropertyCells(rowNumber++, parameter.getPropertyName(), parameter.getPropertyValue(), false);
                rowNumber++;
            }
            this.dataStartRowIndex = rowNumber;
            this.titleStartRowIndex = rowNumber;
            for (int h = 0; h < this.headers.size(); h++) {
                Row row = this.sheet.createRow(rowNumber++);
                for (int i = 0; i < ((ReportRow)this.headers.get(h)).size(); i++) {
                    Cell cellHeader = row.createCell(i);
                    cellHeader.setCellStyle(this.headerStyle);
                    cellHeader.setCellType(1);
                    cellHeader.setCellValue(((ReportCell)((ReportRow)this.headers.get(h)).getCells().get(i)).getCellContent());
                    this.titleEndColIndex = i;
                }
            }
            this.titleEndRowIndex = rowNumber - 1;
            for (Iterator<ReportRow> r = (Iterator<ReportRow>)this.rows.iterator(); ((Iterator)r).hasNext(); ) {
                ReportRow reportRow = ((Iterator<ReportRow>)r).next();
                Row row = this.sheet.createRow(rowNumber++);
                for (int i = 0; i < reportRow.getCells().size(); i++) {
                    Cell cell = row.createCell(i);
                    ReportCell rc = reportRow.getCells().get(i);
                    String cellValue = rc.getCellContent();
                    Integer cellFormat = rc.getCellFormat();
                    cell.setCellType(1);
                    cell.setCellStyle(this.styles.get(rc.getCellStyleName()));
                    if (cellFormat.intValue() == -1) {
                        cell.setCellFormula(rc.getCellContent());
                        cell.setCellType(2);
                    }
                    if (cellFormat.intValue() == 2 && isDate(cellValue))
                        try {
                            SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
                            cell.setCellValue(sdf.parse(cellValue));
                        } catch (Exception ex) {
                            cell.setCellType(1);
                            cell.setCellValue(cellValue);
                        }
                    if (cellFormat.intValue() == 1)
                        try {
                            if (isInteger(cellValue)) {
                                if (cellValue.length() < 11)
                                    cell.setCellValue(Integer.parseInt(cellValue));
                                cell.setCellType(0);
                            } else if (isDecimal(cellValue)) {
                                cell.setCellValue(Double.parseDouble(cellValue));
                                cell.setCellType(0);
                            }
                        } catch (Exception ex) {
                            cell.setCellValue(cellValue);
                            cell.setCellType(1);
                        }
                    if (cellFormat.intValue() == 0) {
                        cell.setCellValue(cellValue);
                        cell.setCellType(1);
                    }
                    if (cellFormat.intValue() == 3) {
                        cell.setCellType(1);
                        cell.setCellValue(rc.getCellContent());
                    }
                    if (rc.isBold())
                        cell.setCellStyle(getBoldStyle(cell));
                    rc.isNoBorder();
                }
                row.setHeightInPoints(row.getHeightInPoints() * 1.2F);
            }
            this.dataEndRowIndex = rowNumber - 1;
            this.dataStartColIndex = 0;
            if (this.printable)
                addSignFooter(rowNumber);
        }
        this.prepared = true;
    }

    public void buildExcel(String filePath) {
        try {
            if (!this.prepared)
                prepareExcel();
            if (this.sheet != null)
                for (int i = 0; i < ((ReportRow)this.headers.get(0)).size(); i++) {
                    this.sheet.autoSizeColumn(i, true);
                    int columnWidth = this.sheet.getColumnWidth(i);
                    System.out.println("" + i + "->" + columnWidth);
                    if (columnWidth <= this.minColumnWidth) {
                        this.sheet.setColumnWidth(i, this.defaultColumnWidth);
                        System.out.println(i + " set ->" + this.sheet.getColumnWidth(i));
                    }
                }
            FileOutputStream fOut = new FileOutputStream(filePath);
            this.workbook.write(fOut);
            fOut.flush();
            fOut.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public void buildExcelEng(String filePath) {
        try {
            if (!this.prepared)
                prepareExcelEng();
            if (this.sheet != null)
                for (int i = 0; i < ((ReportRow)this.headers.get(0)).size(); i++) {
                    this.sheet.autoSizeColumn(i, true);
                    int columnWidth = this.sheet.getColumnWidth(i);
                    System.out.println("" + i + "->" + columnWidth);
                    if (columnWidth <= this.minColumnWidth) {
                        this.sheet.setColumnWidth(i, this.defaultColumnWidth);
                        System.out.println(i + " set ->" + this.sheet.getColumnWidth(i));
                    }
                }
            FileOutputStream fOut = new FileOutputStream(filePath);
            this.workbook.write(fOut);
            fOut.flush();
            fOut.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public void buildExcel2(String filePath, Object[] arrColumnIndex, Object[] arrColumnWidth) {
        try {
            if (!this.prepared)
                prepareExcel();
            engSheet(filePath, arrColumnIndex, arrColumnWidth);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public void buildExcelEng2(String filePath, Object[] arrColumnIndex, Object[] arrColumnWidth) {
        try {
            if (!this.prepared)
                prepareExcelEng();
            engSheet(filePath, arrColumnIndex, arrColumnWidth);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private void engSheet(String filePath, Object[] arrColumnIndex, Object[] arrColumnWidth) throws IOException {
        if (this.sheet != null)
            for (int i = 0; i < ((ReportRow)this.headers.get(0)).size(); i++) {
                this.sheet.autoSizeColumn(i, true);
                System.out.println(i + " default ->" + this.sheet.getColumnWidth(i));
                int columnWidth = this.sheet.getColumnWidth(i);
                if (columnWidth == 238) {
                    int columnIndex = -1;
                    columnIndex = Arrays.<Object>asList(arrColumnIndex).indexOf(Integer.valueOf(i));
                    if (columnIndex >= 0) {
                        this.sheet.setColumnWidth(i, ((Integer)arrColumnWidth[columnIndex]).intValue());
                        System.out.println(i + " set ->" + this.sheet.getColumnWidth(i));
                    }
                }
            }
        FileOutputStream fOut = new FileOutputStream(filePath);
        this.workbook.write(fOut);
        fOut.flush();
        fOut.close();
    }

    public void setFooter(int fontSize, String footerStr) {
        Footer footer = getSheet().getFooter();
        footer.setCenter(String.valueOf(HSSFFooter.fontSize((short)fontSize)) + footerStr);
    }

    public void setFooterCenter(String footerCenter) {
        Footer footer = getSheet().getFooter();
        footer.setCenter(footerCenter);
    }

    public void setPageNumberSizeAndFooter(int fontSize, String str) {
        Footer footer = getSheet().getFooter();
        str = str.replaceAll("#PageNumber#", HSSFFooter.page());
        str = str.replaceAll("#PageCount#", HSSFFooter.numPages());
        footer.setRight(String.valueOf(HSSFFooter.fontSize((short)fontSize)) + str);
    }

    public void setPageNumberFooter() {
        Footer footer = getSheet().getFooter();
        footer.setRight("+ HSSFFooter.page() + " + HSSFFooter.numPages() + "");
    }

    public void addSignFooter(int rowNumber) {}

    public boolean isDecimal(String str) {
        if (str == null || "".equals(str))
            return false;
        Pattern pattern = Pattern.compile("^(-?\\d+)(\\.\\d+)?");
        return pattern.matcher(str).matches();
    }

    public boolean isInteger(String str) {
        if (str == null)
            return false;
        Pattern pattern = Pattern.compile("[0-9]+");
        return pattern.matcher(str).matches();
    }

    public boolean isDate(String str) {
        if (str == null)
            return false;
        Pattern pattern = Pattern.compile("^([1-2]\\d{3})[\\/|\\-](0?[1-9]|10|11|12)[\\/|\\-]([1-2]?[0-9]|0[1-9]|30|31)$");
        return pattern.matcher(str).matches();
    }

    public List<ReportParameter> getParametes() {
        return this.parametes;
    }

    public List<ReportRow> getHeaders() {
        return this.headers;
    }

    public List<ReportRow> getRows() {
        return this.rows;
    }

    public void setParametes(List<ReportParameter> parametes) {
        this.parametes = parametes;
    }

    public void setHeaders(List<ReportRow> headers) {
        this.headers = headers;
    }

    public void setRows(List<ReportRow> rows) {
        this.rows = rows;
    }

    public String getSheetName() {
        return this.sheetName;
    }

    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }

    public String getReportName() {
        return this.reportName;
    }

    public String getOperatorName() {
        return this.operatorName;
    }

    public void setReportName(String reportName) {
        this.reportName = reportName;
    }

    public void setOperatorName(String operatorName) {
        this.operatorName = operatorName;
    }

    public boolean isPrintProperty() {
        return this.printProperty;
    }

    public void setPrintProperty(boolean printProperty) {
        this.printProperty = printProperty;
    }

    public int getHeaderRowIndex() {
        return this.headerRowIndex;
    }

    public void setHeaderRowIndex(int headerRowIndex) {
        this.headerRowIndex = headerRowIndex;
    }

    public void setColumnWidth(int columnIndex, int width) {
        try {
            this.sheet.setColumnWidth(columnIndex, width);
        } catch (Exception exception) {}
    }

    public void setRowHeight(int rowIndex, int height) {
        this.sheet.getRow(rowIndex).setHeight((short)height);
    }

    public SXSSFWorkbook getWorkbook() {
        return this.workbook;
    }

    public Sheet getSheet() {
        return this.sheet;
    }

    public void setWorkbook(SXSSFWorkbook workbook) {
        this.workbook = workbook;
    }

    public void setSheet(Sheet sheet) {
        this.sheet = sheet;
    }

    public void hideRow(int rowIndex) {
        if (!this.prepared)
            prepareExcel();
        this.sheet.getRow(rowIndex).setZeroHeight(true);
    }

    public void setDefaultHeader() {
        Header header = getSheet().getHeader();
        header.setCenter(getReportName());
    }

    public void setHeader(int fontSize, String headerStr) {
        Header header = getSheet().getHeader();
        header.setCenter(String.valueOf(HSSFHeader.fontSize((short)fontSize)) + headerStr);
    }

    public void setDefaultPrintArea() {
        getWorkbook().setPrintArea(0, this.dataStartColIndex, this.dataEndColIndex, this.dataStartRowIndex, this.dataEndRowIndex);
    }

    public void setDefaultRepeatTitle() {
        getWorkbook().setRepeatingRowsAndColumns(0, this.titleStartColIndex, this.titleEndColIndex, this.titleStartRowIndex, this.titleEndRowIndex);
    }

    public void setPageMargin(double left, double top, double right, double bottom) {
        getSheet().setMargin((short)0, left);
        getSheet().setMargin((short)1, right);
        getSheet().setMargin((short)2, top);
        getSheet().setMargin((short)3, bottom);
    }

    public void setDefaultFooter() {
        Footer footer = getSheet().getFooter();
        String footLeft = "";
        footLeft = String.valueOf(footLeft) + "__________________   ";
        footLeft = String.valueOf(footLeft) + "CB__________________   ";
        footLeft = String.valueOf(footLeft) + "HR __________________   ";
        footLeft = String.valueOf(footLeft) + "__________________";
        footer.setCenter(footLeft);
        footer.setRight("+ HSSFFooter.page() + " + HSSFFooter.numPages() + "");
    }

    public void setLandScape(boolean ls) {
        getSheet().getPrintSetup().setLandscape(ls);
    }

    public void setPageSize(int pagesize) {
        getSheet().getPrintSetup().setPaperSize((short)pagesize);
    }

    public PrintSetup getPrintSetup() {
        return getSheet().getPrintSetup();
    }

    public void hideColumn(int colIndex) {
        getSheet().setColumnHidden(colIndex, true);
    }

    public void setOrder(int colIndex) {
        for (int i = 0; i < this.rows.size(); i++)
            ((ReportRow)this.rows.get(i)).setOrder(colIndex);
    }

    public void sort() {
        ReportRowComparator comparator = new ReportRowComparator();
        Collections.sort(this.rows, comparator);
    }

    public int getDataStartRowIndex() {
        return this.dataStartRowIndex;
    }

    public int getDataEndRowIndex() {
        return this.dataEndRowIndex;
    }

    public int getDataStartColIndex() {
        return this.dataStartColIndex;
    }

    public int getDataEndColIndex() {
        return this.dataEndColIndex;
    }

    public int getTitleStartRowIndex() {
        return this.titleStartRowIndex;
    }

    public int getTitleEndRowIndex() {
        return this.titleEndRowIndex;
    }

    public int getTitleStartColIndex() {
        return this.titleStartColIndex;
    }

    public int getTitleEndColIndex() {
        return this.titleEndColIndex;
    }

    public int getDefaultColumnWidth() {
        return this.defaultColumnWidth;
    }

    public int getMinColumnWidth() {
        return this.minColumnWidth;
    }

    public boolean isPrintable() {
        return this.printable;
    }

    public void setDataStartRowIndex(int dataStartRowIndex) {
        this.dataStartRowIndex = dataStartRowIndex;
    }

    public void setDataEndRowIndex(int dataEndRowIndex) {
        this.dataEndRowIndex = dataEndRowIndex;
    }

    public void setDataStartColIndex(int dataStartColIndex) {
        this.dataStartColIndex = dataStartColIndex;
    }

    public void setDataEndColIndex(int dataEndColIndex) {
        this.dataEndColIndex = dataEndColIndex;
    }

    public void setTitleStartRowIndex(int titleStartRowIndex) {
        this.titleStartRowIndex = titleStartRowIndex;
    }

    public void setTitleEndRowIndex(int titleEndRowIndex) {
        this.titleEndRowIndex = titleEndRowIndex;
    }

    public void setTitleStartColIndex(int titleStartColIndex) {
        this.titleStartColIndex = titleStartColIndex;
    }

    public void setTitleEndColIndex(int titleEndColIndex) {
        this.titleEndColIndex = titleEndColIndex;
    }

    public void setPrintable(boolean printable) {
        this.printable = printable;
    }

    public void setDefaultColumnWidth(int columnWidth) {
        this.defaultColumnWidth = columnWidth;
    }

    public void setMinColumnWidth(int columnWidth) {
        this.minColumnWidth = columnWidth;
    }


}