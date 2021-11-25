package mutilExcel;


import java.util.ArrayList;
import java.util.List;

public class ReportRow {
    private List<ReportCell> cells = new ArrayList<>();

    private int rowNumber;

    private String sortKey;

    private int cellWidth;

    public ReportRow() {
        this.sortKey = "";
    }

    public void setWidth(int width) {
        this.cellWidth = width;
    }

    public void addHeaderCell(String cell) {
        ReportCell rc = new ReportCell(cell, 0);
        rc.setCellStyleName(ReportDefinition.HEADER_STYLE);
        rc.setRow(this);
        rc.setColumnWidth(this.cellWidth);
        this.cells.add(rc);
    }

    public void addHeaderCellColor(String cell) {
        ReportCell rc = new ReportCell(cell, 0);
        rc.setCellStyleName(ReportDefinition.HEADER_STYLE_COLOR);
        rc.setRow(this);
        rc.setColumnWidth(cell.getBytes().length*256*2);
        this.cells.add(rc);
    }

    public void addLongString(String cell) {
        ReportCell rc = new ReportCell(cell, 0);
        rc.setCellStyleName(ReportDefinition.LONG_STRING_STYLE);
        rc.setRow(this);
        this.cells.add(rc);
    }

    public void addLongStringColor(String cell) {
        ReportCell rc = new ReportCell(cell, 0);
        rc.setCellStyleName(ReportDefinition.LONG_STRING_STYLE_COLOR);
        rc.setRow(this);
        this.cells.add(rc);
    }

    public void addString(String cell) {
        ReportCell rc = new ReportCell(cell, 0);
        rc.setCellStyleName(ReportDefinition.STRING_STYLE);
        rc.setRow(this);
        this.cells.add(rc);
    }

    public void addStringColor(String cell) {
        ReportCell rc = new ReportCell(cell, 0);
        rc.setCellStyleName(ReportDefinition.STRING_STYLE_COLOR);
        rc.setRow(this);
        this.cells.add(rc);
    }

    public void addNumber(String cell) {
        ReportCell rc = new ReportCell(cell, 1);
        rc.setCellStyleName(ReportDefinition.NUMBER_STYLE);
        rc.setRow(this);
        this.cells.add(rc);
    }

    public void addNumberColor(String cell) {
        ReportCell rc = new ReportCell(cell, 1);
        rc.setCellStyleName(ReportDefinition.NUMBER_STYLE_COLOR);
        rc.setRow(this);
        this.cells.add(rc);
    }

    public void addDate(String cell) {
        ReportCell rc = new ReportCell(cell, 2);
        rc.setCellStyleName(ReportDefinition.DATE_STYLE);
        rc.setRow(this);
        this.cells.add(rc);
    }

    public void addDateColor(String cell) {
        ReportCell rc = new ReportCell(cell, 2);
        rc.setCellStyleName(ReportDefinition.DATE_STYLE_COLOR);
        rc.setRow(this);
        this.cells.add(rc);
    }

    public void addCell(String cell) {
        addString(cell);
    }

    public void addCellColor(String cell) {
//        addStringColor(cell);
        ReportCell rc = new ReportCell(cell, 0);
        rc.setCellStyleName(ReportDefinition.HEADER_STYLE_COLOR);
        rc.setRow(this);
        this.cells.add(rc);
    }

    public void addCellWidth(String cell, int width) {
        ReportCell rc = new ReportCell(cell, 0);
        rc.setCellStyleName(ReportDefinition.STRING_STYLE);
        rc.setRow(this);
        rc.setColumnWidth(width);
        this.cells.add(rc);
    }

    public void addCellWidthColor(String cell, int width) {
        ReportCell rc = new ReportCell(cell, 0);
        rc.setCellStyleName(ReportDefinition.STRING_STYLE_COLOR);
        rc.setRow(this);
        rc.setColumnWidth(width);
        this.cells.add(rc);
    }

    public void addBoldString(String cell) {
        ReportCell rc = new ReportCell(cell, 0);
        rc.setCellStyleName(ReportDefinition.STRING_STYLE);
        rc.setBold(true);
        rc.setRow(this);
        this.cells.add(rc);
    }

    public void addBoldStringColor(String cell) {
        ReportCell rc = new ReportCell(cell, 0);
        rc.setCellStyleName(ReportDefinition.STRING_STYLE_COLOR);
        rc.setBold(true);
        rc.setRow(this);
        this.cells.add(rc);
    }

    public void addCurrencyNumber(String cell) {
        ReportCell rc = new ReportCell(cell, 1);
        rc.setCellStyleName(ReportDefinition.CURRENCY_STYLE);
        rc.setRow(this);
        this.cells.add(rc);
    }

    public void addCurrencyNumberColor(String cell) {
        ReportCell rc = new ReportCell(cell, 1);
        rc.setCellStyleName(ReportDefinition.CURRENCY_STYLE_COLOR);
        rc.setRow(this);
        this.cells.add(rc);
    }

    public void addBoldCurrencyNumber(String cell) {
        ReportCell rc = new ReportCell(cell, 1);
        rc.setCellStyleName(ReportDefinition.CURRENCY_STYLE);
        rc.setBold(true);
        rc.setRow(this);
        this.cells.add(rc);
    }

    public void addBoldCurrencyNumberColor(String cell) {
        ReportCell rc = new ReportCell(cell, 1);
        rc.setCellStyleName(ReportDefinition.CURRENCY_STYLE_COLOR);
        rc.setBold(true);
        rc.setRow(this);
        this.cells.add(rc);
    }

    public void addPercentNumber(String cell) {
        ReportCell rc = new ReportCell(cell, 1);
        rc.setCellStyleName(ReportDefinition.PERCENTAGE_STYLE);
        rc.setRow(this);
        this.cells.add(rc);
    }

    public void addPercentNumberColor(String cell) {
        ReportCell rc = new ReportCell(cell, 1);
        rc.setCellStyleName(ReportDefinition.PERCENTAGE_STYLE_COLOR);
        rc.setRow(this);
        this.cells.add(rc);
    }

    public void addBoldDate(String cell) {
        ReportCell rc = new ReportCell(cell, 2);
        rc.setCellStyleName(ReportDefinition.DATE_STYLE);
        rc.setBold(true);
        rc.setRow(this);
        this.cells.add(rc);
    }

    public void addBoldDateColor(String cell) {
        ReportCell rc = new ReportCell(cell, 2);
        rc.setCellStyleName(ReportDefinition.DATE_STYLE_COLOR);
        rc.setBold(true);
        rc.setRow(this);
        this.cells.add(rc);
    }

    public void addBoldNumber(String cell) {
        ReportCell rc = new ReportCell(cell, 1);
        rc.setCellStyleName(ReportDefinition.NUMBER_STYLE);
        rc.setBold(true);
        rc.setRow(this);
        this.cells.add(rc);
    }

    public void addBoldNumberColor(String cell) {
        ReportCell rc = new ReportCell(cell, 1);
        rc.setCellStyleName(ReportDefinition.NUMBER_STYLE_COLOR);
        rc.setBold(true);
        rc.setRow(this);
        this.cells.add(rc);
    }

    public void addNoborderBoldText(String cell) {
        ReportCell rc = new ReportCell(cell, 0);
        rc.setCellStyleName(ReportDefinition.STRING_STYLE);
        rc.setBold(true);
        rc.setNoBorder(true);
        rc.setRow(this);
        this.cells.add(rc);
    }

    public void addNoborderBoldTextColor(String cell) {
        ReportCell rc = new ReportCell(cell, 0);
        rc.setCellStyleName(ReportDefinition.STRING_STYLE_COLOR);
        rc.setBold(true);
        rc.setNoBorder(true);
        rc.setRow(this);
        this.cells.add(rc);
    }

    public void addNoborderText(String cell) {
        ReportCell rc = new ReportCell(cell, 0);
        rc.setCellStyleName(ReportDefinition.STRING_STYLE);
        rc.setNoBorder(true);
        rc.setRow(this);
        this.cells.add(rc);
    }

    public void addNoborderTextColor(String cell) {
        ReportCell rc = new ReportCell(cell, 0);
        rc.setCellStyleName(ReportDefinition.STRING_STYLE_COLOR);
        rc.setNoBorder(true);
        rc.setRow(this);
        this.cells.add(rc);
    }

    public void addCustomCell(String cell, boolean isBold, int hAlign, int vAlign, int borderWidth, boolean hasBorder, int fontSize) {
        ReportCell rc = new ReportCell(cell, 3);
        rc.setCellStyleName(ReportDefinition.STRING_STYLE);
        rc.setCellContent(cell);
        rc.setBold(isBold);
        rc.setHAlign(hAlign);
        rc.setVAlign(vAlign);
        rc.setBorderWidth(borderWidth);
        rc.setHasBorders(hasBorder);
        rc.setFontSize(fontSize);
        rc.setRow(this);
        this.cells.add(rc);
    }

    public void addCustomCellColor(String cell, boolean isBold, int hAlign, int vAlign, int borderWidth, boolean hasBorder, int fontSize) {
        ReportCell rc = new ReportCell(cell, 3);
        rc.setCellStyleName(ReportDefinition.STRING_STYLE_COLOR);
        rc.setCellContent(cell);
        rc.setBold(isBold);
        rc.setHAlign(hAlign);
        rc.setVAlign(vAlign);
        rc.setBorderWidth(borderWidth);
        rc.setHasBorders(hasBorder);
        rc.setFontSize(fontSize);
        rc.setRow(this);
        this.cells.add(rc);
    }

    public void sumColumn(int colIndex, int startRow, int endRow, boolean currencyFormat) {
        String columnName = ColumnName(colIndex);
        String formula = "SUM(" + columnName + (startRow + 1) + ":" + columnName + (endRow + 1) + ")";
        ReportCell rc = new ReportCell(formula, -1);
        if (currencyFormat) {
            rc.setCellStyleName(ReportDefinition.CURRENCY_STYLE);
        } else {
            rc.setCellStyleName(ReportDefinition.NUMBER_STYLE);
        }
        this.cells.add(rc);
    }

    public void sumColumnColor(int colIndex, int startRow, int endRow, boolean currencyFormat) {
        String columnName = ColumnName(colIndex);
        String formula = "SUM(" + columnName + (startRow + 1) + ":" + columnName + (endRow + 1) + ")";
        ReportCell rc = new ReportCell(formula, -1);
        if (currencyFormat) {
            rc.setCellStyleName(ReportDefinition.CURRENCY_STYLE_COLOR);
        } else {
            rc.setCellStyleName(ReportDefinition.NUMBER_STYLE_COLOR);
        }
        this.cells.add(rc);
    }

    public void sumAbove(int startRow, int endRow, boolean currencyFormat) {
        int colIndex = this.cells.size() + 1;
        sumColumn(colIndex, startRow, endRow, currencyFormat);
    }

    public void sumAboveColor(int startRow, int endRow, boolean currencyFormat) {
        int colIndex = this.cells.size() + 1;
        sumColumnColor(colIndex, startRow, endRow, currencyFormat);
    }

    public void addFormulaNumber(String formula) {
        ReportCell rc = new ReportCell(formula, -1);
        rc.setCellStyleName(ReportDefinition.NUMBER_STYLE);
        this.cells.add(rc);
    }

    public void addFormulaNumberColor(String formula) {
        ReportCell rc = new ReportCell(formula, -1);
        rc.setCellStyleName(ReportDefinition.NUMBER_STYLE_COLOR);
        this.cells.add(rc);
    }

    public void addFormulaPercent(String formula) {
        ReportCell rc = new ReportCell(formula, -1);
        rc.setCellStyleName(ReportDefinition.PERCENTAGE_STYLE);
        this.cells.add(rc);
    }

    public void addFormulaPercentColor(String formula) {
        ReportCell rc = new ReportCell(formula, -1);
        rc.setCellStyleName(ReportDefinition.PERCENTAGE_STYLE_COLOR);
        this.cells.add(rc);
    }

    public void addFormulaNumber(String formula, boolean currencyFormat) {
        ReportCell rc = new ReportCell(formula, -1);
        if (currencyFormat) {
            rc.setCellStyleName(ReportDefinition.CURRENCY_STYLE);
        } else {
            rc.setCellStyleName(ReportDefinition.NUMBER_STYLE);
        }
        this.cells.add(rc);
    }

    public void addFormulaNumberColor(String formula, boolean currencyFormat) {
        ReportCell rc = new ReportCell(formula, -1);
        if (currencyFormat) {
            rc.setCellStyleName(ReportDefinition.CURRENCY_STYLE_COLOR);
        } else {
            rc.setCellStyleName(ReportDefinition.NUMBER_STYLE_COLOR);
        }
        this.cells.add(rc);
    }

    public void addFormulaString(String formula) {
        ReportCell rc = new ReportCell(formula, -1);
        rc.setCellStyleName(ReportDefinition.STRING_STYLE);
        this.cells.add(rc);
    }

    public void addFormulaStringColor(String formula) {
        ReportCell rc = new ReportCell(formula, -1);
        rc.setCellStyleName(ReportDefinition.STRING_STYLE_COLOR);
        this.cells.add(rc);
    }

    public void addFormulaString(String formula, int vAlign) {
        ReportCell rc = new ReportCell(formula, -1);
        rc.setCellStyleName(ReportDefinition.STRING_STYLE);
        rc.setVAlign(vAlign);
        this.cells.add(rc);
    }

    public void addFormulaStringColor(String formula, int vAlign) {
        ReportCell rc = new ReportCell(formula, -1);
        rc.setCellStyleName(ReportDefinition.STRING_STYLE_COLOR);
        rc.setVAlign(vAlign);
        this.cells.add(rc);
    }

    public void setOrder(int colIndex) {
        this.sortKey = String.valueOf(this.sortKey) + ((ReportCell)this.cells.get(colIndex)).getCellContent();
    }

    public List<ReportCell> getCells() {
        return this.cells;
    }

    public void setCells(List<ReportCell> cells) {
        this.cells = cells;
    }

    public int size() {
        return this.cells.size();
    }

    public int getRowNumber() {
        return this.rowNumber;
    }

    public void setRowNumber(int rowNumber) {
        this.rowNumber = rowNumber;
    }

    public String getSortKey() {
        return this.sortKey;
    }

    public void setSortKey(String sortKey) {
        this.sortKey = sortKey;
    }

    public static String GenerateLetter(int number) {
        String letter = "";
        switch (number) {
            case 0:
                letter = "Z";
                return letter;
            case 1:
                letter = "A";
                return letter;
            case 2:
                letter = "B";
                return letter;
            case 3:
                letter = "C";
                return letter;
            case 4:
                letter = "D";
                return letter;
            case 5:
                letter = "E";
                return letter;
            case 6:
                letter = "F";
                return letter;
            case 7:
                letter = "G";
                return letter;
            case 8:
                letter = "H";
                return letter;
            case 9:
                letter = "I";
                return letter;
            case 10:
                letter = "J";
                return letter;
            case 11:
                letter = "K";
                return letter;
            case 12:
                letter = "L";
                return letter;
            case 13:
                letter = "M";
                return letter;
            case 14:
                letter = "N";
                return letter;
            case 15:
                letter = "O";
                return letter;
            case 16:
                letter = "P";
                return letter;
            case 17:
                letter = "Q";
                return letter;
            case 18:
                letter = "R";
                return letter;
            case 19:
                letter = "S";
                return letter;
            case 20:
                letter = "T";
                return letter;
            case 21:
                letter = "U";
                return letter;
            case 22:
                letter = "V";
                return letter;
            case 23:
                letter = "W";
                return letter;
            case 24:
                letter = "X";
                return letter;
            case 25:
                letter = "Y";
                return letter;
        }
        return "Sorry,there is no answer!";
    }

    public static String ColumnName(int columnNum) {
        String columnName = "";
        int i = columnNum / 26;
        int j = columnNum % 26;
        String k = "";
        if (i == 0) {
            columnName = GenerateLetter(j);
        } else {
            k = GenerateLetter(j);
            if (j == 0) {
                if (i == 1)
                    return columnName = k;
                return columnName = String.valueOf(ColumnName(i - 1)) + k;
            }
            columnName = String.valueOf(ColumnName(i)) + k;
        }
        return columnName;
    }

    public static void main(String[] args) {
        System.out.println(ColumnName(2));
        System.out.println(ColumnName(12));
        System.out.println(ColumnName(22));
        System.out.println(ColumnName(26));
        System.out.println(ColumnName(52));
        System.out.println(ColumnName(62));
        System.out.println(ColumnName(72));
        System.out.println(ColumnName(82));
    }
}
