package mutilExcel;

import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.util.Date;

public class TestUtil {

    public static void main(String[] args) {
//
//      多sheet附件
        Date now = new Date();
        ReportDefinitionX definition = new ReportDefinitionX();
        definition.setCellColor(IndexedColors.LIGHT_ORANGE.getIndex());

//        definition.setActiveSheetAt(0);
//        definition.setActiveSheet();
        definition.setOperatorName("2021");
        definition.setReportName("财务结账-其他奖金计提调整");
        definition.addParameter("年:", "2021");

        SXSSFWorkbook wb = new SXSSFWorkbook();
        Sheet sheet1 = wb.createSheet("整体");
        sheet1.setDisplayGridlines(false);
        Sheet sheet2 = wb.createSheet("部门明细表");
        sheet2.setDisplayGridlines(false);
        Sheet sheet3 = wb.createSheet("个人明细表");
        sheet3.setDisplayGridlines(false);
        definition.setWorkbook(wb);


        definition.initSheet(sheet1);
        ReportRowX headerRow2 = new ReportRowX();
        headerRow2.addCell("子公司2");
        headerRow2.addCell("一级部门2");
        headerRow2.addCell("预评职等2");
        headerRow2.addCellColor("二级部门2----------测试长度");
        definition.addHeaderRow(headerRow2);
        ReportRowX dataRow2 = new ReportRowX();
        dataRow2.addString("集团总部2");
        dataRow2.addNumber("2");
        dataRow2.addNumber("2222");
        dataRow2.addString("实际生效日期2015-12-1，申请补扣差额\r南二大区销售违规通222");
        definition.addRow(dataRow2);
        definition.prepared = false;
        definition.prepareExcel();

        definition.initSheet(sheet2);
        ReportRowX headerRow3 = new ReportRowX();
        headerRow3.addCell("子公司3");
        headerRow3.addCell("一级部门3");
        headerRow3.addCellColor("预评职等3------描述长度");
        headerRow3.addCell("二级部门3");
        definition.addHeaderRow(headerRow3);
        ReportRowX dataRow3 = new ReportRowX();
        dataRow3.addStringColor("集团总部3");
        dataRow3.addNumber("333");
        dataRow3.addNumberColor("33333");
        dataRow3.addString("实际生效日期2015-12-1，申请补扣差额\r南二大区销售违规通");
        definition.addRow(dataRow3);
        definition.prepared = false;
        definition.prepareExcel();

        definition.initSheet(sheet3);
        ReportRowX headerRow1 = new ReportRowX();
        headerRow1.addCellColor("团队");
        headerRow1.addCell("职能编码");
        headerRow1.addCell("一级部门");
        headerRow1.addCell("二级部门");
        headerRow1.addCell("三级部门");
        headerRow1.addCell("四级部门");
        headerRow1.addCell("五级部门");
        headerRow1.addCell("团队人数\n（含本人，不含实习生）");
        headerRow1.addCell("团队平均工时\n（不含新员工）");
        headerRow1.addCellColor("团队平无工时\n（含新员工）");
        definition.addHeaderRow(headerRow1);
        ReportRowX dataRow1 = new ReportRowX();
        dataRow1.addStringColor("姜羽");
        dataRow1.addString("PD");
        dataRow1.addString("网站运营中心");
        dataRow1.addString("内部系统");
        dataRow1.addString("商业产品部");
        dataRow1.addString("");
        dataRow1.addString("");
        dataRow1.addNumber("24");
        dataRow1.addNumber("20");
        dataRow1.addNumber("24");
        definition.addRow(dataRow1);
        definition.prepared = false;
        definition.prepareExcel();

        definition.buildExcel("D:/test1.xlsx");
//        definition.getSheet().autoSizeColumn(3);
        Date end = new Date();
        System.out.println("耗时(S)=" + ((end.getTime() - now.getTime()) / 1000L));
//
////       生成单sheet附件
//        Date nows = new Date();
//        ReportDefinition definitions = new ReportDefinition();
//        definitions.setCellColor(52);
////        definitions.setColumnWidth(4,9000);
////        CellStyle style = definitions.generateCellStyle();
////        style.setFillPattern((short) 1);
////        style.setFillForegroundColor((short) 16);
////        Font font = definitions.getWorkbook().createFont();
////        font.setColor(IndexedColors.WHITE.getIndex());
////        font.setBoldweight((short) 700);
////        style.setFont(font);
//        definitions.setOperatorName("管理员");
//        definitions.setReportName("Test Report");
//        ReportRow headerRows = new ReportRow();
////        headerRows.setWidth(100);
//        headerRows.addCellColor("LABEL1");
//        headerRows.addCell("LABEL2");
//        headerRows.addCell("LABEL3");
//        headerRows.addCell("LABEL4");
//        headerRows.addHeaderCellColor("LABEL5 测试一下label的长度");
////        headerRows.addCellWidthColor("LABEL5", 1000);
//
//        definitions.addHeaderRow(headerRows);
//        ReportRow dataRows = new ReportRow();
//        dataRows.addStringColor("11");
//        dataRows.addString("22");
//        dataRows.addString("33");
//        dataRows.addString("44");
//        dataRows.addBoldStringColor("55 测试一下label的长度测试一下label的长度");
//
//        definitions.addRow(dataRows);
//        definitions.prepareExcel();
////        definitions.setCellStyle(0, 0, style);
////        definitions.setPageNumberSizeAndFooter(16, "");
//
//        definitions.buildExcel("D:/sample1.xlsx");
////        definitions.getSheet().autoSizeColumn(5);
//        Date ends = new Date();
//        System.out.println("" + ((ends.getTime() - nows.getTime()) / 1000L));



        Date now11 = new Date();
        ReportDefinition definitionNew = new ReportDefinition();
        definitionNew.setCellColor(IndexedColors.LIGHT_YELLOW.getIndex());
        definitionNew.renameSheet(0, "renameSheet");
        definitionNew.setActiveSheetAt(0);
        definitionNew.setActiveSheet();
        System.out.println(definitionNew.getSheetName());
        definitionNew.setOperatorName("SHEET1");
        definitionNew.setReportName("Test Report");
        ReportRow headerRow11 = new ReportRow();
        headerRow11.addCell("Header1");
        headerRow11.addHeaderCell("Header2");
        headerRow11.addHeaderCellColor("Header3-测试一下下");
        headerRow11.addCellColor("Header4");
        definitionNew.addHeaderRow(headerRow11);
        ReportRow dataRow11 = new ReportRow();
        dataRow11.addString("Cell1");
        dataRow11.addString("Cell2");
        dataRow11.addString("Cell3");
        dataRow11.addString("Cell4");
        definitionNew.addRow(dataRow11);
        dataRow11 = new ReportRow();
        dataRow11.addString("Cell1");
        dataRow11.addString("Cell2");
        dataRow11.addStringColor("2021-11-23");
        dataRow11.addDate("2021-11-23");
        definitionNew.addRow(dataRow11);

        System.out.println(definitionNew.getCellColor());
        definitionNew.prepareExcel();
        definitionNew.CreateSheet("testSheetMul");
        definitionNew.setActiveSheetAt(1);
        definitionNew.setActiveSheet();
        System.out.println(definitionNew.getSheetName());
        definitionNew.setOperatorName("SHEET2");
        definitionNew.setReportName("Test Report");
        ReportRow headerRow21 = new ReportRow();
        headerRow21.addHeaderCellColor("测试1");
        headerRow21.addCell("测试2");
        headerRow21.addCell("测试3");
        headerRow21.addCellColor("测试4------测试长度");
        definitionNew.addHeaderRow(headerRow21);
        ReportRow dataRow21 = new ReportRow();
        dataRow21.addString("data1");
        dataRow21.addString("data2");
        dataRow21.addString("data3");
        dataRow21.addString("data4");
        definitionNew.addRow(dataRow21);
        dataRow21 = new ReportRow();
        dataRow21.addStringColor("data1");
        dataRow21.addString("data2");
        dataRow21.addString("data3");
        dataRow21.addString("data4");
        definitionNew.addRow(dataRow21);

        definitionNew.prepareExcel();
        definitionNew.setActiveSheetAt(0);
        definitionNew.buildExcel("D:/22222.xlsx");
        definitionNew.getSheet().autoSizeColumn(3);
        Date end11 = new Date();
        System.out.println("" + ((end11.getTime() - now11.getTime()) / 1000L));
    }
}
