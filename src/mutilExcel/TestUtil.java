package mutilExcel;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.util.Date;

public class TestUtil {

    public static void main(String[] args) {

//      多sheet附件
        Date now = new Date();
        ReportDefinitionX definition = new ReportDefinitionX();
        SXSSFWorkbook wb = new SXSSFWorkbook();
        Sheet sheet1 = wb.createSheet("整体");
        sheet1.setDisplayGridlines(false);
        Sheet sheet2 = wb.createSheet("部门明细表");
        sheet2.setDisplayGridlines(false);
        Sheet sheet3 = wb.createSheet("个人明细表");
        sheet3.setDisplayGridlines(false);
        definition.setWorkbook(wb);
        definition.initSheet(sheet1);
        ReportRowX headerRow1 = new ReportRowX();
        headerRow1.addCell("团队");
        headerRow1.addCell("职能编码");
        headerRow1.addCell("一级部门");
        headerRow1.addCell("二级部门");
        headerRow1.addCell("三级部门");
        headerRow1.addCell("四级部门");
        headerRow1.addCell("五级部门");
        headerRow1.addCell("团队人数\n（含本人，不含实习生）");
        headerRow1.addCell("团队平均工时\n（不含新员工）");
        headerRow1.addCell("团队平无工时\n（含新员工）");
        definition.addHeaderRow(headerRow1);
        ReportRowX dataRow1 = new ReportRowX();
        dataRow1.addString("姜羽");
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
        definition.initSheet(sheet2);
        ReportRowX headerRow2 = new ReportRowX();
        headerRow2.addCell("子公司2");
        headerRow2.addCell("一级部门2");
        headerRow2.addCell("预评职等2");
        headerRow2.addCell("二级部门2");
        definition.addHeaderRow(headerRow2);
        ReportRowX dataRow2 = new ReportRowX();
        dataRow2.addString("集团总部2");
        dataRow2.addNumber("2");
        dataRow2.addNumber("2222");
        dataRow2.addString("实际生效日期2015-12-1，申请补扣差额\r南二大区销售违规通222");
        definition.addRow(dataRow2);
        definition.prepared = false;
        definition.prepareExcel();
        definition.initSheet(sheet3);
        ReportRowX headerRow3 = new ReportRowX();
        headerRow3.addCell("子公司3");
        headerRow3.addCell("一级部门3");
        headerRow3.addCell("预评职等3");
        headerRow3.addCell("二级部门3");
        definition.addHeaderRow(headerRow3);
        ReportRowX dataRow3 = new ReportRowX();
        dataRow3.addString("集团总部3");
        dataRow3.addNumber("333");
        dataRow3.addNumber("33333");
        dataRow3.addString("实际生效日期2015-12-1，申请补扣差额\r南二大区销售违规通");
        definition.addRow(dataRow3);
        definition.prepared = false;
        definition.prepareExcel();
        definition.buildExcel("D:/test1.xlsx");
        definition.getSheet().autoSizeColumn(3);
        Date end = new Date();
        System.out.println("耗时(S)=" + ((end.getTime() - now.getTime()) / 1000L));

//       生成单sheet附件
        Date nows = new Date();
        ReportDefinition definitions = new ReportDefinition();
        CellStyle style = definitions.generateCellStyle();
        style.setFillPattern((short) 1);
        style.setFillForegroundColor((short) 16);
        Font font = definitions.getWorkbook().createFont();
        font.setColor(IndexedColors.WHITE.getIndex());
        font.setBoldweight((short) 700);
        style.setFont(font);
        definitions.setOperatorName("管理员");
        definitions.setReportName("Test Report");
        ReportRow headerRows = new ReportRow();
//        headerRows.setWidth(100);
        headerRows.addCell("LABEL1");
        headerRows.addCell("LABEL2");
        headerRows.addCell("LABEL3");
        headerRows.addCell("LABEL4");
//        headerRows.addHeaderCell("LABEL5");
        headerRows.addCellWidth("LABEL5", 1000);

        definitions.addHeaderRow(headerRows);
        ReportRow dataRows = new ReportRow();
        dataRows.addString("11");
        dataRows.addString("22");
        dataRows.addString("33");
        dataRows.addString("44");
        dataRows.addString("55");

        definitions.addRow(dataRows);
        definitions.prepareExcel();
        definitions.setCellStyle(0, 0, style);
        definitions.setPageNumberSizeAndFooter(16, "");
        definitions.buildExcel("D:/sample.xlsx");
        definitions.getSheet().autoSizeColumn(5);
        Date ends = new Date();
        System.out.println("" + ((ends.getTime() - nows.getTime()) / 1000L));

    }
}
