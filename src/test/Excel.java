package test;

import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import util.excel.ExCell;
import util.excel.ExRow;
import util.excel.ExSheet;
import util.excel.ExWorkBook;
import util.excel.ExWorkBook.StyleCell;
import util.excel.ExWorkBook.StyleDataFormat;

public class Excel {
  private static final SimpleDateFormat sdfTh = new SimpleDateFormat("dd/MM/yyyy", new Locale("th", "TH"));

  public static void main(String[] args) {

    long startTime = System.nanoTime();

    ExWorkBook exWorkbook = new ExWorkBook();

//    Font fontBold = exWorkbook.getWorkbook().createFont();
//    CellStyle cs = exWorkbook.getWorkbook().createCellStyle();
//    cs.setFont(fontBold);
//    cs.setWrapText(true);
//    cs.setAlignment(CellStyle.ALIGN_CENTER);
//    cs.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
//    cs.setBorderLeft(CellStyle.BORDER_THIN);
//    cs.setBorderTop(CellStyle.BORDER_THIN);
//    cs.setBorderRight(CellStyle.BORDER_THIN);
//    cs.setBorderBottom(CellStyle.BORDER_THIN);

    // -------------------------- sheet 1 --------------------------
    ExSheet sheet = exWorkbook.createSheet("sheet1");
    int rowindex = 0;
    int sum = 0;
    int year = 2561;

    String leftBanner, centerBanner, rightBanner;
    leftBanner = "report Excel 002 test";
    centerBanner = "บริษัทอาคเนย์ประกันภัย จำกัด (มหาชน)" + "\r\nใบสรุปค่าบริการ รายตัวแทน"
        + "\r\nประจำงวด 1 เดือน กุมภาพันธ์ ปี 2561\r\nตั้งแต่ 01/02/2561 - 15/02/2561";
    rightBanner = "วันที่พิมพ์: " + sdfTh.format(new Date());
    sheet.createRowBannerReport(rowindex++, leftBanner, centerBanner, rightBanner, 0, 3, 11, 16, 4);

    // row find data
    ExRow row = sheet.createRow(rowindex++);
    ExCell cell = row.createCell(0);
    cell.setStyle(StyleCell.CONTENTRIGHT);
    cell.setValue("Represent Code: ");
    row.createCell(1).setValue("");

    row = sheet.createRow(rowindex++);
    cell = row.createCell(0);
    cell.setStyle(StyleCell.CONTENTRIGHT);
    cell.setValue("Represent Name: ");
    row.createCell(1).setValue("");

    row = sheet.createRow(rowindex++);
    cell = row.createCell(0);
    cell.setStyle(StyleCell.CONTENTRIGHT);
    cell.setValue("Master Agent Code: ");
    row.createCell(1).setValue("MsAgentName");


    row = sheet.createRow(rowindex++);
    cell = row.createCell(0);
    cell.setStyle(StyleCell.CONTENTRIGHT);
    cell.setValue("Agent Code: ");
    row.createCell(1).setValue("AgentName");

    row = sheet.createRow(rowindex++);
    cell = row.createCell(0);
    cell.setStyle(StyleCell.CONTENTRIGHT);
    cell.setValue("ที่อยู่: ");
    row.createCell(1).setValue("AgentAddr");
    // end row find data

    // row header content --------------------------------
    List<String> headers = new ArrayList<>();
    headers.add("ชื่อผู้เอาประกันภัย");
    headers.add("กรมธรรม์");
    headers.add("Product Type");
    headers.add("Product Group");
    headers.add("Product Code");
    headers.add("Product Plan");
    headers.add("วันที่ออกกรมธรรม์");
    headers.add("วันที่คุ้มครอง");
    headers.add("วันที่ชำระ");
    headers.add("SUB UNIT");
    headers.add("CHASSI");
    headers.add("REGNO");
    headers.add("เบี้ยประกันสุทธิ");
    headers.add("ค่าบริการ %");
    headers.add("จำนวนเงิน");
    headers.add("ค่าบริการจ่าย");
    headers.add("คงเหลือ");
    sheet.createRowHeaderReport(rowindex++, headers);
    // end row header content ----------------------------

    // row content -----------------------------------------
    for (int i = 0; i < 10_000; i++) { // error 1818 : row 1825
      row = sheet.createRow(rowindex++);
      row.createCell().setStyle(StyleCell.CONTENTLEFT).setValue("1111111111111111");
      row.createCell().setStyle(StyleCell.CONTENTRIGHT).setValue("2222222222222");
      row.createCell().setValue("333333333333333");
      row.createCell().setValue("Product Group");
      row.createCell().setValue("Product Code");
      row.createCell().setValue("Product Plan");
      row.createCell().setValue("01/01/2561");
      row.createCell().setValue("01/01/2561");
      row.createCell().setValue("01/01/2561");
      row.createCell().setValue("");
      row.createCell().setValue("");
      row.createCell().setStyle(StyleDataFormat.CURRENCY).setValue((1000));
      row.createCell().setStyle(StyleDataFormat.CURRENCY).setValue((2000));
      row.createCell().setStyle(StyleDataFormat.CURRENCY).setValue((1000000));
      row.createCell().setStyle(StyleDataFormat.CURRENCY).setValue((1100000000));
      row.createCell().setStyle(StyleDataFormat.CURRENCY).setValue((999999999));
      row.createCell().setValue("");
      row.createCell().setValue("");
    }
    // end row content -------------------------------------

    // sheet
    Map<Integer, Integer> map = new HashMap<Integer, Integer>();
    map.put(0, 8_000);
    map.put(1, 5_000);
    map.put(2, 5_000);
    map.put(3, 5_000);
    map.put(4, 5_000);
    map.put(5, 5_000);
    sheet.autoSizeColumns(new int[]{6, 7, 8, 11, 15});
    sheet.setColumnsWidth(map);
    sheet.lenderMergeCell();
    sheet.lenderBorderMergeCell();

    sum = rowindex;
    // -------------------------- sheet 2 --------------------------
//    sheet = exWorkbook.createSheet("sheet2");
//    rowindex = 0;
//    leftBanner = "report Excel 002 test";
//    centerBanner = "บริษัทอาคเนย์ประกันภัย จำกัด (มหาชน)" + "\n\rใบสรุปค่าบริการ รายตัวแทน"
//        + "\n\rประจำงวด 1 เดือน กุมภาพันธ์ ปี 2561\n\r ตั้งแต่ 01/02/2561 - 15/02/2561";
//    rightBanner = "วันที่พิมพ์: " + sdfTh.format(new Date());
//    sheet.createRowBannerReport(rowindex++, leftBanner, centerBanner, rightBanner, 0, 3, 11, 16, 4);
//
//    // row find data
//    row = sheet.createRow(rowindex++);
//    cell = row.createCell(0);
//    cell.setStyleHorizontalRight();
//    cell.setValue("Represent Code: ");
//    row.createCell(1).setValue("");
//
//    row = sheet.createRow(rowindex++);
//    cell = row.createCell(0);
//    cell.setStyleHorizontalRight();
//    cell.setValue("Represent Name: ");
//    row.createCell(1).setValue("");
//
//    row = sheet.createRow(rowindex++);
//    cell = row.createCell(0);
//    cell.setStyleHorizontalRight();
//    cell.setValue("Master Agent Code: ");
//    row.createCell(1).setValue("MsAgentName");
//
//
//    row = sheet.createRow(rowindex++);
//    cell = row.createCell(0);
//    cell.setStyleHorizontalRight();
//    cell.setValue("Agent Code: ");
//    row.createCell(1).setValue("AgentName");
//
//    row = sheet.createRow(rowindex++);
//    cell = row.createCell(0);
//    cell.setStyleHorizontalRight();
//    cell.setValue("ที่อยู่: ");
//    row.createCell(1).setValue("AgentAddr");
//    // end row find data
//
//    // row header content --------------------------------
//    headers = new ArrayList<>();
//    headers.add("ชื่อผู้เอาประกันภัย");
//    headers.add("กรมธรรม์");
//    headers.add("Product Type");
//    headers.add("Product Group");
//    headers.add("Product Code");
//    headers.add("Product Plan");
//    headers.add("ันที่ออกกรมธรรม์");
//    headers.add("วันที่คุ้มครอง");
//    headers.add("วันที่ชำระ");
//    headers.add("SUB UNIT");
//    headers.add("CHASSI");
//    headers.add("REGNO");
//    headers.add("เบี้ยประกันสุทธิ");
//    headers.add("ค่าบริการ %");
//    headers.add("จำนวนเงิน");
//    headers.add("ค่าบริการจ่าย");
//    headers.add("คงเหลือ");
//    sheet.createRowHeaderReport(rowindex++, headers);
//    // end row header content ----------------------------
//
//    // row content -----------------------------------------
//    for (int i = 0; i < 10000; i++) { // error 1818 : row 1825
//      row = sheet.createRow(rowindex++);
//      row.createCell().setStyleHorizontalLeft().setValue("1111111111111111");
//      row.createCell().setStyleHorizontalCenter().setValue("2222222222222");
//      row.createCell().setStyleHorizontalCenter().setValue("333333333333333");
//      row.createCell().setValue("Product Group");
//      row.createCell().setValue("Product Code");
//      row.createCell().setValue("Product Plan");
//      row.createCell().setStyleHorizontalCenter().setValue("01/01/2561");
//      row.createCell().setStyleHorizontalCenter().setValue("01/01/2561");
//      row.createCell().setStyleHorizontalCenter().setValue("01/01/2561");
//      row.createCell().setValue("");
//      row.createCell().setValue("");
//      row.createCell().setTypeCurrency().setValue(1000);
//      row.createCell().setTypeCurrency().setValue(2000);
//      row.createCell().setTypeCurrency().setValue(1000000);
//      row.createCell().setTypeCurrency().setValue(1100000000);
//      row.createCell().setTypeCurrency().setValue(999999999);
//      row.createCell().setValue("");
//      row.createCell().setValue("");
//    }
//    // end row content -------------------------------------
//
//    // sheet
//    map = new HashMap<Integer, Integer>();
//    map.put(0, 8_000);
//    map.put(1, 5_000);
//    map.put(2, 5_000);
//    map.put(3, 5_000);
//    map.put(4, 5_000);
//    map.put(5, 5_000);
////    sheet.autoSizeColumns(new int[]{6, 7, 8, 11, 15});
//    sheet.setColumnsWidth(map);
//    sheet.lenderMergeCell();
//    sheet.lenderBorderMergeCell();
//    sum += rowindex;

    System.out.println(("Complete Set Row : Index = [" + sum + "]"));
    exWorkbook.exportFile("T0000001");
    System.out.println((
        "Generated Excel In : " + ((System.nanoTime() - startTime) / 1000000000.0) + " Seconds"));
  }

}
