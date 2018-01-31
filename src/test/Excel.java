package test;

import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import util.excel.ExCell;
import util.excel.ExRow;
import util.excel.ExSheet;
import util.excel.ExWorkBook;

public class Excel {
  private static final SimpleDateFormat sdfTh = new SimpleDateFormat("dd/MM/yyyy", new Locale("th", "TH"));
  public static void main(String[] args) {

    long startTime = System.nanoTime();

    ExWorkBook exWorkbook = new ExWorkBook();
    ExSheet sheet = exWorkbook.createSheet("ใบสรุปค่าบริการ รายตัวแทน ประจำงวด 1");
    int rowindex = 0;
    int year = 2561;

    String leftBanner, centerBanner, rightBanner;
    leftBanner = "report Excel 002 test";
    centerBanner = "บริษัทอาคเนย์ประกันภัย จำกัด (มหาชน)" + "\nใบสรุปค่าบริการ รายตัวแทน"
        + "\nประจำงวด 1 เดือน กุมภาพันธ์ ปี 2561\n ตั้งแต่ 01/02/2561 - 15/02/2561";
    rightBanner = "วันที่พิมพ์: " + sdfTh.format(new Date());
    sheet.createRowBannerReport(rowindex++, leftBanner, centerBanner, rightBanner, 0, 3, 11, 16, 4);

    // row find data
    ExRow row = sheet.createRow(rowindex++);
    ExCell cell = row.createCell(0);
    cell.setStyleHorizontalRight();
    cell.setValue("Represent Code: ");
    row.createCell(1).setValue("");

    row = sheet.createRow(rowindex++);
    cell = row.createCell(0);
    cell.setStyleHorizontalRight();
    cell.setValue("Represent Name: ");
    row.createCell(1).setValue("");

    row = sheet.createRow(rowindex++);
    cell = row.createCell(0);
    cell.setStyleHorizontalRight();
    cell.setValue("Master Agent Code: ");
    row.createCell(1).setValue("MsAgentName");


    row = sheet.createRow(rowindex++);
    cell = row.createCell(0);
    cell.setStyleHorizontalRight();
    cell.setValue("Agent Code: ");
    row.createCell(1).setValue("AgentName");

    row = sheet.createRow(rowindex++);
    cell = row.createCell(0);
    cell.setStyleHorizontalRight();
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
    headers.add("ันที่ออกกรมธรรม์");
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
    for (int i = 0; i < 2000; i++) {
      row = sheet.createRow(rowindex++);
      row.createCell().setStyleHorizontalLeft().setValue("1111111111111111");
      row.createCell().setStyleHorizontalCenter().setValue("2222222222222");
      row.createCell().setStyleHorizontalCenter().setValue("333333333333333");
      row.createCell().setValue("Product Group");
      row.createCell().setValue("Product Code");
      row.createCell().setValue("Product Plan");
      row.createCell().setStyleHorizontalCenter().setValue("01/01/2561");
      row.createCell().setStyleHorizontalCenter().setValue("01/01/2561");
      row.createCell().setStyleHorizontalCenter().setValue("01/01/2561");
      row.createCell().setValue("");
      row.createCell().setValue("");
      row.createCell().setTypeCurrency().setValue(1000);
      row.createCell().setTypeCurrency().setValue(2000);
      row.createCell().setTypeCurrency().setValue(1000000);
      row.createCell().setTypeCurrency().setValue(1100000000);
      row.createCell().setTypeCurrency().setValue(999999999);
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
    sheet.autoSizeColumns(new int[]{6, 7, 11, 15});
    sheet.setColumnsWidth(map);
    sheet.lenderMergeCell();
    sheet.lenderBorderMergeCell();

    System.out.println(("Complete Set Row : Index = [" + rowindex + "]"));
    System.out.println((
        "Generated Excel In : " + ((System.nanoTime() - startTime) / 1000000000.0) + " Seconds"));

    exWorkbook.exportFile("T0000001");
  }

}
