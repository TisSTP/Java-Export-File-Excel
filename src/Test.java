import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Locale;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFDataFormat;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Test {
  private final SimpleDateFormat sdfTh = new SimpleDateFormat("dd/MM/yyyy", new Locale("th", "TH"));

  public static void main(String[] args) {
    System.out.println("Hello World!");

    XSSFWorkbook workbook = new XSSFWorkbook();
    XSSFSheet sheet = workbook.createSheet("new sheet");

    XSSFCellStyle cs = workbook.createCellStyle();
    XSSFDataFormat df = workbook.createDataFormat();
    cs.setDataFormat(df.getFormat("#,##0.00"));

    XSSFRow row = sheet.createRow((short) 0);

    XSSFCell cell = row.createCell((short) 0);
    cell.setCellValue(11111.1);
    cell.setCellStyle(cs);

    cell = row.createCell((short) 1);
    cell.setCellValue("01/09/2559");
    XSSFCellStyle csDate = workbook.createCellStyle();
    CreationHelper createHelper = workbook.getCreationHelper();
    csDate.setDataFormat(createHelper.createDataFormat().getFormat("dd/mm/yyyy"));
    cell.setCellStyle(csDate);

    cell = row.createCell((short) 2);
    cell.setCellValue(new Date());
    cell.setCellStyle(csDate);

    cell = row.createCell((short) 3);
    cell.setCellValue(111222333.99);
    cell.setCellStyle(cs);
    cell.setCellType(CellType.NUMERIC); // set type numeric


    // increase row height to accomodate two lines of text
    //row.setHeightInPoints((3 * spreadsheet.getDefaultRowHeightInPoints()));

    // merge
    CellRangeAddress mergedCellCenter = new CellRangeAddress(
        0, // first row (0-based)
        0, // last row (0-based)
        1, // first column (0-based)
        2 // last column (0-based)
    );
//        spreadsheet.addMergedRegion(mergedCellCenter);

    addBorderMergeStyle(sheet, mergedCellCenter, BorderStyle.THIN);

    // set widht cell
//      spreadsheet.setColumnWidth(1, 8000);

    // adjust column width to fit the content
    sheet.autoSizeColumn((short) 0);
    sheet.autoSizeColumn((short) 1);
    sheet.autoSizeColumn((short) 2);

    // Write the output to a file
    try {
//      String filenameExcel = "workbook.xlsx";
      FileOutputStream fileOut = new FileOutputStream("D:/projects-test/file/test-"+new Date().getTime()+".xlsx");
      workbook.write(fileOut);
      fileOut.close();
      System.out.println("D:/projects-test/file/test.xlsx");
    } catch (IOException e) {
      System.out.println("An error occurred. Stack trace is: ");
      e.getStackTrace();
    }

  }


  private static void addBorderMergeStyle(XSSFSheet sheet, CellRangeAddress mergedCell, BorderStyle border) {
    RegionUtil.setBorderTop(border, mergedCell, sheet);
    RegionUtil.setBorderBottom(border, mergedCell, sheet);
    RegionUtil.setBorderLeft(border, mergedCell, sheet);
    RegionUtil.setBorderRight(border, mergedCell, sheet);
  }

  private CellStyle borderCellStyle(XSSFWorkbook workbook, Boolean left, Boolean top, Boolean right, Boolean bottom, short weight, CellStyle style) {
    if (style == null) {
      style = workbook.createCellStyle();
    }

    if(left) {
      style.setBorderLeft(BorderStyle.valueOf(weight));
    }
    if (top) {
      style.setBorderTop(BorderStyle.valueOf(weight));
    }
    if(right) {
      style.setBorderRight(BorderStyle.valueOf(weight));
    }
    if(bottom) {
      style.setBorderBottom(BorderStyle.valueOf(weight));
    }

    return style;
  }

  private void writeCell(XSSFWorkbook workbook, XSSFCell cell, Object key, XSSFCellStyle style) {
    if(style != null) {
      cell.setCellStyle(style);
    } else {
      // create data format currency
      XSSFDataFormat df = workbook.createDataFormat();
      XSSFCellStyle csCurrency = workbook.createCellStyle();
      csCurrency.setDataFormat(df.getFormat("#,##0.00"));
      cell.setCellStyle(csCurrency);
    }
    if (key instanceof Integer) {
      if (key != null) {
        cell.setCellType(XSSFCell.CELL_TYPE_NUMERIC);
        cell.setCellValue((Integer) key);
      }
    } else if (key instanceof Double) {
      if (key != null) {
        cell.setCellType(XSSFCell.CELL_TYPE_NUMERIC);
        cell.setCellValue((Double) key);
      }
    } else if (key instanceof BigDecimal) {
      cell.setCellType(XSSFCell.CELL_TYPE_NUMERIC);
      if (key != null) {
        cell.setCellValue(((BigDecimal) key).doubleValue());
      } else {
        cell.setCellValue(0);
      }
    } else if (key instanceof Date) {
      if (key != null) {
        cell.setCellValue(chkDateDefault((Date) key));
      }
    } else if (key instanceof Long) {
      if (key != null) {
        cell.setCellType(XSSFCell.CELL_TYPE_NUMERIC);
        cell.setCellValue(((Long) key));
      }
    } else if (key instanceof String) {
      if (key != null) {
        cell.setCellType(XSSFCell.CELL_TYPE_STRING);
        cell.setCellValue((String) key);
      }
    }
  }

  private String chkDateDefault(Date date) {
    try {
      return sdfTh.format(date);
    } catch (Exception e) {
      //e.printStackTrace();
      return "xx/xx/xxxx";
    }
  }
}
