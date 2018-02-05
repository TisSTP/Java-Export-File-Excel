package util.excel;

import java.io.ByteArrayOutputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.HashMap;
import java.util.Locale;
import java.util.Map;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

public class ExWorkBook {
  private Workbook workbook;
  private int numSheet = 0;
  private ByteArrayOutputStream outStream;
  private Map<StyleCell, CellStyle> styleList;
  private Map<StyleDataFormat, CellStyle> dataFormatList;
  private Map<StyleBorder, CellStyle> borderList;
  private Font fontBold;
  private DataFormat df;
  private CreationHelper createHelper;

  // config & default
  public final SimpleDateFormat SDF_TH = new SimpleDateFormat("dd/MM/yyyy", new Locale("th", "TH"));
  private final String FORMAT_CURRENCY_STR = "#,##0.00";
  private final String FORMAT_DATE_STR = "dd/mm/yyyy";
  public final String DEFAULT_DATE_STR = "xx/xx/xxxx";
  private final short weightBorder = CellStyle.BORDER_THIN;

  public enum StyleCell {
    BANNERLEFT,
    BANNERCENTER,
    BANNERRIGHT,
    HEADERLEFT,
    HEADERCENTER,
    HEADERRIGHT,
    FOOTERLEFT,
    FOOTERCENTER,
    FOOTERRIGHT,
    CONTENTLEFT,
    CONTENTCENTER,
    CONTENTCENTER_FB,
    CONTENTRIGHT,
    DEFAULT
  }

  public enum StyleDataFormat {
    CURRENCY,
    CURRENCY_FB,
    CURRENCY_FB_UND, // Font Bold & Underline
    DATETHAI
  }

  public enum StyleBorder {
    ALL
  }

  public ExWorkBook() {
    this(new SXSSFWorkbook(), new ByteArrayOutputStream());
  }

  public ExWorkBook(Workbook workbook, ByteArrayOutputStream outStream) {
    styleList = new HashMap<>();
    dataFormatList = new HashMap<>();
    borderList = new HashMap<>();
    numSheet = 0;

    if(workbook != null) {
      this.workbook = workbook;
    } else {
      this.workbook = new SXSSFWorkbook();
    }
    if(outStream != null) {
      this.outStream = outStream;
    } else {
      this.outStream = new ByteArrayOutputStream();
    }

    this.setDefaultStyle();
  }

  public Workbook getWorkbook() {
    return workbook;
  }

  public void setWorkbook(Workbook workbook) {
    this.workbook = workbook;
  }

  public ExSheet createSheet(String name) {
    numSheet++;
    if(name == null || name.trim().equals("")) {
      return new ExSheet(this, "sheet".concat(String.valueOf(numSheet)));
    }
    return new ExSheet(this, name);
  }

  public byte[] exportBytes() {
    byte[] bytes = null;
    try {
      workbook.write(outStream);
      bytes = outStream.toByteArray();
    } catch (Exception e) {
      e.printStackTrace();
    } finally {
      IOUtils.closeQuietly(outStream);
      try {
        workbook.close();
      } catch (IOException e) {
        e.printStackTrace();
      }
    }
    return bytes;
  }

  public void exportFile(String fileName) {
    try {
      // Write the output to a file
      FileOutputStream fileOut = new FileOutputStream(fileName + ".xlsx");
      this.workbook.write(fileOut);
      fileOut.close();
    } catch (Exception e) {
      e.printStackTrace();
    }
  }
  public CellStyle getDefaultStyle(StyleDataFormat styleDataFormat) {
    return dataFormatList.get(styleDataFormat);
  }
  public CellStyle getDefaultStyle(StyleBorder styleBorder) {
    return borderList.get(styleBorder);
  }
  public CellStyle getDefaultStyle(StyleCell styleCell) {
    return styleList.get(styleCell);
  }

  private void setDefaultStyle() {
    // ----------- Create Font ------------
    this.fontBold = workbook.createFont();
    fontBold.setBold(true);
    Font fntBoldUndline = workbook.createFont();
    fntBoldUndline.setBold(true);
    fntBoldUndline.setUnderline(Font.U_SINGLE_ACCOUNTING);

    // ---------- Data Format ------------
    this.df = workbook.createDataFormat();

    // ---------- Create Helper ----------
    this.createHelper = workbook.getCreationHelper();

    // ------------------- Cell Style List -------------------
    // DEFAULT
    CellStyle cs = workbook.createCellStyle();
    cs.setWrapText(true);
    cs.setAlignment(CellStyle.ALIGN_CENTER);
    cs.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
    this.styleList.put(StyleCell.DEFAULT, cs);

    // BANNER LEFT
    cs = workbook.createCellStyle();
    cs.setFont(fontBold);
    cs.setWrapText(true);
    cs.setAlignment(CellStyle.ALIGN_LEFT);
    cs.setVerticalAlignment(CellStyle.VERTICAL_TOP);
    this.styleList.put(StyleCell.BANNERLEFT, cs);

    // BANNER CENTER
    cs = workbook.createCellStyle();
    cs.setFont(fontBold);
    cs.setWrapText(true);
    cs.setAlignment(CellStyle.ALIGN_CENTER);
    cs.setVerticalAlignment(CellStyle.VERTICAL_TOP);
    this.styleList.put(StyleCell.BANNERCENTER, cs);

    // BANNER RIGHT
    cs = workbook.createCellStyle();
    cs.setFont(fontBold);
    cs.setWrapText(true);
    cs.setAlignment(CellStyle.ALIGN_RIGHT);
    cs.setVerticalAlignment(CellStyle.VERTICAL_TOP);
    this.styleList.put(StyleCell.BANNERRIGHT, cs);

    // HEADER LEFT
    cs = workbook.createCellStyle();
    cs.setFont(fontBold);
    cs.setWrapText(true);
    cs.setAlignment(CellStyle.ALIGN_LEFT);
    cs.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
    cs.setBorderLeft(weightBorder);
    cs.setBorderTop(weightBorder);
    cs.setBorderRight(weightBorder);
    cs.setBorderBottom(weightBorder);
    this.styleList.put(StyleCell.HEADERLEFT, cs);

    // HEADER CENTER
    cs = workbook.createCellStyle();
    cs.setFont(fontBold);
    cs.setWrapText(true);
    cs.setAlignment(CellStyle.ALIGN_CENTER);
    cs.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
    cs.setBorderLeft(weightBorder);
    cs.setBorderTop(weightBorder);
    cs.setBorderRight(weightBorder);
    cs.setBorderBottom(weightBorder);
    this.styleList.put(StyleCell.HEADERCENTER, cs);

    // HEADER RIGHT
    cs = workbook.createCellStyle();
    cs.setFont(fontBold);
    cs.setWrapText(true);
    cs.setAlignment(CellStyle.ALIGN_RIGHT);
    cs.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
    cs.setBorderLeft(weightBorder);
    cs.setBorderTop(weightBorder);
    cs.setBorderRight(weightBorder);
    cs.setBorderBottom(weightBorder);
    this.styleList.put(StyleCell.HEADERRIGHT, cs);

    // FOOTER LEFT
    cs = workbook.createCellStyle();
    cs.setFont(fontBold);
    cs.setWrapText(true);
    cs.setAlignment(CellStyle.ALIGN_LEFT);
    cs.setVerticalAlignment(CellStyle.VERTICAL_BOTTOM);
    this.styleList.put(StyleCell.FOOTERLEFT, cs);

    // FOOTER CENTER
    cs = workbook.createCellStyle();
    cs.setFont(fontBold);
    cs.setWrapText(true);
    cs.setAlignment(CellStyle.ALIGN_CENTER);
    cs.setVerticalAlignment(CellStyle.VERTICAL_BOTTOM);
    this.styleList.put(StyleCell.FOOTERCENTER, cs);

    // FOOTER RIGHT
    cs = workbook.createCellStyle();
    cs.setFont(fontBold);
    cs.setWrapText(true);
    cs.setAlignment(CellStyle.ALIGN_RIGHT);
    cs.setVerticalAlignment(CellStyle.VERTICAL_BOTTOM);
    this.styleList.put(StyleCell.FOOTERRIGHT, cs);

    // CONTENT LEFT
    cs = workbook.createCellStyle();
    cs.setWrapText(true);
    cs.setAlignment(CellStyle.ALIGN_LEFT);
    cs.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
    this.styleList.put(StyleCell.CONTENTLEFT, cs);

    // CONTENT CENTER
    cs = workbook.createCellStyle();
    cs.setWrapText(true);
    cs.setAlignment(CellStyle.ALIGN_CENTER);
    cs.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
    this.styleList.put(StyleCell.CONTENTCENTER, cs);

    // CONTENT CENTER FONT BOLD
    cs = workbook.createCellStyle();
    cs.setFont(fontBold);
    cs.setWrapText(true);
    cs.setAlignment(CellStyle.ALIGN_CENTER);
    cs.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
    this.styleList.put(StyleCell.CONTENTCENTER_FB, cs);

    // CONTENT RIGHT
    cs = workbook.createCellStyle();
    cs.setWrapText(true);
    cs.setAlignment(CellStyle.ALIGN_RIGHT);
    cs.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
    this.styleList.put(StyleCell.CONTENTRIGHT, cs);

    // ---------------- Data Format List ------------------
    // Style Data Format : CURRENCY
    cs = workbook.createCellStyle();
    cs.setWrapText(true);
    cs.setAlignment(CellStyle.ALIGN_RIGHT);
    cs.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
    cs.setDataFormat(df.getFormat(FORMAT_CURRENCY_STR));
    this.dataFormatList.put(StyleDataFormat.CURRENCY, cs);

    // Style Data Format : CURRENCY
    cs = workbook.createCellStyle();
    cs.setFont(fontBold);
    cs.setWrapText(true);
    cs.setAlignment(CellStyle.ALIGN_RIGHT);
    cs.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
    cs.setDataFormat(df.getFormat(FORMAT_CURRENCY_STR));
    this.dataFormatList.put(StyleDataFormat.CURRENCY_FB, cs);

    // Style Data Format : CURRENCY
    cs = workbook.createCellStyle();
    cs.setFont(fntBoldUndline);
    cs.setWrapText(true);
    cs.setAlignment(CellStyle.ALIGN_RIGHT);
    cs.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
    cs.setDataFormat(df.getFormat(FORMAT_CURRENCY_STR));
    this.dataFormatList.put(StyleDataFormat.CURRENCY_FB_UND, cs);

    // Style Data Format : DATETHAI
    cs = workbook.createCellStyle();
    cs.setWrapText(true);
    cs.setAlignment(CellStyle.ALIGN_CENTER);
    cs.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
    cs.setDataFormat(createHelper.createDataFormat().getFormat(FORMAT_DATE_STR));
    this.dataFormatList.put(StyleDataFormat.DATETHAI, cs);

    // StyleBorder : ALL
    cs = workbook.createCellStyle();
    cs.setWrapText(true);
    cs.setAlignment(CellStyle.ALIGN_CENTER);
    cs.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
    cs.setBorderLeft(weightBorder);
    cs.setBorderTop(weightBorder);
    cs.setBorderRight(weightBorder);
    cs.setBorderBottom(weightBorder);
    this.borderList.put(StyleBorder.ALL, cs);
  }
}
