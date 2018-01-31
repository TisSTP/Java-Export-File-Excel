package util.excel;

import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Locale;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import util.Utility;

public class ExCell {
  private Workbook workbook;
  private Sheet sheet;
  private Row row;
  private Cell cell;
  private CellStyle style;
  private Font font;
  private DataFormat df;
  private CreationHelper createHelper;
  private short borderWeight;

  // config
  private final SimpleDateFormat SDF_TH = new SimpleDateFormat("dd/MM/yyyy", new Locale("th", "TH"));
  private final String FORMAT_DATE_STR = "dd/mm/yyyy";
  private final String DEFAULT_DATE_STR = "xx/xx/xxxx";
  private final String DEFAULT_CURRENCY_STR = "#,##0.00";

  public ExCell(Workbook workbook, Sheet sheet, Row row, int index) {
    this.workbook = workbook;
    this.sheet = sheet;
    this.row = row;
    this.cell = this.row.createCell(index);
    this.font = this.workbook.createFont();
    this.df = this.workbook.createDataFormat();
    this.style = this.workbook.createCellStyle();
    this.createHelper = this.workbook.getCreationHelper();
    this.borderWeight = CellStyle.BORDER_THIN;
    this.setStyleHorizontalCenter();
    this.setStyleVerticalCenter();
    this.setStyleWrapText();
  }

  // ------------------------ STYLE ----------------------------
  public CellStyle getStyle() {
    return style;
  }
  public void setStyle(CellStyle style) {
    this.style = style;
  }
  // Alignment
  public ExCell setStyleHorizontalLeft() {
    this.style.setAlignment(CellStyle.ALIGN_LEFT);
    return this;
  }
  public ExCell setStyleHorizontalCenter() {
    this.style.setAlignment(CellStyle.ALIGN_CENTER);
    return this;
  }
  public ExCell setStyleHorizontalRight() {
    this.style.setAlignment(CellStyle.ALIGN_RIGHT);
    return this;
  }
  public ExCell setStyleHorizontalJustify() {
    this.style.setAlignment(CellStyle.ALIGN_JUSTIFY);
    return this;
  }

  // Vertical Alignment
  public ExCell setStyleVerticalTop() {
    this.style.setVerticalAlignment(CellStyle.VERTICAL_TOP);
    return this;
  }
  public ExCell setStyleVerticalCenter() {
    this.style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
    return this;
  }
  public ExCell setStyleVerticalBottom() {
    this.style.setVerticalAlignment(CellStyle.VERTICAL_BOTTOM);
    return this;
  }
  public ExCell setStyleVerticalJustify() {
    this.style.setVerticalAlignment(CellStyle.VERTICAL_JUSTIFY);
    return this;
  }

  // Wrap Text
  public ExCell setStyleWrapText() {
    this.style.setWrapText(true);
    return this;
  }

  // Font
  public ExCell setFontBold() {
    this.font.setBold(true);
    return this;
  }
  public ExCell setFontUnderlineSingleA() {
    this.font.setUnderline(Font.U_SINGLE_ACCOUNTING);
    return this;
  }

  /**
   *
   * @param fontUnderline Font
   * @return
   */
  public ExCell setFontUnderline(byte fontUnderline) {
    this.font.setUnderline(fontUnderline);
    return this;
  }

  // Format Data
  public void setStyleDfCurrency() {
    this.style.setDataFormat(df.getFormat(DEFAULT_CURRENCY_STR));
  }
  public void setStyleDfDate() {
    this.style.setDataFormat(createHelper.createDataFormat().getFormat(FORMAT_DATE_STR));
  }

  // Border
  public short getBorderWeight() {
    return borderWeight;
  }
  public ExCell setBorderWeight(short borderWeight) {
    this.borderWeight = borderWeight;
    return this;
  }
  public ExCell setBorderAll() {
    this.setBorderLeft();
    this.setBorderRight();
    this.setBorderTop();
    this.setBorderBottom();
    return this;
  }
  public ExCell setBorderTop() {
    this.style.setBorderTop(borderWeight);
    return this;
  }
  public ExCell setBorderBottom() {
    this.style.setBorderBottom(borderWeight);
    return this;
  }
  public ExCell setBorderLeft() {
    this.style.setBorderLeft(borderWeight);
    return this;
  }
  public ExCell setBorderRight() {
    this.style.setBorderRight(borderWeight);
    return this;
  }

  // ---------------------- END STYLE --------------------------

  // Convert Data
  public String convertDateToThai(Date date) {
    try {
      return SDF_TH.format(date);
    } catch (Exception e) {
      return DEFAULT_DATE_STR;
    }
  }

  // ------------------------ Type --------------------------
  public ExCell setTypeCurrency() {
    this.cell.setCellType(Cell.CELL_TYPE_NUMERIC);
    this.setStyleDfCurrency();
    return this;
  }
  public void setTypeNumeric() {
    this.cell.setCellType(Cell.CELL_TYPE_NUMERIC);
  }
  public void setTypeString() {
    this.cell.setCellType(Cell.CELL_TYPE_STRING);
  }
  public void setTypeDate() {
    this.setStyleDfDate();
  }
  // ---------------------- End Type ------------------------


  // Set Value
  public void setValue(Object value) {
    this.writeCell(value);
    this.style.setFont(this.font);
    this.cell.setCellStyle(this.style);
  }

  private void writeCell(Object value) {
    if(value instanceof Integer) {
      this.setTypeNumeric();
      if (Utility.isValid(value)) {
        cell.setCellValue((Integer) value);
      } else {
        cell.setCellValue(0);
      }
    } else if(value instanceof Long) {
      this.setTypeNumeric();
      if (Utility.isValid(value)) {
        cell.setCellValue(((Long) value));
      }
    } else if(value instanceof Double) {
      this.setTypeNumeric();
      if (Utility.isValid(value)) {
        cell.setCellValue((Double) value);
      } else {
        cell.setCellValue(0.0);
      }
    } else if(value instanceof BigDecimal) {
      this.setTypeCurrency();
      this.setStyleHorizontalRight();
      if (Utility.isValid(value)) {
        cell.setCellValue(((BigDecimal) value).doubleValue());
      } else {
        cell.setCellValue(0);
      }
    } else if(value instanceof Date) {
      this.setTypeDate();
      if (Utility.isValid(value)) {
        cell.setCellValue(convertDateToThai((Date) value));
      }
    } else if(value instanceof String) {
      this.setTypeString();
      if (Utility.isValid(value)) {
        cell.setCellValue(((String) value).trim());
      }
    } else {
      cell.setCellValue("");
    }
  }

}
