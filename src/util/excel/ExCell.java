package util.excel;

import java.math.BigDecimal;
import java.util.Date;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import util.Utility;
import util.excel.ExWorkBook.StyleBorder;
import util.excel.ExWorkBook.StyleCell;
import util.excel.ExWorkBook.StyleDataFormat;

public class ExCell {
  private ExWorkBook exWorkBook;
  private ExSheet exSheet;
  private Row row;
  private Cell cell;
  private CellStyle style;

  public ExCell(ExWorkBook exWorkBook, ExSheet exSheet, Row row, int index) {
    this.exWorkBook = exWorkBook;
    this.exSheet = exSheet;
    this.row = row;
    this.cell = this.row.createCell(index);
  }

  public Cell getCell() {
    return cell;
  }

  // ------------------------ STYLE ----------------------------
  public CellStyle getStyle() {
    return style;
  }
  public ExCell setStyle(CellStyle style) {
    this.style = style;
    return this;
  }
  public ExCell setStyle(StyleCell style) {
    this.style = this.exWorkBook.getDefaultStyle(style);
    return this;
  }
  public ExCell setStyle(StyleDataFormat style) {
    this.style = this.exWorkBook.getDefaultStyle(style);
    return this;
  }
  public ExCell setStyle(StyleBorder style) {
    this.style = this.exWorkBook.getDefaultStyle(style);
    return this;
  }
  // ---------------------- END STYLE --------------------------

  // Convert Data
  public String convertDateToThai(Date date) {
    try {
      return exWorkBook.SDF_TH.format(date);
    } catch (Exception e) {
      return exWorkBook.DEFAULT_DATE_STR;
    }
  }

  // ------------------------ Type --------------------------
  public void setTypeCurrency() {
    this.cell.setCellType(Cell.CELL_TYPE_NUMERIC);
    this.setStyle(this.exWorkBook.getDefaultStyle(StyleDataFormat.CURRENCY));
  }
  public void setTypeNumeric() {
    this.cell.setCellType(Cell.CELL_TYPE_NUMERIC);
  }
  public void setTypeString() {
    this.cell.setCellType(Cell.CELL_TYPE_STRING);
  }
  public void setTypeDate() {
    this.setStyle(this.exWorkBook.getDefaultStyle(StyleDataFormat.DATETHAI));
  }
  // ---------------------- End Type ------------------------

  // Set Value
  public void setValue(Object value) {
    this.writeCell(value);
    if(style == null) {
      // default
      this.setStyle(this.exWorkBook.getDefaultStyle(StyleCell.DEFAULT));
    }
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
