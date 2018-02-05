package util.excel;

import org.apache.poi.ss.usermodel.Row;

public class ExRow {
  private ExWorkBook exWorkBook;
  private ExSheet exSheet;
  private Row row;
  private int cellInt;

  public ExRow(ExWorkBook exWorkBook, ExSheet exSheet, int index) {
    this.exWorkBook = exWorkBook;
    this.exSheet = exSheet;
    this.row = this.exSheet.getSheet().createRow(index);
    this.cellInt = 0;
  }

  public ExCell createCell(int index) {
    ExCell cell = new ExCell(this.exWorkBook, this.exSheet, this.row, index);
    this.cellInt = index + 1;
    return cell;
  }

  public ExCell createCell() {
    return this.createCell(this.cellInt++);
  }

  public void setHeightInPoints(int line) {
    this.row.setHeightInPoints((line * this.exSheet.getSheet().getDefaultRowHeightInPoints()));
  }

  public Row getRow() {
    return row;
  }
}
