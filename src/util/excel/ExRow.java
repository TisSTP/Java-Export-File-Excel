package util.excel;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class ExRow {
  private Workbook workbook;
  private Sheet sheet;
  private Row row;
  private int cellInt;

  public ExRow(Workbook workbook, Sheet sheet, int index) {
    this.workbook = workbook;
    this.sheet = sheet;
    this.row = this.sheet.createRow(index);
    this.cellInt = 0;
  }

  public ExCell createCell(int index) {
    ExCell cell = new ExCell(this.workbook, this.sheet, this.row, index);
    this.cellInt = index + 1;
    return cell;
  }

  public void setHeightInPoints(int line) {
    this.row.setHeightInPoints((line * this.sheet.getDefaultRowHeightInPoints()));
  }

  public ExCell createCell() {
    return this.createCell(this.cellInt++);
  }

}
