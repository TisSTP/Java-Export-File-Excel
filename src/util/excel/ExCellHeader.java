package util.excel;

import java.util.ArrayList;
import java.util.List;

public class ExCellHeader {
  private int rowId;
  private int rowBegin;
  private int rowEnd;
  private ExSheet sheet;
  private List<String> names;
  private List<int[]> merges;
  private boolean isAddFirst;

  public ExCellHeader(ExSheet sheet, int rowId) {
    this.sheet = sheet;
    this.rowId = rowId;
    this.names = new ArrayList<>();
    this.merges = new ArrayList<>();
    this.isAddFirst = true;
  }

  public int getRowId() {
    this.process();
    return rowId;
  }

  /**
   * create header
   * @param title String : data show
   * @param fRow Integer : first row index
   * @param lRow Integer : last row index
   * @param fCol Integer : first column index
   * @param lCol Integer : last column index
   */
  public void createHeader(String title, int fRow, int lRow, int fCol, int lCol) {
    this.names.add(title);
    this.merges.add(new int[]{fRow, lRow, fCol, lCol});
    // check row
    if(isAddFirst) {
      this.rowBegin = fRow;
      this.rowEnd = lRow;
      this.isAddFirst = false;
    } else {
      if(rowBegin > fRow) {
        this.rowBegin = fRow;
      }
      if(rowEnd < lRow) {
        this.rowEnd = lRow;
      }
    }
  }

  private void process() {
    int numRow = rowEnd - rowBegin + 1;

    ExRow rows[] = new ExRow[numRow];
    // create row
    for (int i = rowBegin; i <= rowEnd; i++) {
      rows[i - rowBegin] = sheet.createRow(i);
      rowId++;
    }
    // row
    if(merges != null && !merges.isEmpty() && names != null && !names.isEmpty()
        && merges.size() == names.size()) {
      int[] c;
      for(int i = 0; i < merges.size(); i++) {
        // check create row
        c = merges.get(i);
        if(c[0] == c[1] && c[2] == c[3]) {
          rows[c[0] - rowBegin].createCell(c[2])
              .setValue(names.get(i));
        } else {
          rows[c[0] - rowBegin].createCell(c[2])
              .setValue(names.get(i));
          // set merge cell
          this.sheet.setMergeCell(c[0], c[1], c[2], c[3]);
          this.sheet.setBorderMergeCell(c[0], c[1], c[2], c[3]);
        }
      }
    }
  }
}
