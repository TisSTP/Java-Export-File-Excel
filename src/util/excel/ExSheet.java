package util.excel;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;

public class ExSheet {
  private Workbook workbook;
  private Sheet sheet;
  private int rowInt;
  private List<Integer[]> mergeLists;
  private List<Integer[]> borderMergeLists;

  public ExSheet(Workbook workbook) {
    this(workbook, null);
  }

  public ExSheet(Workbook workbook, String name) {
    this.workbook = workbook;
    this.clearMergeLists();
    this.clearBorderMergeLists();
    this.rowInt = 0;
    if(name != null) {
      this.sheet = this.workbook.createSheet(name);
    }
  }

  public Sheet getSheet() {
    return sheet;
  }

  /**
   * create row in sheet : run auto row index
   * @return ExRow
   */
  public ExRow createRow() {
    return this.createRow(this.rowInt++);
  }

  /**
   * create row in sheet
   * @param index Integer : row index
   * @return ExRow
   */
  public ExRow createRow(int index) {
    return new ExRow(this.workbook, this.sheet, index);
  }

  // ------------------------------- merge cell -------------------------------
  /**
   * add merge cell to Lists
   * @param fRow Integer : first row index
   * @param lRow Integer : last row index
   * @param fCol Integer : first column index
   * @param lCol Integer : last column index
   */
  public void setMergeCell(int fRow, int lRow, int fCol, int lCol) {
    Integer i[] = new Integer[4];
    i[0] = fRow;
    i[1] = lRow;
    i[2] = fCol;
    i[3] = lCol;
    this.mergeLists.add(i);
  }

  public ExCellHeader setMergeCellHeader(int rowId) {
    return new ExCellHeader(this, rowId);
  }

  /**
   * process add Merged Region
   */
  public void lenderMergeCell() {
    // add merge
    if(this.mergeLists != null && this.mergeLists.size() > 0) {
      for (Integer[] merge : this.mergeLists) {
        CellRangeAddress cellRangeAddress = new CellRangeAddress(merge[0], merge[1], merge[2], merge[3]);
        this.sheet.addMergedRegion(cellRangeAddress);
      }
    }
  }
  // border style
  /**
   * add border merge cell to Lists
   * @param fRow Integer : first row index
   * @param lRow Integer : last row index
   * @param fCol Integer : first column index
   * @param lCol Integer : last column index
   */
  public void setBorderMergeCell(int fRow, int lRow, int fCol, int lCol) {
    Integer i[] = new Integer[4];
    i[0] = fRow;
    i[1] = lRow;
    i[2] = fCol;
    i[3] = lCol;
    this.borderMergeLists.add(i);
  }

  /**
   * process draw border merge cell
   */
  public void lenderBorderMergeCell() {
    this.lenderBorderMergeCell(CellStyle.BORDER_THIN);
  }

  /**
   * process draw border merge cell
   * @param border CellStyle
   * Example use CellStyle.BORDER_THIN
   */
  public void lenderBorderMergeCell(int border) {
    // add merge
    if(this.borderMergeLists != null && !this.borderMergeLists.isEmpty()) {
      CellRangeAddress cellRangeAddress;
      for (Integer[] merge : this.borderMergeLists) {
        cellRangeAddress = new CellRangeAddress(merge[0], merge[1], merge[2], merge[3]);
        RegionUtil.setBorderTop(border, cellRangeAddress, sheet, workbook);
        RegionUtil.setBorderBottom(border, cellRangeAddress, sheet, workbook);
        RegionUtil.setBorderLeft(border, cellRangeAddress, sheet, workbook);
        RegionUtil.setBorderRight(border, cellRangeAddress, sheet, workbook);
      }
    }
  }
  // ------------------------------- merge cell -------------------------------

  // ------------------------------- width column -------------------------------

  /**
   * auto Size Columns(Cell) : auto width Text
   * @param column Array Integer
   */
  public void autoSizeColumns(int column[]) {
    if(column != null && column.length > 0) {
      for (int cell : column) {
        this.sheet.autoSizeColumn(cell);
      }
    }
  }

  /**
   * Map<Integer, Integer> map = new HashMap<Integer, Integer>();
   * map.put(1, 8000);
   * @param map
   */
  public void setColumnsWidth(Map<Integer, Integer> map) {
    //loop a Map
    for (Map.Entry<Integer, Integer> entry : map.entrySet()) {
      this.sheet.setColumnWidth(entry.getKey(), entry.getValue());
    }
  }
  // ---------------------------- end width column -----------------------------

  /**
   * clear merge array list
  */
  public void clearMergeLists() {
    if(this.mergeLists != null) {
      this.mergeLists.clear();
    } else {
      this.mergeLists = new ArrayList<>();
    }
  }

  /**
   * clear border merge array list
   */
  public void clearBorderMergeLists() {
    if(this.borderMergeLists != null) {
      this.borderMergeLists.clear();
    } else {
      this.borderMergeLists = new ArrayList<>();
    }
  }

  // ------------------ Report Excel -----------------
  /**
   * report Excel Banner
   * @param rowId Integer : row index.
   * @param left Object : Data Show Left.
   * @param center Object : Data Show Center
   * @param right Object : Data Show Right
   * @param lCol Integer : first index of cell show data left
   * @param cCol Integer : first index of cell show data center
   * @param rCol Integer : first index of cell show data right
   * @param lastCol Integer : last index of cell show data right
   * @param heightInPoints Integer : Height in Row (line)
   */
  public void createRowBannerReport(int rowId, Object left, Object center, Object right, int lCol, int cCol, int rCol, int lastCol, int heightInPoints) {
    ExRow row = this.createRow(rowId);
    // left
    row.createCell(lCol)
        .setStyleVerticalTop()
        .setStyleHorizontalLeft()
        .setFontBold()
        .setValue(left);

    // center
    row.createCell(cCol)
        .setStyleVerticalTop()
        .setStyleHorizontalCenter()
        .setFontBold()
        .setValue(center);

    // right
    row.createCell(rCol)
        .setStyleVerticalTop()
        .setStyleHorizontalRight()
        .setFontBold()
        .setValue(right);

    // increase row height to accomodate two lines of text
    row.setHeightInPoints(heightInPoints);

    // add merge
    this.setMergeCell(rowId, rowId, lCol, cCol - 1);
    this.setMergeCell(rowId, rowId, cCol, rCol - 1);
    this.setMergeCell(rowId, rowId, rCol, lastCol);
  }

  /**
   * report Excel Footer
   * @param rowId Integer : row index.
   * @param left Object : Data Show Left.
   * @param center Object : Data Show Center
   * @param right Object : Data Show Right
   * @param lCol Integer : first index of cell show data left
   * @param cCol Integer : first index of cell show data center
   * @param rCol Integer : first index of cell show data right
   * @param lastCol Integer : last index of cell show data right
   * @param heightInPoints Integer : Height in Row (line)
   */
  public void createRowFooterReport(int rowId, Object left, Object center, Object right, int lCol, int cCol, int rCol, int lastCol, int heightInPoints) {
    ExRow row = this.createRow(rowId);
    // left
    row.createCell(lCol)
        .setStyleVerticalBottom()
        .setStyleHorizontalLeft()
        .setFontBold()
        .setValue(left);

    // center
    row.createCell(cCol)
        .setStyleVerticalBottom()
        .setStyleHorizontalCenter()
        .setFontBold()
        .setValue(center);

    // right
    row.createCell(rCol)
        .setStyleVerticalBottom()
        .setStyleHorizontalRight()
        .setFontBold()
        .setValue(right);

    // increase row height to accomodate two lines of text
    row.setHeightInPoints(heightInPoints);

    // add merge
    this.setMergeCell(rowId, rowId, lCol, cCol - 1);
    this.setMergeCell(rowId, rowId, cCol, rCol - 1);
    this.setMergeCell(rowId, rowId, rCol, lastCol);
  }

  /**
   * report Excel Header : row 1 don't merge in row or column
   * @param rowId Integer : row index
   * @param lists List<String> : Data Header
   */
  public void createRowHeaderReport(int rowId, List<String> lists){
    if(lists != null && !lists.isEmpty()) {
      ExRow row = this.createRow(rowId);
      for (String name : lists) {
        row.createCell()
            .setFontBold()
            .setBorderAll()
            .setValue(name);
      }
    }
  }
}
