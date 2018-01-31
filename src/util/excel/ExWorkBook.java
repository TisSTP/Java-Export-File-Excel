package util.excel;

import java.io.ByteArrayOutputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

public class ExWorkBook {
  private Workbook workbook;
  private ByteArrayOutputStream outStream;

  public ExWorkBook() {
    this(new SXSSFWorkbook(), new ByteArrayOutputStream());
  }

  public ExWorkBook(Workbook workbook, ByteArrayOutputStream outStream) {
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
  }

  public Workbook getWorkbook() {
    return workbook;
  }

  public void setWorkbook(Workbook workbook) {
    this.workbook = workbook;
  }

  public ExSheet createSheet(String name) {
    return new ExSheet(workbook, name);
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
}
