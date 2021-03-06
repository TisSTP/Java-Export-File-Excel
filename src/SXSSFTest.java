import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.logging.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class SXSSFTest {
  public static void main(String[] args) throws IOException {
    long startTime = System.nanoTime();
//    FileOutputStream inMemoryOut = new FileOutputStream(new File("inMemoryWorkbook.xlsx"));
//    XSSFWorkbook workbook = new XSSFWorkbook();
//    WorkbookExample example = new WorkbookExample(workbook, inMemoryOut);
//    example.export();

    FileOutputStream streamOut = new FileOutputStream(new File("streamWorkbook.xlsx"));
    SXSSFWorkbook streamWorkbook = new SXSSFWorkbook();
    WorkbookExample streamExample = new WorkbookExample(streamWorkbook, streamOut);
    streamExample.export();
    System.out.println("[Process] Save and Commit all record in : "
        + ((System.nanoTime() - startTime) / 1000000000.0) + " Seconds");
  }
}

class WorkbookExample {

  private Logger logger = Logger.getLogger(WorkbookExample.class.getName());
  private Workbook workbook;
  private OutputStream out;

  public WorkbookExample(Workbook workbook, OutputStream out) {
    this.workbook = workbook;
    this.out = out;
  }

  public void export() throws IOException {
    logger.info("export start for " + workbook.getClass().getName());

    List<Person> persons = new ArrayList<Person>();
    for (int i = 0; i < 100_000; i++) {
      persons.add(new Person(String.valueOf("user_" + i)));
    }

    Sheet sheet = workbook.createSheet();
    for (int i = 0; i < persons.size(); i++) {
      Person p = persons.get(i);
      Row row = sheet.createRow(i);
      Cell cell = row.createCell(0);
      cell.setCellValue(p.getName());
    }
    workbook.write(out);
    logger.info("Is row 1 accessible after writing to output stream? " + String.valueOf(sheet.getRow(1) != null));
    out.close();
    workbook.close();

    logger.info("export finished for " + workbook.getClass().getName());
  }

  public static class Person {

    private String name;

    public Person(String name) {
      this.name = name;
    }

    public String getName() {
      return name;
    }

    public void setName(String name) {
      this.name = name;
    }

  }

}
