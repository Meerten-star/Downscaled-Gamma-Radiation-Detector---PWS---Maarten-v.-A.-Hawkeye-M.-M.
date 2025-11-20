import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.*;
import java.util.*;

public class BackgroundSubstractionSheet extends ExcelSheet {

    ExcelSheet sheet;

    public BackgroundSubstractionSheet(XSSFWorkbook workbook, File backgroundFile, File sourceFile) throws FileNotFoundException {
        super(workbook, new File[]{backgroundFile, sourceFile}, "BackgroundSubstraction");

    }

    public void scaleCalibration() {}

    public void writeSubstractionColumn() {}

    public int calculateAge() {
        return 0;
    }

    public void writeAge() {}




}
