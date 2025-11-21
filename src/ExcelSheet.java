import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xssf.usermodel.*;

import java.io.File;
import java.io.FileNotFoundException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Scanner;

public class ExcelSheet {
    static int dataStartRowIndex = 4;

    private XSSFSheet sheet;
    private ArrayList<Spectrum> spectra;
    private ArrayList<Double> baseCalibration;

    public ExcelSheet(XSSFWorkbook workbook, File[] dataFiles, String sheetName) throws FileNotFoundException {
        try {
            sheet = workbook.createSheet("Data of " + sheetName);
            System.out.println("Creating new sheet...");
        } catch (IllegalArgumentException e) {
            sheet = workbook.getSheet("Data of " + sheetName);
            System.out.println("Replacing sheet...");
        }

        this.spectra = new ArrayList<>();
        int fileNo = 0;
        for (File dataFile : dataFiles) {
            if (dataFile.getName().contains(".spe")) { // get correct files
                spectra.add(new Spectrum(dataFile, fileNo));

                if (spectra.size() == 1) {
                    baseCalibration = spectra.getFirst().getCalibration();
                } else if (spectra.size() > 1) {
                    spectra.get(fileNo).checkCalibration(baseCalibration);
                }

                fileNo++;
            }
        }
    }

    public void writeHeaders() {
        XSSFRow headerRowDosage1 = ExportMain.getOrElseCreateRow(sheet, 1);
        XSSFRow headerRowDosage2 = ExportMain.getOrElseCreateRow(sheet, 2);
        headerRowDosage1.createCell(1).setCellValue("Total Dosage (in nS/h)");
        headerRowDosage2.createCell(1).setCellValue("Detector Face Dosage (in nS/h)");

        XSSFRow headerRowData = ExportMain.getOrElseCreateRow(sheet, dataStartRowIndex - 1);
        headerRowData.createCell(0).setCellValue("Channel");
        headerRowData.createCell(1).setCellValue("Energy (in keV)");
    }

    public void writeChannelEnergy() {
        double startValue = baseCalibration.get(0);
        double stepValue = baseCalibration.get(1);
        for (int channelNo = 0; channelNo < Spectrum.numberOfChannels; channelNo++) {
            XSSFRow row = ExportMain.getOrElseCreateRow(sheet, channelNo + dataStartRowIndex);
            row.createCell(0).setCellValue(channelNo);
            row.createCell(1).setCellValue(startValue + stepValue * channelNo);
        }
    }

    public void writeData() {
        for (Spectrum spectrum: spectra) {
            spectrum.writeDataColumn(sheet);
            spectrum.createExcelChart(sheet);
        }
    }

    public void sizeColumns() {
        for (int i = 0; i < sheet.getRow(0).getLastCellNum(); i++) {
            sheet.autoSizeColumn(i);
        }
    }





}
