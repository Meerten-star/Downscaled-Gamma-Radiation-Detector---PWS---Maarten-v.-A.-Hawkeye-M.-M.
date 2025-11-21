import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.io.*;
import java.util.*;

public class Spectrum {
    // detector constants
    static double detectorEfficiency = 0.10;
    static double detectorDiameter = 5; // in cm
    static double detectorFaceArea = Math.PI * Math.pow(detectorDiameter / 2, 2); // in cm2
    static double detectorDistanceToSource = 5; // in cm
    static double totalDosageSurfaceArea = 4 * Math.PI * Math.pow(detectorDistanceToSource, 2);
    static int numberOfChannels = 2048;

    static int dataStartRowIndex = 4;
    // sizes of created column charts
    static int chartWidth = 8;
    static int chartHeight = 20;
    static int chartGap = 1;
    static int chartStartX = 6;

    // file info
    private final int fileNo;
    private final String fileName;

    // dataInfo
    private double liveTime;
    private double realTime;
    private double startValue;
    private double stepValue;

    private double totalSphereDose = 0;
    private double detectorFaceDose = 0;

    private final ArrayList<Integer> countsList = new ArrayList<>();

    public Spectrum(File file, int fileNo) throws FileNotFoundException {
        this.fileName = file.getName();
        this.fileNo = fileNo;
        Scanner fileScanner = new Scanner(file);
        readDataInfo(fileScanner);
        readDataList(fileScanner);
        fileScanner.close();

        if (!fileName.toLowerCase().contains("alpha") && !fileName.toLowerCase().contains("beta")) {
            calculateDosage();
        }
    }


    public void readDataInfo(Scanner fileScanner) {
        while (fileScanner.hasNextLine()) {
            String line = fileScanner.nextLine();
            if (line.contains("$MEAS_TIM:")) {
                String[] dataLine = fileScanner.nextLine().trim().split("\\s+");
                liveTime = Double.parseDouble(dataLine[0]);
                realTime = Double.parseDouble(dataLine[1]);
            } else if (line.contains("$SPEC_CAL:")) {
                String[] dataLine = fileScanner.nextLine().trim().split("\\s+");
                startValue = Double.parseDouble(dataLine[0]);
                stepValue = Double.parseDouble(dataLine[1]);
            } else if (line.contains("$DATA:")) {
                fileScanner.nextLine();
                break;
            }
        }
    }

    public void readDataList(Scanner fileScanner) { // copying the data row by row
        while (fileScanner.hasNextLine()) {
            String line = fileScanner.nextLine();
            line = line.trim();

            String[] dataPoints = line.split("\\s+");
            for (String point : dataPoints) {
                if (!point.isEmpty()) {
                    countsList.add(Integer.parseInt(point));
                }
            }
        }
    }

    private double round(double value, int decimals) {
        double format = Math.pow(10, decimals);
        return Math.round(value * format) / format;
    }

    public void calculateDosage() {
        double energyKey;
        double x;
        for (int i = 0; i < countsList.size(); i++) {
            energyKey = startValue + stepValue * (i - 1);
            if (energyKey < 400) {
                x = 0.0013 * energyKey - 0.01;
            } else {
                x = 0.0019 * energyKey - 0.143;
            }
            double cps = countsList.get(i) / (liveTime * detectorEfficiency);
            if (x > 0) {
                double dosagePoint = cps * x; // cps * nSv/h/cps
                totalSphereDose += dosagePoint * totalDosageSurfaceArea / detectorFaceArea;
                detectorFaceDose += dosagePoint;
            }
        }
        if (fileName.toLowerCase().contains("background")) { // not a point source
            totalSphereDose = detectorFaceDose;
        }
        totalSphereDose = round(totalSphereDose, 3);
        detectorFaceDose = round(detectorFaceDose, 3);
    }

    public void checkCalibration(ArrayList<Double> baseCalibration) { // checks if the calibration is the same for each measurement
        double baseStartValue = baseCalibration.get(0);
        double baseStepValue = baseCalibration.get(1);
        if (baseStartValue != startValue || baseStepValue != stepValue) {
            System.out.println("!WARNING!" +
                    "\nERROR: Calibrations not aligned! file number 0 and " + fileNo + " do not contain the same calibration." +
                    "\nbase={" + baseStartValue + ", " + baseStepValue + "}  !=  file" + fileNo + "={" + startValue + ", " + stepValue + "}" +
                    "\n!WARNING!");
        }
    }

    public void writeDataColumn(XSSFSheet sheet) throws IndexOutOfBoundsException {

        // Header
        XSSFRow row = ExportMain.getOrElseCreateRow(sheet, 0);
        Cell header = row.createCell(fileNo + 2);
        header.setCellValue(fileName);

        // Dosage
        XSSFRow row1 = sheet.getRow(1);
        XSSFRow row2 = sheet.getRow(2);
        Cell dosageCell1 = row1.createCell(fileNo + 2);
        Cell dosageCell2 = row2.createCell(fileNo + 2);
        if (totalSphereDose == 0) {
            dosageCell1.setCellValue("None");
            dosageCell2.setCellValue("None");
        } else { // beta or alpha spectrometry -> no dosage calculation
            dosageCell1.setCellValue(totalSphereDose);
            dosageCell2.setCellValue(detectorFaceDose);
        }


        // Data
        for (int rowIndex = dataStartRowIndex; rowIndex < countsList.size() + dataStartRowIndex; rowIndex++) {
            Cell dataCell = ExportMain.getOrElseCreateRow(sheet, rowIndex).createCell(fileNo + 2);
            dataCell.setCellValue(countsList.get(rowIndex - dataStartRowIndex));
        }
    }

    public void createExcelChart(XSSFSheet sheet) {
        //set location of chart
        XSSFChart chart = sheet.createDrawingPatriarch().createChart(new XSSFClientAnchor(0, 0, 0, 0, 2 + (chartWidth + chartGap) * fileNo + chartStartX, 2,  2 + (chartWidth + chartGap) * fileNo + chartWidth + chartStartX, 2 + chartHeight));
        // chart settings
        chart.setTitleText("bMCA Spectrum: " + fileName + " (live/real = " + liveTime + "/" + realTime + "s)");
        chart.setTitleOverlay(false);
        XDDFValueAxis xAxis = chart.createValueAxis(AxisPosition.BOTTOM);
        xAxis.setTitle("Energy (in keV)");
        xAxis.setCrosses(AxisCrosses.AUTO_ZERO);
        xAxis.setMinimum(startValue);
        xAxis.setMaximum(startValue + 2048 * stepValue);
        xAxis.setMajorUnit(250);
        XDDFValueAxis yAxis = chart.createValueAxis(AxisPosition.LEFT);
        yAxis.setTitle("Counts");

        // cell ranges
        XDDFDataSource<Double> xValues = XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(dataStartRowIndex, 2047 + dataStartRowIndex, 1, 1));
        XDDFNumericalDataSource<Double> yValues = XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(dataStartRowIndex, 2047 + dataStartRowIndex, fileNo + 2, fileNo + 2));
        // create the bar chart and link them with the cell ranges
        XDDFBarChartData columnChart = (XDDFBarChartData) chart.createData(ChartTypes.BAR, xAxis, yAxis);
        XDDFBarChartData.Series series = (XDDFBarChartData.Series) columnChart.addSeries(xValues, yValues);
        series.setTitle("Yippee!!"); // easter egg
        // adjust chart settings
        columnChart.setBarDirection(BarDirection.COL);
        columnChart.setVaryColors(false);

        // plot the chart
        chart.plot(columnChart);
    }

    public ArrayList<Double> getCalibration() {
        return new ArrayList<>(Arrays.asList(startValue, stepValue));
    }

    public int getUraniumCount() {
        return 0;
    }

    public int getBismuthCount() {
        return 0;
    }

}
