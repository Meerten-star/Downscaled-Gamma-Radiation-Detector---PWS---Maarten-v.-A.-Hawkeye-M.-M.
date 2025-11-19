import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xssf.usermodel.*;

import java.io.*;
import java.util.*;


public class Spectra {
    static File[] source_files = new File("C:/Users/12687/OneDrive - Atheneum College Hageveld/PWS/Appendix/" +
            "Spectra").listFiles();
    static File output_file = new File("C:/Users/12687/OneDrive - Atheneum College Hageveld/PWS/Appendix/" +
            "Spectra/Spectra Graphs.xlsx");

    static XSSFWorkbook workbook;
    static Scanner requestScanner = new Scanner(System.in);

    static int dataStartRowIndex = 4;
    // sizes of created column charts
    static int chartWidth = 8;
    static int chartHeight = 20;
    static int chartGap = 1;
    static int chartStartX = 6;
    // detector constants
    static double detectorEfficiency = 0.10;
    static double detectorDiameter = 5; // in cm
    static double detectorFaceArea = Math.PI * Math.pow(detectorDiameter / 2, 2); // in cm2
    static double detectorDistanceToSource = 5; // in cm
    static double totalDosageSurfaceArea = 4 * Math.PI * Math.pow(detectorDistanceToSource, 2);

    static List<String> yes_types = new ArrayList<>(Arrays.asList("yes", "ja", "yea", "yeah", "Absolutely!", "ye", "y",  "Undoubtedly", "da", "si", "oui"));

    public static void main(String[] args) throws Exception {
        List<File[]> directories = new ArrayList<>();
        List<String> directory_names = new ArrayList<>();
        for (File source_file : source_files) {
            if (source_file.isDirectory()) {
                directories.add(source_file.listFiles());
                directory_names.add(source_file.getName());
            }
        }

        if (output_file.exists()) {
            workbook = new XSSFWorkbook(new FileInputStream(output_file));
        } else {
            workbook = new XSSFWorkbook();
        }

        System.out.println("""
                            Which sheet(s) would you like to update?
                               (0) All   (1) New day   (a-z) Specific""");
        String response = requestScanner.nextLine();
        switch (response) {
            case "0" -> { // all
                for (File[] data_files : directories) {
                    String directoryName = directory_names.get(directories.indexOf(data_files));
                    createSheetMain(data_files, directoryName);
                }
            }
            case "1" -> {
                File[] data_files = directories.getLast();
                String directoryName = directory_names.getLast();
                createSheetMain(data_files, directoryName);
            }
            default -> {
                int sheetIndex = stringToIndex(response, directories.size() - 1);

                File[] data_files = directories.get(sheetIndex);
                String directoryName = directory_names.get(sheetIndex);
                createSheetMain(data_files, directoryName);

            }
        }

        workbook.close();

    }

    static int stringToIndex(String s, int maxIndex) throws Exception {
        int charInt = s.charAt(0);
        if (charInt > 96) { // Letter (lowercase)
            charInt -= 97;
        } else if (charInt > 64) { // Letter (uppercase)
            charInt -= 65;
        } else { // Space (question isn't answered)
            throw new Exception();
        }
        if (charInt > maxIndex) {
            charInt = maxIndex;
        }
        return charInt;
    }

    static void createSheetMain(File[] dataFiles, String sheetName) throws IOException {
        XSSFSheet sheet;
        try {
            sheet = workbook.createSheet("Data of " + sheetName);
            System.out.println("Created new sheet");
        } catch (IllegalArgumentException e) {
            System.out.println("The program will replace the sheet with the name: Data of " + sheetName);
            System.out.println("Are you sure you want to proceed?");
            String response = requestScanner.nextLine();
            if (!yes_types.contains(response)) {
                System.out.println("The sheet will not be replaced.");
                return;
            }
            sheet = workbook.getSheet("Data of " + sheetName);
            System.out.println("Replacing existing sheet...");
        }

        // headers
        XSSFRow headerRowDosage1 = getOrElseCreateRow(sheet, 1);
        XSSFRow headerRowDosage2 = getOrElseCreateRow(sheet, 2);
        headerRowDosage1.createCell(1).setCellValue("Total Dosage (in nS/h)");
        headerRowDosage2.createCell(1).setCellValue("Detector Face Dosage (in nS/h)");

        XSSFRow headerRowData = getOrElseCreateRow(sheet, dataStartRowIndex - 1);
        headerRowData.createCell(0).setCellValue("Channel");
        headerRowData.createCell(1).setCellValue("Energy (in keV)");

        // data
        int fileNo = -1;
        for (File dataFile : dataFiles) {
            // get correct files
            String fileName = dataFile.getName();
            if (!fileName.contains(".spe")) {
                continue;
            }
            fileNo++;

            // get all data out of each file
            Scanner fileScanner = new Scanner(dataFile);
            ArrayList<Double> infoList = getDataInfo(fileScanner);
            ArrayList<Integer> dataList = getDataList(fileScanner);
            fileScanner.close();
            // calculate dosage
            double[] dosage = getDosage(dataList, infoList, fileName.toLowerCase());

            if (fileNo == 0) { // write channel and energy columns (only once)
                writeChannelEnergy(sheet, dataList.size(), infoList);
            }

            // output the data in the table & append dosage
            writeDataColumn(sheet, fileNo, fileName, dosage, dataList);
            // create chart with data
            createExcelChart(sheet, fileNo, fileName, infoList);

            checkCalibration(sheet, fileNo, infoList);
        }

        for (int i = 0; i < sheet.getRow(0).getLastCellNum(); i++) {
            sheet.autoSizeColumn(i);
        }

        try (FileOutputStream out = new FileOutputStream(output_file)) {
            workbook.write(out);
        }
    }

    static ArrayList<Double> getDataInfo(Scanner fileScanner) {
        ArrayList<Double> dataInfo = new ArrayList<>();
        while (fileScanner.hasNextLine()) {
            String line = fileScanner.nextLine();
            if (line.contains("$MEAS_TIM:") || line.contains("$SPEC_CAL:")) {
                String[] dataLine = fileScanner.nextLine().trim().split("\\s+");
                dataInfo.add(Double.parseDouble(dataLine[0]));
                dataInfo.add(Double.parseDouble(dataLine[1]));
            } else if (line.contains("$DATA:")) {
                fileScanner.nextLine();
                break;
            }
        }
        return dataInfo; //{liveTime, realTime, startValue, stepValue};
    }

    static ArrayList<Integer> getDataList(Scanner fileScanner) {
        // copying the data row by row
        ArrayList<Integer> dataList = new ArrayList<>();
        while (fileScanner.hasNextLine()) {
            String line = fileScanner.nextLine();
            line = line.trim();

            String[] dataPoints = line.split("\\s+");
            for (String point : dataPoints) {
                if (!point.isEmpty()) {
                    dataList.add(Integer.parseInt(point));
                }
            }
        }
        fileScanner.close();
        return dataList;
    }

    static double[] getDosage(List<Integer> countsList, List<Double> infoList, String fileNameLowerCase) {
        double energyKey;
        double x;
        double totalSphereDose = 0;
        double detectorFaceDose = 0;
        for (int i = 0; i < countsList.size(); i++) {
            energyKey = infoList.get(2) + infoList.get(3) * (i - 1);
            if (energyKey < 400) {
                x = 0.0013 * energyKey - 0.01;
            } else {
                x = 0.0019 * energyKey - 0.143;
            }
            double cps = countsList.get(i) / (infoList.getFirst() * detectorEfficiency);
            if (x > 0) {
                double dosagePoint = cps * x; // cps * nSv/h/cps
                totalSphereDose += dosagePoint * totalDosageSurfaceArea / detectorFaceArea;
                detectorFaceDose += dosagePoint;
            }
        }
        double[] dosages = new double[2];
        if (fileNameLowerCase.contains("alpha") || fileNameLowerCase.contains("beta")) { // not gamma
            return dosages;
        } else if (fileNameLowerCase.contains("background")) { // not a point source
            totalSphereDose = detectorFaceDose;
        }
        dosages[0] = round(totalSphereDose, 3);
        dosages[1] = round(detectorFaceDose, 3);
        return dosages;
    }

    static void writeChannelEnergy(XSSFSheet sheet, int numberOfChannels, ArrayList<Double> infoList) {
        for (int channelNo = 0; channelNo < numberOfChannels; channelNo++) {
            XSSFRow row = getOrElseCreateRow(sheet, channelNo + dataStartRowIndex);
            row.createCell(0).setCellValue(channelNo);
            row.createCell(1).setCellValue(infoList.get(2) + infoList.get(3) * channelNo);
        }
    }

    static XSSFRow getOrElseCreateRow(XSSFSheet sheet, int rownum) {
        XSSFRow row = sheet.getRow(rownum);
        if (row == null) {
            row = sheet.createRow(rownum);
            System.out.println("Row " + rownum +  " created");
        }
        return row;
    }

    static void writeDataColumn(XSSFSheet sheet, int fileNo, String fileName, double[] dosages, ArrayList<Integer> dataList) throws IndexOutOfBoundsException {
        // Header
        XSSFRow row = getOrElseCreateRow(sheet, 0);
        Cell header = row.createCell(fileNo + 2);
        header.setCellValue(fileName);

        // Dosage
        for (int i = 0; i < dosages.length; i++) {
            row = sheet.getRow(i + 1);
            Cell dosageCell = row.createCell(fileNo + 2);
            if (dosages[i] != 0) {
                dosageCell.setCellValue(dosages[i]);
            } else { // beta or alpha spectrometry -> no dosage calculation
                dosageCell.setCellValue("None");
            }
        }

        // Data
        for (int rowIndex = dataStartRowIndex; rowIndex < dataList.size() + dataStartRowIndex; rowIndex++) {
            Cell dataCell = getOrElseCreateRow(sheet, rowIndex).createCell(fileNo + 2);
            dataCell.setCellValue(dataList.get(rowIndex - dataStartRowIndex));
        }
    }

    static void createExcelChart(XSSFSheet sheet, int fileNo, String fileName, ArrayList<Double> dataInfoList) {
        //set location of chart
        XSSFChart chart = sheet.createDrawingPatriarch().createChart(new XSSFClientAnchor(0, 0, 0, 0, 2 + (chartWidth + chartGap) * fileNo + chartStartX, 2,  2 + (chartWidth + chartGap) * fileNo + chartWidth + chartStartX, 2 + chartHeight));
        // chart settings
        chart.setTitleText("bMCA Spectrum: " + fileName + " (live/real = " + dataInfoList.get(0) + "/" + dataInfoList.get(1) + "s)");
        chart.setTitleOverlay(false);
        XDDFValueAxis xAxis = chart.createValueAxis(AxisPosition.BOTTOM);
        xAxis.setTitle("Energy (in keV)");
        xAxis.setCrosses(AxisCrosses.AUTO_ZERO);
        xAxis.setMinimum(dataInfoList.get(2));
        xAxis.setMaximum(dataInfoList.get(2) + 2048 * dataInfoList.get(3));
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

    static double round(double value, int decimals) {
        double format = Math.pow(10, decimals);
        return Math.round(value * format) / format;
    }

    static void checkCalibration(XSSFSheet sheet, int fileNo, ArrayList<Double> infoList) { // checks if the calibration is the same for each measurement
        Cell channel0 = sheet.getRow(dataStartRowIndex).getCell(1);
        Cell channel1 = sheet.getRow(dataStartRowIndex + 1).getCell(1);
        double baseStartValue = round(channel0.getNumericCellValue(), 6);
        double baseStepValue = round(channel1.getNumericCellValue() - channel0.getNumericCellValue(), 6);
        if (baseStartValue != infoList.get(2) || baseStepValue != infoList.get(3)) {
            System.out.println("!WARNING!");
            System.out.println("ERROR: Calibrations not aligned! file number 0 and " + fileNo + " do not contain the same calibration.");
            System.out.println("base={" + baseStartValue + ", " + baseStepValue + "}  !=  file" + fileNo + "={" + infoList.get(2) + ", " + infoList.get(3) + "}");
            System.out.println("!WARNING!");
        }
    }

}
