import java.util.*;
import java.io.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xssf.usermodel.*;


public class Spectra_OLD_VERSION {
    // fill in the date you want to export the files off.
    static boolean updateAll; // else update latest

    static File[] source_files = new File("C:/Users/12687/OneDrive - Atheneum College Hageveld/PWS/Appendix/" +
            "Spectra").listFiles();
    static File output_file = new File("C:/Users/12687/OneDrive - Atheneum College Hageveld/PWS/Appendix/" +
            "Spectra/Spectra Graphs - OLD VERSION.xlsx");

    static XSSFWorkbook workbook;


    // sizes of created column charts
    static int chartWidth = 8;
    static int chartHeight = 20;
    static int chartGap = 1;
    static List<String> yes_types = new ArrayList<>(Arrays.asList("yes", "ja", "yea", "yeah", "Absolutely!", "ye", "y",  "Undoubtedly", "da", "si", "oui"));

    public static void main(String[] args) throws IOException {
        List<File[]> directories = new ArrayList<>();
        List<String> directory_names = new ArrayList<>();
        for (File source_file : source_files) {
            if (source_file.isDirectory()) {
                directories.add(source_file.listFiles());
                directory_names.add(source_file.getName());
            }
        }

        Scanner request = new Scanner(System.in);
        System.out.println("Do you want to update all (1) or only update the latest (0) ?");
        updateAll = request.nextLine().equals("1");
        request.close();

        if (updateAll) {
            workbook = new XSSFWorkbook();
            for (File[] data_files : directories) {
                String directoryName = directory_names.get(directories.indexOf(data_files));
                createSheetMain(data_files, directoryName);
            }
        } else {
            try {
                workbook = new XSSFWorkbook(new FileInputStream(output_file));
            } catch (FileNotFoundException e) {
                System.out.println("Update all first");
            }
            File[] data_files = directories.getLast();
            String directoryName = directory_names.getLast();
            createSheetMain(data_files, directoryName);
        }
        workbook.close();

    }

    static void createSheetMain(File[] data_files, String sheetName) throws IOException {
        XSSFSheet sheet;
        List<List<Integer>> dataLists = new ArrayList<>();
        List<String> fileNames = new ArrayList<>();
        List<List<Double>> dataInfoList = new ArrayList<>();
        List<Double> dosageList = new ArrayList<>();
        for (File data_file : data_files) {
            // get correct files
            if (!data_file.getName().contains(".spe")) {
                continue;
            }
            // get all data out of each file
            dataLists.add(getDataList(data_file));
            fileNames.add(data_file.getName());
            dataInfoList.add(getDataInfo(data_file));
            dosageList.add(getDosage(dataLists.getLast(), dataInfoList.getFirst()));
        }
        try {
            sheet = workbook.createSheet("Data of " + sheetName);
            System.out.println("Created new sheet");
        } catch (IllegalArgumentException e) {
            System.out.println("The program will replace the latest sheet." +
                    "\nAre you sure you want to proceed?");
            Scanner inputScanner = new Scanner(System.in);
            if (!yes_types.contains(inputScanner.nextLine())) {
                System.exit(1);
            }
            sheet = workbook.getSheet("Data of " + sheetName);
            System.out.println("Replacing existing sheet");
        }
        createDataTables(sheet, dataLists, fileNames, dataInfoList, dosageList);

        createExcelCharts(sheet, dataLists, fileNames, dataInfoList);

        FileOutputStream out = new FileOutputStream(output_file);
        workbook.write(out);
        out.close();


        checkCalibration(dataInfoList);
    }

    static List<Integer> getDataList(File file) throws FileNotFoundException {
        // skips to where the data begins
        Scanner file_scanner = new Scanner(file);
        while (file_scanner.hasNextLine()) {
            String line = file_scanner.nextLine();
            if (line.contains("$DATA:")) {
                file_scanner.nextLine();
                break;
            }
        }
        // copying the data row by row
        List<Integer> dataList = new ArrayList<>();
        while (file_scanner.hasNextLine()) {
            String line = file_scanner.nextLine();
            line = line.trim();

            String[] dataPoints = line.split("\\s+");
            for (String point : dataPoints) {
                if (!point.isEmpty()) {
                    dataList.add(Integer.parseInt(point));
                }
            }
        }
        file_scanner.close();
        return dataList;
    }

    static ArrayList<Double> getDataInfo(File file) throws FileNotFoundException {
        Scanner file_scanner = new Scanner(file);
        ArrayList<Double> dataInfo = new ArrayList<>();
        while (file_scanner.hasNextLine()) {
            String line = file_scanner.nextLine();
            if (line.contains("$MEAS_TIM:") || line.contains("$SPEC_CAL:")) {
                String[] dataLine = file_scanner.nextLine().trim().split("\\s+");
                dataInfo.add(Double.parseDouble(dataLine[0]));
                dataInfo.add(Double.parseDouble(dataLine[1]));
            } else if (line.contains("$DATA:")) {
                break;
            }
        }
        file_scanner.close();
        return dataInfo; //{liveTime, realTime, startValue, stepValue};
    }

    static double getDosage(List<Integer> countsList, List<Double> infoList) {
        double energyKey;
        double x;
        double totalDose = 0;
        for (int i = 0; i < countsList.size(); i++) {
            energyKey = infoList.get(2) + infoList.get(3) * (i - 1);
            if (energyKey < 400) {
                x = 0.0013 * energyKey - 0.01;
            } else {
                x = 0.0019 * energyKey - 0.143;
            }
            totalDose += x * countsList.get(i);
        }
        return totalDose;
    }

    static void createDataTables(XSSFSheet sheet, List<List<Integer>> dataLists, List<String> fileNames, List<List<Double>> dataInfoList, List<Double> dosageList) throws IndexOutOfBoundsException {
        XSSFRow row;

        // Headers
        row = sheet.createRow(0);
        int columnIndex = 0;
        Cell header = row.createCell(columnIndex++);
        header.setCellValue("Channel");
        header = row.createCell(columnIndex++);
        header.setCellValue("Energy (in keV)");
        for (String name : fileNames) {
            header = row.createCell(columnIndex++);
            header.setCellValue(name);
        }

        // Data
        for (int rowIndex = 1; rowIndex < dataLists.getFirst().size() + 1; rowIndex++) {
            columnIndex = 0;
            row = sheet.createRow(rowIndex);
            // Channels
            Cell xCell = row.createCell(columnIndex++);
            xCell.setCellValue(rowIndex - 1);
            // Energy
            xCell = row.createCell(columnIndex++);
            xCell.setCellValue(dataInfoList.getFirst().get(2) + dataInfoList.getFirst().get(3) * (rowIndex - 1));
            // Data
            for (List<Integer> dataList : dataLists) {
                Cell dataCell = row.createCell(columnIndex++);
                dataCell.setCellValue(dataList.get(rowIndex - 1));
            }
        }
        for (columnIndex = 0; columnIndex < dosageList.size(); columnIndex++) {
            row = sheet.createRow(1);
            Cell cell = row.createCell(columnIndex + 1 + dataLists.size());
            cell.setCellValue(dosageList.get(columnIndex));
        }

    }

    static void createExcelCharts(XSSFSheet sheet, List<List<Integer>> dataLists, List<String> fileNames, List<List<Double>> dataInfoList) { // creates accompanying charts of the data (energy in x-axis)
        for (int i = 0; i < dataLists.size(); i++) {
            //set location of chart
            XSSFChart chart = sheet.createDrawingPatriarch().createChart(new XSSFClientAnchor(0, 0, 0, 0, 2 + (chartWidth + chartGap) * i, 2,  2 + (chartWidth + chartGap) * i + chartWidth, 2 + chartHeight));
            // chart settings
            chart.setTitleText("bMCA Spectrum: " + fileNames.get(i) + " (live/real = " + dataInfoList.get(i).get(0) + "/" + dataInfoList.get(i).get(1) + "s)");
            chart.setTitleOverlay(false);
            XDDFValueAxis xAxis = chart.createValueAxis(AxisPosition.BOTTOM);
            xAxis.setTitle("Energy (in keV)");
            xAxis.setCrosses(AxisCrosses.AUTO_ZERO);
            xAxis.setMinimum(dataInfoList.get(i).get(2));
            xAxis.setMaximum(dataInfoList.get(i).get(2) + 2048 * dataInfoList.get(i).get(3));
            xAxis.setMajorUnit(250);
            XDDFValueAxis yAxis = chart.createValueAxis(AxisPosition.LEFT);
            yAxis.setTitle("Counts");

            // cell ranges
            XDDFDataSource<Double> xValues = XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(1, 2049, 1, 1));
            XDDFNumericalDataSource<Double> yValues = XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(1, 2049, i + 2, i + 2));
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
    }

    static void printData(List<Integer> dataList) {
        // Prints the size of the data and the first 256 values.
        System.out.println("Total data points: " + dataList.size());
        for (int i = 0; i < 256; i++) {
            System.out.println("Channel " + i + ": " + dataList.get(i));
        }
    }

    static void checkCalibration(List<List<Double>> dataInfoList) { // check if the calibration is the same for each measurement (has to be last operation)
        List<Double> lastCalibration = new ArrayList<>();
        for (List<Double> calibration : dataInfoList) {
            calibration.removeFirst();
            calibration.removeFirst();
            if (!lastCalibration.equals(calibration) && !lastCalibration.isEmpty()) {
                System.out.println("!WARNING!");
                System.out.println("ERROR: Calibrations not aligned!");
                System.out.println("" + lastCalibration + calibration);
                System.out.println("!WARNING!");
            }
            lastCalibration = new ArrayList<>(calibration);
        }
    }
}
