import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.*;

public class ExportMain {
    static File[] source_files = new File("C:/Users/12687/OneDrive - Atheneum College Hageveld/PWS/Appendix/" +
            "Spectra").listFiles();
    static String outputFileName = "Spectra Graphs.xlsx";
    static File output_file = new File("C:/Users/12687/OneDrive - Atheneum College Hageveld/PWS/Appendix/" +
            "Spectra/" + outputFileName);

    static XSSFWorkbook workbook = new XSSFWorkbook();
    static Scanner requestScanner = new Scanner(System.in);


    public static void main(String[] args) throws Exception {
        List<File[]> directories = new ArrayList<>();
        List<String> directory_names = new ArrayList<>();
        for (File source_file : source_files) {
            if (source_file.isDirectory()) {
                directories.add(source_file.listFiles());
                directory_names.add(source_file.getName());
            }
        }


        System.out.println("""
                            Which sheet(s) would you like to update?
                               (0) All   (1) New day   (a-z) Specific""");
        String response = requestScanner.nextLine();
        switch (response) {
            case "0" -> { // all
                for (File[] data_files : directories) {
                    String directoryName = directory_names.get(directories.indexOf(data_files));
                    ExcelSheet sheet = new ExcelSheet(workbook, data_files, directoryName);
                    writeToSheet(sheet);
                }
            }
            case "1" -> {
                File[] data_files = directories.getLast();
                String directoryName = directory_names.getLast();
                ExcelSheet sheet = new ExcelSheet(workbook, data_files, directoryName);
                writeToSheet(sheet);
            }
            default -> {
                int sheetIndex = stringToIndex(response, directories.size() - 1);

                File[] data_files = directories.get(sheetIndex);
                String directoryName = directory_names.get(sheetIndex);
                ExcelSheet sheet = new ExcelSheet(workbook, data_files, directoryName);
                writeToSheet(sheet);

            }
        }

        if (output_file.exists()) {
            System.out.println("\nThe program has exported the data to the file '" + outputFileName + "' in 'PWS/Appendix/Spectra'.");
        } else {
            System.out.println("\nThe program will generate a new file with the name '" + outputFileName + "' in 'PWS/Appendix/Spectra'.");
        }

        try (FileOutputStream out = new FileOutputStream(output_file)) {
            workbook.write(out);
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

    static void writeToSheet(ExcelSheet sheet) {
        sheet.writeHeaders();
        sheet.writeChannelEnergy();
        sheet.writeData();
        sheet.sizeColumns();
    }




}
