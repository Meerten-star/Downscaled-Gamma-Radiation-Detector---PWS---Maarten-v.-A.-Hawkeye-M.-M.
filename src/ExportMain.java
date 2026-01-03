import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.*;

public class ExportMain {
    static List<String> yes_types = new ArrayList<>(Arrays.asList("yes", "ja", "yea", "yeah", "Absolutely!", "ye", "y",  "Undoubtedly", "da", "si", "oui"));

    static File[] source_files = new File("C:/Users/12687/OneDrive - Atheneum College Hageveld/PWS/Appendix/" +
            "Spectra").listFiles();
    static String outputFileName = "Spectra Graphs.xlsx";
    static File output_file = new File("C:/Users/12687/OneDrive - Atheneum College Hageveld/PWS/Appendix/" +
            "Spectra/" + outputFileName);

    static XSSFWorkbook workbook;
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

        try {
            workbook = new XSSFWorkbook(new FileInputStream(output_file));
        } catch (FileNotFoundException e) {
            workbook = new XSSFWorkbook();
        }


        System.out.println("""
                            Which sheet(s) would you like to update?
                               (0) All   (1) New day   (a-z) Specific   (2) Age determination""");
        String response = requestScanner.nextLine();
        switch (response) {
            case "0" -> { // all
                for (File[] data_files : directories) {
                    String directoryName = directory_names.get(directories.indexOf(data_files));
                    XSSFSheet sheetCheck = workbook.getSheet("Data of " + directoryName);
                    boolean writeSheetBoolean;
                    if (sheetCheck != null) {
                        writeSheetBoolean = requestConfirmation(directoryName);
                    } else {
                        writeSheetBoolean = true;
                    }
                    if (writeSheetBoolean) {
                        ExcelSheet sheet = new ExcelSheet(workbook, data_files, directoryName);
                        writeToSheet(sheet);
                    }
                }
            }
            case "1" -> {
                File[] data_files = directories.getLast();
                String directoryName = directory_names.getLast();
                ExcelSheet sheet = new ExcelSheet(workbook, data_files, directoryName);
                writeToSheet(sheet);
            }
            case "2" -> {
                ;
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


    static XSSFRow getOrElseCreateRow(XSSFSheet sheet, int rownum) {
        XSSFRow row = sheet.getRow(rownum);
        if (row == null) {
            row = sheet.createRow(rownum);
        }
        return row;
    }

    static boolean requestConfirmation(String sheetName) {
        System.out.println("The program will replace the sheet with the name: Data of " + sheetName);
        System.out.println("Are you sure you want to proceed?");
        Scanner requestScanner = new Scanner(System.in);
        String response = requestScanner.nextLine();
        if (yes_types.contains(response)) {
            return true;
        } else {
            System.out.println("The sheet will not be replaced.");
            return false;
        }

    }

}
