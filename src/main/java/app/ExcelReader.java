package app;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import java.io.*;
import java.net.URL;
import java.nio.charset.Charset;
import java.security.CodeSource;
import java.util.*;

public class ExcelReader {

    public static final String SAMPLE_XLSX_FILE_PATH = "Data.xlsx";

    public static void main(String[] args) throws IOException, InvalidFormatException {

        // Creating a Workbook from an Excel file (.xls or .xlsx)
        CodeSource src = ExcelReader.class.getProtectionDomain().getCodeSource();
        URL url = new URL(src.getLocation(), SAMPLE_XLSX_FILE_PATH);

        Workbook workbook = WorkbookFactory.create(new File(url.getPath()));

        // Retrieving the number of sheets in the Workbook
        System.out.println("Workbook has " + workbook.getNumberOfSheets() + " Sheets : ");

        /*
           =============================================================
           Iterating over all the sheets in the workbook (Multiple ways)
           =============================================================
        */

        System.out.println("Retrieving Sheets");
        workbook.forEach(sheet -> {
            System.out.println("=> " + sheet.getSheetName());
        });

        /*
           ==================================================================
           Iterating over all the rows and columns in a Sheet (Multiple ways)
           ==================================================================
        */

        // Getting the Sheet at index zero
        Sheet sheet = workbook.getSheetAt(0);

        // Create a DataFormatter to format and get each cell's value as String
        DataFormatter dataFormatter = new DataFormatter();

        Map<String, Integer> firstColumn = new HashMap<>();

        // Evaluate the first column
        sheet.forEach(row -> {
            String number = dataFormatter.formatCellValue(row.getCell(0));
            try {
                Integer.parseInt(number);
            }catch (Exception e){return;}

            String cellValue = dataFormatter.formatCellValue(row.getCell(1));

            if(firstColumn.containsKey(cellValue))
                firstColumn.put(cellValue, firstColumn.get(cellValue)+1);
            else
                firstColumn.put(cellValue, 1);
        });

        Map<String, Integer> thirdColumn = new HashMap<>();
        // Evaluate the second column
        sheet.forEach(row -> {
            String number = dataFormatter.formatCellValue(row.getCell(0));
            try {
                Integer.parseInt(number);
            }catch (Exception e){return;}

            String cellValue = dataFormatter.formatCellValue(row.getCell(2));
            if(firstColumn.containsKey(cellValue))
                thirdColumn.put(cellValue, firstColumn.get(cellValue));

        });

        List<String> newList = new ArrayList<>();

        thirdColumn.forEach((k,v)->{
            for (int i = 0; i < v; i++) {
                newList.add(k);
            }
        });

        System.out.println("new list size: "+newList.size());

        // Generate the third column
        try {
            sheet.forEach(row -> {
                String number = dataFormatter.formatCellValue(row.getCell(0));
                try {
                    Integer.parseInt(number);
                } catch (Exception e) {
                    return;
                }

                if (newList.isEmpty())
                    return;

                String cellphone = newList.remove(0);
//                System.out.println(cellphone);
                row.createCell(3).setCellValue(cellphone);

            });
        }catch (NullPointerException e){
            e.printStackTrace();
        }
        // Closing the workbook

        URL url2 = new URL(src.getLocation(), "output.xls");
        FileOutputStream fileOut = new FileOutputStream(url2.getPath());
        workbook.write(fileOut);
        fileOut.close();
        workbook.close();
    }
}