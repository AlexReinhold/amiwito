package app;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import java.io.*;
import java.util.*;
import javax.swing.JTextArea;

        
public class ExcelReader {

    //public static final String SAMPLE_XLSX_FILE_PATH = "Data.xlsx";

    public void process(String file_path) throws IOException, InvalidFormatException {

        // Creating a Workbook from an Excel file (.xls or .xlsx)
        //CodeSource src = ExcelReader.class.getProtectionDomain().getCodeSource();
        //URL url = new URL(src.getLocation(), file_path);
        Main.output.setText("");
//        File currFile = new File(url.getPath());
        FileInputStream currFile = new FileInputStream(file_path);
       
        Workbook workbook = WorkbookFactory.create(currFile);

        // Retrieving the number of sheets in the Workbook
        Main.output.append("Workbook has " + workbook.getNumberOfSheets() + " Sheets : "+"\n");
        /*
           =============================================================
           Iterating over all the sheets in the workbook (Multiple ways)
           =============================================================
        */

        Main.output.append("Retrieving Sheets"+"\n");
        workbook.forEach(sheet -> {
            Main.output.append("=> " + sheet.getSheetName()+"\n");
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

        Main.output.append("Reading first column"+"\n");
        // Evaluate the first column
        sheet.forEach(row -> {
            String number = dataFormatter.formatCellValue(row.getCell(0));
            try {
                Integer.parseInt(number);
            }catch (Exception e){return;}

            String cellValue = dataFormatter.formatCellValue(row.getCell(1)).trim();
            if(cellValue.isEmpty())
                return;

            if(firstColumn.containsKey(cellValue))
                firstColumn.put(cellValue, firstColumn.get(cellValue)+1);
            else
                firstColumn.put(cellValue, 1);
        });
        Main.output.append("Done"+"\n");
        Main.output.append("**********************"+"\n");
        
        Main.output.append("Reading second column"+"\n");
        
        Map<String, Integer> thirdColumn = new HashMap<>();
        // Evaluate the second column
        sheet.forEach(row -> {
            String number = dataFormatter.formatCellValue(row.getCell(0));
            try {
                Integer.parseInt(number);
            }catch (Exception e){return;}

            String cellValue = dataFormatter.formatCellValue(row.getCell(2)).trim();
            if(cellValue.isEmpty())
                return;
            if(firstColumn.containsKey(cellValue))
                thirdColumn.put(cellValue, firstColumn.get(cellValue));

        });
        Main.output.append("Done"+"\n");
        Main.output.append("**********************"+"\n");
        
        Main.output.append("Generating new data"+"\n");
        List<String> newList = new ArrayList<>();
        thirdColumn.forEach((k,v)->{
            for (int i = 0; i < v; i++) {
                newList.add(k);
            }
        });
        Main.output.append("Done"+"\n");
        Main.output.append("**********************"+"\n");
        
        Main.output.append("Writing third column > "+newList.size()+"\n");
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
        
        Main.output.append("Done"+"\n");
        Main.output.append("**********************"+"\n");
        
        // Closing the workbook
        currFile.close();

        workbook.write(new FileOutputStream(file_path));
        workbook.close();

        Main.output.append("Finished Process"+"\n");
        
    }
}