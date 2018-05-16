package app;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import java.io.*;
import java.util.*;
import javax.swing.JTextPane;


public class ExcelReader extends Thread{

    //public static final String SAMPLE_XLSX_FILE_PATH = "Data.xlsx";
    private String file_path;
    
    public ExcelReader(String file_path) {
        this.file_path = file_path;
    }

    public void run()
    {
        Main.updateText("Opening the file");
        try{
            FileInputStream currFile = new FileInputStream(file_path);
            Workbook workbook = WorkbookFactory.create(currFile);

            // Retrieving the number of sheets in the Workbook
            Main.updateText("Workbook has " + workbook.getNumberOfSheets() + " Sheets : ");

            /*
               =============================================================
               Iterating over all the sheets in the workbook (Multiple ways)
               =============================================================
            */
            Main.updateText("Retrieving Sheets");
            workbook.forEach(sheet -> {
                Main.updateText("=> " + sheet.getSheetName());
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

            Map<String, Integer> firstColumn = new LinkedHashMap<>();

            Main.updateText("Reading first column");
            // Evaluate the first column
            sheet.forEach(row -> {
                String number = dataFormatter.formatCellValue(row.getCell(0));
                try {
                    Integer.parseInt(number);
                }catch (Exception e){return;}

                Cell currCell = row.getCell(1);
                if (currCell == null || currCell.getCellType() == Cell.CELL_TYPE_BLANK)
                    return;

                currCell.setCellType(Cell.CELL_TYPE_STRING);

                String cellValue = dataFormatter.formatCellValue(currCell).trim();
                if(cellValue.isEmpty())
                    return;

                System.out.println("first column > "+cellValue);

                if(firstColumn.containsKey(cellValue))
                    firstColumn.replace(cellValue, firstColumn.get(cellValue)+1);
                else
                    firstColumn.put(cellValue, 1);
            });
            Main.updateText("Done");
            Main.updateText("**********************");

            Main.updateText("Reading second column");

            Map<String, Integer> thirdColumn = new LinkedHashMap<>();
            // Evaluate the second column
            sheet.forEach(row -> {
                String number = dataFormatter.formatCellValue(row.getCell(0));
                try {
                    Integer.parseInt(number);
                }catch (Exception e){return;}

                Cell currCell = row.getCell(2);

                if (currCell == null || currCell.getCellType() == Cell.CELL_TYPE_BLANK)
                    return;

                currCell.setCellType(Cell.CELL_TYPE_STRING);

                String cellValue = dataFormatter.formatCellValue(currCell).trim();
                if(cellValue.isEmpty())
                    return;
                System.out.println("second column > "+cellValue);

                if(firstColumn.containsKey(cellValue))
                    thirdColumn.put(cellValue, firstColumn.get(cellValue));

            });
            Main.updateText("Done");
            Main.updateText("**********************");

            Main.updateText("Generating new data");
            List<String> newList = new LinkedList<>();
            thirdColumn.forEach((k,v)->{
                System.out.println("New list > "+k);
                for (int i = 0; i < v; i++) {
                    newList.add(k);
                }
            });
            Main.updateText("Done");
            Main.updateText("**********************");

            Main.updateText("Writing third column > "+newList.size()+" rows");
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
//                    int last = newList.size() - 1;
                    String cellphone = newList.remove(0);
                    System.out.println(cellphone);
                    row.createCell(3).setCellValue(cellphone);

                });
            }catch (NullPointerException e){
                e.printStackTrace();
            }

            Main.updateText("Done");
            Main.updateText("**********************");

            // Closing the workbook
            currFile.close();

            workbook.write(new FileOutputStream(file_path));
            workbook.close();

            Main.updateText("Finished Process");
        }catch (Exception e){
            e.printStackTrace();
            Main.updateText(e.getMessage());
        }

        Main.busy = false;

    }
}