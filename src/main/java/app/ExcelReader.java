package app;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.omg.CORBA.DATA_CONVERSION;

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

            Map<Integer, String> firstColumn = new LinkedHashMap<>();

            Main.updateText("Reading first column");
            // Evaluate the first column
            sheet.forEach(row -> {
                String number = dataFormatter.formatCellValue(row.getCell(0));
                int index;
                try {
                    index = Integer.parseInt(number);
                }catch (Exception e){return;}

                Cell currCell = row.getCell(1);
                if (currCell == null || currCell.getCellType() == Cell.CELL_TYPE_BLANK)
                    return;

                currCell.setCellType(Cell.CELL_TYPE_STRING);

                String cellValue = dataFormatter.formatCellValue(currCell).trim();
                if(cellValue.isEmpty())
                    return;

                firstColumn.put(index, cellValue);

            });
            Main.updateText("Done");
            Main.updateText("**********************");

//            System.out.println("first column > "+ firstColumn.size());

            Main.updateText("Reading second column");

            Map<String,Integer> secondColumn = new LinkedHashMap<>();

            // Evaluate the second column
            List<Data> whatsapp = new LinkedList<>();
            List<Data> nowhatsapp = new LinkedList<>();

            sheet.forEach(row -> {
                String number = dataFormatter.formatCellValue(row.getCell(0));
                int index;
                try {
                    index = Integer.parseInt(number);
                }catch (Exception e){return;}

                Cell currCell = row.getCell(2);

                if (currCell == null || currCell.getCellType() == Cell.CELL_TYPE_BLANK)
                    return;

                currCell.setCellType(Cell.CELL_TYPE_STRING);

                String cellValue = dataFormatter.formatCellValue(currCell).trim();
                if(cellValue.isEmpty())
                    return;
//                System.out.println("second column > "+cellValue);

                secondColumn.put(cellValue,index);

            });
            Main.updateText("Done");
            Main.updateText("**********************");

            Main.updateText("Comparing data");

            firstColumn.forEach((k,v)->{
                if(secondColumn.containsKey(v))
                    whatsapp.add(new Data(k,v).setType(Data.WHATSAPP));
                else
                    nowhatsapp.add(new Data(k,v).setType(Data.NO_WHATSAPP));
            });
            Main.updateText("Done");
            Main.updateText("**********************");

//            Main.updateText("Writing third column > "+newList.size()+" rows");
            // Generate the third column
//            try {
//                sheet.forEach(row -> {
//                    String number = dataFormatter.formatCellValue(row.getCell(0));
//                    try {
//                        Integer.parseInt(number);
//                    } catch (Exception e) {
//                        return;
//                    }
//
//                    if (newList.isEmpty())
//                        return;
////                    int last = newList.size() - 1;
//                    String cellphone = newList.remove(0);
//                    System.out.println(cellphone);
//                    row.createCell(3).setCellValue(cellphone);
//
//                });
//            }catch (NullPointerException e){
//                e.printStackTrace();
//            }

            Main.updateText("Generating New Data");

            Row firstrow = sheet.getRow(0);
            firstrow.createCell(4).setCellValue("Index");
            firstrow.createCell(5).setCellValue("Data C");
            firstrow.createCell(6).setCellValue("Status");

            int i = 1;
            for (Data w : whatsapp) {
                Row currRow = sheet.getRow(i);
                currRow.createCell(4).setCellValue(w.getId());
                currRow.createCell(5).setCellValue(w.getNumber());
                currRow.createCell(6).setCellValue(w.getType());
                i++;
            }

            for (Data nw : nowhatsapp) {
                Row currRow = sheet.getRow(i);
                currRow.createCell(6).setCellValue(nw.getType());
                currRow.createCell(5).setCellValue(nw.getNumber());
                currRow.createCell(4).setCellValue(nw.getId());
                i++;
            }


            Main.updateText("Done");
            Main.updateText("**********************");

            // Closing the workbook
            currFile.close();

            workbook.write(new FileOutputStream(file_path));
            workbook.close();

            Main.updateText("Finished Process");
            Main.updateText("**********************");
        }catch (Exception e){
            e.printStackTrace();
            Main.updateText("Error");
            Main.updateText(e.getMessage());
            Main.updateText("**********************");
        }

        Main.busy = false;

    }
}