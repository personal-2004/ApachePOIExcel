package org.o7planning.apachepoiexcel.demo;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class CopyPriceERC {
    public static void main(String[] args) throws IOException {
        // Read XSL file
        FileInputStream inputStream = new FileInputStream(new File("C:/demo/erc_selected_vendors_20200611_10h09m.xls"));
        // Get the workbook instance for XLS file
        HSSFWorkbook workbook = new HSSFWorkbook(inputStream);
        // Get first sheet from the workbook
        HSSFSheet sheet = workbook.getSheetAt(0);
        //Rename first sheet from the workbook// workbook.setSheetName(0,"ERC");
        List<Cell> retail = new ArrayList<Cell>();
       for (int i = 0; i < 50 ; i++) {
            Row row = sheet.getRow(i);
            Cell cOld = row.getCell(6);
            retail.add(cOld);
         }
        inputStream.close();
        // Read XSL file
        FileInputStream inputStream1 = new FileInputStream(new File("C:/demo/Копия xls-blank.xls"));
        // Get the workbook instance for XLS file
        HSSFWorkbook workbook1 = new HSSFWorkbook(inputStream1);
        // Get first sheet from the workbook
        HSSFSheet sheet1 = workbook1.getSheetAt(0);
        Cell cNew ;
        int i = 0;
        for (Cell rperc:retail ) {
           Row row = sheet1.getRow(i);
           cNew = row.createCell(5);
           cNew.setCellValue(String.valueOf(rperc));
           System.out.println("rperc = " + rperc);
           i++;
        }
        
        File file = new File("C:/demo/Копия xls-blank.xls");
        // file.getParentFile().mkdirs();
        FileOutputStream outFile = new FileOutputStream(file);
        workbook1.write(outFile);
        System.out.println("Created file: " + file.getAbsolutePath());
    }
}
