package org.o7planning.apachepoiexcel.demo;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;

public class CreateExcelDemoLight {


    public static void main(String[] args) throws IOException {
      // HSSFWorkbook workbook = new HSSFWorkbook();
      //  HSSFSheet sheet = workbook.createSheet("ERC");

        // Read XSL file
        FileInputStream inputStream = new FileInputStream(new File("C:/demo/employee.xls"));
        // Get the workbook instance for XLS file
        HSSFWorkbook workbook = new HSSFWorkbook(inputStream);
       // Get first sheet from the workbook
        HSSFSheet sheet = workbook.getSheetAt(0);
        //Rename first sheet from the workbook
        workbook.setSheetName(0,"ERC");
        //delete Row index:
       // sheet.removeRow(sheet.getRow(1));
        //delete column

        for (int i = 0; i < 4 ; i++) {
            Row row = sheet.getRow(i);
            Cell cOld = row.getCell(3);
            row.removeCell(cOld);
//            cOld = row.createCell(3, CellType.STRING);
//            cOld.setCellValue("55555");
             }

      /*  // Get iterator to all the rows in current sheet
        Iterator<Row> rowIterator = sheet.iterator();

        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            // Get iterator to all cells of current row
            Iterator<Cell> cellIterator = row.cellIterator();

            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();

                // Change to getCellType() if using POI 4.x
                CellType cellType = cell.getCellTypeEnum();*/
       // private void deleteColumn(Sheet sheet, int columnToDelete) {
           /* for (int rId = 0; rId < sheet.getLastRowNum(); rId++) {
                Row row = sheet.getRow(rId);
                for (int cID = columnToDelete; cID < row.getLastCellNum(); cID++) {
                    Cell cOld = row.getCell(cID);
                    if (cOld != null) {
                        row.removeCell(cOld);
                    }*/
                   /* Cell cNext = row.getCell(cID + 1);*/
                 /*   if (cNext != null) {
                        Cell cNew = row.createCell(cID, cNext.getCellTypeEnum());
                        cloneCell(cNew, cNext);
                        //Set the column width only on the first row.
                        //Other wise the second row will overwrite the original column width set previously.
                        if(rId == 0) {
                            sheet.setColumnWidth(cID, sheet.getColumnWidth(cID + 1));

                        }
                    }*/
               // }



        File file = new File("C:/demo/ERC.xls");
       // file.getParentFile().mkdirs();
        FileOutputStream outFile = new FileOutputStream(file);
        workbook.write(outFile);
        System.out.println("Created file: " + file.getAbsolutePath());
    }
}
