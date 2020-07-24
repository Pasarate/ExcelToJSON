package com.drive.google;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

public class DataType {

    public static void main(String[] args) throws IOException {

        String path = System.getProperty("user.dir") + "\\src\\main\\Resource\\leads.xlsx";

        File file = new File(path);
        FileInputStream fis = new FileInputStream(file);

        XSSFWorkbook wb = new XSSFWorkbook(fis);
        wb.setMissingCellPolicy(Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);


        XSSFSheet sheet = wb.getSheetAt(0);

        Iterator<Row> row = sheet.iterator();

        int rowNumber = 0;


        while (row.hasNext()) {

            Row cRow = row.next();

            if (rowNumber == 0) {
                rowNumber++;
                continue;
            }


            Iterator<Cell> cellInRow = cRow.iterator();

            int cellNumber = 0;

            for (int cn = 0; cn <= cRow.getLastCellNum(); cn++) {

                Cell cell = cRow.getCell(cn);
                System.out.println(cell + " <- ");
                if (cell == null) {

                } else {
                    System.out.println(cell.getCellType() + " ***** " + cellNumber);
                }
                cellNumber++;
            }


            if (rowNumber == 1) {
                rowNumber++;
                break;
            }

        }
    }
}