package com.drive.google;

import org.apache.poi.hslf.record.CString;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.*;

public final class READFILE {
    private static File file;
    private static FileInputStream fis;
    private static XSSFWorkbook wb;
    private static XSSFSheet sheet;

    private static Map<String, String> leadmap = new HashMap();
    private static List<String> header = new ArrayList();
    private static String[] headerName;

    public static void getFile() {
        String path = System.getProperty("user.dir") + "\\src\\main\\Resource\\leads.xlsx";
        file = new File(path);
    }

    public static void readFile() throws IOException {

        FileInputStream fis = new FileInputStream(file);
        XSSFWorkbook wb = new XSSFWorkbook(fis);
        XSSFSheet sheet = wb.getSheetAt(0);

        int noOfRows = sheet.getPhysicalNumberOfRows();
        Row firstRow = sheet.getRow(0);
        int noOfCells = firstRow.getLastCellNum();

        headerName = new String[noOfCells];

        for (int i = 0; i < noOfCells; i++) {
            Cell cell1 = firstRow.getCell(i);
            headerName[i] = cell1.getStringCellValue();
        }
        for (String i : headerName) {
            System.out.println(i);
        }

    }

    public static void getRow() {

    }


    public static void closeFile() throws IOException {


    }

}
