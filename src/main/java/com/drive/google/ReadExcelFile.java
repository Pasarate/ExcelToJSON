package com.drive.google;


import com.fasterxml.jackson.databind.ObjectMapper;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;


public class ReadExcelFile {

    public static void main(String[] args) throws IOException {
        String path = System.getProperty("user.dir") + "\\src\\main\\Resource\\leads.xlsx";
        List<LEADS_POJO> leads_pojo = readExcelFIle(path);

        ObjectMapper mapper = new ObjectMapper();
        String jsonString = mapper.writeValueAsString(leads_pojo);
        System.out.println(jsonString);

    }

    private static List<LEADS_POJO> readExcelFIle(String path) throws IOException {

        File file = new File(path);
        FileInputStream fis = new FileInputStream(file);
        XSSFWorkbook wb = new XSSFWorkbook(fis);
        wb.setMissingCellPolicy(Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);

        XSSFSheet sheet = wb.getSheetAt(0);

        Iterator<Row> row = sheet.iterator();

        List<LEADS_POJO> list_leads_pojo = new ArrayList<LEADS_POJO>();
        int rowNumber = 0;
        while (row.hasNext()) {

            Row cRow = row.next();

            if (rowNumber == 0) {
                rowNumber++;
                continue;
            }
            if (rowNumber == 15) {
                rowNumber++;
                break;
            }


            Iterator<Cell> cellInRow = cRow.iterator();
            LEADS_POJO obj = new LEADS_POJO();
            int cellNumber = 0;


            for (int cn = 0; cn <= cRow.getLastCellNum(); cn++) {

                Cell cCell = cRow.getCell(cn);

                if (cCell == null) {

                } else {
                    if (cellNumber == 0) {
                        obj.setProjectName(cCell.getStringCellValue());
                    } else if (cellNumber == 1) {
                        obj.setProjectType(cCell.getStringCellValue());
                    } else if (cellNumber == 2) {
                        obj.setDescription(cCell.getStringCellValue());
                    } else if (cellNumber == 3) {
                        obj.setSqft(cCell.getStringCellValue());
                    } else if (cellNumber == 4) {
                        obj.setEstimatedProjectCost(cCell.getStringCellValue());
                    } else if (cellNumber == 5) {
                        obj.setPermitNumber(cCell.getStringCellValue());
                    } else if (cellNumber == 6) {
                        obj.setNoticeType(cCell.getStringCellValue());
                    } else if (cellNumber == 7) {
                        obj.setStreet(cCell.getStringCellValue());
                    } else if (cellNumber == 8) {
                        obj.setCity(cCell.getStringCellValue());
                    } else if (cellNumber == 9) {
                        obj.setState(cCell.getStringCellValue());
                    } else if (cellNumber == 10) {
                        obj.setZipcode(cCell.getStringCellValue());
                    } else if (cellNumber == 11) {
                        obj.setContactInfo(cCell.getStringCellValue());
                    } else if (cellNumber == 12) {
                        obj.setContactPhone(cCell.getStringCellValue());
                    } else if (cellNumber == 13) {
                        obj.setContactAddress(cCell.getStringCellValue());
                    } else if (cellNumber == 14) {
                        obj.setContactEmail(cCell.getStringCellValue());
                    } else if (cellNumber == 15) {
                        obj.setOwner(cCell.getStringCellValue());
                    } else if (cellNumber == 16) {
                        obj.setArchitect(cCell.getStringCellValue());
                    } else if (cellNumber == 17) {
                        obj.setApplicationDate(cCell.getStringCellValue());
                    } else if (cellNumber == 18) {
                        obj.setUploadDate(cCell.getStringCellValue());
                    } else if (cellNumber == 19) {
                        obj.setStatus(cCell.getStringCellValue());
                    } else if (cellNumber == 20) {
                        obj.setCloseDate(cCell.getStringCellValue());
                    } else if (cellNumber == 21) {
                        obj.setLink(cCell.getStringCellValue());
                    } else if (cellNumber == 22) {
                        obj.setSource(cCell.getStringCellValue());
                    } else if (cellNumber == 23) {
                        obj.setConstructionStartDate((int) cCell.getNumericCellValue());
                    }
                }
                cellNumber++;
            }

            list_leads_pojo.add(obj);
            if (rowNumber == 1) {
                rowNumber++;
                break;
            }
        }


        return list_leads_pojo;
    }

}
