package com.example.springbootexceldemo;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class ExcelHelper {
    Logger logger = LoggerFactory.getLogger(ExcelHelper.class);

    public void readExcel(String fileLocation) throws IOException {
        List<List<String>> rowsData = new ArrayList<>();

        FileInputStream file = new FileInputStream(new File(fileLocation));
        Workbook workbook = new XSSFWorkbook(file);
        Sheet sheet = workbook.getSheetAt(0);

        for (int i = 0; i < sheet.getPhysicalNumberOfRows(); i++) {
            List<String> cellsData = new ArrayList<>();
            Row row = sheet.getRow(i);

            for (int j = 0; j < row.getPhysicalNumberOfCells(); j++) {
                Cell cell = row.getCell(j);
                cellsData.add(cell.getStringCellValue());
            }

            rowsData.add(cellsData);
        }

        logger.info("Row length: " + rowsData.size());
        logger.info("Cell length at row 0: " + rowsData.get(0).size());
        workbook.close();
    }

    public void writeExcel(String fileLocation) throws IOException {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("DemoSheet");

        Row rowHeader = sheet.createRow(0);
        Cell cellHeader0 = rowHeader.createCell(0);
        Cell cellHeader1 = rowHeader.createCell(1);
        cellHeader0.setCellValue("NO");
        cellHeader1.setCellValue("NAME");

        Row row1 = sheet.createRow(1);
        Cell row1Cell0 = row1.createCell(0);
        Cell row1Cell1 = row1.createCell(1);
        row1Cell0.setCellValue("1");
        row1Cell1.setCellValue("TED");

        Row row2 = sheet.createRow(2);
        Cell row2Cell0 = row2.createCell(0);
        Cell row2Cell1 = row2.createCell(1);
        row2Cell0.setCellValue("2");
        row2Cell1.setCellValue("TAIY");

        FileOutputStream outputStream = new FileOutputStream(fileLocation);
        workbook.write(outputStream);

        workbook.close();

        logger.info("Written excel to " + fileLocation);
    }
}
