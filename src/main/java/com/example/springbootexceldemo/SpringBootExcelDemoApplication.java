package com.example.springbootexceldemo;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import java.io.File;
import java.io.IOException;

@SpringBootApplication
public class SpringBootExcelDemoApplication {

    public static void main(String[] args) {
        SpringApplication.run(SpringBootExcelDemoApplication.class, args);

        File currDir = new File(".");
        String path = currDir.getAbsolutePath();
        String fileLocation = path.substring(0, path.length() - 1) + "temp.xlsx";

        ExcelHelper excelHelper = new ExcelHelper();

        //Write excel file
        try {
            excelHelper.writeExcel(fileLocation);
        } catch (IOException e) {
            e.printStackTrace();
        }

        //Read excel file
        try {
            excelHelper.readExcel(fileLocation);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

}
