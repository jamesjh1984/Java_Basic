import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joda.time.DateTime;
import org.junit.jupiter.api.Test;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

public class ExcelWriteTest {

    String path = "D:\\IT\\Projects\\GitHub\\Java_Basic\\POI\\src\\main\\resources\\";

    @Test
    public void xlsWrite() throws Exception {

        // 1. create a workbook
        Workbook workbook = new HSSFWorkbook();

        // 2. create a sheet
        Sheet sheet = workbook.createSheet("xls");

        // 3. create first row and its cell, begin with 0
        Row row0 = sheet.createRow(0);

        Cell cell00 = row0.createCell(0);
        cell00.setCellValue("Name");
        Cell cell01 = row0.createCell(1);
        cell01.setCellValue("James");


        // 4. create second row and its cell, begin with 1
        Row row1 = sheet.createRow(1);

        Cell cell10 = row1.createCell(0);
        cell10.setCellValue("Date");
        Cell cell11 = row1.createCell(1);
        String time = new DateTime().toString("yyyy-MM-dd HH:mm:ss");
        cell11.setCellValue(time);

        // 6. create a File
        FileOutputStream fileOutputStream = new FileOutputStream(path + "xlsWrite.xls");

        workbook.write(fileOutputStream);
        // close
        fileOutputStream.close();

        System.out.println("xlsWrite.xls has been created.");
    }



    @Test
    public void xlsxWrite() throws Exception {

        // 1. create a workbook
        Workbook workbook = new XSSFWorkbook();

        // 2. create a sheet
        Sheet sheet = workbook.createSheet("xlsx");

        // 3. create first row and its cell, begin with 0
        Row row0 = sheet.createRow(0);

        Cell cell00 = row0.createCell(0);
        cell00.setCellValue("Name");
        Cell cell01 = row0.createCell(1);
        cell01.setCellValue("James");


        // 4. create second row and its cell, begin with 1
        Row row1 = sheet.createRow(1);

        Cell cell10 = row1.createCell(0);
        cell10.setCellValue("Date");
        Cell cell11 = row1.createCell(1);
        String time = new DateTime().toString("yyyy-MM-dd HH:mm:ss");
        cell11.setCellValue(time);

        // 6. create a File
        FileOutputStream fileOutputStream = new FileOutputStream(path + "xlsxWrite.xlsx");

        workbook.write(fileOutputStream);
        // close
        fileOutputStream.close();

        System.out.println("xlsxWrite.xlsx has been created.");
    }




}
