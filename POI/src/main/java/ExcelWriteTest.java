import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joda.time.DateTime;
import org.junit.jupiter.api.Test;

import java.io.FileOutputStream;


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
    public void xlsBigDataWrite() throws Exception {

        long begin = System.currentTimeMillis();


        // 1. create a workbook
        Workbook workbook = new HSSFWorkbook();

        // 2. create a sheet
        Sheet sheet = workbook.createSheet("xlsBigData");

        // 3. for loop to create cell, begin with 0
        // java.lang.IllegalArgumentException: Invalid row number (65536) outside allowable range (0..65535)
        for (int i = 0; i < 65536; i++) {
            Row row = sheet.createRow(i);
            for (int j = 0; j < 10; j++) {
                Cell cell = row.createCell(j);
                cell.setCellValue(j);
            }
        }

        // 4. create a File
        FileOutputStream fileOutputStream = new FileOutputStream(path + "xlsBigDataWrite.xls");

        workbook.write(fileOutputStream);
        // close
        fileOutputStream.close();

        System.out.println("xlsBigDataWrite.xls has been created.");


        long end = System.currentTimeMillis();


        System.out.println("Time:" + (double) (end - begin)/1000);

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


    // take long time, Time:19.509s for 100000
    @Test
    public void xlsxBigDataWrite() throws Exception {

        long begin = System.currentTimeMillis();


        // 1. create a workbook
        Workbook workbook = new XSSFWorkbook();

        // 2. create a sheet
        Sheet sheet = workbook.createSheet("xlsxBigData");

        // 3. for loop to create cell, begin with 0
        // java.lang.IllegalArgumentException: Invalid row number (65536) outside allowable range (0..65535)
        for (int i = 0; i < 100000; i++) {
            Row row = sheet.createRow(i);
            for (int j = 0; j < 10; j++) {
                Cell cell = row.createCell(j);
                cell.setCellValue(j);
            }
        }

        // 4. create a File
        FileOutputStream fileOutputStream = new FileOutputStream(path + "xlsxBigDataWrite.xlsx");

        workbook.write(fileOutputStream);
        // close
        fileOutputStream.close();

        System.out.println("xlsxBigDataWrite.xlsx has been created.");


        long end = System.currentTimeMillis();


        System.out.println("Time:" + (double) (end - begin)/1000);

    }





    // take less time, Time:6.049 for 100000, will generate a temporary file
    @Test
    public void xlsxBigDataEnhanceWrite() throws Exception {

        long begin = System.currentTimeMillis();


        // 1. create a workbook
        Workbook workbook = new SXSSFWorkbook();

        // 2. create a sheet
        Sheet sheet = workbook.createSheet("xlsxBigDataEnhance");

        // 3. for loop to create cell, begin with 0
        // java.lang.IllegalArgumentException: Invalid row number (65536) outside allowable range (0..65535)
        for (int i = 0; i < 100000; i++) {
            Row row = sheet.createRow(i);
            for (int j = 0; j < 10; j++) {
                Cell cell = row.createCell(j);
                cell.setCellValue(j);
            }
        }

        // 4. create a File
        FileOutputStream fileOutputStream = new FileOutputStream(path + "xlsxBigDataEnhanceWrite.xlsx");

        workbook.write(fileOutputStream);
        // close temporary file
        fileOutputStream.close();
        // clean
        ((SXSSFWorkbook) workbook).dispose();

        System.out.println("xlsxBigDataEnhanceWrite.xlsx has been created.");


        long end = System.currentTimeMillis();


        System.out.println("Time:" + (double) (end - begin)/1000);

    }


}
