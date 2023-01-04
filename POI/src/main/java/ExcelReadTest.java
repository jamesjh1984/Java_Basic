import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joda.time.DateTime;
import org.junit.jupiter.api.Test;

import java.io.FileInputStream;
import java.io.FileOutputStream;

public class ExcelReadTest {

    String path = "D:\\IT\\Projects\\GitHub\\Java_Basic\\POI\\src\\main\\resources\\";


    @Test
    public void xlsRead() throws Exception {

        // 1. get FileInputStream
        FileInputStream fileInputStream = new FileInputStream(path + "xlsWrite.xls");

        // 2. get a sheet
        Workbook workbook = new HSSFWorkbook(fileInputStream);
        Sheet sheet = workbook.getSheetAt(0);

        // 3. get the 1st row and cell
        Row row0 = sheet.getRow(0);
        Cell cell00 = row0.getCell(0);
        Cell cell01 = row0.getCell(1);
        System.out.println(cell00.getStringCellValue() + "|" + cell01.getStringCellValue());

        // 4. get the 2nd row and cell
        Row row1 = sheet.getRow(1);
        Cell cell10 = row1.getCell(0);
        Cell cell11 = row1.getCell(1);
        System.out.println(cell10.getStringCellValue() + "|" + cell11.getStringCellValue());

        // 4. get the 3rd row and cell
        Row row2 = sheet.getRow(2);
        Cell cell20 = row2.getCell(0);
        Cell cell21 = row2.getCell(1);
        System.out.println(cell20.getStringCellValue() + "|" + cell21.getNumericCellValue());

        fileInputStream.close();
    }



    @Test
    public void xlsxRead() throws Exception {

        // 1. get FileInputStream
        FileInputStream fileInputStream = new FileInputStream(path + "xlsxWrite.xlsx");

        // 2. get a sheet
        Workbook workbook = new XSSFWorkbook(fileInputStream);
        Sheet sheet = workbook.getSheetAt(0);

        // 3. get the 1st row and cell
        Row row0 = sheet.getRow(0);
        Cell cell00 = row0.getCell(0);
        Cell cell01 = row0.getCell(1);
        System.out.println(cell00.getStringCellValue() + "|" + cell01.getStringCellValue());

        // 4. get the 2nd row and cell
        Row row1 = sheet.getRow(1);
        Cell cell10 = row1.getCell(0);
        Cell cell11 = row1.getCell(1);
        System.out.println(cell10.getStringCellValue() + "|" + cell11.getStringCellValue());

        // 4. get the 3rd row and cell
        Row row2 = sheet.getRow(2);
        Cell cell20 = row2.getCell(0);
        Cell cell21 = row2.getCell(1);
        System.out.println(cell20.getStringCellValue() + "|" + cell21.getNumericCellValue());

        fileInputStream.close();
    }


}
