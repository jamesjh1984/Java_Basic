import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFDataValidation;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joda.time.DateTime;
import org.junit.jupiter.api.Test;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Date;

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





    @Test
    public void PlayerXLSXRead() throws Exception {

        // 1. get FileInputStream
        FileInputStream fileInputStream = new FileInputStream(path + "Player.xlsx");


        // 2. get a sheet
        Workbook workbook = new XSSFWorkbook(fileInputStream);
        Sheet sheet = workbook.getSheetAt(0);
        int NumberOfRows = sheet.getPhysicalNumberOfRows(); // row count
        System.out.println("NumberOfRows => " + NumberOfRows);

        // 3. get row header
        Row header = sheet.getRow(0);
        if (header!=null) {
            String headerString = new String();
            int NumberOfCells = header.getPhysicalNumberOfCells(); // 一行中Cell的个数
            for (int i = 0; i < NumberOfCells; i++) {
                Cell cell = header.getCell(i);
                if(cell!=null) {
                    //int cellType = cell.getCellType();
                    String CellValue = cell.getStringCellValue();
                    headerString = headerString + "|" + CellValue;
                }
            }
            System.out.println(headerString.substring(1,headerString.length()));
        }


        // 4. get each row value
        for (int i = 1; i < NumberOfRows; i++) {
            String rowString = new String();
            Row rowData = sheet.getRow(i);
            if(rowData!=null){
                int NumberOfCells = header.getPhysicalNumberOfCells(); // 一行中Cell的个数
                // System.out.println("NumberOfCells => " + NumberOfCells);

                for (int j = 0; j < NumberOfCells; j++) {
                    Cell cell = rowData.getCell(j);
                    if(cell!=null) {
                        int cellType = cell.getCellType();
                        String CellValue = "";

                        switch (cellType) {
                            case Cell.CELL_TYPE_STRING:
                                CellValue = cell.getStringCellValue();
                                break;
                            case Cell.CELL_TYPE_BOOLEAN:
                                CellValue = String.valueOf(cell.getBooleanCellValue());
                                break;
                            case Cell.CELL_TYPE_BLANK:
                                break;
                            case Cell.CELL_TYPE_NUMERIC:
                                if (HSSFDateUtil.isCellDateFormatted(cell)) {
                                    Date date = cell.getDateCellValue();
                                    CellValue = new DateTime(date).toString("yyyyMMdd");
                                } else {
                                    CellValue = cell.toString();
                                }
                                break;
                            case Cell.CELL_TYPE_ERROR:
                                System.out.println("Cell Type is error.");
                                break;
                        }
                        rowString = rowString + "|" + CellValue;
                    }
                }
                //rowString = rowString + "\n";
                System.out.println(rowString.substring(1,rowString.length()));
            }
        }


        fileInputStream.close();
    }


}
