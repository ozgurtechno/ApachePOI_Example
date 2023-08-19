package Excel_ApachiPOI;

import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;
import java.io.IOException;

public class Simple_Example {

    public static void main(String[] args) throws IOException {

        FileInputStream inputStream = new FileInputStream("src/test/java/Excel_ApachiPOI/students.xlsx");
        Workbook workbook= WorkbookFactory.create(inputStream);
        Sheet sheet = workbook.getSheet("students");
        Row row = sheet.getRow(1);
        Cell cell = row.getCell(1);

        System.out.println(sheet.getRow(1).getCell(1));
        System.out.println(row.getCell(0));
        System.out.println(sheet.getFirstRowNum());
        System.out.println(sheet.getLastRowNum());
        System.out.println(sheet.getPhysicalNumberOfRows());
        System.out.println(sheet.getRow(2).getPhysicalNumberOfCells());
        System.out.println(sheet.getRow(2).getLastCellNum());

        Cell x = sheet.getRow(1).getCell(4);
        Cell y = sheet.getRow(1).getCell(3);
        System.out.println("x.getCellType() = " + x.getCellType());
        System.out.println("y.getCellType() = " + y.getCellType());

        System.out.println((int) x.getNumericCellValue());
        System.out.println(sheet.getRow(2).getCell(4).getStringCellValue());



       System.out.println(sheet.getRow(2).getCell(2).getSheet());


    }
}
