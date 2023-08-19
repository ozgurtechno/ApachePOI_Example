package Excel_ApachiPOI;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

public class WriteToExcel {

    public static void main(String[] args) throws IOException {

        XSSFWorkbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Persons");

        int rowCount = 0;
        for (int i = 1; i <= 9; i++) {

            for (int j = 1; j <= 9; j++) {

                Row row = sheet.createRow(rowCount++);
                Cell cell1 = row.createCell(0);
                cell1.setCellValue(i);

                Cell cell2 = row.createCell(1);
                cell2.setCellValue(" x ");

                Cell cell3 = row.createCell(2);
                cell3.setCellValue(j);

                Cell cell4 = row.createCell(3);
                cell4.setCellValue(" = ");

                Cell cell5 = row.createCell(4);
                cell5.setCellValue((i) * (j));

            }
        }

        FileOutputStream outputStream = new FileOutputStream("./temp.xlsx");
        workbook.write(outputStream);
        workbook.close();
    }

}
