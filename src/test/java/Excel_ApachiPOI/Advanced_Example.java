package Excel_ApachiPOI;

import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.TreeMap;

public class Advanced_Example {

    static String TestDataPath = "src/test/java/Excel_ApachiPOI/students.xlsx";
    static HashMap<String, LinkedHashMap<String, String>> hm1 = new HashMap<>();
    // Tom={firstname=Tom, lastname=cruise, }, Maria={}


    public static void main(String[] args) throws IOException {
        ReadTestData("students");
        System.out.println(hm1);
    }


    public static void ReadTestData(String sheetName) throws IOException {

        FileInputStream inputStream = new FileInputStream(TestDataPath);
        Workbook workbook= WorkbookFactory.create(inputStream);
        Sheet sheet = workbook.getSheet(sheetName);
        Row HeaderRow = sheet.getRow(0);

        String s3 = "";

        for (int i = 1; i < 3; i++) {
            Row currentRow = sheet.getRow(i);

            LinkedHashMap<String, String> currentHash = new LinkedHashMap<>(); // ordered- insertion order   --->
//            TreeMap<String, String> currentHash = new TreeMap<>(); // sorted   --->
//            HashMap<String, String> currentHash = new HashMap<>(); // unordered   --->
            for (int j = 0; j < currentRow.getPhysicalNumberOfCells() - 1; j++) {

                Cell currentCell1 = currentRow.getCell(0);

                switch (currentCell1.getCellType()) {
                    case STRING:
                        s3 = currentCell1.getStringCellValue();
                        break;
                    case NUMERIC:
                        s3 = String.valueOf(currentCell1.getNumericCellValue());
                        break;
                }

                Cell currentCell = currentRow.getCell(j);
                switch (currentCell.getCellType()) {
                    case STRING:
                        currentHash.put(HeaderRow.getCell(j).getStringCellValue(),
                                currentCell.getStringCellValue());
                        break;
                    case NUMERIC:
                        currentHash.put(HeaderRow.getCell(j).getStringCellValue(),
                                String.valueOf(currentCell.getNumericCellValue()));
                        break;
                }
            }
            hm1.put(s3, currentHash);
        }
    }
}
