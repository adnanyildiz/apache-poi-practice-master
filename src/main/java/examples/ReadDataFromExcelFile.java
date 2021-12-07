package examples;

import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

public class ReadDataFromExcelFile {

        public static void main(String[] args) throws IOException {

            String filePath = "data.xlsx";
            String sheetName = "data";

            InputStream in = new FileInputStream(filePath);//from java
            Workbook workbook = WorkbookFactory.create(in);//from poi
            Sheet sheet = workbook.getSheet(sheetName);//from poi

            Row firstRow = sheet.getRow(0);

            Cell firstCell = firstRow.getCell(0);
            System.out.println("Value is: " + firstCell.getStringCellValue());//Value is: id

            Cell secondCell = firstRow.getCell(1);
            System.out.println("Value of 2nd cell is: " + secondCell.getStringCellValue());//Value of 2nd cell is: first_name

            System.out.println("Number of cells in the row " + firstRow.getPhysicalNumberOfCells());//6

            for (Cell cell : firstRow) {//firstrow.iter
                System.out.println(cell);
            }

           // firstRow.cellIterator().forEachRemaining(System.out::println); same with above


            int NumberOfRows = sheet.getPhysicalNumberOfRows();
            System.out.println("Number of rows in the sheet " + sheet.getPhysicalNumberOfRows());//Number of rows in the sheet 21

            sheet.forEach(row -> {
                row.cellIterator().forEachRemaining(cell -> System.out.print(cell.toString()+ " "));
                System.out.println("");
            });
        }

    }

