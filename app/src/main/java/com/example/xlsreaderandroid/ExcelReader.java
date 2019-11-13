package com.example.xlsreaderandroid;

import android.content.Context;
import android.util.Log;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

public class ExcelReader {

    List<String> emailList = new ArrayList<>();

    public List<String> getEmailList(Context context) {
        InputStream myInput = null;
        OPCPackage myFileSystem = null;

        //  Don't forget to Change to your assets folder excel sheet
        try {
            myInput = context.getResources().openRawResource(R.raw.email);
            File tempFile = File.createTempFile("email", "xlsx");
            copyFile(myInput, new FileOutputStream(tempFile));

            // Create a OPCPackage object
            myFileSystem = OPCPackage.open(tempFile);
        } catch (InvalidFormatException e) {
            e.printStackTrace();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }


        // Creating a Workbook from an Excel file (.xls or .xlsx)
        Workbook workbook = null;
        try {
            workbook = WorkbookFactory.create(myFileSystem);
        } catch (IOException e) {
            e.printStackTrace();
            return null;
        }

        // Retrieving the number of sheets in the Workbook
        System.out.println("Workbook has " + workbook.getNumberOfSheets() + " Sheets : ");


        // 1. You can obtain a sheetIterator and iterate over it
        Iterator<Sheet> sheetIterator = workbook.sheetIterator();
        System.out.println("Retrieving Sheets using Iterator");
        while (sheetIterator.hasNext()) {
            Sheet sheet = sheetIterator.next();
            System.out.println("=> " + sheet.getSheetName());
        }

        // Getting the Sheet at index zero
        Sheet sheet = workbook.getSheetAt(0);

        Iterator<Row> rowIterator = sheet.iterator(); // Traversing over each row of XLSX file
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next(); // For each row, iterate through each columns
            Iterator<Cell> cellIterator = row.cellIterator();
            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();
                if (cell.getColumnIndex() == 1 && cell.getRowIndex() != 0) {
                    Log.d("Emails = ", cell.getStringCellValue());
                    emailList.add(cell.getStringCellValue());
                }

            }
        }

        /*for (Sheet sheet : workbook) {
            System.out.println("=> " + sheet.getSheetName());
            for (Row row: sheet) {
                for(Cell cell: row) {
                    String cellValue = dataFormatter.formatCellValue(cell);
                    System.out.print(cellValue + "\t");
                }
                System.out.println();
            }
        }*/

        return emailList;
    }

    private void copyFile(InputStream in, OutputStream out) throws IOException {
        byte[] buffer = new byte[1024];
        int read;
        while ((read = in.read(buffer)) != -1) {
            out.write(buffer, 0, read);
        }
    }

    private File createFileFromInputStream(InputStream inputStream) {

        try {
            String excelFilePath = "email.xlsx";

            File f = new File(excelFilePath);
            OutputStream outputStream = new FileOutputStream(f);
            byte buffer[] = new byte[1024];
            int length = 0;

            while ((length = inputStream.read(buffer)) > 0) {
                outputStream.write(buffer, 0, length);
            }

            outputStream.close();
            inputStream.close();

            return f;
        } catch (IOException e) {
            //Logging exception
        }

        return null;
    }

}
