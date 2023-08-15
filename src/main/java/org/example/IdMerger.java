package org.example;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbookFactory;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Collections;

// Press Shift twice to open the Search Everywhere dialog and type `show whitespaces`,
// then press Enter. You can now see whitespace characters in your code.
public class IdMerger {
    public static void main(String[] args) throws Exception {
        try {
            IdMerger merger = new IdMerger();
            File file = new File("src/POC.xlsx");
            OPCPackage pkg = OPCPackage.open(file.getAbsolutePath());
            XSSFWorkbook wb = new XSSFWorkbook(pkg);
            // Get Excel file and input sheets
            XSSFSheet sheet1 = wb.getSheetAt(0);
            XSSFSheet sheet2 = wb.getSheetAt(1);
            // Variables for customer id array for both sheets
            int size1 = sheet1.getLastRowNum();
            int size2 = sheet2.getLastRowNum();
            ArrayList<Integer> id1 = new ArrayList<>();
            ArrayList<Integer> id2 = new ArrayList<>();

            // Remove duplicate ids, then combine arrays and sort in order
            ArrayList<Integer> finalIdList = merger.getMergedList(id1, id2);

            // Write into new final sheet where the ids are string values
            merger.mergeFinalSheet("src/POX.xlsx", finalIdList);

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public ArrayList<Integer> getMergedList(ArrayList<Integer> id1, ArrayList<Integer> id2) {
        id1.removeAll(id2);
        id1.addAll(id2);
        Collections.sort(id1);
        return id1;
    }

    public void mergeFinalSheet(String file, ArrayList<Integer> finalIdList) throws IOException, InvalidFormatException {
        File file2 = new File("src/POC.xlsx");
        Workbook wb2 = WorkbookFactory.create(file2);
        wb2.createSheet("final");
        XSSFSheet finalSheet = (XSSFSheet) wb2.getSheet("final");
        finalSheet.createRow(0);
        finalSheet.getRow(0).createCell(0);
        finalSheet.getRow(0).getCell(0).setCellValue("CustomerId");
        for (int i = 0; i < finalIdList.size(); i++) {
            finalSheet.createRow(i + 1);
            finalSheet.getRow(i + 1).createCell(0);
            finalSheet.getRow(i + 1).getCell(0)
                    .setCellValue(Integer.parseInt(String.valueOf(finalIdList.get(i))));
        }
    }
}