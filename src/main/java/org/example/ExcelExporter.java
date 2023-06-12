package org.example;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFTable;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.example.model.Artist;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.List;

public class ExcelExporter {

    public void export(String filePath, String sheetName, String tableName, List<Artist> artists) {

        try (FileInputStream inputStream = new FileInputStream(filePath);) {
            XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
            XSSFSheet sheet = workbook.getSheet(sheetName);

            XSSFTable table = sheet.getTables()
                    .stream()
                    .filter(t -> t.getName().equals(tableName))
                    .findFirst()
                    .orElse(null);

            if (table != null) {
                for (Artist artist : artists) {
                    int startColIndex = table.getStartColIndex();
                    int startRowIndex = table.getStartRowIndex();

                    XSSFRow row = sheet.getRow(startRowIndex);

                    XSSFCellStyle cellStyle1 = sheet.getRow(startRowIndex).getCell(startColIndex).getCellStyle();
                    XSSFCellStyle cellStyle2 = sheet.getRow(startRowIndex).getCell(startColIndex+1).getCellStyle();
                    XSSFCellStyle cellStyle3 = sheet.getRow(startRowIndex).getCell(startColIndex+2).getCellStyle();

                    Row dataRow = sheet.createRow(sheet.getLastRowNum() + 1);
                    Cell cell1 = dataRow.createCell(startColIndex++);
                    cell1.setCellValue(artist.getId());
                    cell1.setCellStyle(cellStyle1);


                    dataRow.createCell(startColIndex++).setCellValue(artist.getArtistName());
                    dataRow.createCell(startColIndex).setCellValue(artist.getDateOfBirth());
                }
            }

            try (FileOutputStream outputStream = new FileOutputStream(filePath)) {
                workbook.write(outputStream);
            }
            System.out.println("Data exported");

        } catch (Exception e) {
            System.out.println("Error exporting data: " + e.getMessage());
        }

    }

}
