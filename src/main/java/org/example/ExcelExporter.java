package org.example;

import org.apache.poi.ss.formula.FormulaParser;
import org.apache.poi.ss.formula.FormulaRenderer;
import org.apache.poi.ss.formula.FormulaShifter;
import org.apache.poi.ss.formula.FormulaType;
import org.apache.poi.ss.formula.ptg.AreaPtgBase;
import org.apache.poi.ss.formula.ptg.Ptg;
import org.apache.poi.ss.formula.ptg.RefPtgBase;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FontScheme;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFEvaluationWorkbook;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFTable;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.example.model.Artist;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
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

                    XSSFCellStyle cellStyle1 = sheet.getRow(startRowIndex+1).getCell(startColIndex).getCellStyle();
                    XSSFCellStyle cellStyle2 = sheet.getRow(startRowIndex+1).getCell(startColIndex + 1).getCellStyle();
                    XSSFCellStyle cellStyle3 = sheet.getRow(startRowIndex+1).getCell(startColIndex + 2).getCellStyle();
                    XSSFCellStyle cellStyle4 = sheet.getRow(startRowIndex+1).getCell(startColIndex + 3).getCellStyle();

                    XSSFCell formulaCell = sheet.getRow(startRowIndex + 1).getCell(startColIndex + 3);
                    String formula = sheet.getRow(startRowIndex + 1).getCell(startColIndex + 3).getCellFormula();


                    Row dataRow = sheet.createRow(sheet.getLastRowNum() + 1);
                    Cell cell1 = dataRow.createCell(startColIndex++);
                    cell1.setCellValue(artist.getId());
                    cell1.setCellStyle(cellStyle1);

                    Cell cell2 = dataRow.createCell(startColIndex++);
                    cell2.setCellValue(artist.getArtistName());
                    cell2.setCellStyle(cellStyle2);

                    Cell cell3 = dataRow.createCell(startColIndex++);
                    cell3.setCellValue(artist.getDateOfBirth());
                    cell3.setCellStyle(cellStyle3);

                    Cell cell4 = dataRow.createCell(startColIndex);
                    String newformula = copyFormula(sheet, formula, 0, 1);
                    cell4.setCellFormula(newformula);
                    cell4.setCellStyle(cellStyle4);
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

    public XSSFWorkbook export1(FileInputStream stream, String sheetName, String tableName, List<Artist> artists) {
        try {
            XSSFWorkbook workbook = new XSSFWorkbook(stream);
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



                    XSSFCellStyle cellStyle1 = sheet.getRow(startRowIndex+2).getCell(startColIndex).getCellStyle();
                    XSSFCellStyle cellStyle2 = sheet.getRow(startRowIndex+2).getCell(startColIndex + 1).getCellStyle();
                    XSSFCellStyle cellStyle3 = sheet.getRow(startRowIndex+2).getCell(startColIndex + 2).getCellStyle();

                    Row dataRow = sheet.createRow(sheet.getLastRowNum() + 1);
                    Cell cell1 = dataRow.createCell(startColIndex++);
                    cell1.setCellValue(artist.getId());
                    cell1.setCellStyle(cellStyle1);


                    Cell cell2 = dataRow.createCell(startColIndex++);
                    cell2.setCellValue(artist.getArtistName());
                    cell2.setCellStyle(cellStyle2);


                    Cell cell3 = dataRow.createCell(startColIndex);
                    cell3.setCellValue(artist.getDateOfBirth());
                    cell3.setCellStyle(cellStyle3);
                }
            }
            return workbook;
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    private static String copyFormula(XSSFSheet sheet, String formula, int coldiff, int rowdiff) {

        XSSFEvaluationWorkbook workbookWrapper =
                XSSFEvaluationWorkbook.create((XSSFWorkbook) sheet.getWorkbook());
        Ptg[] ptgs = FormulaParser.parse(formula, workbookWrapper, FormulaType.CELL
                , sheet.getWorkbook().getSheetIndex(sheet));

        for (int i = 0; i < ptgs.length; i++) {
            if (ptgs[i] instanceof RefPtgBase) { // base class for cell references
                RefPtgBase ref = (RefPtgBase) ptgs[i];
                if (ref.isColRelative())
                    ref.setColumn(ref.getColumn() + coldiff);
                if (ref.isRowRelative())
                    ref.setRow(ref.getRow() + rowdiff);
            }
            else if (ptgs[i] instanceof AreaPtgBase) { // base class for range references
                AreaPtgBase ref = (AreaPtgBase) ptgs[i];
                if (ref.isFirstColRelative())
                    ref.setFirstColumn(ref.getFirstColumn() + coldiff);
                if (ref.isLastColRelative())
                    ref.setLastColumn(ref.getLastColumn() + coldiff);
                if (ref.isFirstRowRelative())
                    ref.setFirstRow(ref.getFirstRow() + rowdiff);
                if (ref.isLastRowRelative())
                    ref.setLastRow(ref.getLastRow() + rowdiff);
            }
        }

        formula = FormulaRenderer.toFormulaString(workbookWrapper, ptgs);
        return formula;
    }


}
