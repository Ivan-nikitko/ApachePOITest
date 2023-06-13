package org.example;

import com.nhl.dflib.DataFrame;
import com.nhl.dflib.Series;
import com.nhl.dflib.row.RowProxy;
import org.apache.poi.ss.formula.FormulaParser;
import org.apache.poi.ss.formula.FormulaRenderer;
import org.apache.poi.ss.formula.FormulaType;
import org.apache.poi.ss.formula.ptg.AreaPtgBase;
import org.apache.poi.ss.formula.ptg.Ptg;
import org.apache.poi.ss.formula.ptg.RefPtgBase;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFEvaluationWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFTable;
import org.apache.poi.xssf.usermodel.XSSFTableColumn;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Iterator;
import java.util.List;

public class ExcelExporter {
//
//    public void export(String filePath, String sheetName, String tableName, List<Artist> artists) {
//
//        try (FileInputStream inputStream = new FileInputStream(filePath);) {
//            XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
//            XSSFSheet sheet = workbook.getSheet(sheetName);
//
//            XSSFTable table = sheet.getTables()
//                    .stream()
//                    .filter(t -> t.getName().equals(tableName))
//                    .findFirst()
//                    .orElse(null);
//
//            if (table != null) {
//
//
//                int startColIndex = table.getStartColIndex();
//                int startRowIndex = table.getStartRowIndex();
//
//                int currentColIndex = startColIndex;
//                int currentRowIndex = startRowIndex;
//
//                for (Artist artist : artists) {
//
//                    String formula = sheet.getRow(startRowIndex + 1).getCell(startColIndex + 3).getCellFormula();
//
//                    Row dataRow = sheet.createRow(sheet.getLastRowNum());
//
//                    Cell cell1 = dataRow.createCell(startColIndex);
//                    cell1.setCellValue(artist.getId());
//                    cell1.setCellStyle(getCellStyle(sheet, startColIndex++, startRowIndex));
//
//                    Cell cell2 = dataRow.createCell(startColIndex);
//                    cell2.setCellValue(artist.getArtistName());
//                    cell2.setCellStyle(getCellStyle(sheet, startColIndex++, startRowIndex));
//
//                    Cell cell3 = dataRow.createCell(startColIndex);
//                    cell3.setCellValue(artist.getDateOfBirth());
//                    cell3.setCellStyle(getCellStyle(sheet, startColIndex++, startRowIndex));
//
//                    Cell cell4 = dataRow.createCell(startColIndex);
//                    //  String newFormula = copyFormula(sheet, formula, 0, i+1);
//                    cell4.setCellValue(formula);
//                    cell4.setCellStyle(getCellStyle(sheet, startColIndex, startRowIndex));
//                }
//            }
//
//            try (FileOutputStream outputStream = new FileOutputStream(filePath)) {
//                workbook.write(outputStream);
//            }
//            System.out.println("Data exported");
//
//        } catch (Exception e) {
//            System.out.println("Error exporting data: " + e.getMessage());
//        }
//
//    }

    public void export(String filePath, String tableName, DataFrame dataFrame) {

        try (FileInputStream inputStream = new FileInputStream(filePath);) {
            XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
            XSSFSheet sheet = workbook.getSheet(tableName);

            int lastRowNum = sheet.getLastRowNum();

            XSSFTable table = sheet.getTables()
                    .stream()
                    .filter(t -> t.getName().equals(tableName))
                    .findFirst()
                    .orElse(null);

            if (table != null) {
                int startColIndex = table.getStartColIndex();
                int startRowIndex = table.getStartRowIndex();
//
//                int currentColIndex = startColIndex;
//                int currentRowIndex = startRowIndex;


                List<RowProxy> rowProxies = new ArrayList<>();
                dataFrame.iterator().forEachRemaining(rowProxies::add);


//                String table1 = Printers.tabular.toString(dataFrame);
//                System.out.println(table1);
                Iterator<RowProxy> iterator = dataFrame.iterator();

//                while (dataFrame.iterator().hasNext()) {
//                    RowProxy next = iterator.next();
//
//
//                }
                List<XSSFTableColumn> columns = table.getColumns();
                table.setDataRowCount(rowProxies.size());


                List<String> frameColumnsNames = Arrays.asList(dataFrame.getColumnsIndex().getLabels());

                for (int i = 0; i < rowProxies.size(); i++) {
                    Row dataRow = sheet.createRow(sheet.getLastRowNum() + 1);
                    for (XSSFTableColumn tableColumn : columns) {
                        Cell cell = dataRow.createCell(table.getStartColIndex() + tableColumn.getColumnIndex());
                        if (frameColumnsNames.contains(tableColumn.getName())) {
                            Series<Object> frameColumn = dataFrame.getColumn(tableColumn.getName());
                            cell.setCellValue(frameColumn.get(i).toString());
                        } else {
                            XSSFCell formulaCandidateCell = sheet.getRow(startRowIndex + 1).getCell(startColIndex+tableColumn.getColumnIndex());
                            if(formulaCandidateCell.getCellType().equals(CellType.FORMULA)){
                               cell.setCellFormula(formulaCandidateCell.getCellFormula());

                            } else {
                                cell.setCellValue("Default");
                            }
                        }
                        cell.setCellStyle(getCellStyleFromTemplate(sheet,startColIndex+tableColumn.getColumnIndex(),startRowIndex));
                    }
                }


                //   }


                //           String formula = sheet.getRow(startRowIndex + 1).getCell(startColIndex + 3).getCellFormula();
//
//                Row dataRow = sheet.createRow(sheet.getLastRowNum());
//
//                Cell cell1 = dataRow.createCell(startColIndex);
//                cell1.setCellValue(artist.getId());
//                cell1.setCellStyle(getCellStyle(sheet, startColIndex++, startRowIndex));
//
//                Cell cell2 = dataRow.createCell(startColIndex);
//                cell2.setCellValue(artist.getArtistName());
//                cell2.setCellStyle(getCellStyle(sheet, startColIndex++, startRowIndex));
//
//                Cell cell3 = dataRow.createCell(startColIndex);
//                cell3.setCellValue(artist.getDateOfBirth());
//                cell3.setCellStyle(getCellStyle(sheet, startColIndex++, startRowIndex));
//
//                Cell cell4 = dataRow.createCell(startColIndex);
//                //  String newFormula = copyFormula(sheet, formula, 0, i+1);
//                cell4.setCellValue(formula);
//                cell4.setCellStyle(getCellStyle(sheet, startColIndex, startRowIndex));

            }


            try (FileOutputStream outputStream = new FileOutputStream(filePath)) {
                workbook.write(outputStream);
            }
            System.out.println("Data exported");

        } catch (Exception e) {
            System.out.println("Error exporting data: " + e.getMessage());
        }
    }


    private XSSFCellStyle getCellStyleFromTemplate(XSSFSheet sheet, int startColIndex, int startRowIndex) {
        return sheet.getRow(startRowIndex + 1).getCell(startColIndex).getCellStyle();
    }

    private static String copyFormula(XSSFSheet sheet, String formula, int colDiff, int rowDiff) {

        XSSFEvaluationWorkbook workbookWrapper =
                XSSFEvaluationWorkbook.create(sheet.getWorkbook());
        Ptg[] ptgs = FormulaParser.parse(formula, workbookWrapper, FormulaType.CELL
                , sheet.getWorkbook().getSheetIndex(sheet));

        for (int i = 0; i < ptgs.length; i++) {
            if (ptgs[i] instanceof RefPtgBase) { // base class for cell references
                RefPtgBase ref = (RefPtgBase) ptgs[i];
                if (ref.isColRelative())
                    ref.setColumn(ref.getColumn() + colDiff);
                if (ref.isRowRelative())
                    ref.setRow(ref.getRow() + rowDiff);
            } else if (ptgs[i] instanceof AreaPtgBase) { // base class for range references
                AreaPtgBase ref = (AreaPtgBase) ptgs[i];
                if (ref.isFirstColRelative())
                    ref.setFirstColumn(ref.getFirstColumn() + colDiff);
                if (ref.isLastColRelative())
                    ref.setLastColumn(ref.getLastColumn() + colDiff);
                if (ref.isFirstRowRelative())
                    ref.setFirstRow(ref.getFirstRow() + rowDiff);
                if (ref.isLastRowRelative())
                    ref.setLastRow(ref.getLastRow() + rowDiff);
            }
        }

        formula = FormulaRenderer.toFormulaString(workbookWrapper, ptgs);
        return formula;
    }


}
