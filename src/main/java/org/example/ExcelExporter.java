package org.example;

import com.nhl.dflib.DataFrame;
import com.nhl.dflib.Series;
import com.nhl.dflib.row.RowProxy;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
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
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFTable;
import org.apache.poi.xssf.usermodel.XSSFTableColumn;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

public class ExcelExporter {
    public ExcelExporter() {
    }

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

                List<RowProxy> rowProxies = new ArrayList<>();
                dataFrame.iterator().forEachRemaining(rowProxies::add);


                Iterator<RowProxy> iterator = dataFrame.iterator();

//                while (dataFrame.iterator().hasNext()) {
//                    RowProxy next = iterator.next();
//
//
//                }
                List<XSSFTableColumn> columns = table.getColumns();
                int startColIndex = table.getStartColIndex();
                int startRowIndex = table.getStartRowIndex();

                Map<String, XSSFCellStyle> columnStylesMap = getCellStyleMap(sheet, columns, startColIndex, startRowIndex);


                table.setDataRowCount(rowProxies.size()+1);

                List<String> frameColumnsNames = getFrameColumnsNames(dataFrame);

                for (int i = 0; i < rowProxies.size(); i++) {
                    Row dataRow = sheet.createRow(sheet.getLastRowNum() + 1);
                    for (XSSFTableColumn tableColumn : columns) {
                        Cell cell = dataRow.createCell(table.getStartColIndex() + tableColumn.getColumnIndex());
                        if (frameColumnsNames.contains(tableColumn.getName())) {
                            processMappedColumn(dataFrame, i, tableColumn, cell);
                        } else {
                            processUnmappedColumn(sheet, startColIndex, startRowIndex, tableColumn, cell);
                        }
                        cell.setCellStyle(columnStylesMap.get(tableColumn.getName()));
                    }
                }
              //  sheet.shiftRows(startRowIndex , lastRowNum, -1);
            }

          //  removeRow(sheet, 1);

            try (FileOutputStream outputStream = new FileOutputStream(filePath)) {
                workbook.write(outputStream);
            }
            System.out.println("Data exported");

        } catch (Exception e) {
            System.out.println("Error exporting data: " + e.getMessage());
        }
    }


    public void removeRow(XSSFSheet sheet, int rowIndex) {
        int lastRowNum = sheet.getLastRowNum();
        if (rowIndex >= 0 && rowIndex < lastRowNum) {
            sheet.shiftRows(rowIndex + 1, lastRowNum, -1);
        }
        if (rowIndex == lastRowNum) {
            XSSFRow removingRow = sheet.getRow(rowIndex);
            if (removingRow != null) {
                sheet.removeRow(removingRow);
            }
        }
    }

    private void processMappedColumn(DataFrame dataFrame, int i, XSSFTableColumn tableColumn, Cell cell) {
        Series<Object> frameColumnValues = dataFrame.getColumn(tableColumn.getName());
        cell.setCellValue(frameColumnValues.get(i).toString());
    }

    private void processUnmappedColumn(XSSFSheet sheet, int startColIndex, int startRowIndex, XSSFTableColumn tableColumn, Cell cell) {
        XSSFCell formulaCandidateCell = sheet.getRow(startRowIndex + 1).getCell(startColIndex + tableColumn.getColumnIndex());
        if (formulaCandidateCell.getCellType().equals(CellType.FORMULA)) {
            cell.setCellFormula(formulaCandidateCell.getCellFormula());
        } else {
            cell.setCellValue("Default");
        }
    }


    private List<String> getFrameColumnsNames(DataFrame dataFrame) {
        return Arrays.asList(dataFrame.getColumnsIndex().getLabels());
    }

    private Map<String, XSSFCellStyle> getCellStyleMap(XSSFSheet sheet, List<XSSFTableColumn> columns, int startColIndex, int startRowIndex) {
        Map<String, XSSFCellStyle> columnStylesMap = new HashMap<>();

        for (XSSFTableColumn column : columns) {
            XSSFCell templateStyleCell = sheet.getRow(startRowIndex + 1).getCell(startColIndex + column.getColumnIndex());
            XSSFCellStyle cellStyle = templateStyleCell.getCellStyle();
            columnStylesMap.put(column.getName(), cellStyle);
        }
        return columnStylesMap;
    }


    private String copyFormula(XSSFSheet sheet, String formula, int colDiff, int rowDiff) {

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
