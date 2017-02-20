/**
 * Created by Nikolay Mishin on 20.02.2017.
 * Read Exlx file
 */

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

class ReadFile {
    public static void main(String[] args) {
        String fileName = "c:\\Users\\ira\\Documents\\генеалогия\\github\\mishin_family.xlsx";
        File file = new File(fileName);
        FileInputStream fileInputStream;
        Workbook workbook = null;
        Sheet sheet;
        Iterator<Row> rowIterator;
        try {
            workbook = getWorkbook(fileName, file, workbook);
            FormulaEvaluator evaluator;
            if (workbook != null) evaluator = workbook.getCreationHelper().createFormulaEvaluator();
            else evaluator = null;
            if (workbook != null) sheet = workbook.getSheetAt(0);
            else sheet = null;
            rowIterator = (sheet != null) ? sheet.iterator() : null;

//            for (int i = 0; i < 23; i++) {
//                System.out.println(row1.getCell(i));
//            }
            for (int j = 1; j < 603; j++) {
                Row row1 = sheet.getRow(j);
                System.out.println("old: " + row1.getCell(0) + " :" + j + " new: ind000" + j);
            }


//            ReadRows(rowIterator, evaluator);
//System.out.println(sheet);
        } catch (FileNotFoundException e) {
// TODO Auto-generated catch block
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void ReadRows(Iterator<Row> rowIterator, FormulaEvaluator evaluator) {
        while (rowIterator != null && rowIterator.hasNext()) {
            Row row = rowIterator.next();
            Iterator<Cell> cellIterator = row.cellIterator();
//                row.getCell();
            ShowCells(evaluator, cellIterator);
            System.out.println("\n");
        }
    }

    private static void ShowCells(FormulaEvaluator evaluator, Iterator<Cell> cellIterator) {
        while (cellIterator.hasNext()) {
            Cell cell = cellIterator.next();
            //Check the cell type after evaluating formulae
            //If it is formula cell, it will be evaluated otherwise no change will happen
            switch (evaluator != null ? evaluator.evaluateInCell(cell).getCellType() : 0) {
                case Cell.CELL_TYPE_NUMERIC:
                    System.out.print(cell.getNumericCellValue() + " ");
                    break;
                case Cell.CELL_TYPE_STRING:
                    System.out.print(cell.getStringCellValue() + " ");
                    break;
                case Cell.CELL_TYPE_FORMULA:
//                                Not again
                    break;
                case Cell.CELL_TYPE_BLANK:
                    break;
            }
        }
    }

    private static Workbook getWorkbook(String fileName, File file, Workbook workbook) throws IOException {
        FileInputStream fileInputStream;
        fileInputStream = new FileInputStream(file);
        String fileExtension = fileName.substring(fileName.indexOf("."));
//        System.out.println(fileExtension);
        if (fileExtension.equals(".xls")) {
            workbook = new HSSFWorkbook(new POIFSFileSystem(fileInputStream));
        } else if (fileExtension.equals(".xlsx")) {
            workbook = new XSSFWorkbook(fileInputStream);
        } else {
            System.out.println("Wrong File Type");
        }
        return workbook;
    }
}
