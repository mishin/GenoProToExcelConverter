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
    public static void main(String[] args) throws IOException {

        String fileName = "c:\\Users\\ira\\Downloads\\mishin_family.xlsx";
        File file = new File(fileName);
        FileInputStream fileInputStream;
        Workbook workbook = null;
        Sheet sheet;
        Iterator<Row> rowIterator;
        try {
            fileInputStream = new FileInputStream(file);
            String fileExtension = fileName.substring(fileName.indexOf("."));
            System.out.println(fileExtension);
            if (fileExtension.equals(".xls")) {
                workbook = new HSSFWorkbook(new POIFSFileSystem(fileInputStream));
            } else if (fileExtension.equals(".xlsx")) {
                workbook = new XSSFWorkbook(fileInputStream);
            } else {
                System.out.println("Wrong File Type");
            }
            FormulaEvaluator evaluator;
            if (workbook != null) evaluator = workbook.getCreationHelper().createFormulaEvaluator();
            else evaluator = null;
            if (workbook != null) sheet = workbook.getSheetAt(0);
            else sheet = null;
            rowIterator = (sheet != null) ? sheet.iterator() : null;
            while (rowIterator != null && rowIterator.hasNext()) {
                Row row = rowIterator.next();
                Iterator<Cell> cellIterator = row.cellIterator();
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
                System.out.println("\n");
            }
//System.out.println(sheet);
        } catch (FileNotFoundException e) {
// TODO Auto-generated catch block
            e.printStackTrace();
        }
/*        catch (IOException e) {
            e.printStackTrace();
        }​*/

/*            FileInputStream file = new FileInputStream(
                    new File("c:\\Users\\ira\\Downloads\\mishin_family.xls.xlsx")
            );*/
/*
            Workbook[] wbs = new Workbook[] { new HSSFWorkbook(), new XSSFWorkbook() };
            for(int i=0; i<wbs.length; i++) {
                Workbook wb = wbs[i];
                CreationHelper createHelper = wb.getCreationHelper();

                // create a new sheet
                Sheet s = wb.createSheet();
                // declare a row object reference
                Row r = null;
                // declare a cell object reference
                Cell c = null;
                // create 2 cell styles
                CellStyle cs = wb.createCellStyle();
                CellStyle cs2 = wb.createCellStyle();
                DataFormat df = wb.createDataFormat();

                // create 2 fonts objects
                Font f = wb.createFont();
                Font f2 = wb.createFont();

                // Set font 1 to 12 point type, blue and bold
                f.setFontHeightInPoints((short) 12);
                f.setColor( IndexedColors.RED.getIndex() );
                f.setBoldweight(Font.BOLDWEIGHT_BOLD);

                // Set font 2 to 10 point type, red and bold
                f2.setFontHeightInPoints((short) 10);
                f2.setColor( IndexedColors.RED.getIndex() );
                f2.setBoldweight(Font.BOLDWEIGHT_BOLD);

                // Set cell style and formatting
                cs.setFont(f);
                cs.setDataFormat(df.getFormat("#,##0.0"));

                // Set the other cell style and formatting
                cs2.setBorderBottom(cs2.BORDER_THIN);
                cs2.setDataFormat(df.getFormat("text"));
                cs2.setFont(f2);


                // Define a few rows
                for(int rownum = 0; rownum < 30; rownum++) {
                    Row r = s.createRow(rownum);
                    for(int cellnum = 0; cellnum < 10; cellnum += 2) {
                        Cell c = r.createCell(cellnum);
                        Cell c2 = r.createCell(cellnum+1);

                        c.setCellValue((double)rownum + (cellnum/10));
                        c2.setCellValue(
                                createHelper.createRichTextString("Hello! " + cellnum)
                        );
                    }
                }

                // Save
                String filename = "workbook.xls";
                if(wb instanceof XSSFWorkbook) {
                    filename = filename + "x";
                }

                FileOutputStream out = new FileOutputStream(filename);
                wb.write(out);
                out.close();
            }

            Workbook wb = WorkbookFactory.create(
                    new File("c:\\Users\\ira\\Downloads\\mishin_family.xls.xlsx")
            );*/
        /*    Sheet mySheet = wb.getSheetAt(0);
            Iterator<Row> rowIter = mySheet.rowIterator();
            System.out.println(mySheet.getRow(1).getCell(0));

            //Get the workbook instance for XLS file
            XSSFWorkbook workbook = new XSSFWorkbook(file);
            OPCPackage pkg = OPCPackage.open(new File("file.xlsx"));

            //Get first sheet from the workbook
            XSSFSheet sheet = workbook.getSheetAt(0);

            //Iterate through each rows from first sheet
            Iterator<Row> rowIterator = sheet.iterator();
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();

                //For each row, iterate through each columns
                Iterator<Cell> cellIterator = row.cellIterator();
                while (cellIterator.hasNext()) {

                    Cell cell = cellIterator.next();

                    switch (cell.getCellType()) {
                        case Cell.CELL_TYPE_BOOLEAN:
                            System.out.print(cell.getBooleanCellValue() + "\t\t");
                            break;
                        case Cell.CELL_TYPE_NUMERIC:
                            System.out.print(cell.getNumericCellValue() + "\t\t");
                            break;
                        case Cell.CELL_TYPE_STRING:
                            System.out.print(cell.getStringCellValue() + "\t\t");
                            break;
                    }
                }
                System.out.println("");
            }
            file.close();
            FileOutputStream out =
                    new FileOutputStream(new File("C:\\test.xls"));
            workbook.write(out);
            out.close();

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }*/
    }
}
