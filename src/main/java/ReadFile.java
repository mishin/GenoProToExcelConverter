/**
 * Created by Nikolay Mishin on 20.02.2017.
 * Read Exlx file
 */

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;

class ReadXlsxFile {
    public static void main(String[] args) {
        ReadXlsxFile xread2 = new ReadXlsxFile();
        xread2.readFile();
    }

    private ReadXlsxFile() {
    }

    private void readFile() {
        String fileName = "c:\\Users\\Mishin737\\Documents\\work_line\\20022017\\readXlsxInJava-master\\mishin_family.xlsx";
        String fileForWrite = "c:\\Users\\Mishin737\\Documents\\work_line\\20022017\\readXlsxInJava-master\\padigree.xlsx";
        try {
            Sheet sheet = getSheet(fileName);
            List<Padigree> xlsxData = readSheetPadigree(sheet);
            writePadigreeListToExcel(xlsxData, fileForWrite);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private List<Padigree> readSheetPadigree(Sheet sheet) {
        /**
         * All rows is 603
         * All columns is 23
         * father's ID	mother's ID
         * 8            9
         * */
        HashMap<String, String> oldVsNewCode = new HashMap<>();
//        int numberOfColumns = 23;
        int numberOfRows = 526;
//        String[][] data = new String[numberOfColumns][numberOfRows];
        List<Padigree> padigreeList = new ArrayList<>();

        for (int j = 1; j < numberOfRows; j++) {
            Row row1 = sheet.getRow(j);
            String newCode = String.format("ind%06d", j);
            String oldCode = ReadCell(row1.getCell(0));
            System.out.println(String.format("old: %s: , new: ind%s", oldCode, newCode));
            oldVsNewCode.put(oldCode, newCode);
        }

        for (int j = 1; j < numberOfRows; j++) {
            Padigree padigree = setPadigree(sheet, j, oldVsNewCode);
            padigreeList.add(padigree);
        }
        return padigreeList;
    }

    private Padigree setPadigree(Sheet sheet, int j, HashMap<String, String> oldVsNewCode) {
        Padigree padigree = new Padigree();
        padigree.setID(oldVsNewCode.get(readCell(sheet, j, 0)));
        padigree.setTitle(readCell(sheet, j, 1));
        padigree.setPrefix(readCell(sheet, j, 2));
        padigree.setFirstName(readCell(sheet, j, 3));
        padigree.setMidName(readCell(sheet, j, 4));
        padigree.setLastName(readCell(sheet, j, 5));
        padigree.setSuffix(readCell(sheet, j, 6));
        padigree.setNickname(readCell(sheet, j, 7));
        padigree.setFatherId(oldVsNewCode.get(readCell(sheet, j, 8)));
        padigree.setMotherId(oldVsNewCode.get(readCell(sheet, j, 9)));
        padigree.setEmail(readCell(sheet, j, 10));
        padigree.setWebpage(readCell(sheet, j, 11));
        padigree.setDateOfBirth(readCell(sheet, j, 12));
        padigree.setDateOfDeath(readCell(sheet, j, 13));
        padigree.setGender(readCell(sheet, j, 14));
        padigree.setIsLiving(readCell(sheet, j, 15));
        padigree.setPlaceOfBirth(readCell(sheet, j, 16));
        padigree.setPlaceOfDdeath(readCell(sheet, j, 17));
        padigree.setCemetery(readCell(sheet, j, 18));
        padigree.setSchools(readCell(sheet, j, 19));
        padigree.setJobs(readCell(sheet, j, 20));
        padigree.setWorkPlaces(readCell(sheet, j, 21));
        padigree.setPlacesOfLiving(readCell(sheet, j, 22));
        padigree.setGeneral(readCell(sheet, j, 23));
        return padigree;
    }

    private static String readCell(Sheet sheet, int j, int i) {
        return ReadCell(sheet.getRow(j).getCell(i));
    }

    private static void writePadigreeListToExcel(List<Padigree> padigreeList, String fileForWrite) {
        // Using XSSF for xlsx format, for xls use HSSF
        Workbook workbook = new XSSFWorkbook();
        Sheet padigreeSheet = workbook.createSheet("Padigree");

        int rowIndex = 0;
        for (Padigree padigree : padigreeList) {
            Row row = padigreeSheet.createRow(rowIndex++);
            int cellIndex = 0;
            row.createCell(cellIndex++).setCellValue(padigree.getID());
            row.createCell(cellIndex++).setCellValue(padigree.getTitle());
            row.createCell(cellIndex++).setCellValue(padigree.getPrefix());
            row.createCell(cellIndex++).setCellValue(padigree.getFirstName());
            row.createCell(cellIndex++).setCellValue(padigree.getMidName());
            row.createCell(cellIndex++).setCellValue(padigree.getLastName());
            row.createCell(cellIndex++).setCellValue(padigree.getSuffix());
            row.createCell(cellIndex++).setCellValue(padigree.getNickname());
            row.createCell(cellIndex++).setCellValue(padigree.getFatherId());
            row.createCell(cellIndex++).setCellValue(padigree.getMotherId());
            row.createCell(cellIndex++).setCellValue(padigree.getEmail());
            row.createCell(cellIndex++).setCellValue(padigree.getWebpage());
            row.createCell(cellIndex++).setCellValue(padigree.getDateOfBirth());
            row.createCell(cellIndex++).setCellValue(padigree.getDateOfDeath());
            row.createCell(cellIndex++).setCellValue(padigree.getGender());
            row.createCell(cellIndex++).setCellValue(padigree.getIsLiving());
            row.createCell(cellIndex++).setCellValue(padigree.getPlaceOfBirth());
            row.createCell(cellIndex++).setCellValue(padigree.getPlaceOfDdeath());
            row.createCell(cellIndex++).setCellValue(padigree.getCemetery());
            row.createCell(cellIndex++).setCellValue(padigree.getSchools());
            row.createCell(cellIndex++).setCellValue(padigree.getJobs());
            row.createCell(cellIndex++).setCellValue(padigree.getWorkPlaces());
            row.createCell(cellIndex++).setCellValue(padigree.getPlacesOfLiving());
            row.createCell(cellIndex++).setCellValue(padigree.getGeneral());

        }

        //write this workbook in excel file.
        try {
            FileOutputStream fos = new FileOutputStream(fileForWrite);
            workbook.write(fos);
            fos.close();

            System.out.println(fileForWrite + " is successfully written");
        } catch (IOException e) {
            e.printStackTrace();
        }


    }

    private static String ReadCell(Cell cell) {
        if (cell != null) {
            switch (cell.getCellType()) {
                case Cell.CELL_TYPE_NUMERIC:
                    return String.valueOf(cell.getNumericCellValue());
                case Cell.CELL_TYPE_STRING:
                    return cell.getStringCellValue();
                case Cell.CELL_TYPE_FORMULA:
                    break;
                case Cell.CELL_TYPE_BLANK:
                    break;
            }
        }
        return null;
    }

    private static Sheet getSheet(String fileName) throws IOException {
        File file = new File(fileName);
        Workbook workbook = null;
        FileInputStream fileInputStream = new FileInputStream(file);
        String fileExtension = fileName.substring(fileName.indexOf("."));
//        System.out.println(fileExtension);
        switch (fileExtension) {
            case ".xls":
                workbook = new HSSFWorkbook(new POIFSFileSystem(fileInputStream));
                break;
            case ".xlsx":
                workbook = new XSSFWorkbook(fileInputStream);
                break;
            default:
                System.out.println("Wrong File Type");
                break;
        }
        Sheet sheet;
        if (workbook != null) sheet = workbook.getSheetAt(0);
        else sheet = null;
        return sheet;
    }
}
