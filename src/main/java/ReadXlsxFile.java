/**
 * Created by Nikolay Mishin on 20.02.2017.
 * Read Excel file
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
        ReadXlsxFile xlsxRead = new ReadXlsxFile();
        xlsxRead.readFile();
    }

    private ReadXlsxFile() {
    }

    private void readFile() {
        String fileName = "c:\\Users\\Mishin737\\Documents\\work_line\\20022017\\readXlsxInJava-master\\mishin_family.xlsx";
        String fileForWrite = "c:\\Users\\Mishin737\\Documents\\work_line\\20022017\\readXlsxInJava-master\\pedigree.xlsx";
        try {
            Sheet sheet = getSheet(fileName);
            List<Pedigree> xlsxData = readSheetPedigree(sheet);
            writePedigreeListToExcel(xlsxData, fileForWrite);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private List<Pedigree> readSheetPedigree(Sheet sheet) {
        /**
         * All rows is 603
         * All columns is 23
         * father's ID	mother's ID
         * 8            9
         * */
        HashMap<String, String> oldVsNewCode = new HashMap<>();
        HashMap<String, String> familyCodeMap = new HashMap<>();
//        int numberOfColumns = 23;
        int numberOfRows = 526;
//        String[][] data = new String[numberOfColumns][numberOfRows];
        List<Pedigree> pedigreeList = new ArrayList<>();
//        HashSet<String> uniqueValues = new HashSet<>(values);
        for (int j = 1; j < numberOfRows; j++) {
            Row row1 = sheet.getRow(j);
            String newCode = String.format("ind%06d", j);
            String oldCode = ReadCell(row1.getCell(0));
            System.out.println(String.format("old: %s: , new: ind%s", oldCode, newCode));
            oldVsNewCode.put(oldCode, newCode);
        }

        for (int j = 1; j < numberOfRows; j++) {
            Pedigree pedigree = setPedigree(sheet, j, oldVsNewCode);
            pedigreeList.add(pedigree);
        }

        int z = 0;
        for (Pedigree pedigree : pedigreeList) {
            if (pedigree.getFatherId() != null && pedigree.getMotherId() != null) {
                String familyString = getFamilyString(pedigree);
                if (!familyCodeMap.containsKey(familyString)) {
                    String famCode = String.format("fam%06d", ++z);
                    familyCodeMap.put(familyString, famCode);
                }
                pedigree.setFamily(familyCodeMap.get(familyString));
            }
        }

        return pedigreeList;
    }

    private String getFamilyString(Pedigree pedigree) {
        return pedigree.getFatherId() + "," + pedigree.getMotherId();
    }

    private Pedigree setPedigree(Sheet sheet, int j, HashMap<String, String> oldVsNewCode) {
        Pedigree pedigree = new Pedigree();
        pedigree.setID(oldVsNewCode.get(readCell(sheet, j, 0)));
        pedigree.setTitle(readCell(sheet, j, 1));
        pedigree.setPrefix(readCell(sheet, j, 2));
        pedigree.setFirstName(readCell(sheet, j, 3));
        pedigree.setMidName(readCell(sheet, j, 4));
        pedigree.setLastName(readCell(sheet, j, 5));
        pedigree.setSuffix(readCell(sheet, j, 6));
        pedigree.setNickname(readCell(sheet, j, 7));
        pedigree.setFatherId(oldVsNewCode.get(readCell(sheet, j, 8)));
        pedigree.setMotherId(oldVsNewCode.get(readCell(sheet, j, 9)));
        pedigree.setEmail(readCell(sheet, j, 10));
        pedigree.setWebPage(readCell(sheet, j, 11));
        pedigree.setDateOfBirth(readCell(sheet, j, 12));
        pedigree.setDateOfDeath(readCell(sheet, j, 13));
        pedigree.setGender(readCell(sheet, j, 14));
        pedigree.setIsLiving(readCell(sheet, j, 15));
        pedigree.setPlaceOfBirth(readCell(sheet, j, 16));
        pedigree.setPlaceOfDeath(readCell(sheet, j, 17));
        pedigree.setCemetery(readCell(sheet, j, 18));
        pedigree.setSchools(readCell(sheet, j, 19));
        pedigree.setJobs(readCell(sheet, j, 20));
        pedigree.setWorkPlaces(readCell(sheet, j, 21));
        pedigree.setPlacesOfLiving(readCell(sheet, j, 22));
        pedigree.setGeneral(readCell(sheet, j, 23));
        return pedigree;
    }

    private static String readCell(Sheet sheet, int j, int i) {
        return ReadCell(sheet.getRow(j).getCell(i));
    }

    private static void writePedigreeListToExcel(List<Pedigree> pedigreeList, String fileForWrite) {
        Workbook workbook = new XSSFWorkbook();
        Sheet pedigreeSheet = workbook.createSheet("Pedigree");

        int rowIndex = 0;
        for (Pedigree pedigree : pedigreeList) {
            Row row = pedigreeSheet.createRow(rowIndex++);
            int cellIndex = 0;
            row.createCell(cellIndex++).setCellValue(pedigree.getID());
            row.createCell(cellIndex++).setCellValue(pedigree.getTitle());
            row.createCell(cellIndex++).setCellValue(pedigree.getPrefix());
            row.createCell(cellIndex++).setCellValue(pedigree.getFirstName());
            row.createCell(cellIndex++).setCellValue(pedigree.getMidName());
            row.createCell(cellIndex++).setCellValue(pedigree.getLastName());
            row.createCell(cellIndex++).setCellValue(pedigree.getSuffix());
            row.createCell(cellIndex++).setCellValue(pedigree.getNickname());
            row.createCell(cellIndex++).setCellValue(pedigree.getFatherId());
            row.createCell(cellIndex++).setCellValue(pedigree.getMotherId());
            row.createCell(cellIndex++).setCellValue(pedigree.getEmail());
            row.createCell(cellIndex++).setCellValue(pedigree.getWebPage());
            row.createCell(cellIndex++).setCellValue(pedigree.getDateOfBirth());
            row.createCell(cellIndex++).setCellValue(pedigree.getDateOfDeath());
            row.createCell(cellIndex++).setCellValue(pedigree.getGender());
            row.createCell(cellIndex++).setCellValue(pedigree.getIsLiving());
            row.createCell(cellIndex++).setCellValue(pedigree.getPlaceOfBirth());
            row.createCell(cellIndex++).setCellValue(pedigree.getPlaceOfDeath());
            row.createCell(cellIndex++).setCellValue(pedigree.getCemetery());
            row.createCell(cellIndex++).setCellValue(pedigree.getSchools());
            row.createCell(cellIndex++).setCellValue(pedigree.getJobs());
            row.createCell(cellIndex++).setCellValue(pedigree.getWorkPlaces());
            row.createCell(cellIndex++).setCellValue(pedigree.getPlacesOfLiving());
            row.createCell(cellIndex++).setCellValue(pedigree.getGeneral());
            //noinspection UnusedAssignment,UnusedAssignment,UnusedAssignment
            row.createCell(cellIndex++).setCellValue(pedigree.getFamily());

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
