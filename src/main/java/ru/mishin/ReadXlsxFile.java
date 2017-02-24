package ru.mishin;
/**
 * Created by Nikolay Mishin on 20.02.2017.
 * Read Excel file
 */

import freemarker.template.Configuration;
import freemarker.template.Template;
import freemarker.template.TemplateException;
import freemarker.template.TemplateExceptionHandler;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.logging.Logger;

import static java.util.Locale.ENGLISH;

class ReadXlsxFile {

    private  List<String> familyCode =  new ArrayList<>();
    private final static Logger log = Logger.getLogger(String.valueOf(ReadXlsxFile.class));

    public static void main(String[] args) {
        ReadXlsxFile xlsxRead = new ReadXlsxFile();
        xlsxRead.readFile();
    }

    private ReadXlsxFile() {
    }

    private void readFile() {
        Properties prop = readProperties();
        String root = prop.getProperty("root");//"c:\\Users\\ira\\Documents\\генеалогия\\github\\";
        String fileName = root + prop.getProperty("readFile");//"mishin_family.xlsx";
        System.out.println("fileName: " + fileName);
        String fileForWrite = root + prop.getProperty("writeFile");//"pedigree.xlsx";
        try {
            Sheet sheet = getSheet(fileName);
            List<Pedigree> xlsxData = readSheetPedigree(sheet);
            writePedigreeListToExcel(xlsxData, fileForWrite);
            fillTemplate(prop, xlsxData);
        } catch (IOException | TemplateException e) {
            e.printStackTrace();
        }
    }

    private Properties readProperties() {
        Properties prop = new Properties();
        InputStream input = null;

        try {
            input = new FileInputStream("src/main/resources/config.properties");
            InputStreamReader isr = new InputStreamReader(input, "UTF-8");
            prop.load(isr);
            // load a properties file
            prop.load(input);
        } catch (IOException ex) {
            ex.printStackTrace();
        } finally {
            if (input != null) {
                try {
                    input.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
        return prop;
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
        int numberOfRows = 603;//526;
        List<Pedigree> pedigreeList = new ArrayList<>();
        for (int j = 1; j < numberOfRows; j++) {
            Row row = sheet.getRow(j);
            String newCode = String.format("ind%06d", j);
            String oldCode = readCell(row.getCell(0));
//            System.out.println(String.format("mv %s.jpg %s.jpg", oldCode.toLowerCase(), newCode.toLowerCase()));
            oldVsNewCode.put(oldCode, newCode);
        }

        for (int j = 1; j < numberOfRows; j++) {
            Pedigree pedigree = setPedigree(sheet, j, oldVsNewCode);
            pedigreeList.add(pedigree);
        }

        int z = 0;
        for (Pedigree pedigree : pedigreeList) {
            String familyString = getFamilyString(pedigree);
            if (!familyCodeMap.containsKey(familyString)) {
                String famCode = String.format("fam%06d", ++z);
                familyCodeMap.put(familyString, famCode);
            }
            pedigree.setFamilyId(familyCodeMap.get(familyString));
        }

        return pedigreeList;
    }

    private String getFamilyString(Pedigree pedigree) {
        if (pedigree.getFatherId() != null || pedigree.getMotherId() != null) {
            return pedigree.getFatherId() + "," + pedigree.getMotherId();
        } else {
            return pedigree.getID();
        }
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
        String dateOfBirth = readCell(sheet, j, 12);
        if (dateOfBirth != null) {
            DateTimeFormatter formatter = DateTimeFormatter.ofPattern("d,MM,yyyy");
            LocalDate localDate = LocalDate.parse(dateOfBirth, formatter);
            pedigree.setDateOfBirth(localDate.format(DateTimeFormatter.ofPattern("d MMM yyyy", ENGLISH)));
        } else {
            pedigree.setDateOfBirth("");
        }
        pedigree.setDateOfDeath(readCell(sheet, j, 13));
        String gender = readCell(sheet, j, 14);
        switch (gender) {
            case "1.0":
                pedigree.setGender("F");
                break;
            case "0.0":
                pedigree.setGender("M");
                break;
            default:
                pedigree.setGender("M");
                break;
        }
        ;
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
        return readCell(sheet.getRow(j).getCell(i));
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
            row.createCell(cellIndex++).setCellValue(pedigree.getFamilyId());

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

    void fillTemplate(Properties prop, List<Pedigree> pedigreeList) throws IOException, TemplateException {
        Configuration cfg = getConfiguration(prop);


         /* Get the template (uses cache internally) */
        Template template = null;
        try {
            template = cfg.getTemplate(prop.getProperty("patternMap"));
        } catch (IOException e) {
            e.printStackTrace();
        }

//        makeExampleTemplate(cfg);
        List<PedigreeLink> pedigreeLinkList = getPedigreeLinks(pedigreeList);

//        System.out.println(pedigreeLinkList);
        // Build the data-model
        Map<String, Object> data = new HashMap<>();
        data.put("individuals", pedigreeList);
        data.put("pedigreelinks", pedigreeLinkList);

/*        List<String> familyCode = null;
        familyCode.add("fam000001");
        //Java 8 only, forEach and Lambda
        for (Pedigree pedigree : pedigreeList) {
            String family = pedigree.getFamilyId();
//            if (family != null && !familyCode.contains(family)) {
                System.out.println(family);
//                familyCode.add(k.getFamilyId());
//            }
        }*/
            /*if (set != null)
                familyCode.add(set);
        pedigreeLinkList.forEach((k) -> {
            if (!familyCode.contains(k.getFamilyId())) {
                System.out.println(k.getFamilyId());
//                familyCode.add(k.getFamilyId());
            }
            ;
        });*/

        data.put("families", familyCode);

        // Console output
        Writer out = new OutputStreamWriter(System.out);
        try {
            template.process(data, out);
        } catch (TemplateException | IOException e) {
            e.printStackTrace();
        }
        try {
            out.flush();
        } catch (IOException e) {
            e.printStackTrace();
        }

    }

    private List<PedigreeLink> getPedigreeLinks(List<Pedigree> pedigreeList) {
        List<PedigreeLink> pedigreeLinkList = new ArrayList<>();

        /*<PedigreeLink PedigreeLink="Parent" Family="fam00002" Individual="ind00004"/>
        <PedigreeLink PedigreeLink="Parent" Family="fam00002" Individual="ind00005"/>
        <PedigreeLink PedigreeLink="Biological" Family="fam00002" Individual="ind00006"/>*/

//        List<String> familyCode =  new ArrayList<>();

        for (Pedigree pedigree : pedigreeList) {
            familyCode.add(pedigree.getFamilyId());
            if (pedigree.getID() != null) {
                PedigreeLink pedigreeLink = new PedigreeLink();
                pedigreeLink.setParentOrChild("Biological");
                pedigreeLink.setFamilyId(pedigree.getFamilyId());
                pedigreeLink.setIndividualId(pedigree.getID());
                pedigreeLinkList.add(pedigreeLink);
            }
            if (pedigree.getMotherId() != null) {
                PedigreeLink pedigreeLink = new PedigreeLink();
                pedigreeLink.setParentOrChild("Parent");
                pedigreeLink.setFamilyId(pedigree.getFamilyId());
                pedigreeLink.setIndividualId(pedigree.getMotherId());
                pedigreeLinkList.add(pedigreeLink);
            }
            if (pedigree.getFatherId() != null) {
                PedigreeLink pedigreeLink = new PedigreeLink();
                pedigreeLink.setParentOrChild("Parent");
                pedigreeLink.setFamilyId(pedigree.getFamilyId());
                pedigreeLink.setIndividualId(pedigree.getFatherId());
                pedigreeLinkList.add(pedigreeLink);
            }
        }
        return pedigreeLinkList;
    }

    private void makeExampleTemplate(Configuration cfg) throws IOException, TemplateException {
        // Create the root hash. We use a Map here, but it could be a JavaBean too.
        Map<String, Object> root = new HashMap<>();

// Put string "user" into the root
        root.put("user", "Big Joe");

// Create the "latestProduct" hash. We use a JavaBean here, but it could be a Map too.
        Product latest = new Product();
        latest.setUrl("products/greenmouse.html");
        latest.setName("green mouse");
// and put it into the root
        root.put("latestProduct", latest);

        Template temp = cfg.getTemplate("test.ftlh");

        Writer out = new OutputStreamWriter(System.out);
        temp.process(root, out);
    }

    private Configuration getConfiguration(Properties prop) {
    /* ------------------------------------------------------------------------ */
        /* You should do this ONLY ONCE in the whole application life-cycle:        */

        /* Create and adjust the configuration singleton */
        Configuration cfg = new Configuration(Configuration.VERSION_2_3_25);
        try {
            cfg.setDirectoryForTemplateLoading(new File(prop.getProperty("inPatternDir")));
        } catch (IOException e) {
            e.printStackTrace();
        }

        cfg.setDefaultEncoding("UTF-8");
        cfg.setTemplateExceptionHandler(TemplateExceptionHandler.RETHROW_HANDLER);
        cfg.setLogTemplateExceptions(false);

 /* ------------------------------------------------------------------------ */
        /* You usually do these for MULTIPLE TIMES in the application life-cycle:   */
        return cfg;
    }

    private static String readCell(Cell cell) {
        if (cell != null) {
            switch (cell.getCellType()) {
                case Cell.CELL_TYPE_NUMERIC:
                    return String.valueOf(cell.getNumericCellValue());
                case Cell.CELL_TYPE_STRING:
                    return cell.getStringCellValue().trim();
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
