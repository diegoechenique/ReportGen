/*
* To change this license header, choose License Headers in Project Properties.
* To change this template file, choose Tools | Templates
* and open the template in the editor.
*/
package facade;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.GregorianCalendar;
import java.util.HashSet;
import java.util.Iterator;
import java.util.List;
import java.util.Locale;
import java.util.Set;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import res.Strings;
import vo.RefRecord;
import vo.Table8Line;

/**
 *
 * @author diego
 */
public class PoiHelper {
    
    private File psw;
    private File graphs;
    private ArrayList<RefRecord> recordList;
    private XSSFWorkbook pswWorkbook;
    private XSSFWorkbook graphsWorkbook;
    
    
    
    public PoiHelper(File psw, File graphs) {
        
        this.psw = psw;
        this.graphs = graphs;
        loadWorkbooks();
        
    }
    
    public void init(){
        getRefRecords();
    }
    
    private void loadWorkbooks(){
        try {
            this.pswWorkbook = new XSSFWorkbook(new FileInputStream(psw));
            this.graphsWorkbook = new XSSFWorkbook(new FileInputStream(graphs));
        } catch (IOException ex) {
            //--Do something
        }
    }
    /**
     * Retrieves the Column no. of a cell by the value of it's content
     *
     * @param str
     * @param sheet
     * @return
     */
    private static int getCellColumnByString(String str, XSSFSheet sheet) {
        int columnNumber = 0;
        
        for (Row r : sheet) {
            for (Cell c : r) {
                String cellValueStr = "";
                try {
                    cellValueStr = c.getStringCellValue();
                } catch (IllegalStateException e) {
                }
                if (cellValueStr.equals(str)) {
                    
                    columnNumber = c.getColumnIndex();
                    
                }
            }
        }
        return columnNumber;
    }
    
    
    
    public void getGraph0() {
        
        int fCount = 0;
        int mCount = 0;
        int uCount = 0;
        double wssxFemaleCount;
        double wssxMaleCount;
        double wssxUnknownCount;
        
        XSSFSheet refSheet = pswWorkbook.getSheetAt(pswWorkbook.getSheetIndex(Strings.PSW_SHEET_REFERRALS));
        XSSFSheet genderSheet = pswWorkbook.getSheetAt(pswWorkbook.getSheetIndex(Strings.PSW_COLUMN_GENDER));
        int columnNumber = getCellColumnByString(Strings.PSW_COLUMN_GENDER, refSheet);
        
        Iterator<Row> rowIterator = refSheet.iterator();
        
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            Cell cell = CellUtil.getCell(row, columnNumber);
            
            if (cell.getStringCellValue().equals(Strings.PSW_GENDER_F)) {
                fCount++;
            } else if (cell.getStringCellValue().equals(Strings.PSW_GENDER_M)) {
                mCount++;
            }
        }
        
        Row femaleRow = genderSheet.getRow(2);
        Row unknownRow = genderSheet.getRow(3);
        Row maleRow = genderSheet.getRow(4);
        
        wssxFemaleCount = femaleRow.getCell(1).getNumericCellValue();
        wssxMaleCount = maleRow.getCell(1).getNumericCellValue();
        wssxUnknownCount = unknownRow.getCell(1).getNumericCellValue();
        
        graphsWorkbook = new XSSFWorkbook();
        XSSFSheet sheet;
        if(graphsWorkbook.getNumberOfSheets()==0){
            sheet = graphsWorkbook.createSheet(Strings.GRAPH_SHEET_1);
        }
        else{
            sheet = graphsWorkbook.getSheetAt(0);
        }
        
        Row titlesRow = sheet.createRow(0);
        Row fRow = sheet.createRow(1);
        Row mRow = sheet.createRow(2);
        Row uRow = sheet.createRow(3);
        Cell cell = titlesRow.createCell(1);
        cell.setCellValue(Strings.GRAPH_COLUMN_REF_COUNT);
        cell = titlesRow.createCell(2);
        cell.setCellValue(Strings.GRAPH_COLUMN_WSSX_COUNT);
        Cell fCell = fRow.createCell(0);
        fCell.setCellValue(Strings.PSW_COLUMN_FEMALE);
        fCell = fRow.createCell(1);
        fCell.setCellValue(fCount);
        fCell = fRow.createCell(2);
        fCell.setCellValue(wssxFemaleCount);
        Cell mCell = mRow.createCell(0);
        mCell.setCellValue(Strings.PSW_COLUMN_MALE);
        mCell = mRow.createCell(1);
        mCell.setCellValue(mCount);
        mCell = mRow.createCell(2);
        mCell.setCellValue(wssxMaleCount);
        Cell uCell = uRow.createCell(0);
        uCell.setCellValue(Strings.PSW_COLUMN_UNKNOWN);
        uCell = uRow.createCell(1);
        uCell.setCellValue(uCount);
        uCell = uRow.createCell(2);
        uCell.setCellValue(wssxUnknownCount);
        
        try{
            FileOutputStream fileOut = new FileOutputStream(graphs);
            graphsWorkbook.write(fileOut);
            fileOut.close();
        }catch (IOException ex) {
            Logger.getLogger(PoiHelper.class.getName()).log(Level.SEVERE, null, ex);
        }
        
        
    }
    
    /**
     * Generates the table for the second graph in Graphs.xlsx
     *
     * @param file
     * @param graphs
     */
    public void getGraph1() {
        
        int capabilityMCount = 0;
        int capabilityFCount = 0;
        int anxietyMCount = 0;
        int anxietyFCount = 0;
        int carreerMCount = 0;
        int carreerFCount = 0;
        int clinicalMCount = 0;
        int clinicalFCount = 0;
        int communicationMCount = 0;
        int communicationFCount = 0;
        int conductMCount = 0;
        int conductFCount = 0;
        int culturalMCount = 0;
        int culturalFCount = 0;
        int examMCount = 0;
        int examFCount = 0;
        int phHhealthMCount = 0;
        int phHealthFCount = 0;
        int menHealthMCount = 0;
        int menHealthFCount = 0;
        int languageMCount = 0;
        int languageFCount = 0;
        int profMCount = 0;
        int profFCount = 0;
        int adhdMCount = 0;
        int adhdFCount = 0;
        int asdMCount = 0;
        int asdFCount = 0;
        int dyslexiaMCount = 0;
        int dyslexiaFCount = 0;
        int dyspraxiaMCount = 0;
        int dyspraxiaFCount = 0;
        int srttMCount = 0;
        int srttFCount = 0;
        int teamMCount = 0;
        int teamFCount = 0;
        int timeMCount = 0;
        int timeFCount = 0;
        int otherMCount = 0;
        int otherFCount = 0;
        
        for(RefRecord record:recordList){
            if (record.isAnxiety()) {
                if (record.getGender().equals(Strings.PSW_GENDER_F)) {
                    anxietyFCount++;
                } else if (record.getGender().equals(Strings.PSW_GENDER_M)) {
                    anxietyMCount++;
                }
            }
            if (record.isCapability()) {
                if (record.getGender().equals(Strings.PSW_GENDER_F)) {
                    capabilityFCount++;
                } else if (record.getGender().equals(Strings.PSW_GENDER_M)) {
                    capabilityMCount++;
                }
            }
            if (record.isCarreer()) {
                if (record.getGender().equals(Strings.PSW_GENDER_F)) {
                    carreerFCount++;
                } else if (record.getGender().equals(Strings.PSW_GENDER_M)) {
                    carreerMCount++;
                }
            }
            if (record.isClinSkills()) {
                if (record.getGender().equals(Strings.PSW_GENDER_F)) {
                    clinicalFCount++;
                } else if (record.getGender().equals(Strings.PSW_GENDER_M)) {
                    clinicalMCount++;
                }
            }
            if (record.isCommunication()) {
                if (record.getGender().equals(Strings.PSW_GENDER_F)) {
                    communicationFCount++;
                } else if (record.getGender().equals(Strings.PSW_GENDER_M)) {
                    communicationMCount++;
                }
            }
            if (record.isConduct()) {
                if (record.getGender().equals(Strings.PSW_GENDER_F)) {
                    conductFCount++;
                } else if (record.getGender().equals(Strings.PSW_GENDER_M)) {
                    conductMCount++;
                }
            }
            if (record.isCultural()) {
                if (record.getGender().equals(Strings.PSW_GENDER_F)) {
                    culturalFCount++;
                } else if (record.getGender().equals(Strings.PSW_GENDER_M)) {
                    culturalMCount++;
                }
            }
            if (record.isExam()) {
                if (record.getGender().equals(Strings.PSW_GENDER_F)) {
                    examFCount++;
                } else if (record.getGender().equals(Strings.PSW_GENDER_M)) {
                    examMCount++;
                }
            }
            if (record.isHealthMental()) {
                if (record.getGender().equals(Strings.PSW_GENDER_F)) {
                    menHealthFCount++;
                } else if (record.getGender().equals(Strings.PSW_GENDER_M)) {
                    menHealthMCount++;
                }
            }
            if (record.isHealthPhysical()) {
                if (record.getGender().equals(Strings.PSW_GENDER_F)) {
                    phHealthFCount++;
                } else if (record.getGender().equals(Strings.PSW_GENDER_M)) {
                    phHhealthMCount++;
                }
            }
            if (!record.isLanguage()) {
                if (record.getGender().equals(Strings.PSW_GENDER_F)) {
                    languageFCount++;
                } else if (record.getGender().equals(Strings.PSW_GENDER_M)) {
                    languageMCount++;
                }
            }
            if (record.isProfessionalism()) {
                if (record.getGender().equals(Strings.PSW_GENDER_F)) {
                    profFCount++;
                } else if (record.getGender().equals(Strings.PSW_GENDER_M)) {
                    profMCount++;
                }
            }
            if (record.isAdhd()) {
                if (record.getGender().equals(Strings.PSW_GENDER_F)) {
                    adhdFCount++;
                } else if (record.getGender().equals(Strings.PSW_GENDER_M)) {
                    adhdMCount++;
                }
            }
            if (record.isAsd()) {
                if (record.getGender().equals(Strings.PSW_GENDER_F)) {
                    asdFCount++;
                } else if (record.getGender().equals(Strings.PSW_GENDER_M)) {
                    asdMCount++;
                }
            }
            if (record.isDyslexia()) {
                if (record.getGender().equals(Strings.PSW_GENDER_F)) {
                    dyslexiaFCount++;
                } else if (record.getGender().equals(Strings.PSW_GENDER_M)) {
                    dyslexiaMCount++;
                }
            }
            if (record.isDyspraxia()) {
                if (record.getGender().equals(Strings.PSW_GENDER_F)) {
                    dyspraxiaFCount++;
                } else if (record.getGender().equals(Strings.PSW_GENDER_M)) {
                    dyspraxiaMCount++;
                }
            }
            if (record.isSrtt()) {
                if (record.getGender().equals(Strings.PSW_GENDER_F)) {
                    srttFCount++;
                } else if (record.getGender().equals(Strings.PSW_GENDER_M)) {
                    srttMCount++;
                }
            }
            if (record.isTeam()) {
                if (record.getGender().equals(Strings.PSW_GENDER_F)) {
                    anxietyFCount++;
                } else if (record.getGender().equals(Strings.PSW_GENDER_M)) {
                    anxietyMCount++;
                }
            }
            if (record.isTime()) {
                if (record.getGender().equals(Strings.PSW_GENDER_F)) {
                    timeFCount++;
                } else if (record.getGender().equals(Strings.PSW_GENDER_M)) {
                    timeMCount++;
                }
            }
            if (record.isOtherRefReason()){
                if (record.getGender().equals(Strings.PSW_GENDER_F)) {
                    otherFCount++;
                } else if (record.getGender().equals(Strings.PSW_GENDER_M)) {
                    otherMCount++;
                }
            }
        }
        
        XSSFSheet sheet;
        
        if (graphsWorkbook.getNumberOfSheets() == 1) {
            sheet = graphsWorkbook.createSheet(Strings.GRAPH_SHEET_2);
        } else {
            sheet = graphsWorkbook.getSheetAt(1);
        }
        
        Row titlesRow = sheet.createRow(0);
        Row fRow = sheet.createRow(1);
        Row mRow = sheet.createRow(2);
        Cell cell = titlesRow.createCell(1);
        cell.setCellValue(Strings.PSW_COLUMN_ANXIETY);
        cell = titlesRow.createCell(2);
        cell.setCellValue(Strings.PSW_COLUMN_CAPABILITY);
        cell = titlesRow.createCell(3);
        cell.setCellValue(Strings.PSW_COLUMN_CARREER);
        cell = titlesRow.createCell(4);
        cell.setCellValue(Strings.PSW_COLUMN_CLINICAL_SKILLS);
        cell = titlesRow.createCell(5);
        cell.setCellValue(Strings.PSW_COLUMN_COMMUNICATION);
        cell = titlesRow.createCell(6);
        cell.setCellValue(Strings.PSW_COLUMN_CONDUCT);
        cell = titlesRow.createCell(7);
        cell.setCellValue(Strings.PSW_COLUMN_CULTURAL);
        cell = titlesRow.createCell(8);
        cell.setCellValue(Strings.PSW_COLUMN_EXAM);
        cell = titlesRow.createCell(9);
        cell.setCellValue(Strings.PSW_COLUMN_MENTAL);
        cell = titlesRow.createCell(10);
        cell.setCellValue(Strings.PSW_COLUMN_PHYSICAL);
        cell = titlesRow.createCell(11);
        cell.setCellValue(Strings.PSW_COLUMN_LANGUAGE);
        cell = titlesRow.createCell(12);
        cell.setCellValue(Strings.PSW_COLUMN_PROFFESSIONALISM);
        cell = titlesRow.createCell(13);
        cell.setCellValue(Strings.PSW_COLUMN_ADHD);
        cell = titlesRow.createCell(14);
        cell.setCellValue(Strings.PSW_COLUMN_ASD);
        cell = titlesRow.createCell(15);
        cell.setCellValue(Strings.PSW_COLUMN_DYSLEXIA);
        cell = titlesRow.createCell(16);
        cell.setCellValue(Strings.PSW_COLUMN_DYSPRAXIA);
        cell = titlesRow.createCell(17);
        cell.setCellValue(Strings.PSW_COLUMN_SRTT);
        cell = titlesRow.createCell(18);
        cell.setCellValue(Strings.PSW_COLUMN_TEAM);
        cell = titlesRow.createCell(19);
        cell.setCellValue(Strings.PSW_COLUMN_TIME);
        cell = titlesRow.createCell(20);
        cell.setCellValue(Strings.GRAPH_COLUMN_OTHER);
        
        Cell fCell = fRow.createCell(0);
        fCell.setCellValue(Strings.PSW_COLUMN_FEMALE);
        Cell mCell = mRow.createCell(0);
        mCell.setCellValue(Strings.PSW_COLUMN_MALE);
        fCell = fRow.createCell(1);
        fCell.setCellValue(anxietyFCount);
        fCell = fRow.createCell(2);
        fCell.setCellValue(capabilityFCount);
        fCell = fRow.createCell(3);
        fCell.setCellValue(carreerFCount);
        fCell = fRow.createCell(4);
        fCell.setCellValue(clinicalFCount);
        fCell = fRow.createCell(5);
        fCell.setCellValue(communicationFCount);
        fCell = fRow.createCell(6);
        fCell.setCellValue(conductFCount);
        fCell = fRow.createCell(7);
        fCell.setCellValue(culturalFCount);
        fCell = fRow.createCell(8);
        fCell.setCellValue(examFCount);
        fCell = fRow.createCell(9);
        fCell.setCellValue(phHealthFCount);
        fCell = fRow.createCell(10);
        fCell.setCellValue(menHealthFCount);
        fCell = fRow.createCell(11);
        fCell.setCellValue(languageFCount);
        fCell = fRow.createCell(12);
        fCell.setCellValue(profFCount);
        fCell = fRow.createCell(13);
        fCell.setCellValue(adhdFCount);
        fCell = fRow.createCell(14);
        fCell.setCellValue(asdFCount);
        fCell = fRow.createCell(15);
        fCell.setCellValue(dyslexiaFCount);
        fCell = fRow.createCell(16);
        fCell.setCellValue(dyspraxiaFCount);
        fCell = fRow.createCell(17);
        fCell.setCellValue(srttFCount);
        fCell = fRow.createCell(18);
        fCell.setCellValue(teamFCount);
        fCell = fRow.createCell(19);
        fCell.setCellValue(timeFCount);
        fCell = fRow.createCell(20);
        fCell.setCellValue(otherFCount);
        
        mCell = mRow.createCell(1);
        mCell.setCellValue(anxietyMCount);
        mCell = mRow.createCell(2);
        mCell.setCellValue(capabilityMCount);
        mCell = mRow.createCell(3);
        mCell.setCellValue(carreerMCount);
        mCell = mRow.createCell(4);
        mCell.setCellValue(clinicalMCount);
        mCell = mRow.createCell(5);
        mCell.setCellValue(communicationMCount);
        mCell = mRow.createCell(6);
        mCell.setCellValue(conductMCount);
        mCell = mRow.createCell(7);
        mCell.setCellValue(culturalMCount);
        mCell = mRow.createCell(8);
        mCell.setCellValue(examMCount);
        mCell = mRow.createCell(9);
        mCell.setCellValue(phHhealthMCount);
        mCell = mRow.createCell(10);
        mCell.setCellValue(menHealthMCount);
        mCell = mRow.createCell(11);
        mCell.setCellValue(languageMCount);
        mCell = mRow.createCell(12);
        mCell.setCellValue(profMCount);
        mCell = mRow.createCell(13);
        mCell.setCellValue(adhdMCount);
        mCell = mRow.createCell(14);
        mCell.setCellValue(asdMCount);
        mCell = mRow.createCell(15);
        mCell.setCellValue(dyslexiaMCount);
        mCell = mRow.createCell(16);
        mCell.setCellValue(dyspraxiaMCount);
        mCell = mRow.createCell(17);
        mCell.setCellValue(srttMCount);
        mCell = mRow.createCell(18);
        mCell.setCellValue(teamMCount);
        mCell = mRow.createCell(19);
        mCell.setCellValue(timeMCount);
        mCell = mRow.createCell(20);
        mCell.setCellValue(otherMCount);
        try{
            FileOutputStream fileOut = new FileOutputStream(graphs);
            graphsWorkbook.write(fileOut);
            fileOut.close();
            
        } catch (IOException ex) {
            Logger.getLogger(PoiHelper.class.getName()).log(Level.SEVERE, null, ex);
        }
        
    }
    
    /**
     * Generates the table for the third graph in Graphs.xlsx
     *
     * @param psw
     * @param graphs
     */
    public void getGraph2() {
        
        int mentalFcount = 0;
        int physicalFcount = 0;
        int bothFcount = 0;
        int mentalMcount = 0;
        int physicalMcount = 0;
        int bothMcount = 0;
        
        XSSFSheet refSheet = pswWorkbook.getSheetAt(pswWorkbook.getSheetIndex(Strings.PSW_SHEET_REFERRALS));
        int mentalHealthColNo = getCellColumnByString(Strings.PSW_COLUMN_MENTAL, refSheet);
        int physicalHealthColNo = getCellColumnByString(Strings.PSW_COLUMN_PHYSICAL, refSheet);
        int genderColNo = getCellColumnByString(Strings.PSW_COLUMN_GENDER, refSheet);
        
        Iterator<Row> rowIterator = refSheet.iterator();
        
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            
            Cell mentalHealthCell = CellUtil.getCell(row, mentalHealthColNo);
            Cell physicalHealthCell = CellUtil.getCell(row, physicalHealthColNo);
            
            if (!row.getCell(mentalHealthColNo).getStringCellValue().equals("") && row.getCell(genderColNo).getStringCellValue().equals("F") && row.getCell(physicalHealthColNo).getStringCellValue().equals("")) {
                mentalFcount++;
            }
            if (!row.getCell(physicalHealthColNo).getStringCellValue().equals("") && row.getCell(genderColNo).getStringCellValue().equals("F") && row.getCell(mentalHealthColNo).getStringCellValue().equals("")) {
                physicalFcount++;
            }
            if (!row.getCell(mentalHealthColNo).getStringCellValue().equals("") && row.getCell(genderColNo).getStringCellValue().equals("F") && !row.getCell(physicalHealthColNo).getStringCellValue().equals("")) {
                bothFcount++;
            }
            if (!row.getCell(mentalHealthColNo).getStringCellValue().equals("") && row.getCell(genderColNo).getStringCellValue().equals("M") && !row.getCell(physicalHealthColNo).getStringCellValue().equals("")) {
                bothMcount++;
            }
            if (!row.getCell(physicalHealthColNo).getStringCellValue().equals("") && row.getCell(genderColNo).getStringCellValue().equals("M") && row.getCell(mentalHealthColNo).getStringCellValue().equals("")) {
                physicalFcount++;
            }
            if (!row.getCell(mentalHealthColNo).getStringCellValue().equals("") && row.getCell(genderColNo).getStringCellValue().equals("M") && row.getCell(physicalHealthColNo).getStringCellValue().equals("")) {
                mentalMcount++;
            }
        }
        
        XSSFSheet sheet;
        
        if (graphsWorkbook.getNumberOfSheets() == 2) {
            sheet = graphsWorkbook.createSheet(Strings.GRAPH_SHEET_3);
        } else {
            sheet = graphsWorkbook.getSheetAt(2);
        }
        
        Row titlesRow = sheet.createRow(0);
        Row fRow = sheet.createRow(1);
        Row mRow = sheet.createRow(2);
        Cell cell = titlesRow.createCell(1);
        cell.setCellValue(Strings.GRAPH_COLUMN_MENTAL);
        cell = titlesRow.createCell(2);
        cell.setCellValue(Strings.GRAPH_COLUMN_PHYSICAL);
        cell = titlesRow.createCell(3);
        cell.setCellValue(Strings.GRAPH_COLUMN_BOTH);
        Cell fcell = fRow.createCell(0);
        fcell.setCellValue(Strings.PSW_COLUMN_FEMALE);
        fcell = fRow.createCell(1);
        fcell.setCellValue(mentalFcount);
        fcell = fRow.createCell(2);
        fcell.setCellValue(physicalFcount);
        fcell = fRow.createCell(3);
        fcell.setCellValue(bothFcount);
        Cell mcell = mRow.createCell(0);
        mcell.setCellValue(Strings.PSW_COLUMN_MALE);
        mcell = mRow.createCell(1);
        mcell.setCellValue(mentalMcount);
        mcell = mRow.createCell(2);
        mcell.setCellValue(physicalMcount);
        mcell = mRow.createCell(3);
        mcell.setCellValue(bothMcount);
        
        try{
            FileOutputStream fileOut = new FileOutputStream(graphs);
            graphsWorkbook.write(fileOut);
            fileOut.close();
            
        } catch (IOException ex) {
            Logger.getLogger(PoiHelper.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
    
    /**
     * Generates the table for the fourth graph in Graphs.xlsx
     *
     * @param psw
     * @param graphs
     */
    public void getGraph3() {
        
        int f1Count = 0;
        int ct1Count = 0;
        int st3Count = 0;
        int st6Count = 0;
        
        
        XSSFSheet refSheet = pswWorkbook.getSheetAt(pswWorkbook.getSheetIndex(Strings.PSW_SHEET_REFERRALS));
        int gradeColNo = getCellColumnByString(Strings.PSW_COLUMN_GRADE, refSheet);
        
        Iterator<Row> rowIterator = refSheet.iterator();
        
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            
            Cell gradeCell = CellUtil.getCell(row, gradeColNo);
            
            Set<String> f1grades = new HashSet<>();
            f1grades.add("F1");
            f1grades.add("F2");
            f1grades.add("FY1");
            f1grades.add("FY2");
            
            Set<String> ct1grades = new HashSet<>();
            ct1grades.add("CT1");
            ct1grades.add("CT2");
            ct1grades.add("ST1");
            ct1grades.add("ST2");
            
            Set<String> st3grades = new HashSet<>();
            st3grades.add("ST3");
            st3grades.add("ST4");
            st3grades.add("ST5");
            st3grades.add("CT3");
            
            Set<String> st6grades = new HashSet<>();
            st6grades.add("ST6");
            st6grades.add("ST7");
            st6grades.add("ST8");
            
            if (f1grades.contains(gradeCell.getStringCellValue())) {
                f1Count++;
            }
            
            if (ct1grades.contains(gradeCell.getStringCellValue())) {
                ct1Count++;
            }
            if (st3grades.contains(gradeCell.getStringCellValue())) {
                st3Count++;
            }
            if (st6grades.contains(gradeCell.getStringCellValue())) {
                st6Count++;
            }
        }
        
        XSSFSheet sheet;
        
        if (graphsWorkbook.getNumberOfSheets() == 3) {
            sheet = graphsWorkbook.createSheet(Strings.GRAPH_SHEET_4);
        } else {
            sheet = graphsWorkbook.getSheetAt(3);
        }
        
        double totalCount = f1Count + ct1Count + st3Count + st6Count;
        double f1perc = Math.round(f1Count / totalCount * 100);
        double ct1perc = Math.round(ct1Count / totalCount * 100);
        double st3perc = Math.round(st3Count / totalCount * 100);
        double st6perc = Math.round(st6Count / totalCount * 100);
        
        Row titlesRow = sheet.createRow(0);
        Row f1Row = sheet.createRow(1);
        Row ct1Row = sheet.createRow(2);
        Row st3Row = sheet.createRow(3);
        Row st6Row = sheet.createRow(4);
        Cell cell = titlesRow.createCell(1);
        cell.setCellValue(Strings.GRAPH_COLUMN_TOTAL);
        Cell f1cell = f1Row.createCell(0);
        f1cell.setCellValue(Strings.GRAPH_ROW_F1 + " " + f1perc + "%");
        f1cell = f1Row.createCell(1);
        f1cell.setCellValue(f1Count);
        Cell ct1Cell = ct1Row.createCell(0);
        ct1Cell.setCellValue(Strings.GRAPH_ROW_CT1 + " " + ct1perc + "%");
        ct1Cell = ct1Row.createCell(1);
        ct1Cell.setCellValue(ct1Count);
        Cell st3Cell = st3Row.createCell(0);
        st3Cell.setCellValue(Strings.GRAPH_ROW_ST3 + " " + st3perc + "%");
        st3Cell = st3Row.createCell(1);
        st3Cell.setCellValue(st3Count);
        Cell st6Cell = st6Row.createCell(0);
        st6Cell.setCellValue(Strings.GRAPH_ROW_ST6 + " " + st6perc + "%");
        st6Cell = st6Row.createCell(1);
        st6Cell.setCellValue(st6Count);
        try{
            FileOutputStream fileOut = new FileOutputStream(graphs);
            graphsWorkbook.write(fileOut);
            fileOut.close();
            
        } catch (IOException ex) {
            Logger.getLogger(PoiHelper.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
    
    /**
     * Generates the table for the fifth graph in Graphs.xlsx
     *
     * @param psw
     * @param graphs
     */
    public void getGraph4() {
        int stCount = 0;
        int ftCount = 0;
        int gpCount = 0;
        int otherCount = 0;
        
        
        XSSFSheet refSheet = pswWorkbook.getSheetAt(pswWorkbook.getSheetIndex(Strings.PSW_SHEET_REFERRALS));
        int gradeColNo = getCellColumnByString(Strings.PSW_COLUMN_GRADE, refSheet);
        
        Iterator<Row> rowIterator = refSheet.iterator();
        
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            
            Cell cell = CellUtil.getCell(row, gradeColNo);
            
            if (cell.getStringCellValue().equals(Strings.PSW_GRADE_FY1) || cell.getStringCellValue().equals(Strings.PSW_GRADE_FY2)) {
                ftCount++;
            }
            if (cell.getStringCellValue().equals(Strings.PSW_GRADE_ST1) || cell.getStringCellValue().equals(Strings.PSW_GRADE_ST2) || cell.getStringCellValue().equals(Strings.PSW_GRADE_ST3) || cell.getStringCellValue().equals(Strings.PSW_GRADE_ST4) || cell.getStringCellValue().equals(Strings.PSW_GRADE_ST5) || cell.getStringCellValue().equals(Strings.PSW_GRADE_ST6) || cell.getStringCellValue().equals(Strings.PSW_GRADE_ST7) || cell.getStringCellValue().equals(Strings.PSW_GRADE_ST8)) {
                stCount++;
            }
            if (cell.getStringCellValue().equals(Strings.PSW_GRADE_GPST1) || cell.getStringCellValue().equals(Strings.PSW_GRADE_GPST2) || cell.getStringCellValue().equals(Strings.PSW_GRADE_GPST3)) {
                gpCount++;
            }
            if (cell.getStringCellValue().equals(Strings.PSW_GRADE_DCT1) || cell.getStringCellValue().equals(Strings.PSW_GRADE_DCT2) || cell.getStringCellValue().equals(Strings.PSW_GRADE_DF1) || cell.getStringCellValue().equals(Strings.PSW_GRADE_DF2)) {
                otherCount++;
            }
        }
        
        XSSFSheet sheet;
        
        if (graphsWorkbook.getNumberOfSheets() == 4) {
            sheet = graphsWorkbook.createSheet(Strings.GRAPH_SHEET_5);
        } else {
            sheet = graphsWorkbook.getSheetAt(4);
        }
        
        Row titlesRow = sheet.createRow(0);
        Row ftRow = sheet.createRow(1);
        Row stRow = sheet.createRow(2);
        Row gpRow = sheet.createRow(3);
        Row otherRow = sheet.createRow(4);
        Cell cell = titlesRow.createCell(1);
        cell.setCellValue(Strings.GRAPH_COLUMN_TOTAL);
        Cell ftcell = ftRow.createCell(0);
        ftcell.setCellValue(Strings.GRAPH_ROW_FT);
        ftcell = ftRow.createCell(1);
        ftcell.setCellValue(ftCount);
        Cell stCell = stRow.createCell(0);
        stCell.setCellValue(Strings.GRAPH_ROW_ST);
        stCell = stRow.createCell(1);
        stCell.setCellValue(stCount);
        Cell gpCell = gpRow.createCell(0);
        gpCell.setCellValue(Strings.GRAPH_ROW_GP);
        gpCell = gpRow.createCell(1);
        gpCell.setCellValue(gpCount);
        Cell otherCell = otherRow.createCell(0);
        otherCell.setCellValue(Strings.GRAPH_ROW_OTHER);
        otherCell = otherRow.createCell(1);
        otherCell.setCellValue(otherCount);
        
        try{
            FileOutputStream fileOut = new FileOutputStream(graphs);
            graphsWorkbook.write(fileOut);
            fileOut.close();
            
        } catch (IOException ex) {
            Logger.getLogger(PoiHelper.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
    
    /**
     * Generates the table for the sixth graph in Graphs.xlsx
     *
     * @param psw
     * @param graphs
     */
    public void getGraph5() {
        
        
        XSSFSheet ocSheet = pswWorkbook.getSheetAt(pswWorkbook.getSheetIndex(Strings.PSW_SHEET_OPEN_CASES));
        Row ocR0 = ocSheet.getRow(0);
        Row ocR1 = ocSheet.getRow(1);
        Row ocR2 = ocSheet.getRow(2);
        
        XSSFSheet sheet;
        
        if (graphsWorkbook.getNumberOfSheets() == 5) {
            sheet = graphsWorkbook.createSheet(Strings.GRAPH_SHEET_6);
        } else {
            sheet = graphsWorkbook.getSheetAt(5);
        }
        
        Row titlesRow = sheet.createRow(0);
        Row fRow = sheet.createRow(1);
        Row sRow = sheet.createRow(2);
        
        titlesRow.createCell(1).setCellValue(ocR0.getCell(1).getStringCellValue());
        titlesRow.createCell(2).setCellValue(ocR0.getCell(2).getStringCellValue());
        titlesRow.createCell(3).setCellValue(ocR0.getCell(3).getStringCellValue());
        titlesRow.createCell(4).setCellValue(ocR0.getCell(4).getStringCellValue());
        titlesRow.createCell(5).setCellValue(ocR0.getCell(5).getStringCellValue());
        
        CellStyle cellStyle = graphsWorkbook.createCellStyle();
        CreationHelper createHelper = graphsWorkbook.getCreationHelper();
        cellStyle.setDataFormat(
                createHelper.createDataFormat().getFormat("MMM-YY"));
        
        Cell r1c0 = fRow.createCell(0);
        r1c0.setCellValue(ocR1.getCell(0).getDateCellValue());
        r1c0.setCellStyle(cellStyle);
        fRow.createCell(1).setCellValue(ocR1.getCell(1).getNumericCellValue());
        fRow.createCell(2).setCellValue(ocR1.getCell(2).getNumericCellValue());
        fRow.createCell(3).setCellValue(ocR1.getCell(3).getNumericCellValue());
        fRow.createCell(4).setCellValue(ocR1.getCell(4).getNumericCellValue());
        fRow.createCell(5).setCellValue(ocR1.getCell(5).getNumericCellValue());
        
        Cell r2c0 = sRow.createCell(0);
        r2c0.setCellValue(ocR2.getCell(0).getDateCellValue());
        r2c0.setCellStyle(cellStyle);
        sRow.createCell(1).setCellValue(ocR2.getCell(1).getNumericCellValue());
        sRow.createCell(2).setCellValue(ocR2.getCell(2).getNumericCellValue());
        sRow.createCell(3).setCellValue(ocR2.getCell(3).getNumericCellValue());
        sRow.createCell(4).setCellValue(ocR2.getCell(4).getNumericCellValue());
        sRow.createCell(5).setCellValue(ocR2.getCell(5).getNumericCellValue());
        
        try{
            FileOutputStream fileOut = new FileOutputStream(graphs);
            graphsWorkbook.write(fileOut);
            fileOut.close();
            
        } catch (IOException ex) {
            Logger.getLogger(PoiHelper.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
    
    
    public void getGraph6(){
        
        
        
        
        XSSFSheet sheet;
        
        if (graphsWorkbook.getNumberOfSheets() == 6) {
            sheet = graphsWorkbook.createSheet(Strings.GRAPH_SHEET_7);
        } else {
            sheet = graphsWorkbook.getSheetAt(6);
        }
        
        
    }
    
    /**
     * Generates the table for the eighth graph in Graphs.xlsx
     *
     * @param psw
     * @param graphs
     */
    public void getGraph7() {
        
        
        int noConcerns = 0;
        int onGoing = 0;
        int completed = 0;
        int released = 0;
        int resigned = 0;
        int other = 0;
        int death = 0;
        
        XSSFSheet ccSheet = pswWorkbook.getSheetAt(pswWorkbook.getSheetIndex(Strings.PSW_SHEET_CLOSED_CASES));
        int totalsColumn = getCellColumnByString(Strings.PSW_COLUMN_OUTCOME_KEY, ccSheet);
        
        XSSFSheet sheet;
        
        noConcerns = Integer.parseInt(ccSheet.getRow(1).getCell(totalsColumn).getRawValue());
        onGoing = Integer.parseInt(ccSheet.getRow(2).getCell(totalsColumn).getRawValue());
        completed = Integer.parseInt(ccSheet.getRow(3).getCell(totalsColumn).getRawValue());
        released = Integer.parseInt(ccSheet.getRow(4).getCell(totalsColumn).getRawValue());
        resigned = Integer.parseInt(ccSheet.getRow(5).getCell(totalsColumn).getRawValue());
        other = Integer.parseInt(ccSheet.getRow(6).getCell(totalsColumn).getRawValue());
        death = Integer.parseInt(ccSheet.getRow(7).getCell(totalsColumn).getRawValue());
        
        if (graphsWorkbook.getNumberOfSheets() <= 7) {
            sheet = graphsWorkbook.createSheet(Strings.GRAPH_SHEET_8);
        } else {
            sheet = graphsWorkbook.getSheetAt(7);
        }
        
        Row title = sheet.createRow(0);
        title.createCell(1).setCellValue(Strings.GRAPH_COLUMN_TOTAL);
        
        Row noConcernsRow = sheet.createRow(1);
        noConcernsRow.createCell(0).setCellValue(Strings.GRAPH_ROW_RTT_NO_CONCERNS);
        noConcernsRow.createCell(1).setCellValue(noConcerns);
        
        Row onGoingRow = sheet.createRow(2);
        onGoingRow.createCell(0).setCellValue(Strings.GRAPH_ROW_RTT_CONCERNS);
        onGoingRow.createCell(1).setCellValue(onGoing);
        
        Row completedRow = sheet.createRow(3);
        completedRow.createCell(0).setCellValue(Strings.GRAPH_ROW_COMPLETED_TRAINING);
        completedRow.createCell(1).setCellValue(completed);
        
        Row releasedRow = sheet.createRow(4);
        releasedRow.createCell(0).setCellValue(Strings.GRAPH_ROW_RELEASED_TRAINING);
        releasedRow.createCell(1).setCellValue(released);
        
        Row resignedRow = sheet.createRow(5);
        resignedRow.createCell(0).setCellValue(Strings.GRAPH_ROW_RESIGNED_TRAINING);
        resignedRow.createCell(1).setCellValue(resigned);
        
        Row otherRow = sheet.createRow(6);
        otherRow.createCell(0).setCellValue(Strings.GRAPH_ROW_OTHER_TRAINING);
        otherRow.createCell(1).setCellValue(other);
        
        Row deathRow = sheet.createRow(7);
        deathRow.createCell(0).setCellValue(Strings.GRAPH_ROW_DEATH);
        deathRow.createCell(1).setCellValue(death);
        
        try {
            FileOutputStream fileOut = new FileOutputStream(graphs);
            graphsWorkbook.write(fileOut);
            fileOut.close();
        } catch (IOException ex) {
            Logger.getLogger(PoiHelper.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
    
    /**
     * Generates the table for the eigth graph in Graphs.xlsx
     *
     * @param psw
     * @param graphs
     */
    public void getGraph8() {
        
        XSSFSheet refSheet = pswWorkbook.getSheetAt(pswWorkbook.getSheetIndex(Strings.PSW_SHEET_REFERRALS));
        XSSFSheet wssxSheet = pswWorkbook.getSheetAt(pswWorkbook.getSheetIndex(Strings.PSW_SHEET_WESSEX));
        XSSFSheet ccSheet = pswWorkbook.getSheetAt(pswWorkbook.getSheetIndex(Strings.PSW_SHEET_CLOSED_CASES));
        
        int dateOpnRefColumn = getCellColumnByString(Strings.PSW_COLUMN_DATE, ccSheet);
        
        
        int yr = DocHelper.getStartingYear();
        int yp1 = yr+1;
        
        int aprYrCount = 0;
        int mayYrCount = 0;
        int juneYrCount = 0;
        int julyYrCount = 0;
        int augYrCount = 0;
        int septYrCount = 0;
        int octYrCount = 0;
        int novYrcount = 0;
        int decYrCount = 0;
        int janYrp1Count = 0;
        int febYrp1Count = 0;
        int marYrp1Count = 0;
        
        Date aprYr = new GregorianCalendar(yr, 3, 1).getTime();
        Calendar aprYrCal = Calendar.getInstance();
        aprYrCal.setTime(aprYr);
        Date mayYr = new GregorianCalendar(yr, 4, 1).getTime();
        Calendar mayYrCal = Calendar.getInstance();
        mayYrCal.setTime(mayYr);
        Date juneYr = new GregorianCalendar(yr, 5, 1).getTime();
        Calendar juneYrCal = Calendar.getInstance();
        juneYrCal.setTime(juneYr);
        Date julyYr = new GregorianCalendar(yr, 6, 1).getTime();
        Calendar julyYrCal = Calendar.getInstance();
        julyYrCal.setTime(julyYr);
        Date augYr = new GregorianCalendar(yr, 7, 1).getTime();
        Calendar augYrCal = Calendar.getInstance();
        augYrCal.setTime(augYr);
        Date septYr = new GregorianCalendar(yr, 8, 1).getTime();
        Calendar septYrCal = Calendar.getInstance();
        septYrCal.setTime(septYr);
        Date octYr = new GregorianCalendar(yr, 9, 1).getTime();
        Calendar octYrCal = Calendar.getInstance();
        octYrCal.setTime(octYr);
        Date novYr = new GregorianCalendar(yr, 10, 1).getTime();
        Calendar novYrCal = Calendar.getInstance();
        novYrCal.setTime(novYr);
        Date decYr = new GregorianCalendar(yr, 11, 1).getTime();
        Calendar decYrCal = Calendar.getInstance();
        decYrCal.setTime(decYr);
        Date janYp1 = new GregorianCalendar(yp1, 0, 1).getTime();
        Calendar janYp1Cal = Calendar.getInstance();
        janYp1Cal.setTime(janYp1);
        Date febYp1 = new GregorianCalendar(yp1, 1, 1).getTime();
        Calendar febYp1Cal = Calendar.getInstance();
        febYp1Cal.setTime(febYp1);
        Date marYp1 = new GregorianCalendar(yp1, 2, 1).getTime();
        Calendar marYp1Cal = Calendar.getInstance();
        marYp1Cal.setTime(marYp1);
        
        for(RefRecord record : recordList){
            Calendar calendar = Calendar.getInstance();
            calendar.setTime(record.getRefDate());
            if(calendar.get(Calendar.MONTH)==aprYrCal.get(Calendar.MONTH)){
                aprYrCount++;
            }
            if(calendar.get(Calendar.MONTH)==mayYrCal.get(Calendar.MONTH)){
                mayYrCount++;
            }
            if(calendar.get(Calendar.MONTH)==juneYrCal.get(Calendar.MONTH)){
                juneYrCount++;
            }
            if(calendar.get(Calendar.MONTH)==julyYrCal.get(Calendar.MONTH)){
                julyYrCount++;
            }
            if(calendar.get(Calendar.MONTH)==augYrCal.get(Calendar.MONTH)){
                augYrCount++;
            }
            if(calendar.get(Calendar.MONTH)==septYrCal.get(Calendar.MONTH)){
                septYrCount++;
            }
            if(calendar.get(Calendar.MONTH)==octYrCal.get(Calendar.MONTH)){
                octYrCount++;
            }
            if(calendar.get(Calendar.MONTH)==novYrCal.get(Calendar.MONTH)){
                novYrcount++;
            }
            if(calendar.get(Calendar.MONTH)==decYrCal.get(Calendar.MONTH)){
                decYrCount++;
            }
            if(calendar.get(Calendar.MONTH)==janYp1Cal.get(Calendar.MONTH)){
                janYrp1Count++;
            }
            if(calendar.get(Calendar.MONTH)==febYp1Cal.get(Calendar.MONTH)){
                febYrp1Count++;
            }
            if(calendar.get(Calendar.MONTH)==marYp1Cal.get(Calendar.MONTH)){
                marYrp1Count++;
            }
            
            int aprYrFinal = marYrp1Count+aprYrCount;
            int mayYrFinal = aprYrFinal+mayYrCount;
            int juneYrFinal = mayYrFinal + juneYrCount;
            int julyYrFinal = juneYrFinal + julyYrCount;
            int augYrFinal = julyYrFinal + augYrCount;
            int septYrFinal = augYrFinal + septYrCount;
            int octYrFinal = septYrFinal + octYrCount;
            int novYrFinal = octYrFinal + novYrcount;
//            int decYrFinal = novYrFinal + decYrCount;
//            int janYrp1Final = decYrFinal + janYrp1Count;
//            int febYrp1Final = de;
            int marYrp1Final;
            
            XSSFSheet sheet;
            if (graphsWorkbook.getNumberOfSheets() == 8) {
                sheet = graphsWorkbook.createSheet(Strings.GRAPH_SHEET_9);
            } else {
                sheet = graphsWorkbook.getSheetAt(8);
            }
            
            try{
                FileOutputStream fileOut = new FileOutputStream(graphs);
                graphsWorkbook.write(fileOut);
                fileOut.close();
                
            } catch (IOException ex) {
                Logger.getLogger(PoiHelper.class.getName()).log(Level.SEVERE, null, ex);
            }
            
        }
    }
    
    public void getGraph9() {
        
        XSSFSheet graphSheet = pswWorkbook.getSheetAt(pswWorkbook.getSheetIndex(Strings.PSW_SHEET_GRAPHS));
        
        XSSFSheet sheet;
        
        int titleColumnNo = getCellColumnByString(Strings.PSW_COLUMN_TOTAL_TRAINEE_TIME, graphSheet);
        int titleRowNo = getCellRowByString(Strings.PSW_COLUMN_TOTAL_TRAINEE_TIME, graphSheet);
        
        int aprRow = titleRowNo + 1;
        int mayRow = titleRowNo + 2;
        int junRow = titleRowNo + 3;
        int julRow = titleRowNo + 4;
        int augRow = titleRowNo + 5;
        int sepRow = titleRowNo + 6;
        int octRow = titleRowNo + 7;
        int novRow = titleRowNo + 8;
        int decRow = titleRowNo + 9;
        int janRow = titleRowNo + 10;
        int febRow = titleRowNo + 11;
        int marRow = titleRowNo + 12;
        
        double timeApr = graphSheet.getRow(aprRow).getCell(titleColumnNo).getNumericCellValue() * 24;
        int hoursApr = (int) timeApr;
        int minutesApr = (int) (timeApr - hoursApr) * 60;
        
        double timeMay = graphSheet.getRow(mayRow).getCell(titleColumnNo).getNumericCellValue() * 24;
        int hoursMay = (int) timeMay;
        int minutesMay = (int) (timeMay - hoursMay) * 60;
        
        double timeJun = graphSheet.getRow(junRow).getCell(titleColumnNo).getNumericCellValue() * 24;
        int hoursJun = (int) timeJun;
        int minutesJun = (int) (timeJun - hoursJun) * 60;
        
        double timeJul = graphSheet.getRow(julRow).getCell(titleColumnNo).getNumericCellValue() * 24;
        int hoursJul = (int) timeJul;
        int minutesJul = (int) (timeJul - hoursJul) * 60;
        
        double timeAug = graphSheet.getRow(augRow).getCell(titleColumnNo).getNumericCellValue() * 24;
        int hoursAug = (int) timeAug;
        int minutesAug = (int) (timeAug - hoursAug) * 60;
        
        double timeSep = graphSheet.getRow(sepRow).getCell(titleColumnNo).getNumericCellValue() * 24;
        int hoursSep = (int) timeSep;
        int minutesSep = (int) (timeSep - hoursSep) * 60;
        
        double timeOct = graphSheet.getRow(octRow).getCell(titleColumnNo).getNumericCellValue() * 24;
        int hoursOct = (int) timeOct;
        int minutesOct = (int) (timeOct - hoursOct) * 60;
        
        double timeNov = graphSheet.getRow(novRow).getCell(titleColumnNo).getNumericCellValue() * 24;
        int hoursNov = (int) timeNov;
        int minutesNov = (int) (timeNov - hoursNov) * 60;
        
        double timeDec = graphSheet.getRow(decRow).getCell(titleColumnNo).getNumericCellValue() * 24;
        int hoursDec = (int) timeDec;
        int minutesDec = (int) (timeDec - hoursDec) * 60;
        
        double timeJan = graphSheet.getRow(janRow).getCell(titleColumnNo).getNumericCellValue() * 24;
        int hoursJan = (int) timeJan;
        int minutesJan = (int) (timeJan - hoursApr) * 60;
        
        double timeFeb = graphSheet.getRow(febRow).getCell(titleColumnNo).getNumericCellValue() * 24;
        int hoursFeb = (int) timeFeb;
        int minutesFeb = (int) (timeFeb - hoursFeb) * 60;
        
        double timeMar = graphSheet.getRow(marRow).getCell(titleColumnNo).getNumericCellValue() * 24;
        int hoursMar = (int) timeMar;
        int minutesMar = (int) (timeMar - hoursMar) * 60;
        
        if (graphsWorkbook.getNumberOfSheets() <= 9) {
            sheet = graphsWorkbook.createSheet(Strings.GRAPH_SHEET_10);
        } else {
            sheet = graphsWorkbook.getSheetAt(9);
        }
        
        Row title = sheet.createRow(0);
        title.createCell(1).setCellValue(Strings.GRAPH_COLUMN_TOTAL);
        
        Row aprGraphRow = sheet.createRow(1);
        aprGraphRow.createCell(0).setCellValue(Strings.GRAPH_ROW_APR);
        aprGraphRow.createCell(1).setCellValue(hoursApr);
        
        Row mayGraphRow = sheet.createRow(2);
        mayGraphRow.createCell(0).setCellValue(Strings.GRAPH_ROW_MAY);
        mayGraphRow.createCell(1).setCellValue(hoursMay);
        
        Row junGraphRow = sheet.createRow(3);
        junGraphRow.createCell(0).setCellValue(Strings.GRAPH_ROW_JUN);
        junGraphRow.createCell(1).setCellValue(hoursJun);
        
        Row julGraphRow = sheet.createRow(4);
        julGraphRow.createCell(0).setCellValue(Strings.GRAPH_ROW_JUL);
        julGraphRow.createCell(1).setCellValue(hoursJul);
        
        Row augGraphRow = sheet.createRow(5);
        augGraphRow.createCell(0).setCellValue(Strings.GRAPH_ROW_AUG);
        augGraphRow.createCell(1).setCellValue(hoursAug);
        
        Row sepGraphRow = sheet.createRow(6);
        sepGraphRow.createCell(0).setCellValue(Strings.GRAPH_ROW_SEP);
        sepGraphRow.createCell(1).setCellValue(hoursSep);
        
        Row octGraphRow = sheet.createRow(7);
        octGraphRow.createCell(0).setCellValue(Strings.GRAPH_ROW_OCT);
        octGraphRow.createCell(1).setCellValue(hoursOct);
        
        Row novGraphRow = sheet.createRow(8);
        novGraphRow.createCell(0).setCellValue(Strings.GRAPH_ROW_NOV);
        novGraphRow.createCell(1).setCellValue(hoursNov);
        
        Row decGraphRow = sheet.createRow(9);
        decGraphRow.createCell(0).setCellValue(Strings.GRAPH_ROW_DEC);
        decGraphRow.createCell(1).setCellValue(hoursDec);
        
        Row janGraphRow = sheet.createRow(10);
        janGraphRow.createCell(0).setCellValue(Strings.GRAPH_ROW_JAN);
        janGraphRow.createCell(1).setCellValue(hoursJan);
        
        Row febGraphRow = sheet.createRow(11);
        febGraphRow.createCell(0).setCellValue(Strings.GRAPH_ROW_FEB);
        febGraphRow.createCell(1).setCellValue(hoursFeb);
        
        Row marGraphRow = sheet.createRow(12);
        marGraphRow.createCell(0).setCellValue(Strings.GRAPH_ROW_MAR);
        marGraphRow.createCell(1).setCellValue(hoursMar);
        
        try{
            FileOutputStream fileOut = new FileOutputStream(graphs);
            graphsWorkbook.write(fileOut);
            fileOut.close();
            
        } catch (IOException ex) {
            Logger.getLogger(PoiHelper.class.getName()).log(Level.SEVERE, null, ex);
        }
        
    }
    
    public void getGraph10() {
        
        
        XSSFSheet graphSheet = pswWorkbook.getSheetAt(pswWorkbook.getSheetIndex(Strings.PSW_SHEET_GRAPHS));
        XSSFSheet sheet;
        
        int titleColumnNo = getCellColumnByString(Strings.PSW_COLUMN_TOTAL_TRAINEE_COSTS, graphSheet);
        int titleRowNo = getCellRowByString(Strings.PSW_COLUMN_TOTAL_TRAINEE_COSTS, graphSheet);
        int aprRow = titleRowNo + 1;
        int mayRow = titleRowNo + 2;
        int junRow = titleRowNo + 3;
        int julRow = titleRowNo + 4;
        int augRow = titleRowNo + 5;
        int sepRow = titleRowNo + 6;
        int octRow = titleRowNo + 7;
        int novRow = titleRowNo + 8;
        int decRow = titleRowNo + 9;
        int janRow = titleRowNo + 10;
        int febRow = titleRowNo + 11;
        int marRow = titleRowNo + 12;
        
        double aprCount = graphSheet.getRow(aprRow).getCell(titleColumnNo).getNumericCellValue();
        double mayCount = graphSheet.getRow(mayRow).getCell(titleColumnNo).getNumericCellValue();
        double junCount = graphSheet.getRow(junRow).getCell(titleColumnNo).getNumericCellValue();
        double julCount = graphSheet.getRow(julRow).getCell(titleColumnNo).getNumericCellValue();
        double augCount = graphSheet.getRow(augRow).getCell(titleColumnNo).getNumericCellValue();
        double sepCount = graphSheet.getRow(sepRow).getCell(titleColumnNo).getNumericCellValue();
        double octCount = graphSheet.getRow(octRow).getCell(titleColumnNo).getNumericCellValue();
        double novCount = graphSheet.getRow(novRow).getCell(titleColumnNo).getNumericCellValue();
        double decCount = graphSheet.getRow(decRow).getCell(titleColumnNo).getNumericCellValue();
        double janCount = graphSheet.getRow(janRow).getCell(titleColumnNo).getNumericCellValue();
        double febCount = graphSheet.getRow(febRow).getCell(titleColumnNo).getNumericCellValue();
        double marCount = graphSheet.getRow(marRow).getCell(titleColumnNo).getNumericCellValue();
        
        if (graphsWorkbook.getNumberOfSheets() <= 10) {
            sheet = graphsWorkbook.createSheet(Strings.GRAPH_SHEET_11);
        } else {
            sheet = graphsWorkbook.getSheetAt(10);
        }
        
        Row title = sheet.createRow(0);
        title.createCell(1).setCellValue(Strings.GRAPH_COLUMN_TOTAL);
        
        Row aprGraphRow = sheet.createRow(1);
        aprGraphRow.createCell(0).setCellValue(Strings.GRAPH_ROW_APR);
        aprGraphRow.createCell(1).setCellValue(aprCount);
        
        Row mayGraphRow = sheet.createRow(2);
        mayGraphRow.createCell(0).setCellValue(Strings.GRAPH_ROW_MAY);
        mayGraphRow.createCell(1).setCellValue(mayCount);
        
        Row junGraphRow = sheet.createRow(3);
        junGraphRow.createCell(0).setCellValue(Strings.GRAPH_ROW_JUN);
        junGraphRow.createCell(1).setCellValue(junCount);
        
        Row julGraphRow = sheet.createRow(4);
        julGraphRow.createCell(0).setCellValue(Strings.GRAPH_ROW_JUL);
        julGraphRow.createCell(1).setCellValue(julCount);
        
        Row augGraphRow = sheet.createRow(5);
        augGraphRow.createCell(0).setCellValue(Strings.GRAPH_ROW_AUG);
        augGraphRow.createCell(1).setCellValue(augCount);
        
        Row sepGraphRow = sheet.createRow(6);
        sepGraphRow.createCell(0).setCellValue(Strings.GRAPH_ROW_SEP);
        sepGraphRow.createCell(1).setCellValue(sepCount);
        
        Row octGraphRow = sheet.createRow(7);
        octGraphRow.createCell(0).setCellValue(Strings.GRAPH_ROW_OCT);
        octGraphRow.createCell(1).setCellValue(octCount);
        
        Row novGraphRow = sheet.createRow(8);
        novGraphRow.createCell(0).setCellValue(Strings.GRAPH_ROW_NOV);
        novGraphRow.createCell(1).setCellValue(novCount);
        
        Row decGraphRow = sheet.createRow(9);
        decGraphRow.createCell(0).setCellValue(Strings.GRAPH_ROW_DEC);
        decGraphRow.createCell(1).setCellValue(decCount);
        
        Row janGraphRow = sheet.createRow(10);
        janGraphRow.createCell(0).setCellValue(Strings.GRAPH_ROW_JAN);
        janGraphRow.createCell(1).setCellValue(janCount);
        
        Row febGraphRow = sheet.createRow(11);
        febGraphRow.createCell(0).setCellValue(Strings.GRAPH_ROW_FEB);
        febGraphRow.createCell(1).setCellValue(febCount);
        
        Row marGraphRow = sheet.createRow(12);
        marGraphRow.createCell(0).setCellValue(Strings.GRAPH_ROW_MAR);
        marGraphRow.createCell(1).setCellValue(marCount);
        
        try{
            FileOutputStream fileOut = new FileOutputStream(graphs);
            graphsWorkbook.write(fileOut);
            fileOut.close();
            
        } catch (IOException ex) {
            Logger.getLogger(PoiHelper.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
    
    public void getGraph11() {
        XSSFSheet graphSheet = pswWorkbook.getSheetAt(pswWorkbook.getSheetIndex(Strings.PSW_SHEET_GRAPHS));
        XSSFSheet sheet;
        
        int ssgColumnNo = getCellColumnByString(Strings.PSW_COLUMN_SSG_COSTS, graphSheet);
        int cmColumnNo = getCellColumnByString(Strings.PSW_COLUMN_CM_COSTS, graphSheet);
        int titleRowNo = getCellRowByString(Strings.PSW_COLUMN_SSG_COSTS, graphSheet);
        int aprRow = titleRowNo + 1;
        int mayRow = titleRowNo + 2;
        int junRow = titleRowNo + 3;
        int julRow = titleRowNo + 4;
        int augRow = titleRowNo + 5;
        int sepRow = titleRowNo + 6;
        int octRow = titleRowNo + 7;
        int novRow = titleRowNo + 8;
        int decRow = titleRowNo + 9;
        int janRow = titleRowNo + 10;
        int febRow = titleRowNo + 11;
        int marRow = titleRowNo + 12;
        
        double aprSsgCount = graphSheet.getRow(aprRow).getCell(ssgColumnNo).getNumericCellValue();
        double maySsgCount = graphSheet.getRow(mayRow).getCell(ssgColumnNo).getNumericCellValue();
        double junSsgCount = graphSheet.getRow(junRow).getCell(ssgColumnNo).getNumericCellValue();
        double julSsgCount = graphSheet.getRow(julRow).getCell(ssgColumnNo).getNumericCellValue();
        double augSsgCount = graphSheet.getRow(augRow).getCell(ssgColumnNo).getNumericCellValue();
        double sepSsgCount = graphSheet.getRow(sepRow).getCell(ssgColumnNo).getNumericCellValue();
        double octSsgCount = graphSheet.getRow(octRow).getCell(ssgColumnNo).getNumericCellValue();
        double novSsgCount = graphSheet.getRow(novRow).getCell(ssgColumnNo).getNumericCellValue();
        double decSsgCount = graphSheet.getRow(decRow).getCell(ssgColumnNo).getNumericCellValue();
        double janSsgCount = graphSheet.getRow(janRow).getCell(ssgColumnNo).getNumericCellValue();
        double febSsgCount = graphSheet.getRow(febRow).getCell(ssgColumnNo).getNumericCellValue();
        double marSsgCount = graphSheet.getRow(marRow).getCell(ssgColumnNo).getNumericCellValue();
        double totalSsgCount = aprSsgCount+maySsgCount+junSsgCount+julSsgCount+augSsgCount+sepSsgCount+octSsgCount+novSsgCount+decSsgCount+janSsgCount+febSsgCount+marSsgCount;
        
        double aprCmCount = graphSheet.getRow(aprRow).getCell(cmColumnNo).getNumericCellValue();
        double mayCmCount = graphSheet.getRow(mayRow).getCell(cmColumnNo).getNumericCellValue();
        double junCmCount = graphSheet.getRow(junRow).getCell(cmColumnNo).getNumericCellValue();
        double julCmCount = graphSheet.getRow(julRow).getCell(cmColumnNo).getNumericCellValue();
        double augCmCount = graphSheet.getRow(augRow).getCell(cmColumnNo).getNumericCellValue();
        double sepCmCount = graphSheet.getRow(sepRow).getCell(cmColumnNo).getNumericCellValue();
        double octCmCount = graphSheet.getRow(octRow).getCell(cmColumnNo).getNumericCellValue();
        double novCmCount = graphSheet.getRow(novRow).getCell(cmColumnNo).getNumericCellValue();
        double decCmCount = graphSheet.getRow(decRow).getCell(cmColumnNo).getNumericCellValue();
        double janCmCount = graphSheet.getRow(janRow).getCell(cmColumnNo).getNumericCellValue();
        double febCmCount = graphSheet.getRow(febRow).getCell(cmColumnNo).getNumericCellValue();
        double marCmCount = graphSheet.getRow(marRow).getCell(cmColumnNo).getNumericCellValue();
        double totalCmCount = aprCmCount+mayCmCount+junCmCount+julCmCount+augCmCount+sepCmCount+octCmCount+novCmCount+decCmCount+janCmCount+febCmCount+marCmCount;
        
        if (graphsWorkbook.getNumberOfSheets() <= 11) {
            sheet = graphsWorkbook.createSheet(Strings.GRAPH_SHEET_12);
        } else {
            sheet = graphsWorkbook.getSheetAt(11);
        }
        
        Row title = sheet.createRow(0);
        title.createCell(1).setCellValue(Strings.GRAPH_COLUMN_SSG);
        title.createCell(2).setCellValue(Strings.GRAPH_COLUMN_CM);
        
        Row aprGraphRow = sheet.createRow(1);
        aprGraphRow.createCell(0).setCellValue(Strings.GRAPH_ROW_APR);
        aprGraphRow.createCell(1).setCellValue(aprSsgCount);
        aprGraphRow.createCell(2).setCellValue(aprCmCount);
        
        Row mayGraphRow = sheet.createRow(2);
        mayGraphRow.createCell(0).setCellValue(Strings.GRAPH_ROW_MAY);
        mayGraphRow.createCell(1).setCellValue(maySsgCount);
        mayGraphRow.createCell(2).setCellValue(mayCmCount);
        
        Row junGraphRow = sheet.createRow(3);
        junGraphRow.createCell(0).setCellValue(Strings.GRAPH_ROW_JUN);
        junGraphRow.createCell(1).setCellValue(junSsgCount);
        junGraphRow.createCell(2).setCellValue(junCmCount);
        
        Row julGraphRow = sheet.createRow(4);
        julGraphRow.createCell(0).setCellValue(Strings.GRAPH_ROW_JUL);
        julGraphRow.createCell(1).setCellValue(julSsgCount);
        julGraphRow.createCell(2).setCellValue(julCmCount);
        
        Row augGraphRow = sheet.createRow(5);
        augGraphRow.createCell(0).setCellValue(Strings.GRAPH_ROW_AUG);
        augGraphRow.createCell(1).setCellValue(augSsgCount);
        augGraphRow.createCell(2).setCellValue(augCmCount);
        
        Row sepGraphRow = sheet.createRow(6);
        sepGraphRow.createCell(0).setCellValue(Strings.GRAPH_ROW_SEP);
        sepGraphRow.createCell(1).setCellValue(sepSsgCount);
        sepGraphRow.createCell(2).setCellValue(sepCmCount);
        
        Row octGraphRow = sheet.createRow(7);
        octGraphRow.createCell(0).setCellValue(Strings.GRAPH_ROW_OCT);
        octGraphRow.createCell(1).setCellValue(octSsgCount);
        octGraphRow.createCell(2).setCellValue(octCmCount);
        
        Row novGraphRow = sheet.createRow(8);
        novGraphRow.createCell(0).setCellValue(Strings.GRAPH_ROW_NOV);
        novGraphRow.createCell(1).setCellValue(novSsgCount);
        novGraphRow.createCell(2).setCellValue(novCmCount);
        
        Row decGraphRow = sheet.createRow(9);
        decGraphRow.createCell(0).setCellValue(Strings.GRAPH_ROW_DEC);
        decGraphRow.createCell(1).setCellValue(decSsgCount);
        decGraphRow.createCell(2).setCellValue(decCmCount);
        
        Row janGraphRow = sheet.createRow(10);
        janGraphRow.createCell(0).setCellValue(Strings.GRAPH_ROW_JAN);
        janGraphRow.createCell(1).setCellValue(janSsgCount);
        janGraphRow.createCell(2).setCellValue(janCmCount);
        
        Row febGraphRow = sheet.createRow(11);
        febGraphRow.createCell(0).setCellValue(Strings.GRAPH_ROW_FEB);
        febGraphRow.createCell(1).setCellValue(febSsgCount);
        febGraphRow.createCell(2).setCellValue(febCmCount);
        
        Row marGraphRow = sheet.createRow(12);
        marGraphRow.createCell(0).setCellValue(Strings.GRAPH_ROW_MAR);
        marGraphRow.createCell(1).setCellValue(marSsgCount);
        marGraphRow.createCell(2).setCellValue(marCmCount);
        
        Row totalGraphRow = sheet.createRow(13);
        totalGraphRow.createCell(0).setCellValue(Strings.GRAPH_ROW_TOTAL);
        totalGraphRow.createCell(1).setCellValue(totalSsgCount);
        totalGraphRow.createCell(2).setCellValue(totalCmCount);
        
        try{
            FileOutputStream fileOut = new FileOutputStream(graphs);
            graphsWorkbook.write(fileOut);
            fileOut.close();
        } catch (IOException ex) {
            Logger.getLogger(PoiHelper.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
    
    /**
     * Returns the counts for the first table on the document
     *
     * @param file
     * @return
     */
    public List<Integer> countTable0() {
        List<Integer> list = new ArrayList<>();
        
        List<Integer> t1Integers = countT1Integers();
        int stCount = 0;
        int fCount = 0;
        int gpCount = 0;
        int otherCount = 0;
        int totalCount = 0;
        int casesClosed = t1Integers.get(0);
        int casesOClosed = t1Integers.get(1);
        
        XSSFSheet refSheet = pswWorkbook.getSheetAt(pswWorkbook.getSheetIndex(Strings.PSW_SHEET_REFERRALS));
        int gradeColNo = getCellColumnByString(Strings.PSW_COLUMN_GRADE, refSheet);
        Iterator<Row> rowIterator = refSheet.iterator();
        
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            if (!isRowEmpty(row)) {
                if (!isCellEmpty(row.getCell(gradeColNo))) {
                    switch (row.getCell(gradeColNo).getStringCellValue()) {
                        case Strings.PSW_GRADE_FY1:
                        case Strings.PSW_GRADE_FY2:
                            fCount++;
                            break;
                        case Strings.PSW_GRADE_ST1:
                        case Strings.PSW_GRADE_ST2:
                        case Strings.PSW_GRADE_ST3:
                        case Strings.PSW_GRADE_ST4:
                        case Strings.PSW_GRADE_ST5:
                        case Strings.PSW_GRADE_ST6:
                        case Strings.PSW_GRADE_ST7:
                        case Strings.PSW_GRADE_ST8:
                        case Strings.PSW_GRADE_CT1:
                        case Strings.PSW_GRADE_CT2:
                        case Strings.PSW_GRADE_CT3:
                            stCount++;
                            break;
                        case Strings.PSW_GRADE_GPST1:
                        case Strings.PSW_GRADE_GPST2:
                        case Strings.PSW_GRADE_GPST3:
                            gpCount++;
                            break;
                        case Strings.PSW_GRADE_DF1:
                        case Strings.PSW_GRADE_DF2:
                        case Strings.PSW_GRADE_DCT1:
                        case Strings.PSW_GRADE_DCT2:
                        case Strings.PSW_GRADE_PHARMACY:
                            otherCount++;
                            break;
                        default:
                            break;
                    }
                }
            }
        }
        
        totalCount = fCount + stCount + gpCount + otherCount;
        
        list.add(stCount);
        list.add(fCount);
        list.add(gpCount);
        list.add(otherCount);
        list.add(totalCount);
        list.add(casesClosed);
        list.add(casesOClosed);
        return list;
    }
    
    /**
     * Counts the values for table number 2: Referral Reason and gender split
     * for financial years... It uses "Graphs" spreadsheet, since this data is
     * already calculated
     *
     * @param graphs
     * @param psw
     * @return
     */
    public List<Integer> countTable1() {
        
        List<Integer> list = new ArrayList<>();
        
        int capabilityMCount = 0;
        int capabilityFCount = 0;
        int anxietyMCount = 0;
        int anxietyFCount = 0;
        int carreerMCount = 0;
        int carreerFCount = 0;
        int clinicalMCount = 0;
        int clinicalFCount = 0;
        int communicationMCount = 0;
        int communicationFCount = 0;
        int conductMCount = 0;
        int conductFCount = 0;
        int culturalMCount = 0;
        int culturalFCount = 0;
        int examMCount = 0;
        int examFCount = 0;
        int phHealthMCount = 0;
        int phHealthFCount = 0;
        int menHealthMCount = 0;
        int menHealthFCount = 0;
        int languageMCount = 0;
        int languageFCount = 0;
        int profMCount = 0;
        int profFCount = 0;
        int adhdMCount = 0;
        int adhdFCount = 0;
        int asdMCount = 0;
        int asdFCount = 0;
        int dyslexiaMCount = 0;
        int dyslexiaFCount = 0;
        int dyspraxiaMCount = 0;
        int dyspraxiaFCount = 0;
        int srttMCount = 0;
        int srttFCount = 0;
        int teamMCount = 0;
        int teamFCount = 0;
        int timeMCount = 0;
        int timeFCount = 0;
        int otherMCount = 0;
        int otherFCount = 0;
        XSSFSheet graphSheet = graphsWorkbook.getSheetAt(graphsWorkbook.getSheetIndex(Strings.GRAPH_SHEET_2));
        
        int anxietyColNo = getCellColumnByString(Strings.PSW_COLUMN_ANXIETY, graphSheet);
        int capColNo = getCellColumnByString(Strings.PSW_COLUMN_CAPABILITY, graphSheet);
        int carreerColNo = getCellColumnByString(Strings.PSW_COLUMN_CARREER, graphSheet);
        int clinSkillsColNo = getCellColumnByString(Strings.PSW_COLUMN_CLINICAL_SKILLS, graphSheet);
        int communicationColNo = getCellColumnByString(Strings.PSW_COLUMN_COMMUNICATION, graphSheet);
        int conductColNo = getCellColumnByString(Strings.PSW_COLUMN_CONDUCT, graphSheet);
        int culturalColNo = getCellColumnByString(Strings.PSW_COLUMN_CULTURAL, graphSheet);
        int examColNo = getCellColumnByString(Strings.PSW_COLUMN_EXAM, graphSheet);
        int mentalHealthColNo = getCellColumnByString(Strings.PSW_COLUMN_MENTAL, graphSheet);
        int physicalHealthColNo = getCellColumnByString(Strings.PSW_COLUMN_PHYSICAL, graphSheet);
        int languageColNo = getCellColumnByString(Strings.PSW_COLUMN_LANGUAGE, graphSheet);
        int professionalismColNo = getCellColumnByString(Strings.PSW_COLUMN_PROFFESSIONALISM, graphSheet);
        int adhdColNo = getCellColumnByString(Strings.PSW_COLUMN_ADHD, graphSheet);
        int asdColNo = getCellColumnByString(Strings.PSW_COLUMN_ASD, graphSheet);
        int dyslexiaColNo = getCellColumnByString(Strings.PSW_COLUMN_DYSLEXIA, graphSheet);
        int dyspraxiaColNo = getCellColumnByString(Strings.PSW_COLUMN_DYSPRAXIA, graphSheet);
        int srttColNo = getCellColumnByString(Strings.PSW_COLUMN_SRTT, graphSheet);
        int teamColNo = getCellColumnByString(Strings.PSW_COLUMN_TEAM, graphSheet);
        int timeColNo = getCellColumnByString(Strings.PSW_COLUMN_TIME, graphSheet);
        int otherColNo = getCellColumnByString(Strings.GRAPH_COLUMN_OTHER, graphSheet);
        int femRowNo = getCellRowByString(Strings.PSW_COLUMN_FEMALE, graphSheet);
        int maleRowNo = getCellRowByString(Strings.PSW_COLUMN_MALE, graphSheet);
        
        anxietyFCount = (int) graphSheet.getRow(femRowNo).getCell(anxietyColNo).getNumericCellValue();
        capabilityFCount = (int) graphSheet.getRow(femRowNo).getCell(capColNo).getNumericCellValue();
        carreerFCount = (int) graphSheet.getRow(femRowNo).getCell(carreerColNo).getNumericCellValue();
        clinicalFCount = (int) graphSheet.getRow(femRowNo).getCell(clinSkillsColNo).getNumericCellValue();
        communicationFCount = (int) graphSheet.getRow(femRowNo).getCell(communicationColNo).getNumericCellValue();
        conductFCount = (int) graphSheet.getRow(femRowNo).getCell(conductColNo).getNumericCellValue();
        culturalFCount = (int) graphSheet.getRow(femRowNo).getCell(culturalColNo).getNumericCellValue();
        examFCount = (int) graphSheet.getRow(femRowNo).getCell(examColNo).getNumericCellValue();
        menHealthFCount = (int) graphSheet.getRow(femRowNo).getCell(mentalHealthColNo).getNumericCellValue();
        phHealthFCount = (int) graphSheet.getRow(femRowNo).getCell(physicalHealthColNo).getNumericCellValue();
        languageFCount = (int) graphSheet.getRow(femRowNo).getCell(languageColNo).getNumericCellValue();
        profFCount = (int) graphSheet.getRow(femRowNo).getCell(professionalismColNo).getNumericCellValue();
        adhdFCount = (int) graphSheet.getRow(femRowNo).getCell(adhdColNo).getNumericCellValue();
        asdFCount = (int) graphSheet.getRow(femRowNo).getCell(asdColNo).getNumericCellValue();
        dyslexiaFCount = (int) graphSheet.getRow(femRowNo).getCell(dyslexiaColNo).getNumericCellValue();
        dyspraxiaFCount = (int) graphSheet.getRow(femRowNo).getCell(dyspraxiaColNo).getNumericCellValue();
        srttFCount = (int) graphSheet.getRow(femRowNo).getCell(srttColNo).getNumericCellValue();
        teamFCount = (int) graphSheet.getRow(femRowNo).getCell(teamColNo).getNumericCellValue();
        timeFCount = (int) graphSheet.getRow(femRowNo).getCell(timeColNo).getNumericCellValue();
        otherFCount = (int) graphSheet.getRow(femRowNo).getCell(otherColNo).getNumericCellValue();
        
        anxietyMCount = (int) graphSheet.getRow(maleRowNo).getCell(anxietyColNo).getNumericCellValue();
        capabilityMCount = (int) graphSheet.getRow(maleRowNo).getCell(capColNo).getNumericCellValue();
        carreerMCount = (int) graphSheet.getRow(maleRowNo).getCell(carreerColNo).getNumericCellValue();
        clinicalMCount = (int) graphSheet.getRow(maleRowNo).getCell(clinSkillsColNo).getNumericCellValue();
        communicationMCount = (int) graphSheet.getRow(maleRowNo).getCell(communicationColNo).getNumericCellValue();
        conductMCount = (int) graphSheet.getRow(maleRowNo).getCell(conductColNo).getNumericCellValue();
        culturalMCount = (int) graphSheet.getRow(maleRowNo).getCell(culturalColNo).getNumericCellValue();
        examMCount = (int) graphSheet.getRow(maleRowNo).getCell(examColNo).getNumericCellValue();
        menHealthMCount = (int) graphSheet.getRow(maleRowNo).getCell(mentalHealthColNo).getNumericCellValue();
        phHealthMCount = (int) graphSheet.getRow(maleRowNo).getCell(physicalHealthColNo).getNumericCellValue();
        languageMCount = (int) graphSheet.getRow(maleRowNo).getCell(languageColNo).getNumericCellValue();
        profMCount = (int) graphSheet.getRow(maleRowNo).getCell(professionalismColNo).getNumericCellValue();
        adhdMCount = (int) graphSheet.getRow(maleRowNo).getCell(adhdColNo).getNumericCellValue();
        asdMCount = (int) graphSheet.getRow(maleRowNo).getCell(asdColNo).getNumericCellValue();
        dyslexiaMCount = (int) graphSheet.getRow(maleRowNo).getCell(dyslexiaColNo).getNumericCellValue();
        dyspraxiaMCount = (int) graphSheet.getRow(maleRowNo).getCell(dyspraxiaColNo).getNumericCellValue();
        srttMCount = (int) graphSheet.getRow(maleRowNo).getCell(srttColNo).getNumericCellValue();
        teamMCount = (int) graphSheet.getRow(maleRowNo).getCell(teamColNo).getNumericCellValue();
        timeMCount = (int) graphSheet.getRow(maleRowNo).getCell(timeColNo).getNumericCellValue();
        otherMCount = (int) graphSheet.getRow(maleRowNo).getCell(otherColNo).getNumericCellValue();
        
        
        list.add(anxietyFCount);
        list.add(anxietyMCount);
        list.add(capabilityFCount);
        list.add(capabilityMCount);
        list.add(carreerFCount);
        list.add(carreerMCount);
        list.add(clinicalFCount);
        list.add(clinicalMCount);
        list.add(communicationFCount);
        list.add(communicationMCount);
        list.add(conductFCount);
        list.add(conductMCount);
        list.add(culturalFCount);
        list.add(culturalMCount);
        list.add(examFCount);
        list.add(examMCount);
        list.add(menHealthFCount);
        list.add(menHealthMCount);
        list.add(phHealthFCount);
        list.add(phHealthMCount);
        list.add(languageFCount);
        list.add(languageMCount);
        list.add(profFCount);
        list.add(profMCount);
        list.add(adhdFCount);
        list.add(adhdMCount);
        list.add(asdFCount);
        list.add(asdMCount);
        list.add(dyslexiaFCount);
        list.add(dyslexiaMCount);
        list.add(dyspraxiaFCount);
        list.add(dyspraxiaMCount);
        list.add(srttFCount);
        list.add(srttMCount);
        list.add(teamFCount);
        list.add(teamMCount);
        list.add(timeFCount);
        list.add(timeMCount);
        list.add(otherFCount);
        list.add(otherMCount);
        
        return list;
    }
    
    public List<Integer> countTable2() {
        
        List<Integer> list = new ArrayList<>();
        int referredCount = countTotalReferrals();
        int f1TotalCount = 0;
        int f2TotalCount = 0;
        int f1ReferredCount = 0;
        int f2ReferredCount = 0;
        
        try {
            FileInputStream pswFileIn = new FileInputStream(psw);
            XSSFWorkbook pswWorkbook = new XSSFWorkbook(pswFileIn);
            XSSFSheet refSheet = pswWorkbook.getSheetAt(pswWorkbook.getSheetIndex(Strings.PSW_SHEET_REFERRALS));
            int gradeColNo = getCellColumnByString(Strings.PSW_COLUMN_GRADE, refSheet);
            Iterator<Row> rowIteratorRefSheet = refSheet.iterator();
            
            XSSFSheet foundationSheet = pswWorkbook.getSheetAt(pswWorkbook.getSheetIndex(Strings.PSW_SHEET_FOUNDATION));
            
            int countColNo = getCellColumnByString(Strings.PSW_COLUMN_TRAINEE_COUNT, foundationSheet);
            int f1CountRowNo = getCellRowByString(Strings.PSW_ROW_F1COUNT, foundationSheet);
            int f2CountRowNo = getCellRowByString(Strings.PSW_ROW_F2COUNT, foundationSheet);
            
            while (rowIteratorRefSheet.hasNext()) {
                Row row = rowIteratorRefSheet.next();
                if (row.getRowNum() > 3 && row.getCell(gradeColNo) == null) {
                    break;
                }
                if (row.getRowNum() > 3 && !row.getCell(gradeColNo).getStringCellValue().equals("") && row.getCell(gradeColNo).getStringCellValue().equals(Strings.PSW_GRADE_FY1)) {
                    f1ReferredCount++;
                }
                if (row.getRowNum() > 3 && row.getCell(gradeColNo).getStringCellValue().equals(Strings.PSW_GRADE_FY2)) {
                    f2ReferredCount++;
                }
            }
            
            f1TotalCount = (int) foundationSheet.getRow(f1CountRowNo).getCell(countColNo).getNumericCellValue();
            f2TotalCount = (int) foundationSheet.getRow(f2CountRowNo).getCell(countColNo).getNumericCellValue();
            list.add(referredCount);
            list.add(f1TotalCount);
            list.add(f2TotalCount);
            list.add(f1ReferredCount);
            list.add(f2ReferredCount);
            
        } catch (IOException ex) {
            Logger.getLogger(PoiHelper.class.getName()).log(Level.SEVERE, null, ex);
        }
        
        return list;
    }
    
    public List<Integer> countTable3() {
        
        List<Integer> list = new ArrayList<>();
        
        int bournemouthRefNo = 0;
        int dorchesterRefNo = 0;
        int dorsetRefNo = 0;
        int hhftRefNo = 0;
        int iowRefNo = 0;
        int jerseyRefNo = 0;
        int pooleRefNo = 0;
        int portsmouthRefNo = 0;
        int salisburyRefNo = 0;
        int solentRefNo = 0;
        int southamptonRefNo = 0;
        int southernRefNo = 0;
        
        int bournemouthTotal = 0;
        int dorchesterTotal = 0;
        int dorsetTotal = 0;
        int hhftTotal = 0;
        int iowTotal = 0;
        int jerseyTotal = 0;
        int pooleTotal = 0;
        int portsmouthTotal = 0;
        int salisburyTotal = 0;
        int solentTotal = 0;
        int southamptonTotal = 0;
        int southernTotal = 0;
        
        int totalRefs = countTotalReferrals();
        int totalWessex = 0;
        
        XSSFSheet refSheet = pswWorkbook.getSheetAt(pswWorkbook.getSheetIndex(Strings.PSW_SHEET_REFERRALS));
        int trustRColNo = getCellColumnByString(Strings.PSW_COLUMN_TRUST, refSheet);
        Iterator<Row> rowIteratorRefSheet = refSheet.iterator();
        
        XSSFSheet trustSheet = pswWorkbook.getSheetAt(pswWorkbook.getSheetIndex(Strings.PSW_SHEET_TRUST));
        int traineeCountColNo = getCellColumnByString(Strings.PSW_COLUMN_TRAINEE_COUNT, trustSheet);
        int trustColNo = getCellColumnByString(Strings.PSW_COLUMN_TRUST, trustSheet);
        Iterator<Row> rowIteratorTrustSheet = trustSheet.iterator();
        
        totalWessex = countTotalWessex();
        
        while (rowIteratorTrustSheet.hasNext()) {
            Row row = rowIteratorTrustSheet.next();
            switch (row.getCell(trustColNo).getStringCellValue()) {
                case Strings.PSW_ROW_BOURNEMOUTH:
                    bournemouthTotal = (int) row.getCell(traineeCountColNo).getNumericCellValue();
                    break;
                case Strings.PSW_ROW_DORCHESTER:
                    dorchesterTotal = (int) row.getCell(traineeCountColNo).getNumericCellValue();
                    break;
                case Strings.PSW_ROW_DORSET:
                    dorsetTotal = (int) row.getCell(traineeCountColNo).getNumericCellValue();
                    break;
                case Strings.PSW_ROW_HHFT:
                    hhftTotal = (int) row.getCell(traineeCountColNo).getNumericCellValue();
                    break;
                case Strings.PSW_ROW_IOW:
                    iowTotal = (int) row.getCell(traineeCountColNo).getNumericCellValue();
                    break;
                case Strings.PSW_ROW_JERSEY:
                    jerseyTotal = (int) row.getCell(traineeCountColNo).getNumericCellValue();
                    break;
                case Strings.PSW_ROW_POOLE:
                    pooleTotal = (int) row.getCell(traineeCountColNo).getNumericCellValue();
                    break;
                case Strings.PSW_ROW_PORTSMOUTH:
                    portsmouthTotal = (int) row.getCell(traineeCountColNo).getNumericCellValue();
                    break;
                case Strings.PSW_ROW_SALISBURY:
                    salisburyTotal = (int) row.getCell(traineeCountColNo).getNumericCellValue();
                    break;
                case Strings.PSW_ROW_SOLENT:
                    solentTotal = (int) row.getCell(traineeCountColNo).getNumericCellValue();
                    break;
                case Strings.PSW_ROW_SOUTHAMPTON:
                    southamptonTotal = (int) row.getCell(traineeCountColNo).getNumericCellValue();
                    break;
                case Strings.PSW_ROW_SOUTHERN:
                    southernTotal = (int) row.getCell(traineeCountColNo).getNumericCellValue();
                    break;
            }
            
        }
        
        while (rowIteratorRefSheet.hasNext()) {
            Row row = rowIteratorRefSheet.next();
            if (!isRowEmpty(row)) {
                if (!isCellEmpty(row.getCell(trustRColNo))) {
                    
                    if (row.getCell(trustRColNo).getStringCellValue().contains(Strings.PSW_TRUST_BOURNEMOUTH)) {
                        bournemouthRefNo++;
                    } else if (row.getCell(trustRColNo).getStringCellValue().contains(Strings.PSW_TRUST_DORCHESTER)) {
                        dorchesterRefNo++;
                    } else if (row.getCell(trustRColNo).getStringCellValue().contains(Strings.PSW_TRUST_DORSET)) {
                        dorsetRefNo++;
                    } else if (row.getCell(trustRColNo).getStringCellValue().contains(Strings.PSW_TRUST_HHFT)) {
                        hhftRefNo++;
                    } else if (row.getCell(trustRColNo).getStringCellValue().contains(Strings.PSW_TRUST_IOW)) {
                        iowRefNo++;
                    } else if (row.getCell(trustRColNo).getStringCellValue().contains(Strings.PSW_TRUST_JERSEY)) {
                        jerseyRefNo++;
                    } else if (row.getCell(trustRColNo).getStringCellValue().contains(Strings.PSW_TRUST_POOLE)) {
                        pooleRefNo++;
                    } else if (row.getCell(trustRColNo).getStringCellValue().contains(Strings.PSW_TRUST_PORTSMOUTH)) {
                        portsmouthRefNo++;
                    } else if (row.getCell(trustRColNo).getStringCellValue().contains(Strings.PSW_TRUST_SALISBURY)) {
                        salisburyRefNo++;
                    } else if (row.getCell(trustRColNo).getStringCellValue().contains(Strings.PSW_TRUST_SOLENT)) {
                        solentRefNo++;
                    } else if (row.getCell(trustRColNo).getStringCellValue().contains(Strings.PSW_TRUST_SOUTHAMPTON)) {
                        southamptonRefNo++;
                    } else if (row.getCell(trustRColNo).getStringCellValue().contains(Strings.PSW_TRUST_SOUTHERN)) {
                        southernRefNo++;
                    }
                    
                }
            }
        }
        
        
        list.add(bournemouthTotal);
        list.add(dorchesterTotal);
        list.add(dorsetTotal);
        list.add(hhftTotal);
        list.add(iowTotal);
        list.add(jerseyTotal);
        list.add(pooleTotal);
        list.add(portsmouthTotal);
        list.add(salisburyTotal);
        list.add(solentTotal);
        list.add(southamptonTotal);
        list.add(southernTotal);
        list.add(bournemouthRefNo);
        list.add(dorchesterRefNo);
        list.add(dorsetRefNo);
        list.add(hhftRefNo);
        list.add(iowRefNo);
        list.add(jerseyRefNo);
        list.add(pooleRefNo);
        list.add(portsmouthRefNo);
        list.add(salisburyRefNo);
        list.add(solentRefNo);
        list.add(southamptonRefNo);
        list.add(southernRefNo);
        list.add(totalRefs);
        list.add(totalWessex);
        
        return list;
    }
    
    public List<Integer> countTable4() {
        
        List<Integer> list = new ArrayList<>();
        
        int anaestheticsRefNo = 0;
        int dentalRefNo = 0;
        int emergRefNo = 0;
        int foundationRefNo = 0;
        int gpRefNo = 0;
        int medicineRefNo = 0;
        int obsRefNo = 0;
        int occhealthRefNo = 0;
        int paediatricsRefNo = 0;
        int pathologyRefNo = 0;
        int pharmacyRefNo = 0;
        int psychRefNo = 0;
        int pubhealthRefNo = 0;
        int radioRefNo = 0;
        int surgeryRefNo = 0;
        
        double anaestheticsTotal = 0;
        double dentalTotal = 0;
        double emergTotal = 0;
        double foundationTotal = 0;
        double gpTotal = 0;
        double medicineTotal = 0;
        double obsTotal = 0;
        double occhealthTotal = 0;
        double paediatricsTotal = 0;
        double pathologyTotal = 0;
        double pharmacyTotal = 0;
        double psychTotal = 0;
        double pubhealthTotal = 0;
        double radioTotal = 0;
        double surgeryTotal = 0;
        
        int totalRefs = countTotalReferrals();
        int totalWssx = countTotalWessex();
        
        XSSFSheet refSheet = pswWorkbook.getSheetAt(pswWorkbook.getSheetIndex(Strings.PSW_SHEET_REFERRALS));
        Iterator<Row> rowIteratorRefSheet = refSheet.rowIterator();
        int specialtyRColNo = getCellColumnByString(Strings.PSW_COLUMN_SPECIALTY, refSheet);
        
        XSSFSheet specialtySheet = pswWorkbook.getSheetAt(pswWorkbook.getSheetIndex(Strings.PSW_SHEET_SPECIALTY));
        Iterator<Row> rowIteratorSpcSheet = specialtySheet.rowIterator();
        int specialtyColNo = getCellColumnByString(Strings.PSW_COLUMN_SPECIALTY, specialtySheet);
        int traineeCountColNo = getCellColumnByString(Strings.PSW_COLUMN_TRAINEE_COUNT, specialtySheet);
        
        while (rowIteratorRefSheet.hasNext()) {
            Row row = rowIteratorRefSheet.next();
            
            if (!isRowEmpty(row)) {
                if (!isCellEmpty(row.getCell(specialtyRColNo))) {
                    Cell c = row.getCell(specialtyRColNo);
                    if (c.getStringCellValue().contains(Strings.PSW_SPECIALTY_ANAESTHETICS)) {
                        anaestheticsRefNo++;
                    } else if (c.getStringCellValue().contains(Strings.PSW_SPECIALTY_DENTAL)) {
                        dentalRefNo++;
                    } else if (c.getStringCellValue().contains(Strings.PSW_SPECIALTY_EMERGENCY)) {
                        emergRefNo++;
                    } else if (c.getStringCellValue().contains(Strings.PSW_SPECIALTY_FOUNDATION)) {
                        foundationRefNo++;
                    } else if (c.getStringCellValue().contains(Strings.PSW_SPECIALTY_GENERAL_PRACTICE)) {
                        gpRefNo++;
                    } else if (c.getStringCellValue().contains(Strings.PSW_SPECIALTY_MEDICINE)) {
                        medicineRefNo++;
                    } else if (c.getStringCellValue().contains(Strings.PSW_SPECIALTY_OBS)) {
                        obsRefNo++;
                    } else if (c.getStringCellValue().contains(Strings.PSW_SPECIALTY_OCC_HEALTH)) {
                        occhealthRefNo++;
                    } else if (c.getStringCellValue().contains(Strings.PSW_SPECIALTY_PAEDIATRICS)) {
                        paediatricsRefNo++;
                    } else if (c.getStringCellValue().contains(Strings.PSW_SPECIALTY_PATHOLOGY)) {
                        pathologyRefNo++;
                    } else if (c.getStringCellValue().contains(Strings.PSW_SPECIALTY_PHARMACY)) {
                        pharmacyRefNo++;
                    } else if (c.getStringCellValue().contains(Strings.PSW_SPECIALTY_PSYCHIATRY)) {
                        psychRefNo++;
                    } else if (c.getStringCellValue().contains(Strings.PSW_SPECIALTY_PUBLIC_HEALTH)) {
                        pubhealthRefNo++;
                    } else if (c.getStringCellValue().contains(Strings.PSW_SPECIALTY_RADIOLOGY)) {
                        radioRefNo++;
                    } else if (c.getStringCellValue().contains(Strings.PSW_SPECIALTY_SURGERY)) {
                        surgeryRefNo++;
                    }
                }
            }
        }
        
        while (rowIteratorSpcSheet.hasNext()) {
            Row row = rowIteratorSpcSheet.next();
            
            if (!isRowEmpty(row)) {
                if (!isCellEmpty(row.getCell(specialtyColNo))) {
                    switch (row.getCell(specialtyColNo).getStringCellValue()) {
                        case Strings.PSW_ROW_ANAESTHETICS_A:
                        case Strings.PSW_ROW_ANAESTHETICS_B:
                            anaestheticsTotal = anaestheticsTotal + row.getCell(traineeCountColNo).getNumericCellValue();
                            break;
                        case Strings.PSW_ROW_DENTAL_A:
                        case Strings.PSW_ROW_DENTAL_B:
                            dentalTotal = dentalTotal + row.getCell(traineeCountColNo).getNumericCellValue();
                            break;
                        case Strings.PSW_ROW_EMERGENCY_A:
                        case Strings.PSW_ROW_EMERGENCY_B:
                            emergTotal = emergTotal + row.getCell(traineeCountColNo).getNumericCellValue();
                            break;
                        case Strings.PSW_ROW_FOUNDATION:
                            foundationTotal = row.getCell(traineeCountColNo).getNumericCellValue();
                            break;
                        case Strings.PSW_ROW_GP_A:
                        case Strings.PSW_ROW_GP_B:
                        case Strings.PSW_ROW_GP_C:
                        case Strings.PSW_ROW_GP_D:
                        case Strings.PSW_ROW_GP_E:
                        case Strings.PSW_ROW_GP_F:
                        case Strings.PSW_ROW_GP_G:
                        case Strings.PSW_ROW_GP_H:
                        case Strings.PSW_ROW_GP_I:
                            gpTotal = gpTotal + row.getCell(traineeCountColNo).getNumericCellValue();
                            break;
                        case Strings.PSW_ROW_MEDICINE_A:
                        case Strings.PSW_ROW_MEDICINE_B:
                        case Strings.PSW_ROW_MEDICINE_C:
                        case Strings.PSW_ROW_MEDICINE_D:
                        case Strings.PSW_ROW_MEDICINE_E:
                        case Strings.PSW_ROW_MEDICINE_F:
                        case Strings.PSW_ROW_MEDICINE_G:
                        case Strings.PSW_ROW_MEDICINE_H:
                        case Strings.PSW_ROW_MEDICINE_I:
                        case Strings.PSW_ROW_MEDICINE_J:
                            medicineTotal = medicineTotal + row.getCell(traineeCountColNo).getNumericCellValue();
                            break;
                        case Strings.PSW_ROW_OBS:
                            obsTotal = row.getCell(traineeCountColNo).getNumericCellValue();
                            break;
                        case Strings.PSW_ROW_OC_HEALTH:
                            occhealthTotal = row.getCell(traineeCountColNo).getNumericCellValue();
                            break;
                        case Strings.PSW_ROW_PAEDIATRICS_A:
                        case Strings.PSW_ROW_PAEDIATRICS_B:
                            paediatricsTotal = paediatricsTotal + row.getCell(traineeCountColNo).getNumericCellValue();
                            break;
                        case Strings.PSW_ROW_PATHOLOGY_A:
                        case Strings.PSW_ROW_PATHOLOGY_B:
                            pathologyTotal = pathologyTotal + row.getCell(traineeCountColNo).getNumericCellValue();
                            break;
                        case Strings.PSW_ROW_PSYCH_A:
                        case Strings.PSW_ROW_PSYCH_B:
                        case Strings.PSW_ROW_PSYCH_C:
                        case Strings.PSW_ROW_PSYCH_D:
                        case Strings.PSW_ROW_PSYCH_E:
                        case Strings.PSW_ROW_PSYCH_F:
                        case Strings.PSW_ROW_PSYCH_G:
                            psychTotal = psychTotal + pathologyTotal + row.getCell(traineeCountColNo).getNumericCellValue();
                            break;
                        case Strings.PSW_ROW_PUBLIC_HEALTH:
                            pubhealthTotal = row.getCell(traineeCountColNo).getNumericCellValue();
                            break;
                        case Strings.PSW_ROW_RADIOLOGY:
                            radioTotal = row.getCell(traineeCountColNo).getNumericCellValue();
                            break;
                        case Strings.PSW_ROW_SURGERY_A:
                        case Strings.PSW_ROW_SURGERY_B:
                        case Strings.PSW_ROW_SURGERY_C:
                        case Strings.PSW_ROW_SURGERY_D:
                        case Strings.PSW_ROW_SURGERY_E:
                        case Strings.PSW_ROW_SURGERY_F:
                        case Strings.PSW_ROW_SURGERY_G:
                        case Strings.PSW_ROW_SURGERY_H:
                        case Strings.PSW_ROW_SURGERY_I:
                            surgeryTotal = surgeryTotal + row.getCell(traineeCountColNo).getNumericCellValue();
                            break;
                    }
                }
            }
        }
        
        list.add(anaestheticsRefNo);
        list.add(dentalRefNo);
        list.add(emergRefNo);
        list.add(foundationRefNo);
        list.add(gpRefNo);
        list.add(medicineRefNo);
        list.add(obsRefNo);
        list.add(occhealthRefNo);
        list.add(paediatricsRefNo);
        list.add(pathologyRefNo);
        list.add(pharmacyRefNo);
        list.add(psychRefNo);
        list.add(pubhealthRefNo);
        list.add(radioRefNo);
        list.add(surgeryRefNo);
        list.add((int) anaestheticsTotal);
        list.add((int) dentalTotal);
        list.add((int) emergTotal);
        list.add((int) foundationTotal);
        list.add((int) gpTotal);
        list.add((int) medicineTotal);
        list.add((int) obsTotal);
        list.add((int) occhealthTotal);
        list.add((int) paediatricsTotal);
        list.add((int) pathologyTotal);
        list.add((int) pharmacyTotal);
        list.add((int) psychTotal);
        list.add((int) pubhealthTotal);
        list.add((int) radioTotal);
        list.add((int) surgeryTotal);
        list.add(totalRefs);
        list.add(totalWssx);
        
        
        return list;
    }
    
    public ArrayList<RefRecord> getTable5LineByTrust(String trst) {
        
        ArrayList<RefRecord> list = new ArrayList<>();
        
        for(RefRecord record:recordList){
            if ((record.getTrust().equals(trst)||record.getTrust().contains(trst))&&record.isExam()) {
                list.add(record);
            }
        }
        
        return list;
    }
    
    public List<Double> countTable7(){
        List<Double> list = new ArrayList<>();
        
        XSSFSheet costs = graphsWorkbook.getSheet(Strings.GRAPH_SHEET_12);
        Double ssg = costs.getRow(13).getCell(1).getNumericCellValue();
        Double cm = costs.getRow(13).getCell(2).getNumericCellValue();
        Double total = ssg+cm;
        
        list.add(ssg);
        list.add(cm);
        list.add(total);
        
        return list;
    }
    
    public ArrayList<Table8Line> countTable8(){
        ArrayList<Table8Line> list = new ArrayList<>();
        
        Table8Line anaesthetics = new Table8Line();
        Table8Line dental = new Table8Line();
        Table8Line dermatology = new Table8Line();
        Table8Line endocrinology = new Table8Line();
        Table8Line foundation = new Table8Line();
        Table8Line gastroenterology = new Table8Line();
        Table8Line gp = new Table8Line();
        Table8Line haematology = new Table8Line();
        Table8Line histopathology = new Table8Line();
        Table8Line emergMed = new Table8Line();
        Table8Line medicine = new Table8Line();
        Table8Line neurology = new Table8Line();
        Table8Line obs = new Table8Line();
        Table8Line occHealth = new Table8Line();
        Table8Line oncology = new Table8Line();
        Table8Line ophtalmology = new Table8Line();
        Table8Line paediatrics = new Table8Line();
        Table8Line pathology = new Table8Line();
        Table8Line pharmacy = new Table8Line();
        Table8Line psych = new Table8Line();
        Table8Line pubHealth = new Table8Line();
        Table8Line radiology = new Table8Line();
        Table8Line sexHealth = new Table8Line();
        Table8Line rheumathology = new Table8Line();
        Table8Line surgery = new Table8Line();
        
        for(RefRecord record:recordList){
            if(record.getSpecialty().equals(Strings.PSW_SPECIALTY_ANAESTHETICS)){
                anaesthetics = getLineBySpc(anaesthetics, record);
            }
            else if(record.getSpecialty().contains(Strings.PSW_SPECIALTY_DENTAL)){
                dental = getLineBySpc(dental, record);
            }
            else if(record.getSpecialty().contains(Strings.PSW_SPECIALTY_DERMATOLOGY)){
                dermatology = getLineBySpc(dermatology, record);
            }
            else if(record.getSpecialty().contains(Strings.PSW_SPECIALTY_ENDOCRINOLOGY)){
                endocrinology = getLineBySpc(endocrinology, record);
            }
            else if(record.getSpecialty().contains(Strings.PSW_SPECIALTY_FOUNDATION)){
                foundation = getLineBySpc(foundation, record);
            }
            else if(record.getSpecialty().contains(Strings.PSW_SPECIALTY_GASTRO)){
                gastroenterology = getLineBySpc(gastroenterology, record);
            }
            else if(record.getSpecialty().contains(Strings.PSW_SPECIALTY_GENERAL_PRACTICE)){
                gp = getLineBySpc(gp, record);
            }
            else if(record.getSpecialty().contains(Strings.PSW_SPECIALTY_HAEMATOLOGY)){
                haematology = getLineBySpc(haematology, record);
            }
            else if(record.getSpecialty().contains(Strings.PSW_SPECIALTY_HISTOPATHOLOGY)){
                histopathology = getLineBySpc(histopathology, record);
            }
            else if(record.getSpecialty().contains(Strings.PSW_SPECIALTY_EMERGENCY)){
                emergMed = getLineBySpc(emergMed, record);
            }
            else if(record.getSpecialty().contains(Strings.PSW_SPECIALTY_MEDICINE)){
                medicine = getLineBySpc(medicine, record);
            }
            else if(record.getSpecialty().contains(Strings.PSW_SPECIALTY_NEUROLOGY)){
                neurology = getLineBySpc(neurology, record);
            }
            else if(record.getSpecialty().contains(Strings.PSW_SPECIALTY_OBS)){
                obs = getLineBySpc(obs, record);
            }
            else if(record.getSpecialty().contains(Strings.PSW_SPECIALTY_OCC_HEALTH)){
                occHealth = getLineBySpc(occHealth, record);
            }
            else if(record.getSpecialty().contains(Strings.PSW_SPECIALTY_ONCOLOGY)){
                oncology = getLineBySpc(oncology, record);
            }
            else if(record.getSpecialty().contains(Strings.PSW_SPECIALTY_OPHTHALMOLOGY)){
                ophtalmology = getLineBySpc(ophtalmology, record);
            }
            else if(record.getSpecialty().contains(Strings.PSW_SPECIALTY_PAEDIATRICS)){
                paediatrics = getLineBySpc(paediatrics, record);
            }
            else if(record.getSpecialty().contains(Strings.PSW_SPECIALTY_PATHOLOGY)){
                pathology = getLineBySpc(pathology, record);
            }
            else if(record.getSpecialty().contains(Strings.PSW_SPECIALTY_PHARMACY)){
                pharmacy = getLineBySpc(pharmacy, record);
            }
            else if(record.getSpecialty().contains(Strings.PSW_SPECIALTY_PSYCHIATRY)){
                psych = getLineBySpc(psych, record);
            }
            else if(record.getSpecialty().contains(Strings.PSW_SPECIALTY_PUBLIC_HEALTH)){
                pubHealth = getLineBySpc(pubHealth, record);
            }
            else if(record.getSpecialty().contains(Strings.PSW_SPECIALTY_RADIOLOGY)){
                radiology = getLineBySpc(radiology, record);
            }
            else if(record.getSpecialty().contains(Strings.PSW_SPECIALTY_SEXHEALTH)){
                sexHealth = getLineBySpc(sexHealth, record);
            }
            else if(record.getSpecialty().contains(Strings.PSW_SPECIALTY_RHEUMATOLOGY)){
                rheumathology = getLineBySpc(rheumathology, record);
            }
            else if(record.getSpecialty().contains(Strings.PSW_SPECIALTY_SURGERY)){
                surgery = getLineBySpc(surgery, record);
            }
        }
        
        anaesthetics.setTitle(Strings.PSW_SPECIALTY_ANAESTHETICS);
        dental.setTitle(Strings.PSW_SPECIALTY_DENTAL);
        dermatology.setTitle(Strings.PSW_SPECIALTY_DERMATOLOGY);
        endocrinology.setTitle(Strings.PSW_SPECIALTY_ENDOCRINOLOGY);
        foundation.setTitle(Strings.PSW_SPECIALTY_FOUNDATION);
        gastroenterology.setTitle(Strings.PSW_SPECIALTY_GASTRO);
        gp.setTitle(Strings.PSW_SPECIALTY_GENERAL_PRACTICE);
        haematology.setTitle(Strings.PSW_SPECIALTY_HAEMATOLOGY);
        histopathology.setTitle(Strings.PSW_SPECIALTY_HISTOPATHOLOGY);
        emergMed.setTitle(Strings.PSW_SPECIALTY_EMERGENCY);
        medicine.setTitle(Strings.PSW_SPECIALTY_MEDICINE_TITLE);
        neurology.setTitle(Strings.PSW_SPECIALTY_NEUROLOGY);
        obs.setTitle(Strings.PSW_SPECIALTY_OBS);
        occHealth.setTitle(Strings.PSW_SPECIALTY_OCC_HEALTH);
        oncology.setTitle(Strings.PSW_SPECIALTY_ONCOLOGY);
        ophtalmology.setTitle(Strings.PSW_SPECIALTY_OPHTHALMOLOGY);
        paediatrics.setTitle(Strings.PSW_SPECIALTY_PAEDIATRICS);
        pathology.setTitle(Strings.PSW_SPECIALTY_PATHOLOGY);
        psych.setTitle(Strings.PSW_SPECIALTY_PSYCHIATRY);
        pharmacy.setTitle(Strings.PSW_SPECIALTY_PHARMACY);
        pubHealth.setTitle(Strings.PSW_SPECIALTY_PUBLIC_HEALTH);
        radiology.setTitle(Strings.PSW_SPECIALTY_RADIOLOGY);
        sexHealth.setTitle(Strings.PSW_SPECIALTY_SEXHEALTH);
        rheumathology.setTitle(Strings.PSW_SPECIALTY_RHEUMATOLOGY);
        surgery.setTitle(Strings.PSW_SPECIALTY_SURGERY);
        
        
        list.add(anaesthetics);
        list.add(dental);
        list.add(dermatology);
        list.add(endocrinology);
        list.add(foundation);
        list.add(gastroenterology);
        list.add(gp);
        list.add(haematology);
        list.add(histopathology);
        list.add(emergMed);
        list.add(medicine);
        list.add(neurology);
        list.add(obs);
        list.add(occHealth);
        list.add(oncology);
        list.add(ophtalmology);
        list.add(paediatrics);
        list.add(pathology);
        list.add(pharmacy);
        list.add(psych);
        list.add(pubHealth);
        list.add(radiology);
        list.add(sexHealth);
        list.add(rheumathology);
        list.add(surgery);
        
        return list;
    }
    
    public ArrayList<Table8Line> countTable9(){
        ArrayList<Table8Line> list = new ArrayList<>();
        
        Table8Line anaesthetics = new Table8Line();
        Table8Line dental = new Table8Line();
        Table8Line dermatology = new Table8Line();
        Table8Line endocrinology = new Table8Line();
        Table8Line foundation = new Table8Line();
        Table8Line gastroenterology = new Table8Line();
        Table8Line gp = new Table8Line();
        Table8Line haematology = new Table8Line();
        Table8Line histopathology = new Table8Line();
        Table8Line emergMed = new Table8Line();
        Table8Line medicine = new Table8Line();
        Table8Line neurology = new Table8Line();
        Table8Line obs = new Table8Line();
        Table8Line occHealth = new Table8Line();
        Table8Line oncology = new Table8Line();
        Table8Line ophtalmology = new Table8Line();
        Table8Line paediatrics = new Table8Line();
        Table8Line pathology = new Table8Line();
        Table8Line pharmacy = new Table8Line();
        Table8Line psych = new Table8Line();
        Table8Line pubHealth = new Table8Line();
        Table8Line radiology = new Table8Line();
        Table8Line sexHealth = new Table8Line();
        Table8Line rheumathology = new Table8Line();
        Table8Line surgery = new Table8Line();
        
        for(RefRecord record:recordList){
            if(record.getSpecialty().equals(Strings.PSW_SPECIALTY_ANAESTHETICS)){
                anaesthetics = getNonUkLineBySpc(anaesthetics, record);
            }
            else if(record.getSpecialty().contains(Strings.PSW_SPECIALTY_DENTAL)){
                dental = getNonUkLineBySpc(dental, record);
            }
            else if(record.getSpecialty().contains(Strings.PSW_SPECIALTY_DERMATOLOGY)){
                dermatology = getNonUkLineBySpc(dermatology, record);
            }
            else if(record.getSpecialty().contains(Strings.PSW_SPECIALTY_ENDOCRINOLOGY)){
                endocrinology = getNonUkLineBySpc(endocrinology, record);
            }
            else if(record.getSpecialty().contains(Strings.PSW_SPECIALTY_FOUNDATION)){
                foundation = getNonUkLineBySpc(foundation, record);
            }
            else if(record.getSpecialty().contains(Strings.PSW_SPECIALTY_GASTRO)){
                gastroenterology = getNonUkLineBySpc(gastroenterology, record);
            }
            else if(record.getSpecialty().contains(Strings.PSW_SPECIALTY_GENERAL_PRACTICE)){
                gp = getNonUkLineBySpc(gp, record);
            }
            else if(record.getSpecialty().contains(Strings.PSW_SPECIALTY_HAEMATOLOGY)){
                haematology = getNonUkLineBySpc(haematology, record);
            }
            else if(record.getSpecialty().contains(Strings.PSW_SPECIALTY_HISTOPATHOLOGY)){
                histopathology = getNonUkLineBySpc(histopathology, record);
            }
            else if(record.getSpecialty().contains(Strings.PSW_SPECIALTY_EMERGENCY)){
                emergMed = getNonUkLineBySpc(emergMed, record);
            }
            else if(record.getSpecialty().contains(Strings.PSW_SPECIALTY_MEDICINE)){
                medicine = getNonUkLineBySpc(medicine, record);
            }
            else if(record.getSpecialty().contains(Strings.PSW_SPECIALTY_NEUROLOGY)){
                neurology = getNonUkLineBySpc(neurology, record);
            }
            else if(record.getSpecialty().contains(Strings.PSW_SPECIALTY_OBS)){
                obs = getNonUkLineBySpc(obs, record);
            }
            else if(record.getSpecialty().contains(Strings.PSW_SPECIALTY_OCC_HEALTH)){
                occHealth = getNonUkLineBySpc(occHealth, record);
            }
            else if(record.getSpecialty().contains(Strings.PSW_SPECIALTY_ONCOLOGY)){
                oncology = getNonUkLineBySpc(oncology, record);
            }
            else if(record.getSpecialty().contains(Strings.PSW_SPECIALTY_OPHTHALMOLOGY)){
                ophtalmology = getNonUkLineBySpc(ophtalmology, record);
            }
            else if(record.getSpecialty().contains(Strings.PSW_SPECIALTY_PAEDIATRICS)){
                paediatrics = getNonUkLineBySpc(paediatrics, record);
            }
            else if(record.getSpecialty().contains(Strings.PSW_SPECIALTY_PATHOLOGY)){
                pathology = getNonUkLineBySpc(pathology, record);
            }
            else if(record.getSpecialty().contains(Strings.PSW_SPECIALTY_PHARMACY)){
                pharmacy = getNonUkLineBySpc(pharmacy, record);
            }
            else if(record.getSpecialty().contains(Strings.PSW_SPECIALTY_PSYCHIATRY)){
                psych = getNonUkLineBySpc(psych, record);
            }
            else if(record.getSpecialty().contains(Strings.PSW_SPECIALTY_PUBLIC_HEALTH)){
                pubHealth = getNonUkLineBySpc(pubHealth, record);
            }
            else if(record.getSpecialty().contains(Strings.PSW_SPECIALTY_RADIOLOGY)){
                radiology = getNonUkLineBySpc(radiology, record);
            }
            else if(record.getSpecialty().contains(Strings.PSW_SPECIALTY_SEXHEALTH)){
                sexHealth = getNonUkLineBySpc(sexHealth, record);
            }
            else if(record.getSpecialty().contains(Strings.PSW_SPECIALTY_RHEUMATOLOGY)){
                rheumathology = getNonUkLineBySpc(rheumathology, record);
            }
            else if(record.getSpecialty().contains(Strings.PSW_SPECIALTY_SURGERY)){
                surgery = getNonUkLineBySpc(surgery, record);
            }
        }
        
        anaesthetics.setTitle(Strings.PSW_SPECIALTY_ANAESTHETICS);
        dental.setTitle(Strings.PSW_SPECIALTY_DENTAL);
        dermatology.setTitle(Strings.PSW_SPECIALTY_DERMATOLOGY);
        endocrinology.setTitle(Strings.PSW_SPECIALTY_ENDOCRINOLOGY);
        foundation.setTitle(Strings.PSW_SPECIALTY_FOUNDATION);
        gastroenterology.setTitle(Strings.PSW_SPECIALTY_GASTRO);
        gp.setTitle(Strings.PSW_SPECIALTY_GENERAL_PRACTICE);
        haematology.setTitle(Strings.PSW_SPECIALTY_HAEMATOLOGY);
        histopathology.setTitle(Strings.PSW_SPECIALTY_HISTOPATHOLOGY);
        emergMed.setTitle(Strings.PSW_SPECIALTY_EMERGENCY);
        medicine.setTitle(Strings.PSW_SPECIALTY_MEDICINE_TITLE);
        neurology.setTitle(Strings.PSW_SPECIALTY_NEUROLOGY);
        obs.setTitle(Strings.PSW_SPECIALTY_OBS);
        occHealth.setTitle(Strings.PSW_SPECIALTY_OCC_HEALTH);
        oncology.setTitle(Strings.PSW_SPECIALTY_ONCOLOGY);
        ophtalmology.setTitle(Strings.PSW_SPECIALTY_OPHTHALMOLOGY);
        paediatrics.setTitle(Strings.PSW_SPECIALTY_PAEDIATRICS);
        pathology.setTitle(Strings.PSW_SPECIALTY_PATHOLOGY);
        psych.setTitle(Strings.PSW_SPECIALTY_PSYCHIATRY);
        pharmacy.setTitle(Strings.PSW_SPECIALTY_PHARMACY);
        pubHealth.setTitle(Strings.PSW_SPECIALTY_PUBLIC_HEALTH);
        radiology.setTitle(Strings.PSW_SPECIALTY_RADIOLOGY);
        sexHealth.setTitle(Strings.PSW_SPECIALTY_SEXHEALTH);
        rheumathology.setTitle(Strings.PSW_SPECIALTY_RHEUMATOLOGY);
        surgery.setTitle(Strings.PSW_SPECIALTY_SURGERY);
        
        
        list.add(anaesthetics);
        list.add(dental);
        list.add(dermatology);
        list.add(endocrinology);
        list.add(foundation);
        list.add(gastroenterology);
        list.add(gp);
        list.add(haematology);
        list.add(histopathology);
        list.add(emergMed);
        list.add(medicine);
        list.add(neurology);
        list.add(obs);
        list.add(occHealth);
        list.add(oncology);
        list.add(ophtalmology);
        list.add(paediatrics);
        list.add(pathology);
        list.add(pharmacy);
        list.add(psych);
        list.add(pubHealth);
        list.add(radiology);
        list.add(sexHealth);
        list.add(rheumathology);
        list.add(surgery);
        
        return list;
    }
    
    
    private Table8Line getLineBySpc(Table8Line line, RefRecord record){
        
        int male = line.getMale();
        int female = line.getFemale();
        int uk= line.getUk();
        int nonUk= line.getNonUk();
        int age2329 = line.getAge2329();
        int age3035= line.getAge3035();
        int age3540 = line.getAge3540();
        int age40 = line.getAge40();
        int whiteb = line.getWhiteb();
        int whiteo = line.getWhiteo();
        int asian = line.getAsian();
        int african = line.getAfrican();
        int ethOther = line.getEthOther();
        int christian = line.getChristian();
        int islam = line.getIslam();
        int hindu = line.getHindu();
        int atheist = line.getAtheist();
        int sikh = line.getSikh();
        int judaism = line.getJudaism();
        int buddhism = line.getBuddhism();
        int relOther = line.getRelOther();
        int relPNS = line.getRelPNS();
        int yes = line.getYes();
        int no = line.getNo();
        int het = line.getHet();
        int homosexual = line.getHomosexual();
        int bisexual = line.getBisexual();
        int sexOrPNS = line.getSexOrPNS();
        
        
        if(record.getGender().equals(Strings.PSW_GENDER_F)){
            female++;
        }
        else if(record.getGender().equals(Strings.PSW_GENDER_M)){
            male++;
        }
        if(record.getCountry().equals(Strings.PSW_TRAINED_UK)){
            uk++;
        }
        else if(!record.getCountry().equals(Strings.PSW_TRAINED_UK)){
            nonUk++;
        }
        if(record.getAge()>=23&&record.getAge()<=29){
            age2329++;
        }
        else if(record.getAge()>=30&&record.getAge()<=35){
            age3035++;
        }
        else if(record.getAge()>=36&&record.getAge()<=40){
            age3540++;
        }
        else if(record.getAge()>40){
            age40++;
        }
        if(record.getEthnicity().equals(Strings.PSW_ETHNICITY_WHITEB)){
            whiteb++;
        }
        else if(record.getEthnicity().equals(Strings.PSW_ETHNICITY_WHITEO)){
            whiteo++;
        }
        else if(record.getEthnicity().equals(Strings.PSW_ETHNICITY_ASIAN)){
            asian++;
        }
        else if(record.getEthnicity().equals(Strings.PSW_ETHNICITY_AFRICAN)){
            african++;
        }
        else if(!record.getEthnicity().equals("")){
            ethOther++;
        }
        if(record.getReligion().equals(Strings.PSW_RELIGION_CHRISTIAN)){
            christian++;
        }
        else if(record.getReligion().equals(Strings.PSW_RELIGION_ISLAM)){
            islam++;
        }
        else if(record.getReligion().equals(Strings.PSW_RELIGION_HINDU)){
            hindu++;
        }
        else if(record.getReligion().equals(Strings.PSW_RELIGION_ATHEIST)){
            atheist++;
        }
        else if(record.getReligion().equals(Strings.PSW_RELIGION_SIKH)){
            sikh++;
        }
        else if(record.getReligion().equals(Strings.PSW_RELIGION_JUDAISM)){
            judaism++;
        }
        else if(record.getReligion().equals(Strings.PSW_RELIGION_BUDDHIDM)){
            buddhism++;
        }
        else if(!record.getReligion().equals("")){
            relOther++;
        }
        else if(record.getReligion().equals("")||record.getReligion().equals(Strings.PSW_VAR_PNS)){
            relPNS++;
        }
        if(record.getDisability().equals(Strings.PSW_DISABILITY_YES)){
            yes++;
        }
        else if(record.getDisability().equals(Strings.PSW_DISABILITY_NO)){
            no++;
        }
        if(record.getSexOr().equals(Strings.PSW_SEXOR_HET)){
            het++;
        }
        else if(record.getSexOr().equals(Strings.PSW_SEXOR_HOMO)){
            homosexual++;
        }
        else if(record.getSexOr().equals(Strings.PSW_SEXOR_BISEX)){
            bisexual++;
        }
        else if(record.getSexOr().equals("")||record.getSexOr().equals(Strings.PSW_VAR_PNS)){
            sexOrPNS++;
        }
        
        
        line.countTotal();
        line.setMale(male);
        line.setFemale(female);
        line.setUk(uk);
        line.setNonUk(nonUk);
        line.setAge2329(age2329);
        line.setAge3035(age3035);
        line.setAge3540(age3540);
        line.setAge40(age40);
        line.setWhiteb(whiteb);
        line.setWhiteo(whiteo);
        line.setAsian(asian);
        line.setAfrican(african);
        line.setEthOther(ethOther);
        line.setChristian(christian);
        line.setIslam(islam);
        line.setHindu(hindu);
        line.setAtheist(atheist);
        line.setSikh(sikh);
        line.setJudaism(judaism);
        line.setBuddhism(buddhism);
        line.setRelOther(relOther);
        line.setRelPNS(relPNS);
        line.setYes(yes);
        line.setNo(no);
        line.setHet(het);
        line.setHomosexual(homosexual);
        line.setBisexual(bisexual);
        line.setSexOrPNS(sexOrPNS);
        
        return line;
    }
    
    private Table8Line getNonUkLineBySpc(Table8Line line, RefRecord record){
        
        int male = line.getMale();
        int female = line.getFemale();
        int uk= line.getUk();
        int nonUk= line.getNonUk();
        int age2329 = line.getAge2329();
        int age3035= line.getAge3035();
        int age3540 = line.getAge3540();
        int age40 = line.getAge40();
        int whiteb = line.getWhiteb();
        int whiteo = line.getWhiteo();
        int asian = line.getAsian();
        int african = line.getAfrican();
        int ethOther = line.getEthOther();
        int christian = line.getChristian();
        int islam = line.getIslam();
        int hindu = line.getHindu();
        int atheist = line.getAtheist();
        int sikh = line.getSikh();
        int judaism = line.getJudaism();
        int buddhism = line.getBuddhism();
        int relOther = line.getRelOther();
        int relPNS = line.getRelPNS();
        int yes = line.getYes();
        int no = line.getNo();
        int het = line.getHet();
        int homosexual = line.getHomosexual();
        int bisexual = line.getBisexual();
        int sexOrPNS = line.getSexOrPNS();
        
        if(!record.getCountry().equals(Strings.PSW_TRAINED_UK)){
            
            if(record.getGender().equals(Strings.PSW_GENDER_F)){
                female++;
            }
            else if(record.getGender().equals(Strings.PSW_GENDER_M)){
                male++;
            }
            if(record.getCountry().equals(Strings.PSW_TRAINED_UK)){
                uk++;
            }
            else if(!record.getCountry().equals(Strings.PSW_TRAINED_UK)){
                nonUk++;
            }
            if(record.getAge()>=23&&record.getAge()<=29){
                age2329++;
            }
            else if(record.getAge()>=30&&record.getAge()<=35){
                age3035++;
            }
            else if(record.getAge()>=36&&record.getAge()<=40){
                age3540++;
            }
            else if(record.getAge()>40){
                age40++;
            }
            if(record.getEthnicity().equals(Strings.PSW_ETHNICITY_WHITEB)){
                whiteb++;
            }
            else if(record.getEthnicity().equals(Strings.PSW_ETHNICITY_WHITEO)){
                whiteo++;
            }
            else if(record.getEthnicity().equals(Strings.PSW_ETHNICITY_ASIAN)){
                asian++;
            }
            else if(record.getEthnicity().equals(Strings.PSW_ETHNICITY_AFRICAN)){
                african++;
            }
            else if(!record.getEthnicity().equals("")){
                ethOther++;
            }
            if(record.getReligion().equals(Strings.PSW_RELIGION_CHRISTIAN)){
                christian++;
            }
            else if(record.getReligion().equals(Strings.PSW_RELIGION_ISLAM)){
                islam++;
            }
            else if(record.getReligion().equals(Strings.PSW_RELIGION_HINDU)){
                hindu++;
            }
            else if(record.getReligion().equals(Strings.PSW_RELIGION_ATHEIST)){
                atheist++;
            }
            else if(record.getReligion().equals(Strings.PSW_RELIGION_SIKH)){
                sikh++;
            }
            else if(record.getReligion().equals(Strings.PSW_RELIGION_JUDAISM)){
                judaism++;
            }
            else if(record.getReligion().equals(Strings.PSW_RELIGION_BUDDHIDM)){
                buddhism++;
            }
            else if(!record.getReligion().equals("")){
                relOther++;
            }
            else if(record.getReligion().equals("")||record.getReligion().equals(Strings.PSW_VAR_PNS)){
                relPNS++;
            }
            if(record.getDisability().equals(Strings.PSW_DISABILITY_YES)){
                yes++;
            }
            else if(record.getDisability().equals(Strings.PSW_DISABILITY_NO)){
                no++;
            }
            if(record.getSexOr().equals(Strings.PSW_SEXOR_HET)){
                het++;
            }
            else if(record.getSexOr().equals(Strings.PSW_SEXOR_HOMO)){
                homosexual++;
            }
            else if(record.getSexOr().equals(Strings.PSW_SEXOR_BISEX)){
                bisexual++;
            }
            else if(record.getSexOr().equals("")||record.getSexOr().equals(Strings.PSW_VAR_PNS)){
                sexOrPNS++;
            }
        }
        
        
        line.countTotal();
        line.setMale(male);
        line.setFemale(female);
        line.setUk(uk);
        line.setNonUk(nonUk);
        line.setAge2329(age2329);
        line.setAge3035(age3035);
        line.setAge3540(age3540);
        line.setAge40(age40);
        line.setWhiteb(whiteb);
        line.setWhiteo(whiteo);
        line.setAsian(asian);
        line.setAfrican(african);
        line.setEthOther(ethOther);
        line.setChristian(christian);
        line.setIslam(islam);
        line.setHindu(hindu);
        line.setAtheist(atheist);
        line.setSikh(sikh);
        line.setJudaism(judaism);
        line.setBuddhism(buddhism);
        line.setRelOther(relOther);
        line.setRelPNS(relPNS);
        line.setYes(yes);
        line.setNo(no);
        line.setHet(het);
        line.setHomosexual(homosexual);
        line.setBisexual(bisexual);
        line.setSexOrPNS(sexOrPNS);
        
        return line;
    }
    
    private List<Integer> countT1Integers() {
        
        List <Integer> list = new ArrayList<>();
        
        int closedWithinPeriod = 0;
        int openedAndClosed =  0;
        
        XSSFSheet ccSheet = pswWorkbook.getSheet(Strings.PSW_SHEET_CLOSED_CASES);
        
        int dateOpenedColNo = getCellColumnByString(Strings.PSW_COLUMN_DATE_OPENED, ccSheet);
        int dateClosedColNo = getCellColumnByString(Strings.PSW_COLUMN_DATE_CLOSED, ccSheet);
        
        Iterator<Row> rowIterator = ccSheet.iterator();
        System.out.println(DocHelper.getStartingYear());
        while(rowIterator.hasNext()){
            Row r = rowIterator.next();
            if(r.getRowNum()>getCellRowByString(Strings.PSW_COLUMN_DATE_OPENED, ccSheet)){
                Cell dateClosedCell = CellUtil.getCell(r, dateClosedColNo);
                Cell dateOpenedCell = CellUtil.getCell(r, dateOpenedColNo);
                DataFormatter df = new DataFormatter(Locale.UK);
                df.formatCellValue(dateClosedCell);
                df.formatCellValue(dateOpenedCell);
                Date dateClosed = dateClosedCell.getDateCellValue();
                Date dateOpened = dateOpenedCell.getDateCellValue();
                Calendar calClosed = Calendar.getInstance();
                Calendar calOpened = Calendar.getInstance();
                calClosed.setTime(dateClosed);
                calOpened.setTime(dateOpened);
                if((calClosed.get(Calendar.MONTH)>2&&calClosed.get(Calendar.YEAR)==DocHelper.getStartingYear())||(calClosed.get(Calendar.MONTH)<=2&&calClosed.get(Calendar.YEAR)==DocHelper.getStartingYear()+1)){
                    closedWithinPeriod++;
                }
                if(((calClosed.get(Calendar.MONTH)>2&&calClosed.get(Calendar.YEAR)==DocHelper.getStartingYear())
                        ||(calClosed.get(Calendar.MONTH)<=2&&calClosed.get(Calendar.YEAR)==DocHelper.getStartingYear()+1))
                        &&((calOpened.get(Calendar.MONTH)>2&&calOpened.get(Calendar.YEAR)==DocHelper.getStartingYear())
                        ||(calOpened.get(Calendar.MONTH)<2&&calOpened.get(Calendar.YEAR)==DocHelper.getStartingYear()+1))){
                    openedAndClosed++;
                }
                
            }
        }
        
        list.add(closedWithinPeriod);
        list.add(openedAndClosed);
        
        return list;
    }
    
    public Integer countTotalReferrals() {
        int totalRef = 0;
        XSSFSheet refSheet = pswWorkbook.getSheet(Strings.PSW_SHEET_REFERRALS);
        Iterator<Row> rowIterator = refSheet.rowIterator();
        
        int firstRow = getCellRowByString(Strings.PSW_COLUMN_ADHD, refSheet);
        
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            if (!isRowEmpty(row) && row.getRowNum() > firstRow) {
                totalRef++;
            }
        }
        
        return totalRef;
    }
    
    private Integer countTotalWessex() {
        int totalRef = 0;
        XSSFSheet wessexSheet = pswWorkbook.getSheet(Strings.PSW_SHEET_WESSEX);
        int valueColNo = getCellColumnByString(Strings.PSW_COLUMN_VALUE, wessexSheet);
        int countOfTraineesRowNo = getCellRowByString(Strings.PSW_COLUMN_TRAINEE_COUNT, wessexSheet);
        totalRef = (int) wessexSheet.getRow(countOfTraineesRowNo).getCell(valueColNo).getNumericCellValue();
        
        return totalRef;
    }
    
    
    public static boolean isCellEmpty(final Cell cell) {
        if (cell == null || cell.getCellType() == CellType.BLANK) {
            return true;
        }
        if (cell.getCellType() == CellType.STRING && cell.getStringCellValue().isEmpty()) {
            return true;
        }
        return false;
    }
    
    private static boolean isRowEmpty(Row row) {
        for (int c = row.getFirstCellNum(); c < row.getLastCellNum(); c++) {
            Cell cell = row.getCell(c);
            if (cell != null && cell.getCellType() != CellType.BLANK) {
                return false;
            }
        }
        return true;
    }
    
    private static int getCellRowByString(String str, XSSFSheet sheet) {
        int rowNumber = 0;
        
        for (Row r : sheet) {
            for (Cell c : r) {
                String cellValueStr = "";
                try {
                    cellValueStr = c.getStringCellValue();
                } catch (IllegalStateException e) {
                }
                if (cellValueStr.equals(str)) {
                    
                    rowNumber = c.getRowIndex();
                    
                }
            }
        }
        return rowNumber;
    }
    
    public void genGraphFile() {
        try {
            graphs.createNewFile();
        } catch (IOException ex) {
            Logger.getLogger(PoiHelper.class.getName()).log(Level.SEVERE, null, ex);
        }
        
    }
    
    /**
     * Checks if a given String is numeric
     *
     * @param strNum
     * @return
     */
    private static boolean isNumeric(String strNum) {
        if (strNum == null) {
            return false;
        }
        try {
            double d = Double.parseDouble(strNum);
        } catch (NumberFormatException nfe) {
            return false;
        }
        return true;
    }
    
    public void deleteLastSheetGraphs(){
        graphsWorkbook.removeSheetAt(10);
        try{
            FileOutputStream fileOut = new FileOutputStream(graphs);
            graphsWorkbook.write(fileOut);
            fileOut.close();
        } catch (IOException ex) {
            Logger.getLogger(PoiHelper.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
    
    public void getRefRecords(){
        
        ArrayList<RefRecord> records = new ArrayList<>();
        
        XSSFSheet refSheet = pswWorkbook.getSheetAt(pswWorkbook.getSheetIndex(Strings.PSW_SHEET_REFERRALS));
        int spcColumn = getCellColumnByString(Strings.PSW_COLUMN_SPECIALTY, refSheet);
        int genderColumn = getCellColumnByString(Strings.PSW_COLUMN_GENDER, refSheet);
        int trainedColumn = getCellColumnByString(Strings.PSW_COLUMN_COUNTRY, refSheet);
        int ageColumn = getCellColumnByString(Strings.PSW_COLUMN_AGE, refSheet);
        int ethColumn = getCellColumnByString(Strings.PSW_COLUMN_ETHNICITY, refSheet);
        int relColumn = getCellColumnByString(Strings.PSW_COLUMN_RELIGION, refSheet);
        int disabilityColumn = getCellColumnByString(Strings.PSW_COLUMN_DISABILITY, refSheet);
        int sexOrColumn = getCellColumnByString(Strings.PSW_COLUMN_SEXUAL_OR, refSheet);
        int exSupportColumn = getCellColumnByString(Strings.PSW_COLUMN_EXAM, refSheet);
        int anxietyColumn = getCellColumnByString(Strings.PSW_COLUMN_ANXIETY, refSheet);
        int capabilityColumn = getCellColumnByString(Strings.PSW_COLUMN_CAPABILITY, refSheet);
        int carreerColumn = getCellColumnByString(Strings.PSW_COLUMN_CARREER, refSheet);
        int clinicalSkillsColumn = getCellColumnByString(Strings.PSW_COLUMN_CLINICAL_SKILLS, refSheet);
        int communicationColumn = getCellColumnByString(Strings.PSW_COLUMN_COMMUNICATION, refSheet);
        int conductColumn = getCellColumnByString(Strings.PSW_COLUMN_CONDUCT, refSheet);
        int culturalColumn = getCellColumnByString(Strings.PSW_COLUMN_CULTURAL, refSheet);
        int mentalHealthColumn = getCellColumnByString(Strings.PSW_COLUMN_MENTAL, refSheet);
        int physicalHealthColumn = getCellColumnByString(Strings.PSW_COLUMN_PHYSICAL, refSheet);
        int languageColumn = getCellColumnByString(Strings.PSW_COLUMN_LANGUAGE, refSheet);
        int profColumn = getCellColumnByString(Strings.PSW_COLUMN_PROFFESSIONALISM, refSheet);
        int adhdColumn = getCellColumnByString(Strings.PSW_COLUMN_ADHD, refSheet);
        int asdColumn = getCellColumnByString(Strings.PSW_COLUMN_ASD, refSheet);
        int dyslexiaColumn = getCellColumnByString(Strings.PSW_COLUMN_DYSLEXIA, refSheet);
        int dyspraxiaColumn = getCellColumnByString(Strings.PSW_COLUMN_DYSPRAXIA, refSheet);
        int srttColumn = getCellColumnByString(Strings.PSW_COLUMN_SRTT, refSheet);
        int teamColumn = getCellColumnByString(Strings.PSW_COLUMN_TEAM, refSheet);
        int timeColumn = getCellColumnByString(Strings.PSW_COLUMN_TIME, refSheet);
        int otherRefReasonColumn = getCellColumnByString(Strings.GRAPH_COLUMN_OTHER, refSheet);
        int trustColumn = getCellColumnByString(Strings.PSW_COLUMN_TRUST, refSheet);
        int gradeColumn = getCellColumnByString(Strings.PSW_COLUMN_GRADE, refSheet);
        int schoolColumn = getCellColumnByString(Strings.PSW_COLUMN_SPECIALTY, refSheet);
        int dateColumn = getCellColumnByString(Strings.PSW_COLUMN_DATE, refSheet);
        int closedColumn = getCellColumnByString(Strings.PSW_COLUMN_CASE_OPEN, refSheet);
        
        Iterator<Row> rowIterator = refSheet.iterator();
        while(rowIterator.hasNext()){
            
            Row row = rowIterator.next();
            int currentRow = row.getRowNum();
            
            
            Cell spCell = CellUtil.getCell(row, spcColumn);
            Cell genderCell = CellUtil.getCell(row, genderColumn);
            Cell trainedCell = CellUtil.getCell(row, trainedColumn);
            Cell ageCell = CellUtil.getCell(row, ageColumn);
            Cell ethCell = CellUtil.getCell(row, ethColumn);
            Cell relCell = CellUtil.getCell(row, relColumn);
            Cell disabilityCell = CellUtil.getCell(row, disabilityColumn);
            Cell sexOrCell = CellUtil.getCell(row, sexOrColumn);
            Cell anxietyCell = CellUtil.getCell(row, anxietyColumn);
            Cell capabilityCell = CellUtil.getCell(row, capabilityColumn);
            Cell carreerCell = CellUtil.getCell(row, carreerColumn);
            Cell clinSkillsCell = CellUtil.getCell(row, clinicalSkillsColumn);
            Cell communicationCell = CellUtil.getCell(row, communicationColumn);
            Cell conductCell = CellUtil.getCell(row, conductColumn);
            Cell culturalCell = CellUtil.getCell(row, culturalColumn);
            Cell exSupportCell = CellUtil.getCell(row, exSupportColumn);
            Cell mentalHealthCell = CellUtil.getCell(row, mentalHealthColumn);
            Cell physicalHealthCell = CellUtil.getCell(row, physicalHealthColumn);
            Cell languageCell = CellUtil.getCell(row, languageColumn);
            Cell profCell = CellUtil.getCell(row, profColumn);
            Cell adhdCell = CellUtil.getCell(row, adhdColumn);
            Cell asdCell = CellUtil.getCell(row, asdColumn);
            Cell dyslexiaCell = CellUtil.getCell(row, dyslexiaColumn);
            Cell dyspraxiaCell = CellUtil.getCell(row, dyspraxiaColumn);
            Cell srttCell = CellUtil.getCell(row, srttColumn);
            Cell teamCell = CellUtil.getCell(row, teamColumn);
            Cell timeCell = CellUtil.getCell(row, timeColumn);
            Cell otherCell = CellUtil.getCell(row, otherRefReasonColumn);
            Cell trustCell = CellUtil.getCell(row, trustColumn);
            Cell gradeCell = CellUtil.getCell(row, gradeColumn);
            Cell schoolCell = CellUtil.getCell(row, schoolColumn);
            Cell dateCell = CellUtil.getCell(row, dateColumn);
            Cell closedCell = CellUtil.getCell(row, closedColumn);
            
            if(currentRow>3&&!isRowEmpty(row)){
                RefRecord record = new RefRecord();
                record.setSpecialty(spCell.getStringCellValue());
                record.setGender(genderCell.getStringCellValue());
                record.setCountry(trainedCell.getStringCellValue());
                if(ageCell.getCellType()==(CellType.NUMERIC)){
                    record.setAge((int) ageCell.getNumericCellValue());
                }
                
                DataFormatter formatter = new DataFormatter(Locale.UK);
                formatter.formatCellValue(dateCell);
                record.setRefDate(dateCell.getDateCellValue());
                record.setEthnicity(ethCell.getStringCellValue());
                record.setReligion(relCell.getStringCellValue());
                record.setDisability(disabilityCell.getStringCellValue());
                record.setSexOr(sexOrCell.getStringCellValue());
                record.setTrust(trustCell.getStringCellValue());
                record.setGrade(gradeCell.getStringCellValue());
                record.setSchool(schoolCell.getStringCellValue());
                
                if(!isCellEmpty(anxietyCell)){
                    record.setAnxiety(true);
                }
                if(!isCellEmpty(capabilityCell)){
                    record.setCapability(true);
                }
                if(!isCellEmpty(carreerCell)){
                    record.setCarreer(true);
                }
                if(!isCellEmpty(clinSkillsCell)){
                    record.setClinSkills(true);
                }
                if(!isCellEmpty(communicationCell)){
                    record.setCommunication(true);
                }
                if(!isCellEmpty(conductCell)){
                    record.setConduct(true);
                }
                if(!isCellEmpty(culturalCell)){
                    record.setCultural(true);
                }
                if(!isCellEmpty(exSupportCell)){
                    record.setExam(true);
                }
                if(!isCellEmpty(mentalHealthCell)){
                    record.setHealthMental(true);
                }
                if(!isCellEmpty(physicalHealthCell)){
                    record.setHealthPhysical(true);
                }
                if(!isCellEmpty(languageCell)){
                    record.setLanguage(true);
                }
                if(!isCellEmpty(profCell)){
                    record.setProfessionalism(true);
                }
                if(!isCellEmpty(adhdCell)){
                    record.setAdhd(true);
                }
                if(!isCellEmpty(asdCell)){
                    record.setAsd(true);
                }
                if(!isCellEmpty(dyslexiaCell)){
                    record.setDyslexia(true);
                }
                if(!isCellEmpty(dyspraxiaCell)){
                    record.setDyspraxia(true);
                }
                if(!isCellEmpty(srttCell)){
                    record.setSrtt(true);
                }
                if(!isCellEmpty(teamCell)){
                    record.setTeam(true);
                }
                if(!isCellEmpty(timeCell)){
                    record.setTime(true);
                }
                if(!isCellEmpty(otherCell)){
                    record.setOtherRefReason(true);
                }
                if(isCellEmpty(closedCell)){
                    record.setCaseOpen(false);
                }
                else record.setCaseOpen(true);
                records.add(record);
            }
            this.recordList=records;
        }
    }
    
    public ArrayList<File> getFiles(){
        ArrayList<File> list = new ArrayList<>();
        list.add(psw);
        list.add(graphs);
        
        return list;
    }
    
    public void getGraphs(){
        getGraph1();
        getGraph2();
        getGraph3();
        getGraph4();
        getGraph5();
        getGraph6();
        getGraph7();
        getGraph8();
        getGraph9();
        getGraph10();
        getGraph11();
    }
}

//    public ArrayList<String> getTopReasonsReferral(){
//
//    }

