/*
* To change this license header, choose License Headers in Project Properties.
* To change this template file, choose Tools | Templates
* and open the template in the editor.
*/
package facade;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
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
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import res.Strings;

/**
 *
 * @author diego
 */
public class PoiHelper {
    
    private File psw;
    private File graphs;
    
    public PoiHelper(){
        
    }
    
    public PoiHelper(File psw, File graphs){
        this.psw=psw;
        this.graphs=graphs;
    }
    
    
    /**
     * Retrieves the Column no. of a cell by the value of it's content
     * @param str
     * @param sheet
     * @return
     */
    private static int getCellColumnByString(String str, XSSFSheet sheet){
        int columnNumber = 0;
        
        
        for (Row r : sheet) {
            for (Cell c : r) {
                String cellValueStr = "";
                try {
                    cellValueStr=c.getStringCellValue();
                } catch (IllegalStateException e) {
                }
                if(cellValueStr.equals(str)){
                    
                    columnNumber = c.getColumnIndex();
                    
                }
            }
        }
        return columnNumber;
    }
    
    /**
     * Generates the table for the first graph in Graphs.xlsx
     * @param psw
     * @param graphs
     */
    public void getGraph1(){
        
        int fCount = 0;
        int mCount = 0;
        int uCount = 0;
        double wssxFemaleCount;
        double wssxMaleCount;
        double wssxUnknownCount;
        
        FileOutputStream fileOut;
        
        try (FileInputStream fileIn = new FileInputStream(psw)) {
            XSSFWorkbook workbook = new XSSFWorkbook(fileIn);
            XSSFSheet refSheet = workbook.getSheetAt(workbook.getSheetIndex(Strings.PSW_SHEET_REFERRALS));
            XSSFSheet genderSheet = workbook.getSheetAt(workbook.getSheetIndex(Strings.PSW_COLUMN_GENDER));
            int columnNumber = getCellColumnByString(Strings.PSW_COLUMN_GENDER, refSheet);
            
            Iterator<Row> rowIterator = refSheet.iterator();
            
            while (rowIterator.hasNext())
            {
                Row row = rowIterator.next();
                Cell cell = CellUtil.getCell(row, columnNumber);
                
                if(cell.getStringCellValue().equals(Strings.PSW_GENDER_F)){
                    fCount++;
                }
                else if(cell.getStringCellValue().equals(Strings.PSW_GENDER_M)){
                    mCount++;
                }
            }
            
            Row femaleRow = genderSheet.getRow(2);
            Row unknownRow = genderSheet.getRow(3);
            Row maleRow = genderSheet.getRow(4);
            wssxFemaleCount = femaleRow.getCell(1).getNumericCellValue();
            wssxMaleCount = maleRow.getCell(1).getNumericCellValue();
            wssxUnknownCount = unknownRow.getCell(1).getNumericCellValue();
            XSSFWorkbook graphWB = new XSSFWorkbook();
            String sheetName = Strings.GRAPH_SHEET_1;
            XSSFSheet sheet = graphWB.createSheet(sheetName);
            Row titlesRow = sheet.createRow( 0);
            Row fRow = sheet.createRow( 1);
            Row mRow = sheet.createRow( 2);
            Row uRow = sheet.createRow( 3);
            Cell cell = titlesRow.createCell( 1);
            cell.setCellValue(Strings.GRAPH_COLUMN_REF_COUNT);
            cell = titlesRow.createCell( 2);
            cell.setCellValue(Strings.GRAPH_COLUMN_WSSX_COUNT);
            Cell fCell = fRow.createCell( 0);
            fCell.setCellValue(Strings.PSW_COLUMN_FEMALE);
            fCell = fRow.createCell( 1);
            fCell.setCellValue(fCount);
            fCell = fRow.createCell( 2);
            fCell.setCellValue(wssxFemaleCount);
            Cell mCell = mRow.createCell( 0);
            mCell.setCellValue(Strings.PSW_COLUMN_MALE);
            mCell = mRow.createCell( 1);
            mCell.setCellValue(mCount);
            mCell = mRow.createCell( 2);
            mCell.setCellValue(wssxMaleCount);
            Cell uCell = uRow.createCell( 0);
            uCell.setCellValue(Strings.PSW_COLUMN_UNKNOWN);
            uCell = uRow.createCell( 1);
            uCell.setCellValue(uCount);
            uCell = uRow.createCell( 2);
            uCell.setCellValue(wssxUnknownCount);
            fileOut = new FileOutputStream(graphs);
            graphWB.write(fileOut);
            
            fileIn.close();
            fileOut.close();
            
        } catch (IOException ex) {
            Logger.getLogger(PoiHelper.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
    
    /**
     * Generates the table for the second graph in Graphs.xlsx
     * @param file
     * @param graphs
     */
    public  void getGraph2(){
        
        int capabilityMCount = 0;
        int capabilityFCount = 0;
        int examMCount = 0;
        int examFCount = 0;
        int healthMCount = 0;
        int healthFCount = 0;
        int carreerMCount = 0;
        int carreerFCount = 0;
        int conductMCount = 0;
        int conductFCount = 0;
        int otherMCount = 0;
        int otherFCount = 0;
        
        FileInputStream fileOutIn;
        FileOutputStream fileOut;
        
        try (FileInputStream fileIn = new FileInputStream(psw)) {
            XSSFWorkbook workbook = new XSSFWorkbook(fileIn);
            XSSFSheet refSheet = workbook.getSheetAt(workbook.getSheetIndex(Strings.PSW_SHEET_REFERRALS));
            int anxietyColNo = getCellColumnByString(Strings.PSW_COLUMN_ANXIETY, refSheet);
            int carreerColNo = getCellColumnByString(Strings.PSW_COLUMN_CARREER, refSheet);
            int clinicalSkillsColNo = getCellColumnByString(Strings.PSW_COLUMN_CLINICAL_SKILLS, refSheet);
            int communicationColNo = getCellColumnByString(Strings.PSW_COLUMN_COMMUNICATION, refSheet);
            int conductColNo = getCellColumnByString(Strings.PSW_COLUMN_CONDUCT, refSheet);
            int culturalColNo = getCellColumnByString(Strings.PSW_COLUMN_CULTURAL, refSheet);
            int examColNo = getCellColumnByString(Strings.PSW_COLUMN_EXAM, refSheet);
            int mentalHealthColNo = getCellColumnByString(Strings.PSW_COLUMN_MENTAL, refSheet);
            int physicalHealthColNo = getCellColumnByString(Strings.PSW_COLUMN_PHYSICAL, refSheet);
            int languageColNo = getCellColumnByString(Strings.PSW_COLUMN_LANGUAGE, refSheet);
            int profColNo = getCellColumnByString(Strings.PSW_COLUMN_PROFFESSIONALISM, refSheet);
            int adhdColNo = getCellColumnByString(Strings.PSW_COLUMN_ADHD, refSheet);
            int asdColNo = getCellColumnByString(Strings.PSW_COLUMN_ASD, refSheet);
            int dyslexiaColNo = getCellColumnByString(Strings.PSW_COLUMN_DYSLEXIA, refSheet);
            int dyspraxiaColNo = getCellColumnByString(Strings.PSW_COLUMN_DYSPRAXIA, refSheet);
            int srttColNo = getCellColumnByString(Strings.PSW_COLUMN_SRTT, refSheet);
            int teamColNo = getCellColumnByString(Strings.PSW_COLUMN_TEAM, refSheet);
            int timeColNo = getCellColumnByString(Strings.PSW_COLUMN_TIME, refSheet);
            int genderColNo = getCellColumnByString(Strings.PSW_COLUMN_GENDER, refSheet);
            
            Iterator<Row> rowIterator = refSheet.iterator();
            
            while (rowIterator.hasNext())
            {
                Row row = rowIterator.next();
                Cell anxietyCell = CellUtil.getCell(row, anxietyColNo);
                Cell carreerCell = CellUtil.getCell(row, carreerColNo);
                Cell clinicalSkillsCell = CellUtil.getCell(row, clinicalSkillsColNo);
                Cell communicationCell = CellUtil.getCell(row, communicationColNo);
                Cell conductCell = CellUtil.getCell(row, conductColNo);
                Cell culturalCell = CellUtil.getCell(row, culturalColNo);
                Cell examCell = CellUtil.getCell(row, examColNo);
                Cell mentalHealthCell = CellUtil.getCell(row, mentalHealthColNo);
                Cell physicalHealthCell = CellUtil.getCell(row, physicalHealthColNo);
                Cell languageCell = CellUtil.getCell(row, languageColNo);
                Cell profCell = CellUtil.getCell(row, profColNo);
                Cell adhdCell = CellUtil.getCell(row, adhdColNo);
                Cell asdCell = CellUtil.getCell(row, asdColNo);
                Cell dyslexiaCell = CellUtil.getCell(row, dyslexiaColNo);
                Cell dyspraxiaCell = CellUtil.getCell(row, dyspraxiaColNo);
                Cell srttCell = CellUtil.getCell(row, srttColNo);
                Cell teamCell = CellUtil.getCell(row, teamColNo);
                Cell timeCell = CellUtil.getCell(row, timeColNo);
                
                if(!anxietyCell.getStringCellValue().equals("")){
                    
                    if(row.getCell(genderColNo).getStringCellValue().equals(Strings.PSW_GENDER_F)){
                        otherFCount++;
                    }
                    
                    else if(row.getCell(genderColNo).getStringCellValue().equals(Strings.PSW_GENDER_M)){
                        otherMCount++;
                    }
                }
                
                if(!carreerCell.getStringCellValue().equals("")){
                    
                    if(row.getCell(genderColNo).getStringCellValue().equals(Strings.PSW_GENDER_F)){
                        carreerFCount++;
                    }
                    
                    else if(row.getCell(genderColNo).getStringCellValue().equals(Strings.PSW_GENDER_M)){
                        carreerMCount++;
                    }
                }
                
                if(!clinicalSkillsCell.getStringCellValue().equals("")){
                    
                    if(row.getCell(genderColNo).getStringCellValue().equals(Strings.PSW_GENDER_F)){
                        capabilityFCount++;
                    }
                    
                    else if(row.getCell(genderColNo).getStringCellValue().equals(Strings.PSW_GENDER_M)){
                        capabilityMCount++;
                    }
                }
                
                if(!communicationCell.getStringCellValue().equals("")){
                    
                    if(row.getCell(genderColNo).getStringCellValue().equals(Strings.PSW_GENDER_F)){
                        capabilityFCount++;
                    }
                    
                    else if(row.getCell(genderColNo).getStringCellValue().equals(Strings.PSW_GENDER_M)){
                        capabilityMCount++;
                    }
                }
                
                if(!conductCell.getStringCellValue().equals("")){
                    
                    if(row.getCell(genderColNo).getStringCellValue().equals(Strings.PSW_GENDER_F)){
                        conductFCount++;
                    }
                    
                    else if(row.getCell(genderColNo).getStringCellValue().equals(Strings.PSW_GENDER_M)){
                        conductMCount++;
                    }
                }
                
                if(!culturalCell.getStringCellValue().equals("")){
                    
                    if(row.getCell(genderColNo).getStringCellValue().equals(Strings.PSW_GENDER_F)){
                        capabilityFCount++;
                    }
                    
                    else if(row.getCell(genderColNo).getStringCellValue().equals(Strings.PSW_GENDER_M)){
                        capabilityMCount++;
                    }
                }
                
                if(!examCell.getStringCellValue().equals("")){
                    
                    if(row.getCell(genderColNo).getStringCellValue().equals(Strings.PSW_GENDER_F)){
                        examFCount++;
                    }
                    
                    else if(row.getCell(genderColNo).getStringCellValue().equals(Strings.PSW_GENDER_M)){
                        examMCount++;
                    }
                }
                
                if(!mentalHealthCell.getStringCellValue().equals("")){
                    
                    if(row.getCell(genderColNo).getStringCellValue().equals(Strings.PSW_GENDER_F)){
                        healthFCount++;
                    }
                    
                    else if(row.getCell(genderColNo).getStringCellValue().equals(Strings.PSW_GENDER_M)){
                        healthMCount++;
                    }
                }
                
                if(!physicalHealthCell.getStringCellValue().equals("")){
                    
                    if(row.getCell(genderColNo).getStringCellValue().equals(Strings.PSW_GENDER_F)){
                        healthFCount++;
                    }
                    
                    else if(row.getCell(genderColNo).getStringCellValue().equals(Strings.PSW_GENDER_M)){
                        healthMCount++;
                    }
                }
                
                if(!languageCell.getStringCellValue().equals("")){
                    
                    if(row.getCell(genderColNo).getStringCellValue().equals(Strings.PSW_GENDER_F)){
                        capabilityFCount++;
                    }
                    
                    else if(row.getCell(genderColNo).getStringCellValue().equals(Strings.PSW_GENDER_M)){
                        capabilityMCount++;
                    }
                }
                
                if(!profCell.getStringCellValue().equals("")){
                    
                    if(row.getCell(genderColNo).getStringCellValue().equals(Strings.PSW_GENDER_F)){
                        capabilityFCount++;
                    }
                    
                    else if(row.getCell(genderColNo).getStringCellValue().equals(Strings.PSW_GENDER_M)){
                        capabilityMCount++;
                    }
                }
                
                if(!adhdCell.getStringCellValue().equals("")){
                    
                    if(row.getCell(genderColNo).getStringCellValue().equals(Strings.PSW_GENDER_F)){
                        otherFCount++;
                    }
                    
                    else if(row.getCell(genderColNo).getStringCellValue().equals(Strings.PSW_GENDER_M)){
                        otherMCount++;
                    }
                }
                
                if(!asdCell.getStringCellValue().equals("")){
                    
                    if(row.getCell(genderColNo).getStringCellValue().equals(Strings.PSW_GENDER_F)){
                        otherFCount++;
                    }
                    
                    else if(row.getCell(genderColNo).getStringCellValue().equals(Strings.PSW_GENDER_M)){
                        otherMCount++;
                    }
                }
                
                if(!dyslexiaCell.getStringCellValue().equals("")){
                    
                    if(row.getCell(genderColNo).getStringCellValue().equals(Strings.PSW_GENDER_F)){
                        otherFCount++;
                    }
                    
                    else if(row.getCell(genderColNo).getStringCellValue().equals(Strings.PSW_GENDER_M)){
                        otherMCount++;
                    }
                }
                
                if(!dyspraxiaCell.getStringCellValue().equals("")){
                    
                    if(row.getCell(genderColNo).getStringCellValue().equals(Strings.PSW_GENDER_F)){
                        otherFCount++;
                    }
                    
                    else if(row.getCell(genderColNo).getStringCellValue().equals(Strings.PSW_GENDER_M)){
                        otherMCount++;
                    }
                }
                
                if(!srttCell.getStringCellValue().equals("")){
                    
                    if(row.getCell(genderColNo).getStringCellValue().equals(Strings.PSW_GENDER_F)){
                        otherFCount++;
                    }
                    
                    else if(row.getCell(genderColNo).getStringCellValue().equals(Strings.PSW_GENDER_M)){
                        otherMCount++;
                    }
                }
                
                if(!teamCell.getStringCellValue().equals("")){
                    
                    if(row.getCell(genderColNo).getStringCellValue().equals(Strings.PSW_GENDER_F)){
                        capabilityFCount++;
                    }
                    
                    else if(row.getCell(genderColNo).getStringCellValue().equals(Strings.PSW_GENDER_M)){
                        capabilityMCount++;
                    }
                }
                
                if(!timeCell.getStringCellValue().equals("")){
                    
                    if(row.getCell(genderColNo).getStringCellValue().equals(Strings.PSW_GENDER_F)){
                        capabilityFCount++;
                    }
                    
                    else if(row.getCell(genderColNo).getStringCellValue().equals(Strings.PSW_GENDER_M)){
                        capabilityMCount++;
                    }
                }
            }
            
            fileOutIn = new FileInputStream(graphs);
            XSSFSheet sheet;
            XSSFWorkbook graphWB = new XSSFWorkbook(fileOutIn);
            
            if(graphWB.getNumberOfSheets()==1){
                sheet=graphWB.createSheet(Strings.GRAPH_SHEET_2);
            }
            else{
                sheet=graphWB.getSheetAt(1);
            }
            
            Row titlesRow = sheet.createRow( 0);
            Row fRow = sheet.createRow( 1);
            Row mRow = sheet.createRow( 2);
            Cell cell = titlesRow.createCell( 1);
            cell.setCellValue(Strings.PSW_COLUMN_CAPABILITY);
            cell = titlesRow.createCell( 2);
            cell.setCellValue(Strings.GRAPH_COLUMN_EXAM_FAILURE_DIRECT);
            cell = titlesRow.createCell( 3);
            cell.setCellValue(Strings.GRAPH_COLUMN_HEALTH);
            cell = titlesRow.createCell( 4);
            cell.setCellValue(Strings.GRAPH_COLUMN_CARREER_DIRECT);
            cell = titlesRow.createCell( 5);
            cell.setCellValue(Strings.PSW_COLUMN_CONDUCT);
            cell = titlesRow.createCell( 6);
            cell.setCellValue(Strings.GRAPH_COLUMN_OTHER);
            Cell fCell = fRow.createCell( 0);
            fCell.setCellValue(Strings.PSW_COLUMN_FEMALE);
            Cell mCell = mRow.createCell( 0);
            mCell.setCellValue(Strings.PSW_COLUMN_MALE);
            fCell = fRow.createCell( 1);
            fCell.setCellValue(capabilityFCount);
            fCell = fRow.createCell( 2);
            fCell.setCellValue(examFCount);
            fCell = fRow.createCell( 3);
            fCell.setCellValue(healthFCount);
            fCell = fRow.createCell( 4);
            fCell.setCellValue(carreerFCount);
            fCell = fRow.createCell( 5);
            fCell.setCellValue(conductFCount);
            fCell = fRow.createCell( 6);
            fCell.setCellValue(otherFCount);
            mCell = mRow.createCell( 1);
            mCell.setCellValue(capabilityMCount);
            mCell = mRow.createCell( 2);
            mCell.setCellValue(examMCount);
            mCell = mRow.createCell( 3);
            mCell.setCellValue(healthMCount);
            mCell = mRow.createCell( 4);
            mCell.setCellValue(carreerMCount);
            mCell = mRow.createCell( 5);
            mCell.setCellValue(conductMCount);
            mCell = mRow.createCell( 6);
            mCell.setCellValue(otherMCount);
            fileOut = new FileOutputStream(graphs);
            graphWB.write(fileOut);
            
            fileIn.close();
            fileOut.close();
            fileOutIn.close();
            
        }catch (IOException ex) {
            Logger.getLogger(PoiHelper.class.getName()).log(Level.SEVERE, null, ex);
        }
        
    }
    
    /**
     * Generates the table for the third graph in Graphs.xlsx
     * @param psw
     * @param graphs
     */
    public  void getGraph3(){
        
        int mentalFcount = 0;
        int physicalFcount = 0;
        int bothFcount = 0;
        int mentalMcount = 0;
        int physicalMcount = 0;
        int bothMcount = 0;
        
        FileInputStream fileOutIn;
        FileOutputStream fileOut;
        
        try (FileInputStream fileIn = new FileInputStream(psw)) {
            XSSFWorkbook workbook = new XSSFWorkbook(fileIn);
            XSSFSheet refSheet = workbook.getSheetAt(workbook.getSheetIndex(Strings.PSW_SHEET_REFERRALS));
            int mentalHealthColNo = getCellColumnByString(Strings.PSW_COLUMN_MENTAL, refSheet);
            int physicalHealthColNo = getCellColumnByString(Strings.PSW_COLUMN_PHYSICAL, refSheet);
            int genderColNo = getCellColumnByString(Strings.PSW_COLUMN_GENDER, refSheet);
            
            Iterator<Row> rowIterator = refSheet.iterator();
            
            while (rowIterator.hasNext())
            {
                Row row = rowIterator.next();
                
                Cell mentalHealthCell = CellUtil.getCell(row, mentalHealthColNo);
                Cell physicalHealthCell = CellUtil.getCell(row, physicalHealthColNo);
                
                if(!row.getCell(mentalHealthColNo).getStringCellValue().equals("")&&row.getCell(genderColNo).getStringCellValue().equals("F")&&row.getCell(physicalHealthColNo).getStringCellValue().equals("")){
                    mentalFcount++;
                }
                if(!row.getCell(physicalHealthColNo).getStringCellValue().equals("")&&row.getCell(genderColNo).getStringCellValue().equals("F")&&row.getCell(mentalHealthColNo).getStringCellValue().equals("")){
                    physicalFcount++;
                }
                if(!row.getCell(mentalHealthColNo).getStringCellValue().equals("")&&row.getCell(genderColNo).getStringCellValue().equals("F")&&!row.getCell(physicalHealthColNo).getStringCellValue().equals("")){
                    bothFcount++;
                }
                if(!row.getCell(mentalHealthColNo).getStringCellValue().equals("")&&row.getCell(genderColNo).getStringCellValue().equals("M")&&!row.getCell(physicalHealthColNo).getStringCellValue().equals("")){
                    bothMcount++;
                }
                if(!row.getCell(physicalHealthColNo).getStringCellValue().equals("")&&row.getCell(genderColNo).getStringCellValue().equals("M")&&row.getCell(mentalHealthColNo).getStringCellValue().equals("")){
                    physicalFcount++;
                }
                if(!row.getCell(mentalHealthColNo).getStringCellValue().equals("")&&row.getCell(genderColNo).getStringCellValue().equals("M")&&row.getCell(physicalHealthColNo).getStringCellValue().equals("")){
                    mentalMcount++;
                }
            }
            
            fileOutIn = new FileInputStream(graphs);
            XSSFWorkbook graphWB = new XSSFWorkbook(fileOutIn);
            XSSFSheet sheet;
            
            if(graphWB.getNumberOfSheets()==2){
                sheet=graphWB.createSheet(Strings.GRAPH_SHEET_3);
            }
            else{
                sheet=graphWB.getSheetAt(2);
            }
            
            Row titlesRow = sheet.createRow( 0);
            Row fRow = sheet.createRow( 1);
            Row mRow = sheet.createRow( 2);
            Cell cell = titlesRow.createCell( 1);
            cell.setCellValue(Strings.GRAPH_COLUMN_MENTAL);
            cell = titlesRow.createCell( 2);
            cell.setCellValue(Strings.GRAPH_COLUMN_PHYSICAL);
            cell = titlesRow.createCell( 3);
            cell.setCellValue(Strings.GRAPH_COLUMN_BOTH);
            Cell fcell = fRow.createCell( 0);
            fcell.setCellValue(Strings.PSW_COLUMN_FEMALE);
            fcell = fRow.createCell( 1);
            fcell.setCellValue(mentalFcount);
            fcell = fRow.createCell( 2);
            fcell.setCellValue(physicalFcount);
            fcell = fRow.createCell( 3);
            fcell.setCellValue(bothFcount);
            Cell mcell = mRow.createCell( 0);
            mcell.setCellValue(Strings.PSW_COLUMN_MALE);
            mcell = mRow.createCell( 1);
            mcell.setCellValue(mentalMcount);
            mcell = mRow.createCell( 2);
            mcell.setCellValue(physicalMcount);
            mcell = mRow.createCell( 3);
            mcell.setCellValue(bothMcount);
            fileOut = new FileOutputStream(graphs);
            graphWB.write(fileOut);
            
            fileOut.close();
            fileOutIn.close();
            
        }catch (IOException ex) {
            Logger.getLogger(PoiHelper.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
    
    /**
     * Generates the table for the fourth graph in Graphs.xlsx
     * @param psw
     * @param graphs
     */
    public  void getGraph4(){
        
        int f1Count = 0;
        int ct1Count = 0;
        int st3Count = 0;
        int st6Count = 0;
        FileInputStream fileOutIn;
        FileOutputStream fileOut;
        
        try (FileInputStream fileIn = new FileInputStream(psw)) {
            XSSFWorkbook workbook = new XSSFWorkbook(fileIn);
            XSSFSheet refSheet = workbook.getSheetAt(workbook.getSheetIndex(Strings.PSW_SHEET_REFERRALS));
            int gradeColNo = getCellColumnByString(Strings.PSW_COLUMN_GRADE, refSheet);
            
            Iterator<Row> rowIterator = refSheet.iterator();
            
            while (rowIterator.hasNext())
            {
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
                
                if(f1grades.contains(gradeCell.getStringCellValue())){
                    f1Count++;
                }
                
                if(ct1grades.contains(gradeCell.getStringCellValue())){
                    ct1Count++;
                }
                if(st3grades.contains(gradeCell.getStringCellValue())){
                    st3Count++;
                }
                if(st6grades.contains(gradeCell.getStringCellValue())){
                    st6Count++;
                }
            }
            fileOutIn = new FileInputStream(graphs);
            XSSFSheet sheet;
            XSSFWorkbook graphWB = new XSSFWorkbook(fileOutIn);
            
            if(graphWB.getNumberOfSheets()==3){
                sheet=graphWB.createSheet(Strings.GRAPH_SHEET_4);
            }
            else{
                sheet=graphWB.getSheetAt(3);
            }
            
            double totalCount = f1Count+ct1Count+st3Count+st6Count;
            double f1perc = Math.round(f1Count/totalCount*100);
            double ct1perc = Math.round(ct1Count/totalCount*100);
            double st3perc = Math.round(st3Count/totalCount*100);
            double st6perc = Math.round(st6Count/totalCount*100);
            
            Row titlesRow = sheet.createRow(0);
            Row f1Row = sheet.createRow(1);
            Row ct1Row = sheet.createRow(2);
            Row st3Row = sheet.createRow(3);
            Row st6Row = sheet.createRow(4);
            Cell cell = titlesRow.createCell(1);
            cell.setCellValue(Strings.GRAPH_COLUMN_TOTAL);
            Cell f1cell = f1Row.createCell(0);
            f1cell.setCellValue(Strings.GRAPH_ROW_F1+" "+f1perc+"%");
            f1cell = f1Row.createCell(1);
            f1cell.setCellValue(f1Count);
            Cell ct1Cell = ct1Row.createCell(0);
            ct1Cell.setCellValue(Strings.GRAPH_ROW_CT1+" "+ct1perc+"%");
            ct1Cell = ct1Row.createCell(1);
            ct1Cell.setCellValue(ct1Count);
            Cell st3Cell = st3Row.createCell(0);
            st3Cell.setCellValue(Strings.GRAPH_ROW_ST3+" "+st3perc+"%");
            st3Cell = st3Row.createCell(1);
            st3Cell.setCellValue(st3Count);
            Cell st6Cell = st6Row.createCell(0);
            st6Cell.setCellValue(Strings.GRAPH_ROW_ST6+" "+st6perc+"%");
            st6Cell = st6Row.createCell(1);
            st6Cell.setCellValue(st6Count);
            fileOut = new FileOutputStream(graphs);
            graphWB.write(fileOut);
            
            fileOut.close();
            fileOutIn.close();
            
        }catch (IOException ex) {
            Logger.getLogger(PoiHelper.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
    
    /**
     * Generates the table for the fifth graph in Graphs.xlsx
     * @param psw
     * @param graphs
     */
    public  void getGraph5(){
        int stCount = 0;
        int ftCount = 0;
        int gpCount = 0;
        int otherCount = 0;
        
        FileInputStream fileOutIn;
        FileOutputStream fileOut;
        try (FileInputStream fileIn = new FileInputStream(psw)) {
            XSSFWorkbook workbook = new XSSFWorkbook(fileIn);
            XSSFSheet refSheet = workbook.getSheetAt(workbook.getSheetIndex(Strings.PSW_SHEET_REFERRALS));
            int gradeColNo = getCellColumnByString(Strings.PSW_COLUMN_GRADE, refSheet);
            
            Iterator<Row> rowIterator = refSheet.iterator();
            
            while (rowIterator.hasNext())
            {
                Row row = rowIterator.next();
                
                Cell cell = CellUtil.getCell(row, gradeColNo);
                
                if(cell.getStringCellValue().equals(Strings.PSW_GRADE_FY1)||cell.getStringCellValue().equals(Strings.PSW_GRADE_FY2)){
                    ftCount++;
                }
                if(cell.getStringCellValue().equals(Strings.PSW_GRADE_ST1)||cell.getStringCellValue().equals(Strings.PSW_GRADE_ST2)||cell.getStringCellValue().equals(Strings.PSW_GRADE_ST3)||cell.getStringCellValue().equals(Strings.PSW_GRADE_ST4)||cell.getStringCellValue().equals(Strings.PSW_GRADE_ST5)||cell.getStringCellValue().equals(Strings.PSW_GRADE_ST6)||cell.getStringCellValue().equals(Strings.PSW_GRADE_ST7)||cell.getStringCellValue().equals(Strings.PSW_GRADE_ST8)){
                    stCount++;
                }
                if(cell.getStringCellValue().equals(Strings.PSW_GRADE_GPST1)||cell.getStringCellValue().equals(Strings.PSW_GRADE_GPST2)||cell.getStringCellValue().equals(Strings.PSW_GRADE_GPST3)){
                    gpCount++;
                }
                if(cell.getStringCellValue().equals(Strings.PSW_GRADE_DCT1)||cell.getStringCellValue().equals(Strings.PSW_GRADE_DCT2)||cell.getStringCellValue().equals(Strings.PSW_GRADE_DF1)||cell.getStringCellValue().equals(Strings.PSW_GRADE_DF2)){
                    otherCount++;
                }
            }
            
            fileOutIn = new FileInputStream(graphs);
            XSSFSheet sheet;
            XSSFWorkbook graphWB = new XSSFWorkbook(fileOutIn);
            
            if(graphWB.getNumberOfSheets()==4){
                sheet=graphWB.createSheet(Strings.GRAPH_SHEET_5);
            }
            else{
                sheet=graphWB.getSheetAt(4);
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
            
            fileOut = new FileOutputStream(graphs);
            graphWB.write(fileOut);
            
            fileOut.close();
            fileOutIn.close();
            
        }catch (IOException ex) {
            Logger.getLogger(PoiHelper.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
    
    /**
     * Generates the table for the sixth graph in Graphs.xlsx
     * @param psw
     * @param graphs
     */
    public  void getGraph6(){
        
        FileInputStream fileOutIn;
        FileOutputStream fileOut;
        
        try (FileInputStream fileIn = new FileInputStream(psw)) {
            XSSFWorkbook pswworkbook = new XSSFWorkbook(fileIn);
            XSSFSheet ocSheet = pswworkbook.getSheetAt(pswworkbook.getSheetIndex(Strings.PSW_SHEET_OPEN_CASES));
            Row ocR0 = ocSheet.getRow(0);
            Row ocR1 = ocSheet.getRow(1);
            Row ocR2 = ocSheet.getRow(2);
            
            fileOutIn = new FileInputStream(graphs);
            XSSFWorkbook graphWB = new XSSFWorkbook(fileOutIn);
            XSSFSheet sheet;
            
            if(graphWB.getNumberOfSheets()==5){
                sheet=graphWB.createSheet(Strings.GRAPH_SHEET_6);
            }
            else{
                sheet=graphWB.getSheetAt(5);
            }
            
            Row titlesRow = sheet.createRow(0);
            Row fRow = sheet.createRow(1);
            Row sRow = sheet.createRow(2);
            
            titlesRow.createCell(1).setCellValue(ocR0.getCell(1).getStringCellValue());
            titlesRow.createCell(2).setCellValue(ocR0.getCell(2).getStringCellValue());
            titlesRow.createCell(3).setCellValue(ocR0.getCell(3).getStringCellValue());
            titlesRow.createCell(4).setCellValue(ocR0.getCell(4).getStringCellValue());
            titlesRow.createCell(5).setCellValue(ocR0.getCell(5).getStringCellValue());
            
            
            CellStyle cellStyle = graphWB.createCellStyle();
            CreationHelper createHelper = graphWB.getCreationHelper();
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
            
            fileOut = new FileOutputStream(graphs);
            graphWB.write(fileOut);
            
            fileIn.close();
            fileOut.close();
            fileOutIn.close();
            
        } catch (IOException ex) {
            Logger.getLogger(PoiHelper.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
    
    /**
     * Generates the table for the seventh graph in Graphs.xlsx
     * @param psw
     * @param graphs
     */
    public  void getGraph7(){
        
        FileInputStream fileOutIn;
        FileOutputStream fileOut;
        
        int physicalCount = 0;
        int mentalCount = 0;
        int capability = 0;
        
        try (FileInputStream fileIn = new FileInputStream(psw)) {
            
            
            XSSFWorkbook pswworkbook = new XSSFWorkbook(fileIn);
            XSSFSheet ccSheet = pswworkbook.getSheetAt(pswworkbook.getSheetIndex(Strings.PSW_SHEET_CLOSED_CASES));
            int totalsColumn = getCellColumnByString(Strings.PSW_COLUMN_OUTCOME_KEY, ccSheet);
            
            fileOutIn = new FileInputStream(graphs);
            XSSFWorkbook graphWB = new XSSFWorkbook(fileOutIn);
            XSSFSheet sheet;
            
            if(graphWB.getNumberOfSheets()==6){
                sheet=graphWB.createSheet(Strings.GRAPH_SHEET_7);
            }
            else{
                sheet=graphWB.getSheetAt(6);
            }
            
            fileOut = new FileOutputStream(graphs);
            graphWB.write(fileOut);
            
            fileIn.close();
            fileOut.close();
            fileOutIn.close();
            
        } catch (IOException ex) {
            Logger.getLogger(PoiHelper.class.getName()).log(Level.SEVERE, null, ex);
        }
        
        
    }
    
    /**
     * Generates the table for the eigth graph in Graphs.xlsx
     * @param psw
     * @param graphs
     */
    public  void getGraph8() {
        
        FileInputStream fileOutIn;
        FileOutputStream fileOut;
        
        int noConcerns = 0;
        int onGoing = 0;
        int completed = 0;
        int released = 0;
        int resigned = 0;
        int other = 0;
        int death = 0;
        
        try (FileInputStream fileIn = new FileInputStream(psw)) {
            
            
            XSSFWorkbook pswworkbook = new XSSFWorkbook(fileIn);
            XSSFSheet ccSheet = pswworkbook.getSheetAt(pswworkbook.getSheetIndex(Strings.PSW_SHEET_CLOSED_CASES));
            int totalsColumn = getCellColumnByString(Strings.PSW_COLUMN_OUTCOME_KEY, ccSheet);
            
            fileOutIn = new FileInputStream(graphs);
            XSSFWorkbook graphWB = new XSSFWorkbook(fileOutIn);
            XSSFSheet sheet;
            
            noConcerns = Integer.parseInt(ccSheet.getRow(1).getCell(totalsColumn).getRawValue());
            onGoing = Integer.parseInt(ccSheet.getRow(2).getCell(totalsColumn).getRawValue());
            completed = Integer.parseInt(ccSheet.getRow(3).getCell(totalsColumn).getRawValue());
            released = Integer.parseInt(ccSheet.getRow(4).getCell(totalsColumn).getRawValue());
            resigned = Integer.parseInt(ccSheet.getRow(5).getCell(totalsColumn).getRawValue());
            other = Integer.parseInt(ccSheet.getRow(6).getCell(totalsColumn).getRawValue());
            death = Integer.parseInt(ccSheet.getRow(7).getCell(totalsColumn).getRawValue());
            
            if(graphWB.getNumberOfSheets()<=7){
                sheet=graphWB.createSheet(Strings.GRAPH_SHEET_8);
            }
            else{
                sheet=graphWB.getSheetAt(7);
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
            
            fileOut = new FileOutputStream(graphs);
            graphWB.write(fileOut);
            
            fileIn.close();
            fileOut.close();
            fileOutIn.close();
            
        } catch (IOException ex) {
            Logger.getLogger(PoiHelper.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
    
    public  void getGraph9() {
        
        FileInputStream fileOutIn;
        FileOutputStream fileOut;
        
        try (FileInputStream fileIn = new FileInputStream(psw)) {
            
            XSSFWorkbook pswworkbook = new XSSFWorkbook(fileIn);
            XSSFSheet graphSheet = pswworkbook.getSheetAt(pswworkbook.getSheetIndex(Strings.PSW_SHEET_GRAPHS));
            
            fileOutIn = new FileInputStream(graphs);
            XSSFWorkbook graphWB = new XSSFWorkbook(fileOutIn);
            XSSFSheet sheet;
            
            int titleColumnNo = getCellColumnByString(Strings.PSW_COLUMN_TOTAL_TRAINEE_TIME, graphSheet);
            int titleRowNo = getCellRowByString(Strings.PSW_COLUMN_TOTAL_TRAINEE_TIME, graphSheet);
            
            int aprRow = titleRowNo+1;
            int mayRow = titleRowNo+2;
            int junRow = titleRowNo+3;
            int julRow = titleRowNo+4;
            int augRow = titleRowNo+5;
            int sepRow = titleRowNo+6;
            int octRow = titleRowNo+7;
            int novRow = titleRowNo+8;
            int decRow = titleRowNo+9;
            int janRow = titleRowNo+10;
            int febRow = titleRowNo+11;
            int marRow = titleRowNo+12;
            
            double timeApr = graphSheet.getRow(aprRow).getCell(titleColumnNo).getNumericCellValue() * 24;
            int hoursApr = (int)timeApr;
            int minutesApr = (int)(timeApr - hoursApr) * 60;
            
            double timeMay = graphSheet.getRow(mayRow).getCell(titleColumnNo).getNumericCellValue() * 24;
            int hoursMay = (int)timeMay;
            int minutesMay = (int)(timeMay - hoursMay) * 60;
            
            double timeJun = graphSheet.getRow(junRow).getCell(titleColumnNo).getNumericCellValue() * 24;
            int hoursJun = (int)timeJun;
            int minutesJun = (int)(timeJun - hoursJun) * 60;
            
            double timeJul = graphSheet.getRow(julRow).getCell(titleColumnNo).getNumericCellValue() * 24;
            int hoursJul = (int)timeJul;
            int minutesJul = (int)(timeJul - hoursJul) * 60;
            
            double timeAug = graphSheet.getRow(augRow).getCell(titleColumnNo).getNumericCellValue() * 24;
            int hoursAug = (int)timeAug;
            int minutesAug = (int)(timeAug - hoursAug) * 60;
            
            double timeSep = graphSheet.getRow(sepRow).getCell(titleColumnNo).getNumericCellValue() * 24;
            int hoursSep = (int)timeSep;
            int minutesSep = (int)(timeSep - hoursSep) * 60;
            
            double timeOct = graphSheet.getRow(octRow).getCell(titleColumnNo).getNumericCellValue() * 24;
            int hoursOct = (int)timeOct;
            int minutesOct = (int)(timeOct - hoursOct) * 60;
            
            double timeNov = graphSheet.getRow(novRow).getCell(titleColumnNo).getNumericCellValue() * 24;
            int hoursNov = (int)timeNov;
            int minutesNov = (int)(timeNov - hoursNov) * 60;
            
            double timeDec = graphSheet.getRow(decRow).getCell(titleColumnNo).getNumericCellValue() * 24;
            int hoursDec = (int)timeDec;
            int minutesDec = (int)(timeDec - hoursDec) * 60;
            
            double timeJan = graphSheet.getRow(janRow).getCell(titleColumnNo).getNumericCellValue() * 24;
            int hoursJan = (int)timeJan;
            int minutesJan = (int)(timeJan - hoursApr) * 60;
            
            double timeFeb = graphSheet.getRow(febRow).getCell(titleColumnNo).getNumericCellValue() * 24;
            int hoursFeb = (int)timeFeb;
            int minutesFeb = (int)(timeFeb - hoursFeb) * 60;
            
            double timeMar = graphSheet.getRow(marRow).getCell(titleColumnNo).getNumericCellValue() * 24;
            int hoursMar = (int)timeMar;
            int minutesMar = (int)(timeMar - hoursMar) * 60;
            
            if(graphWB.getNumberOfSheets()<=8){
                sheet=graphWB.createSheet(Strings.GRAPH_SHEET_9);
            }
            else{
                sheet=graphWB.getSheetAt(8);
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
            
            
            fileOut = new FileOutputStream(graphs);
            graphWB.write(fileOut);
            
            fileIn.close();
            fileOut.close();
            fileOutIn.close();
            
        }catch (IOException ex) {
            Logger.getLogger(PoiHelper.class.getName()).log(Level.SEVERE, null, ex);
        }
        
    }
    
    public  void getGraph10() {
        
        FileInputStream fileOutIn;
        FileOutputStream fileOut;
        
        try (FileInputStream fileIn = new FileInputStream(psw)) {
            
            XSSFWorkbook pswworkbook = new XSSFWorkbook(fileIn);
            XSSFSheet graphSheet = pswworkbook.getSheetAt(pswworkbook.getSheetIndex(Strings.PSW_SHEET_GRAPHS));
            
            fileOutIn = new FileInputStream(graphs);
            XSSFWorkbook graphWB = new XSSFWorkbook(fileOutIn);
            XSSFSheet sheet;
            
            
            int titleColumnNo = getCellColumnByString(Strings.PSW_COLUMN_TOTAL_TRAINEE_COSTS, graphSheet);
            int titleRowNo = getCellRowByString(Strings.PSW_COLUMN_TOTAL_TRAINEE_COSTS, graphSheet);
            int aprRow = titleRowNo+1;
            int mayRow = titleRowNo+2;
            int junRow = titleRowNo+3;
            int julRow = titleRowNo+4;
            int augRow = titleRowNo+5;
            int sepRow = titleRowNo+6;
            int octRow = titleRowNo+7;
            int novRow = titleRowNo+8;
            int decRow = titleRowNo+9;
            int janRow = titleRowNo+10;
            int febRow = titleRowNo+11;
            int marRow = titleRowNo+12;
            
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
            
            
            if(graphWB.getNumberOfSheets()<=9){
                sheet=graphWB.createSheet(Strings.GRAPH_SHEET_10);
            }
            else{
                sheet=graphWB.getSheetAt(9);
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
            
            
            fileOut = new FileOutputStream(graphs);
            graphWB.write(fileOut);
            
            fileIn.close();
            fileOut.close();
            fileOutIn.close();
            
        }catch (IOException ex) {
            Logger.getLogger(PoiHelper.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
    
    public  void getGraph11() {
        
        FileInputStream fileOutIn;
        FileOutputStream fileOut;
        
        try (FileInputStream fileIn = new FileInputStream(psw)) {
            
            XSSFWorkbook pswworkbook = new XSSFWorkbook(fileIn);
            XSSFSheet graphSheet = pswworkbook.getSheetAt(pswworkbook.getSheetIndex(Strings.PSW_SHEET_GRAPHS));
            
            fileOutIn = new FileInputStream(graphs);
            XSSFWorkbook graphWB = new XSSFWorkbook(fileOutIn);
            XSSFSheet sheet;
            
            int ssgColumnNo = getCellColumnByString(Strings.PSW_COLUMN_SSG_COSTS, graphSheet);
            int cmColumnNo = getCellColumnByString(Strings.PSW_COLUMN_CM_COSTS, graphSheet);
            int titleRowNo = getCellRowByString(Strings.PSW_COLUMN_SSG_COSTS, graphSheet);
            int aprRow = titleRowNo+1;
            int mayRow = titleRowNo+2;
            int junRow = titleRowNo+3;
            int julRow = titleRowNo+4;
            int augRow = titleRowNo+5;
            int sepRow = titleRowNo+6;
            int octRow = titleRowNo+7;
            int novRow = titleRowNo+8;
            int decRow = titleRowNo+9;
            int janRow = titleRowNo+10;
            int febRow = titleRowNo+11;
            int marRow = titleRowNo+12;
            
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
            
            if(graphWB.getNumberOfSheets()<=9){
                sheet=graphWB.createSheet(Strings.GRAPH_SHEET_10);
            }
            else{
                sheet=graphWB.getSheetAt(9);
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
            
            fileOut = new FileOutputStream(graphs);
            graphWB.write(fileOut);
            
            fileIn.close();
            fileOut.close();
            fileOutIn.close();
            
        } catch (IOException ex) {
            Logger.getLogger(PoiHelper.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
    
    /**
     * Returns the counts for the first table on the document
     * @param file
     * @return
     */
    public  List<Integer> countTable0(){
        List<Integer> list = new ArrayList<>();
        
        int stCount = 0;
        int fCount = 0;
        int gpCount = 0;
        int otherCount = 0;
        int totalCount = 0;
        int casesClosed = 0;
        int casesOClosed = 0;
        
        
        try (FileInputStream fileIn = new FileInputStream(psw)) {
            
            
            XSSFWorkbook pswworkbook = new XSSFWorkbook(fileIn);
            XSSFSheet refSheet = pswworkbook.getSheetAt(pswworkbook.getSheetIndex(Strings.PSW_SHEET_REFERRALS));
            int gradeColNo = getCellColumnByString(Strings.PSW_COLUMN_GRADE, refSheet);
            Iterator<Row> rowIterator = refSheet.iterator();
            
            while (rowIterator.hasNext())
            {
                Row row = rowIterator.next();
                if(!isRowEmpty(row)){
                    if(!isCellEmpty(row.getCell(gradeColNo))){
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
            
            totalCount=fCount+stCount+gpCount+otherCount;
            
            list.add(stCount);
            list.add(fCount);
            list.add(gpCount);
            list.add(otherCount);
            list.add(totalCount);
        } catch (IOException ex) {
            Logger.getLogger(PoiHelper.class.getName()).log(Level.SEVERE, null, ex);
        }
        return list;
    }
    
    /**
     * Counts the values for table number 2: Referral Reason and gender split for financial years... It uses "Graphs" spreadsheet, since this data is already calculated
     * @param graphs
     * @param psw
     * @return
     */
    public  List<Integer> countTable1(){
        
        List<Integer> list = new ArrayList<>();
        
        int capFCount = 0;
        int capMCount = 0;
        int exFailFCount = 0;
        int exFailMCount = 0;
        int healthFCount = 0;
        int healthMCount = 0;
        int careerFCount = 0;
        int careerMCount = 0;
        int conductFCount = 0;
        int conductMCount = 0;
        int otherFCount = 0;
        int otherMCount = 0;
        int referredCount = countTotalReferrals();
        
        
        try (FileInputStream graphsFileIn = new FileInputStream(graphs)) {
            
            FileInputStream pswFileIn = new FileInputStream(psw);
            XSSFWorkbook pswWorkbook = new XSSFWorkbook(pswFileIn);
            XSSFSheet refSheet = pswWorkbook.getSheetAt(pswWorkbook.getSheetIndex(Strings.PSW_SHEET_REFERRALS));
            XSSFWorkbook graphsworkbook = new XSSFWorkbook(graphsFileIn);
            XSSFSheet graphSheet = graphsworkbook.getSheetAt(graphsworkbook.getSheetIndex(Strings.GRAPH_SHEET_2));
            int capColNo = getCellColumnByString(Strings.PSW_COLUMN_CAPABILITY, graphSheet);
            int exFailColNo = getCellColumnByString(Strings.GRAPH_COLUMN_EXAM_FAILURE_DIRECT, graphSheet);
            int healthColNo = getCellColumnByString(Strings.GRAPH_COLUMN_HEALTH, graphSheet);
            int carreerColNo = getCellColumnByString(Strings.GRAPH_COLUMN_CARREER_DIRECT, graphSheet);
            int conductColNo = getCellColumnByString(Strings.PSW_COLUMN_CONDUCT, graphSheet);
            int otherColNo = getCellColumnByString(Strings.GRAPH_COLUMN_OTHER, graphSheet);
            int femRowNo = getCellRowByString(Strings.PSW_COLUMN_FEMALE, graphSheet);
            int maleRowNo = getCellRowByString(Strings.PSW_COLUMN_MALE, graphSheet);
            
            capFCount = (int) graphSheet.getRow(femRowNo).getCell(capColNo).getNumericCellValue();
            exFailFCount = (int) graphSheet.getRow(femRowNo).getCell(exFailColNo).getNumericCellValue();
            healthFCount = (int) graphSheet.getRow(femRowNo).getCell(healthColNo).getNumericCellValue();
            careerFCount = (int) graphSheet.getRow(femRowNo).getCell(carreerColNo).getNumericCellValue();
            conductFCount = (int) graphSheet.getRow(femRowNo).getCell(conductColNo).getNumericCellValue();
            otherFCount = (int) graphSheet.getRow(femRowNo).getCell(otherColNo).getNumericCellValue();
            
            capMCount = (int) graphSheet.getRow(maleRowNo).getCell(capColNo).getNumericCellValue();
            exFailMCount = (int) graphSheet.getRow(maleRowNo).getCell(exFailColNo).getNumericCellValue();
            healthMCount = (int) graphSheet.getRow(maleRowNo).getCell(healthColNo).getNumericCellValue();
            careerMCount = (int) graphSheet.getRow(maleRowNo).getCell(carreerColNo).getNumericCellValue();
            conductMCount = (int) graphSheet.getRow(maleRowNo).getCell(conductColNo).getNumericCellValue();
            otherMCount = (int) graphSheet.getRow(maleRowNo).getCell(otherColNo).getNumericCellValue();
            
            list.add(capFCount);
            list.add(capMCount);
            list.add(exFailFCount);
            list.add(exFailMCount);
            list.add(healthFCount);
            list.add(healthMCount);
            list.add(careerFCount);
            list.add(careerMCount);
            list.add(conductFCount);
            list.add(conductMCount);
            list.add(otherFCount);
            list.add(otherMCount);
            list.add(referredCount);
            
        } catch (IOException ex) {
            Logger.getLogger(PoiHelper.class.getName()).log(Level.SEVERE, null, ex);
        }
        
        return list;
    }
    
    public  List<Integer> countTable2(){
        
        List<Integer> list = new ArrayList<>();
        int referredCount = countTotalReferrals();
        int f1TotalCount = 0;
        int f2TotalCount = 0;
        int f1ReferredCount = 0;
        int f2ReferredCount = 0;
        
        try{
            
            FileInputStream pswFileIn = new FileInputStream(psw);
            XSSFWorkbook pswWorkbook = new XSSFWorkbook(pswFileIn);
            XSSFSheet refSheet = pswWorkbook.getSheetAt(pswWorkbook.getSheetIndex(Strings.PSW_SHEET_REFERRALS));
            int gradeColNo = getCellColumnByString(Strings.PSW_COLUMN_GRADE, refSheet);
            Iterator<Row> rowIteratorRefSheet = refSheet.iterator();
            
            XSSFSheet foundationSheet = pswWorkbook.getSheetAt(pswWorkbook.getSheetIndex(Strings.PSW_SHEET_FOUNDATION));
            
            
            int countColNo = getCellColumnByString(Strings.PSW_COLUMN_TRAINEE_COUNT, foundationSheet);
            int f1CountRowNo = getCellRowByString(Strings.PSW_ROW_F1COUNT, foundationSheet);
            int f2CountRowNo = getCellRowByString(Strings.PSW_ROW_F2COUNT, foundationSheet);
            
            
            while(rowIteratorRefSheet.hasNext()){
                Row row = rowIteratorRefSheet.next();
                if(row.getRowNum()>3&&row.getCell(gradeColNo)==null){
                    break;
                }
                if(row.getRowNum()>3&&!row.getCell(gradeColNo).getStringCellValue().equals("")&&row.getCell(gradeColNo).getStringCellValue().equals(Strings.PSW_GRADE_FY1)){
                    f1ReferredCount++;
                }
                if(row.getRowNum()>3&&row.getCell(gradeColNo).getStringCellValue().equals(Strings.PSW_GRADE_FY2)){
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
    
    public  List<Integer> countTable3() {
        
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
        
        try{
            FileInputStream pswFileIn = new FileInputStream(psw);
            XSSFWorkbook pswWorkbook = new XSSFWorkbook(pswFileIn);
            XSSFSheet refSheet = pswWorkbook.getSheetAt(pswWorkbook.getSheetIndex(Strings.PSW_SHEET_REFERRALS));
            int trustRColNo = getCellColumnByString(Strings.PSW_COLUMN_TRUST, refSheet);
            Iterator<Row> rowIteratorRefSheet = refSheet.iterator();
            
            XSSFSheet trustSheet = pswWorkbook.getSheetAt(pswWorkbook.getSheetIndex(Strings.PSW_SHEET_TRUST));
            int traineeCountColNo = getCellColumnByString(Strings.PSW_COLUMN_TRAINEE_COUNT, trustSheet);
            int trustColNo = getCellColumnByString(Strings.PSW_COLUMN_TRUST, trustSheet);
            Iterator<Row> rowIteratorTrustSheet = trustSheet.iterator();
            
            totalWessex = countTotalWessex();
            
            while(rowIteratorTrustSheet.hasNext()){
                Row row = rowIteratorTrustSheet.next();
                switch (row.getCell(trustColNo).getStringCellValue()) {
                    case Strings.PSW_ROW_BOURNEMOUTH:
                        bournemouthTotal = (int)row.getCell(traineeCountColNo).getNumericCellValue();
                        break;
                    case Strings.PSW_ROW_DORCHESTER:
                        dorchesterTotal = (int)row.getCell(traineeCountColNo).getNumericCellValue();
                        break;
                    case Strings.PSW_ROW_DORSET:
                        dorsetTotal = (int)row.getCell(traineeCountColNo).getNumericCellValue();
                        break;
                    case Strings.PSW_ROW_HHFT:
                        hhftTotal = (int)row.getCell(traineeCountColNo).getNumericCellValue();
                        break;
                    case Strings.PSW_ROW_IOW:
                        iowTotal = (int)row.getCell(traineeCountColNo).getNumericCellValue();
                        break;
                    case Strings.PSW_ROW_JERSEY:
                        jerseyTotal = (int)row.getCell(traineeCountColNo).getNumericCellValue();
                        break;
                    case Strings.PSW_ROW_POOLE:
                        pooleTotal = (int)row.getCell(traineeCountColNo).getNumericCellValue();
                        break;
                    case Strings.PSW_ROW_PORTSMOUTH:
                        portsmouthTotal = (int)row.getCell(traineeCountColNo).getNumericCellValue();
                        break;
                    case Strings.PSW_ROW_SALISBURY:
                        salisburyTotal = (int)row.getCell(traineeCountColNo).getNumericCellValue();
                        break;
                    case Strings.PSW_ROW_SOLENT:
                        solentTotal = (int)row.getCell(traineeCountColNo).getNumericCellValue();
                        break;
                    case Strings.PSW_ROW_SOUTHAMPTON:
                        southamptonTotal = (int)row.getCell(traineeCountColNo).getNumericCellValue();
                        break;
                    case Strings.PSW_ROW_SOUTHERN:
                        southernTotal= (int)row.getCell(traineeCountColNo).getNumericCellValue();
                        break;
                }
                
            }
            
            while(rowIteratorRefSheet.hasNext()){
                Row row = rowIteratorRefSheet.next();
                if(!isRowEmpty(row)){
                    if(!isCellEmpty(row.getCell(trustRColNo))){
                        
                        if(row.getCell(trustRColNo).getStringCellValue().contains(Strings.PSW_TRUST_BOURNEMOUTH)){
                            bournemouthRefNo++;
                        }
                        else if(row.getCell(trustRColNo).getStringCellValue().contains(Strings.PSW_TRUST_DORCHESTER)){
                            dorchesterRefNo++;
                        }
                        else if(row.getCell(trustRColNo).getStringCellValue().contains(Strings.PSW_TRUST_DORSET)){
                            dorsetRefNo++;
                        }
                        else if(row.getCell(trustRColNo).getStringCellValue().contains(Strings.PSW_TRUST_HHFT)){
                            hhftRefNo++;
                        }
                        else if(row.getCell(trustRColNo).getStringCellValue().contains(Strings.PSW_TRUST_IOW)){
                            iowRefNo++;
                        }
                        else if(row.getCell(trustRColNo).getStringCellValue().contains(Strings.PSW_TRUST_JERSEY)){
                            jerseyRefNo++;
                        }
                        else if(row.getCell(trustRColNo).getStringCellValue().contains(Strings.PSW_TRUST_POOLE)){
                            pooleRefNo++;
                        }
                        else if(row.getCell(trustRColNo).getStringCellValue().contains(Strings.PSW_TRUST_PORTSMOUTH)){
                            portsmouthRefNo++;
                        }
                        else if(row.getCell(trustRColNo).getStringCellValue().contains(Strings.PSW_TRUST_SALISBURY)){
                            salisburyRefNo++;
                        }
                        else if(row.getCell(trustRColNo).getStringCellValue().contains(Strings.PSW_TRUST_SOLENT)){
                            solentRefNo++;
                        }
                        else if(row.getCell(trustRColNo).getStringCellValue().contains(Strings.PSW_TRUST_SOUTHAMPTON)){
                            southamptonRefNo++;
                        }
                        else if(row.getCell(trustRColNo).getStringCellValue().contains(Strings.PSW_TRUST_SOUTHERN)){
                            southernRefNo++;
                        }
                        
                    }
                }
            }
        } catch (IOException ex) {
            Logger.getLogger(PoiHelper.class.getName()).log(Level.SEVERE, null, ex);
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
    
    public  List<Integer> countTable4(){
        
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
        
        try{
            XSSFWorkbook pswWorkbook = new XSSFWorkbook(new FileInputStream(psw));
            XSSFSheet refSheet = pswWorkbook.getSheet(Strings.PSW_SHEET_REFERRALS);
            Iterator<Row> rowIteratorRefSheet = refSheet.rowIterator();
            int specialtyRColNo = getCellColumnByString(Strings.PSW_COLUMN_SPECIALTY, refSheet);
            
            XSSFSheet specialtySheet = pswWorkbook.getSheet(Strings.PSW_SHEET_SPECIALTY);
            Iterator<Row> rowIteratorSpcSheet = specialtySheet.rowIterator();
            int specialtyColNo = getCellColumnByString(Strings.PSW_COLUMN_SPECIALTY, specialtySheet);
            int traineeCountColNo = getCellColumnByString(Strings.PSW_COLUMN_TRAINEE_COUNT, specialtySheet);
            
            
            while(rowIteratorRefSheet.hasNext()){
                Row row = rowIteratorRefSheet.next();
                
                if(!isRowEmpty(row)){
                    if(!isCellEmpty(row.getCell(specialtyRColNo))){
                        Cell c = row.getCell(specialtyRColNo);
                        if(c.getStringCellValue().contains(Strings.PSW_SPECIALTY_ANAESTHETICS)){
                            anaestheticsRefNo++;
                        }
                        else if(c.getStringCellValue().contains(Strings.PSW_SPECIALTY_DENTAL)){
                            dentalRefNo++;
                        }
                        else if(c.getStringCellValue().contains(Strings.PSW_SPECIALTY_EMERGENCY)){
                            emergRefNo++;
                        }
                        else if(c.getStringCellValue().contains(Strings.PSW_SPECIALTY_FOUNDATION)){
                            foundationRefNo++;
                        }
                        else if(c.getStringCellValue().contains(Strings.PSW_SPECIALTY_GENERAL_PRACTICE)){
                            gpRefNo++;
                        }
                        else if(c.getStringCellValue().contains(Strings.PSW_SPECIALTY_MEDICINE)){
                            medicineRefNo++;
                        }
                        else if(c.getStringCellValue().contains(Strings.PSW_SPECIALTY_OBS)){
                            obsRefNo++;
                        }
                        else if(c.getStringCellValue().contains(Strings.PSW_SPECIALTY_OCC_HEALTH)){
                            occhealthRefNo++;
                        }
                        else if(c.getStringCellValue().contains(Strings.PSW_SPECIALTY_PAEDIATRICS)){
                            paediatricsRefNo++;
                        }
                        else if(c.getStringCellValue().contains(Strings.PSW_SPECIALTY_PATHOLOGY)){
                            pathologyRefNo++;
                        }
                        else if(c.getStringCellValue().contains(Strings.PSW_SPECIALTY_PHARMACY)){
                            pharmacyRefNo++;
                        }
                        else if(c.getStringCellValue().contains(Strings.PSW_SPECIALTY_PSYCHIATRY)){
                            psychRefNo++;
                        }
                        else if(c.getStringCellValue().contains(Strings.PSW_SPECIALTY_PUBLIC_HEALTH)){
                            pubhealthRefNo++;
                        }
                        else if(c.getStringCellValue().contains(Strings.PSW_SPECIALTY_RADIOLOGY)){
                            radioRefNo++;
                        }
                        else if(c.getStringCellValue().contains(Strings.PSW_SPECIALTY_SURGERY)){
                            surgeryRefNo++;
                        }
                    }
                }
            }
            
            while(rowIteratorSpcSheet.hasNext()){
                Row row = rowIteratorSpcSheet.next();
                
                if(!isRowEmpty(row)){
                    if(!isCellEmpty(row.getCell(specialtyColNo))){
                        switch(row.getCell(specialtyColNo).getStringCellValue()){
                            case Strings.PSW_ROW_ANAESTHETICS_A:
                            case Strings.PSW_ROW_ANAESTHETICS_B:
                                anaestheticsTotal=anaestheticsTotal+row.getCell(traineeCountColNo).getNumericCellValue();
                                break;
                            case Strings.PSW_ROW_DENTAL_A:
                            case Strings.PSW_ROW_DENTAL_B:
                                dentalTotal=dentalTotal+row.getCell(traineeCountColNo).getNumericCellValue();
                                break;
                            case Strings.PSW_ROW_EMERGENCY_A:
                            case Strings.PSW_ROW_EMERGENCY_B:
                                emergTotal=emergTotal+row.getCell(traineeCountColNo).getNumericCellValue();
                                break;
                            case Strings.PSW_ROW_FOUNDATION:
                                foundationTotal=row.getCell(traineeCountColNo).getNumericCellValue();
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
                                gpTotal=gpTotal+row.getCell(traineeCountColNo).getNumericCellValue();
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
                                medicineTotal=medicineTotal+row.getCell(traineeCountColNo).getNumericCellValue();
                                break;
                            case Strings.PSW_ROW_OBS:
                                obsTotal=row.getCell(traineeCountColNo).getNumericCellValue();
                                break;
                            case Strings.PSW_ROW_OC_HEALTH:
                                occhealthTotal=row.getCell(traineeCountColNo).getNumericCellValue();
                                break;
                            case Strings.PSW_ROW_PAEDIATRICS_A:
                            case Strings.PSW_ROW_PAEDIATRICS_B:
                                paediatricsTotal=paediatricsTotal+row.getCell(traineeCountColNo).getNumericCellValue();
                                break;
                            case Strings.PSW_ROW_PATHOLOGY_A:
                            case Strings.PSW_ROW_PATHOLOGY_B:
                                pathologyTotal=pathologyTotal+row.getCell(traineeCountColNo).getNumericCellValue();
                                break;
                            case Strings.PSW_ROW_PSYCH_A:
                            case Strings.PSW_ROW_PSYCH_B:
                            case Strings.PSW_ROW_PSYCH_C:
                            case Strings.PSW_ROW_PSYCH_D:
                            case Strings.PSW_ROW_PSYCH_E:
                            case Strings.PSW_ROW_PSYCH_F:
                            case Strings.PSW_ROW_PSYCH_G:
                                psychTotal=psychTotal+pathologyTotal+row.getCell(traineeCountColNo).getNumericCellValue();
                                break;
                            case Strings.PSW_ROW_PUBLIC_HEALTH:
                                pubhealthTotal=row.getCell(traineeCountColNo).getNumericCellValue();
                                break;
                            case Strings.PSW_ROW_RADIOLOGY:
                                radioTotal=row.getCell(traineeCountColNo).getNumericCellValue();
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
                                surgeryTotal=surgeryTotal+row.getCell(traineeCountColNo).getNumericCellValue();
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
            list.add((int)anaestheticsTotal);
            list.add((int)dentalTotal);
            list.add((int)emergTotal);
            list.add((int)foundationTotal);
            list.add((int)gpTotal);
            list.add((int)medicineTotal);
            list.add((int)obsTotal);
            list.add((int)occhealthTotal);
            list.add((int)paediatricsTotal);
            list.add((int)pathologyTotal);
            list.add((int)pharmacyTotal);
            list.add((int)psychTotal);
            list.add((int)pubhealthTotal);
            list.add((int)radioTotal);
            list.add((int)surgeryTotal);
            list.add(totalRefs);
            list.add(totalWssx);
            
            
        } catch (IOException ex) {
            Logger.getLogger(PoiHelper.class.getName()).log(Level.SEVERE, null, ex);
        }
        
        return list;
    }
    
    public List<String> getTable5Line(String trst, String lastLn){
        
        List<String> list = new ArrayList<>();
        
        String trust = trst;
        String gender = "";
        String grade = "";
        String school = "";
        String addRef = "";
        String country = "";
        String age = "";
        String ethicity = "";
        String sexOr = "";
        String religion = "";
        String disability = "";
        String lastLine = lastLn;
        
        try{
            FileInputStream pswFileIn = new FileInputStream(psw);
            XSSFWorkbook pswWorkbook = new XSSFWorkbook(pswFileIn);
            XSSFSheet refSheet = pswWorkbook.getSheetAt(pswWorkbook.getSheetIndex(Strings.PSW_SHEET_REFERRALS));
            int exSupportColNo = getCellColumnByString(Strings.PSW_COLUMN_EXAM, refSheet);
            int anxietyColNo = getCellColumnByString(Strings.PSW_COLUMN_ANXIETY, refSheet);
            int carreerColNo = getCellColumnByString(Strings.PSW_COLUMN_CARREER, refSheet);
            int clinicalSkillsColNo = getCellColumnByString(Strings.PSW_COLUMN_CLINICAL_SKILLS, refSheet);
            int communicationColNo = getCellColumnByString(Strings.PSW_COLUMN_COMMUNICATION, refSheet);
            int conductColNo = getCellColumnByString(Strings.PSW_COLUMN_CONDUCT, refSheet);
            int culturalColNo = getCellColumnByString(Strings.PSW_COLUMN_CULTURAL, refSheet);
            int mentalHealthColNo = getCellColumnByString(Strings.PSW_COLUMN_MENTAL, refSheet);
            int physicalHealthColNo = getCellColumnByString(Strings.PSW_COLUMN_PHYSICAL, refSheet);
            int languageColNo = getCellColumnByString(Strings.PSW_COLUMN_LANGUAGE, refSheet);
            int profColNo = getCellColumnByString(Strings.PSW_COLUMN_PROFFESSIONALISM, refSheet);
            int adhdColNo = getCellColumnByString(Strings.PSW_COLUMN_ADHD, refSheet);
            int asdColNo = getCellColumnByString(Strings.PSW_COLUMN_ASD, refSheet);
            int dyslexiaColNo = getCellColumnByString(Strings.PSW_COLUMN_DYSLEXIA, refSheet);
            int dyspraxiaColNo = getCellColumnByString(Strings.PSW_COLUMN_DYSPRAXIA, refSheet);
            int srttColNo = getCellColumnByString(Strings.PSW_COLUMN_SRTT, refSheet);
            int teamColNo = getCellColumnByString(Strings.PSW_COLUMN_TEAM, refSheet);
            int timeColNo = getCellColumnByString(Strings.PSW_COLUMN_TIME, refSheet);
            int genderColNo = getCellColumnByString(Strings.PSW_COLUMN_GENDER, refSheet);
            int trustColNo = getCellColumnByString(Strings.PSW_COLUMN_TRUST, refSheet);
            int gradeColNo = getCellColumnByString(Strings.PSW_COLUMN_GRADE, refSheet);
            int schoolColNo = getCellColumnByString(Strings.PSW_COLUMN_SPECIALTY, refSheet);
            int countryColNo = getCellColumnByString(Strings.PSW_COLUMN_COUNTRY, refSheet);
            int ageColNo = getCellColumnByString(Strings.PSW_COLUMN_AGE, refSheet);
            int religionColNo = getCellColumnByString(Strings.PSW_COLUMN_RELIGION, refSheet);
            int ethColNo = getCellColumnByString(Strings.PSW_COLUMN_ETHNICITY, refSheet);
            int sexOrColNo = getCellColumnByString(Strings.PSW_COLUMN_SEXUAL_OR, refSheet);
            int disabilityColNo = getCellColumnByString(Strings.PSW_COLUMN_DISABILITY, refSheet);
            
            Iterator<Row> rowIteratorRefSheet = refSheet.iterator();
            while(rowIteratorRefSheet.hasNext()){
                Row r = rowIteratorRefSheet.next();
                if(!isRowEmpty(r)){
                    if(r.getRowNum()>3&&r.getRowNum()>Integer.parseInt(lastLn)&&r.getCell(trustColNo).getStringCellValue().contains(trust)&&!isCellEmpty(r.getCell(exSupportColNo))){
                        
                        lastLine = String.valueOf(r.getRowNum());
                        if(!isCellEmpty(r.getCell(genderColNo))){
                            gender = r.getCell(genderColNo).getStringCellValue();
                        }
                        
                        if(!isCellEmpty(r.getCell(gradeColNo))){
                            grade = r.getCell(gradeColNo).getStringCellValue();
                        }
                        
                        if(!isCellEmpty(r.getCell(schoolColNo))){
                            school = r.getCell(schoolColNo).getStringCellValue();
                        }
                        
                        if(!isCellEmpty(r.getCell(countryColNo))){
                            country = r.getCell(countryColNo).getStringCellValue();
                        }
                        
                        if(!isCellEmpty(r.getCell(ageColNo))){
                            age = String.valueOf(Math.round(r.getCell(ageColNo).getNumericCellValue()));
                        }
                        
                        if(!isCellEmpty(r.getCell(religionColNo))){
                            religion = r.getCell(religionColNo).getStringCellValue();
                        }
                        
                        if(!isCellEmpty(r.getCell(sexOrColNo))){
                            sexOr = r.getCell(sexOrColNo).getStringCellValue();
                        }
                        
                        if(!isCellEmpty(r.getCell(disabilityColNo))){
                            disability = r.getCell(disabilityColNo).getStringCellValue();
                        }
                        
                        if(!isCellEmpty(r.getCell(ethColNo))){
                            ethicity = r.getCell(ethColNo).getStringCellValue();
                        }
                        
                        if(!isCellEmpty(r.getCell(anxietyColNo))){
                            addRef = Strings.PSW_COLUMN_ANXIETY;
                        }
                        if(!isCellEmpty(r.getCell(carreerColNo))){
                            addRef = addRef+"; "+Strings.PSW_COLUMN_CARREER;
                        }
                        if(!isCellEmpty(r.getCell(clinicalSkillsColNo))){
                            addRef = addRef+"; "+Strings.PSW_COLUMN_CLINICAL_SKILLS;
                        }
                        if(!isCellEmpty(r.getCell(communicationColNo))){
                            addRef = addRef+"; "+Strings.PSW_COLUMN_COMMUNICATION;
                        }
                        if(!isCellEmpty(r.getCell(conductColNo))){
                            addRef = addRef+"; "+Strings.PSW_COLUMN_CONDUCT;
                        }
                        if(!isCellEmpty(r.getCell(culturalColNo))){
                            addRef = addRef+"; "+Strings.PSW_COLUMN_CULTURAL;
                        }
                        if(!isCellEmpty(r.getCell(mentalHealthColNo))){
                            addRef = addRef+"; "+Strings.PSW_COLUMN_MENTAL;
                        }
                        if(!isCellEmpty(r.getCell(physicalHealthColNo))){
                            addRef = addRef+"; "+Strings.PSW_COLUMN_PHYSICAL;
                        }
                        if(!isCellEmpty(r.getCell(languageColNo))){
                            addRef = addRef+"; "+Strings.PSW_COLUMN_LANGUAGE;
                        }
                        if(!isCellEmpty(r.getCell(profColNo))){
                            addRef = addRef+"; "+Strings.PSW_COLUMN_PROFFESSIONALISM;
                        }
                        if(!isCellEmpty(r.getCell(adhdColNo))){
                            addRef = addRef+"; "+Strings.PSW_COLUMN_ADHD;
                        }
                        if(!isCellEmpty(r.getCell(asdColNo))){
                            addRef = addRef+"; "+Strings.PSW_COLUMN_ASD;
                        }
                        if(!isCellEmpty(r.getCell(dyslexiaColNo))){
                            addRef = addRef+"; "+Strings.PSW_COLUMN_DYSLEXIA;
                        }
                        if(!isCellEmpty(r.getCell(dyspraxiaColNo))){
                            addRef = addRef+"; "+Strings.PSW_COLUMN_DYSPRAXIA;
                        }
                        if(!isCellEmpty(r.getCell(srttColNo))){
                            addRef = addRef+"; "+Strings.PSW_COLUMN_SRTT;
                        }
                        if(!isCellEmpty(r.getCell(teamColNo))){
                            addRef = addRef+"; "+Strings.PSW_COLUMN_TEAM;
                        }
                        if(!isCellEmpty(r.getCell(timeColNo))){
                            addRef = addRef+"; "+Strings.PSW_COLUMN_TIME;
                        }
                    }
                    if(!gender.equals("")){
                        break;
                    }
                }
            }
            
            list.add(gender);
            list.add(grade);
            list.add(school);
            list.add(addRef);
            list.add(country);
            list.add(age);
            list.add(ethicity);
            list.add(sexOr);
            list.add(religion);
            list.add(disability);
            list.add(lastLine);
            
        }catch (IOException ex) {
            Logger.getLogger(PoiHelper.class.getName()).log(Level.SEVERE, null, ex);
        }
        return list;
    }
    
    public  void countOpenedAndClosed(){
        
        int closedWithinYr = 0;
        int closedLastYr = 0;
        int openedAndClosedWithinYr = 0;
        DataFormatter formatter = new DataFormatter(Locale.UK);
        
        try(XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(psw))){
            
            XSSFSheet closedCasesSheet = wb.getSheet(Strings.PSW_SHEET_CLOSED_CASES);
            int dateOpenedColumn = getCellColumnByString("Date opened", closedCasesSheet);
            int dateClosedColumn = getCellColumnByString("Date Closed", closedCasesSheet);
            
            Iterator<Row> rowIterator = closedCasesSheet.iterator();
            boolean inData = true;
            
            while(rowIterator.hasNext()){
                
                Row row = rowIterator.next();
                
                if (row == null || isCellEmpty(row.getCell(dateOpenedColumn))||isCellEmpty(row.getCell(dateClosedColumn))) {
                    inData = false;
                }
                
                else if(row.getRowNum()>14){
                    
                    Cell dateOpenedCell = row.getCell(dateOpenedColumn);
                    Cell dateClosedCell = row.getCell(dateClosedColumn);
                    
                    Date dateOpened = dateOpenedCell.getDateCellValue();
                    Date dateClosed = dateClosedCell.getDateCellValue();
                    
                    Calendar dateOpenedCal = new GregorianCalendar();
                    dateOpenedCal.setTime(dateOpened);
                    
                    Calendar dateClosedCal = new GregorianCalendar();
                    dateClosedCal.setTime(dateClosed);
                    
                    
                    int dateOpenedYear = dateOpenedCal.get(Calendar.YEAR);
                    int dateOpenedMonth = dateOpenedCal.get(Calendar.MONTH);
                    int dateClosedYear = dateClosedCal.get(Calendar.YEAR);
                    int dateClosedMonth = dateClosedCal.get(Calendar.MONTH);
                    
                    
                    if(dateOpenedYear==DocHelper.getYear()&&dateClosedMonth<=03){
                        openedAndClosedWithinYr++;
                    }
                    if(dateOpenedYear==DocHelper.getYear()-1){
                        if(dateOpenedMonth>03){
                            openedAndClosedWithinYr++;
                        }
                    }
                    if(dateOpenedYear==DocHelper.getYear()-1){
                        if(dateOpenedMonth<03){
                            closedWithinYr++;
                        }
                    }
                    if(dateClosedYear==DocHelper.getYear()-1&&dateClosedMonth>03&&dateOpenedMonth<03&&dateOpenedYear==DocHelper.getYear()-1){
                        closedWithinYr++;
                    }
                    if(dateOpenedYear<DocHelper.getYear()-1&&dateClosedYear>=DocHelper.getYear()-1){
                        closedWithinYr++;
                    }
                }
            }
            
            
        } catch (IOException ex) {
            Logger.getLogger(PoiHelper.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
    
    
    private Integer countTotalReferrals(){
        int totalRef = 0;
        try{
            FileInputStream fis = new FileInputStream(psw);
            XSSFWorkbook wb = new XSSFWorkbook(fis);
            XSSFSheet refSheet = wb.getSheet(Strings.PSW_SHEET_REFERRALS);
            Iterator<Row> rowIterator = refSheet.rowIterator();
            
            int firstRow = getCellRowByString(Strings.PSW_COLUMN_ADHD, refSheet);
            
            while(rowIterator.hasNext()){
                Row row = rowIterator.next();
                if(!isRowEmpty(row)&&row.getRowNum()>firstRow){
                    totalRef++;
                }
            }
            
        } catch (FileNotFoundException ex) {
            Logger.getLogger(PoiHelper.class.getName()).log(Level.SEVERE, null, ex);
        } catch (IOException ex) {
            Logger.getLogger(PoiHelper.class.getName()).log(Level.SEVERE, null, ex);
        }
        return totalRef;
    }
    
    private  Integer countTotalWessex(){
        int totalRef = 0;
        try{
            FileInputStream fis = new FileInputStream(psw);
            XSSFWorkbook wb = new XSSFWorkbook(fis);
            XSSFSheet wessexSheet = wb.getSheet(Strings.PSW_SHEET_WESSEX);
            int valueColNo = getCellColumnByString(Strings.PSW_COLUMN_VALUE, wessexSheet);
            int countOfTraineesRowNo = getCellRowByString(Strings.PSW_COLUMN_TRAINEE_COUNT, wessexSheet);
            totalRef = (int) wessexSheet.getRow(countOfTraineesRowNo).getCell(valueColNo).getNumericCellValue();
            
        } catch (FileNotFoundException ex) {
            Logger.getLogger(PoiHelper.class.getName()).log(Level.SEVERE, null, ex);
        } catch (IOException ex) {
            Logger.getLogger(PoiHelper.class.getName()).log(Level.SEVERE, null, ex);
        }
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
            if (cell != null && cell.getCellType() != CellType.BLANK)
                return false;
        }
        return true;
    }
    
    private static int getCellRowByString(String str, XSSFSheet sheet){
        int rowNumber = 0;
        
        
        for (Row r : sheet) {
            for (Cell c : r) {
                String cellValueStr = "";
                try {
                    cellValueStr=c.getStringCellValue();
                } catch (IllegalStateException e) {
                }
                if(cellValueStr.equals(str)){
                    
                    rowNumber = c.getRowIndex();
                    
                }
            }
        }
        return rowNumber;
    }
    
    public void saveGraphs(){
        try {
            graphs.createNewFile();
        } catch (IOException ex) {
            Logger.getLogger(PoiHelper.class.getName()).log(Level.SEVERE, null, ex);
        }
        
    }
    
    /**
     * Checks if a given String is numeric
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
    
//    public static void genGraphOut(File inFile, File outFile){
//        try {
//            XSSFWorkbook wbOut = new XSSFWorkbook(new FileInputStream(outFile));
//            docOut.write(new FileOutputStream(outFile));
//        } catch (FileNotFoundException ex) {
//            Logger.getLogger(DocHelper.class.getName()).log(Level.SEVERE, null, ex);
//        } catch (IOException ex) {
//            Logger.getLogger(DocHelper.class.getName()).log(Level.SEVERE, null, ex);
//        }
//    }
}
