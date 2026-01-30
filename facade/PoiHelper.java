/*
 * Decompiled with CFR 0.152.
 */
package facade;

import facade.DocHelper;
import gui.MessageDialogs;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.text.SimpleDateFormat;
import java.time.Instant;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.GregorianCalendar;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;
import java.util.List;
import java.util.Locale;
import java.util.Objects;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFCreationHelper;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import vo.ReferralRecord;
import vo.ReferrerFormRecord;
import vo.Table9Line;
import vo.TraineeFormRecord;

public class PoiHelper {
    private File psw;
    private File graphs;
    private File traineeForm;
    private File referrerForm;
    private File mergedForms;
    private File log;
    private ArrayList<ReferralRecord> recordList;
    private XSSFWorkbook pswWorkbook;
    private XSSFWorkbook graphsWorkbook;
    private XSSFWorkbook referrerFormWorkbook;
    private XSSFWorkbook traineeFormWorkbook;
    private XSSFWorkbook mergedFormsWorkbook;

    public PoiHelper() {
    }

    public PoiHelper(File psw, File graphs, ArrayList<ReferralRecord> records) {
        this.psw = psw;
        this.graphs = graphs;
        this.recordList = records;
        this.loadWorkbooks();
    }

    public PoiHelper(File psw, File graphs) {
        this.psw = psw;
        this.graphs = graphs;
        Path path = Paths.get(graphs.getAbsolutePath(), new String[0]);
        this.log = new File(path.getParent().toString() + "/Log - NHS Report Generation Tool.txt");
        this.loadWorkbooks();
    }

    public PoiHelper(File refForm, File trainForm, File mergedOut) {
        this.referrerForm = refForm;
        this.traineeForm = trainForm;
        this.mergedForms = mergedOut;
        Path path = Paths.get(mergedOut.getAbsolutePath(), new String[0]);
        this.log = new File(path.getParent().toString() + "/Log - NHS Report Generation Tool.txt");
        this.loadForms();
    }

    public ArrayList<ReferralRecord> getRecordList() {
        return this.recordList;
    }

    public void setRecordList(ArrayList<ReferralRecord> recordList) {
        this.recordList = recordList;
    }

    private void loadWorkbooks() {
        try {
            this.pswWorkbook = new XSSFWorkbook(new FileInputStream(this.psw));
            this.graphsWorkbook = new XSSFWorkbook();
        }
        catch (IOException iOException) {
            // empty catch block
        }
    }

    public void loadForms() {
        try {
            this.traineeFormWorkbook = new XSSFWorkbook(new FileInputStream(this.traineeForm));
            this.referrerFormWorkbook = new XSSFWorkbook(new FileInputStream(this.referrerForm));
            this.mergedFormsWorkbook = new XSSFWorkbook();
        }
        catch (IOException iOException) {
            // empty catch block
        }
    }

    private void reLoadGraphWorkbook() {
        try {
            this.graphsWorkbook = new XSSFWorkbook(new FileInputStream(this.graphs));
        }
        catch (IOException ex) {
            Logger.getLogger(PoiHelper.class.getName()).log(Level.SEVERE, null, ex);
        }
    }

    private static int getCellColumnByString(String str, XSSFSheet sheet) {
        int columnNumber = 0;
        for (Row r : sheet) {
            for (Cell c : r) {
                String cellValueStr = "";
                try {
                    cellValueStr = c.getStringCellValue();
                }
                catch (IllegalStateException illegalStateException) {
                    // empty catch block
                }
                if (!cellValueStr.equals(str)) continue;
                columnNumber = c.getColumnIndex();
            }
        }
        return columnNumber;
    }

    public void getGraph0() {
        int fCount = 0;
        int mCount = 0;
        int uCount = 0;
        XSSFSheet refSheet = this.pswWorkbook.getSheetAt(this.pswWorkbook.getSheetIndex("Referrals"));
        XSSFSheet genderSheet = this.pswWorkbook.getSheetAt(this.pswWorkbook.getSheetIndex("Gender"));
        int columnNumber = PoiHelper.getCellColumnByString("Gender", refSheet);
        for (Row row : refSheet) {
            Cell cell = CellUtil.getCell(row, columnNumber);
            if (cell.getStringCellValue().equals("Female")) {
                ++fCount;
                continue;
            }
            if (cell.getStringCellValue().equals("Male")) {
                ++mCount;
                continue;
            }
            if (PoiHelper.isCellEmpty(cell) || row.getRowNum() <= PoiHelper.getCellRowByString("Gender", refSheet)) continue;
            ++uCount;
        }
        XSSFRow femaleRow = genderSheet.getRow(2);
        XSSFRow unknownRow = genderSheet.getRow(3);
        XSSFRow maleRow = genderSheet.getRow(4);
        double wssxFemaleCount = femaleRow.getCell(1).getNumericCellValue();
        double wssxMaleCount = maleRow.getCell(1).getNumericCellValue();
        double wssxUnknownCount = unknownRow.getCell(1).getNumericCellValue();
        XSSFSheet sheet = this.graphsWorkbook.getNumberOfSheets() <= 0 ? this.graphsWorkbook.createSheet("1. Gender disclosed") : this.graphsWorkbook.getSheetAt(0);
        XSSFRow titlesRow = sheet.createRow(0);
        XSSFRow fRow = sheet.createRow(1);
        XSSFRow mRow = sheet.createRow(2);
        XSSFRow uRow = sheet.createRow(3);
        Cell cell = titlesRow.createCell(1);
        cell.setCellValue("Referred count");
        cell = titlesRow.createCell(2);
        cell.setCellValue("Wessex count");
        Cell fCell = fRow.createCell(0);
        fCell.setCellValue("Female");
        fCell = fRow.createCell(1);
        fCell.setCellValue(fCount);
        fCell = fRow.createCell(2);
        fCell.setCellValue(wssxFemaleCount);
        Cell mCell = mRow.createCell(0);
        mCell.setCellValue("Male");
        mCell = mRow.createCell(1);
        mCell.setCellValue(mCount);
        mCell = mRow.createCell(2);
        mCell.setCellValue(wssxMaleCount);
        Cell uCell = uRow.createCell(0);
        uCell.setCellValue("Unknown");
        uCell = uRow.createCell(1);
        uCell.setCellValue(uCount);
        uCell = uRow.createCell(2);
        uCell.setCellValue(wssxUnknownCount);
    }

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
        boolean teamMCount = false;
        boolean teamFCount = false;
        int timeMCount = 0;
        int timeFCount = 0;
        int otherMCount = 0;
        int otherFCount = 0;
        boolean i = false;
        for (ReferralRecord record : this.recordList) {
            if (record.isAnxiety()) {
                if (record.getGender().equals("Female")) {
                    ++anxietyFCount;
                } else if (record.getGender().equals("Male")) {
                    ++anxietyMCount;
                }
            }
            if (record.isCapability()) {
                if (record.getGender().equals("Female")) {
                    ++capabilityFCount;
                } else if (record.getGender().equals("Male")) {
                    ++capabilityMCount;
                }
            }
            if (record.isCarreer()) {
                if (record.getGender().equals("Female")) {
                    ++carreerFCount;
                } else if (record.getGender().equals("Male")) {
                    ++carreerMCount;
                }
            }
            if (record.isClinSkills()) {
                if (record.getGender().equals("Female")) {
                    ++clinicalFCount;
                } else if (record.getGender().equals("Male")) {
                    ++clinicalMCount;
                }
            }
            if (record.isCommunication()) {
                if (record.getGender().equals("Female")) {
                    ++communicationFCount;
                } else if (record.getGender().equals("Male")) {
                    ++communicationMCount;
                }
            }
            if (record.isConduct()) {
                if (record.getGender().equals("Female")) {
                    ++conductFCount;
                } else if (record.getGender().equals("Male")) {
                    ++conductMCount;
                }
            }
            if (record.isCultural()) {
                if (record.getGender().equals("Female")) {
                    ++culturalFCount;
                } else if (record.getGender().equals("Male")) {
                    ++culturalMCount;
                }
            }
            if (record.isExam()) {
                if (record.getGender().equals("Female")) {
                    ++examFCount;
                } else if (record.getGender().equals("Male")) {
                    ++examMCount;
                }
            }
            if (record.isHealthMental()) {
                if (record.getGender().equals("Female")) {
                    ++menHealthFCount;
                } else if (record.getGender().equals("Male")) {
                    ++menHealthMCount;
                }
            }
            if (record.isHealthPhysical()) {
                if (record.getGender().equals("Female")) {
                    ++phHealthFCount;
                } else if (record.getGender().equals("Male")) {
                    ++phHhealthMCount;
                }
            }
            if (record.isLanguage()) {
                if (record.getGender().equals("Female")) {
                    ++languageFCount;
                } else if (record.getGender().equals("Male")) {
                    ++languageMCount;
                }
            }
            if (record.isProfessionalism()) {
                if (record.getGender().equals("Female")) {
                    ++profFCount;
                } else if (record.getGender().equals("Male")) {
                    ++profMCount;
                }
            }
            if (record.isAdhd()) {
                if (record.getGender().equals("Female")) {
                    ++adhdFCount;
                } else if (record.getGender().equals("Male")) {
                    ++adhdMCount;
                }
            }
            if (record.isAsd()) {
                if (record.getGender().equals("Female")) {
                    ++asdFCount;
                } else if (record.getGender().equals("Male")) {
                    ++asdMCount;
                }
            }
            if (record.isDyslexia()) {
                if (record.getGender().equals("Female")) {
                    ++dyslexiaFCount;
                } else if (record.getGender().equals("Male")) {
                    ++dyslexiaMCount;
                }
            }
            if (record.isDyspraxia()) {
                if (record.getGender().equals("Female")) {
                    ++dyspraxiaFCount;
                } else if (record.getGender().equals("Male")) {
                    ++dyspraxiaMCount;
                }
            }
            if (record.isSrtt()) {
                if (record.getGender().equals("Female")) {
                    ++srttFCount;
                } else if (record.getGender().equals("Male")) {
                    ++srttMCount;
                }
            }
            if (record.isTeam()) {
                if (record.getGender().equals("Female")) {
                    ++anxietyFCount;
                } else if (record.getGender().equals("Male")) {
                    ++anxietyMCount;
                }
            }
            if (record.isTime()) {
                if (record.getGender().equals("Female")) {
                    ++timeFCount;
                } else if (record.getGender().equals("Male")) {
                    ++timeMCount;
                }
            }
            if (!record.isOtherReason()) continue;
            if (record.getGender().equals("Female")) {
                ++otherFCount;
                continue;
            }
            if (!record.getGender().equals("Male")) continue;
            ++otherMCount;
        }
        XSSFSheet sheet = this.graphsWorkbook.getNumberOfSheets() <= 1 ? this.graphsWorkbook.createSheet("2. Referral Reason") : this.graphsWorkbook.getSheetAt(1);
        XSSFRow titlesRow = sheet.createRow(0);
        XSSFRow fRow = sheet.createRow(1);
        XSSFRow mRow = sheet.createRow(2);
        Cell cell = titlesRow.createCell(1);
        cell.setCellValue("Anxiety / Stress");
        cell = titlesRow.createCell(2);
        cell.setCellValue("Capability");
        cell = titlesRow.createCell(3);
        cell.setCellValue("Career support");
        cell = titlesRow.createCell(4);
        cell.setCellValue("Clinical skills");
        cell = titlesRow.createCell(5);
        cell.setCellValue("Communication / Interpersonal skills");
        cell = titlesRow.createCell(6);
        cell.setCellValue("Conduct");
        cell = titlesRow.createCell(7);
        cell.setCellValue("Cultural factors");
        cell = titlesRow.createCell(8);
        cell.setCellValue("Exam support");
        cell = titlesRow.createCell(9);
        cell.setCellValue("Health Conditions (Mental)");
        cell = titlesRow.createCell(10);
        cell.setCellValue("Health Conditions (Physical)");
        cell = titlesRow.createCell(11);
        cell.setCellValue("Language support");
        cell = titlesRow.createCell(12);
        cell.setCellValue("Professionalism");
        cell = titlesRow.createCell(13);
        cell.setCellValue("ADHD");
        cell = titlesRow.createCell(14);
        cell.setCellValue("ASD");
        cell = titlesRow.createCell(15);
        cell.setCellValue("Dyslexia");
        cell = titlesRow.createCell(16);
        cell.setCellValue("Dyspraxia");
        cell = titlesRow.createCell(17);
        cell.setCellValue("SRTT");
        cell = titlesRow.createCell(18);
        cell.setCellValue("Team working");
        cell = titlesRow.createCell(19);
        cell.setCellValue("Time / Workload Management");
        cell = titlesRow.createCell(20);
        cell.setCellValue("Other");
        Cell fCell = fRow.createCell(0);
        fCell.setCellValue("Female");
        Cell mCell = mRow.createCell(0);
        mCell.setCellValue("Male");
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
        fCell.setCellValue((double)teamFCount);
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
        mCell.setCellValue((double)teamMCount);
        mCell = mRow.createCell(19);
        mCell.setCellValue(timeMCount);
        mCell = mRow.createCell(20);
        mCell.setCellValue(otherMCount);
    }

    public void getGraph2() {
        int mentalFcount = 0;
        int physicalFcount = 0;
        int bothFcount = 0;
        int mentalMcount = 0;
        int physicalMcount = 0;
        int bothMcount = 0;
        XSSFSheet refSheet = this.pswWorkbook.getSheetAt(this.pswWorkbook.getSheetIndex("Referrals"));
        int mentalHealthColNo = PoiHelper.getCellColumnByString("Health Conditions (Mental)", refSheet);
        int physicalHealthColNo = PoiHelper.getCellColumnByString("Health Conditions (Physical)", refSheet);
        int genderColNo = PoiHelper.getCellColumnByString("Gender", refSheet);
        for (Row row : refSheet) {
            Cell mentalHealthCell = CellUtil.getCell(row, mentalHealthColNo);
            Cell physicalHealthCell = CellUtil.getCell(row, physicalHealthColNo);
            if (!row.getCell(mentalHealthColNo).getStringCellValue().equals("") && row.getCell(genderColNo).getStringCellValue().equals("Female") && row.getCell(physicalHealthColNo).getStringCellValue().equals("")) {
                ++mentalFcount;
            }
            if (!row.getCell(physicalHealthColNo).getStringCellValue().equals("") && row.getCell(genderColNo).getStringCellValue().equals("Female") && row.getCell(mentalHealthColNo).getStringCellValue().equals("")) {
                ++physicalFcount;
            }
            if (!row.getCell(mentalHealthColNo).getStringCellValue().equals("") && row.getCell(genderColNo).getStringCellValue().equals("Female") && !row.getCell(physicalHealthColNo).getStringCellValue().equals("")) {
                ++bothFcount;
            }
            if (!row.getCell(mentalHealthColNo).getStringCellValue().equals("") && row.getCell(genderColNo).getStringCellValue().equals("Male") && !row.getCell(physicalHealthColNo).getStringCellValue().equals("")) {
                ++bothMcount;
            }
            if (!row.getCell(physicalHealthColNo).getStringCellValue().equals("") && row.getCell(genderColNo).getStringCellValue().equals("Male") && row.getCell(mentalHealthColNo).getStringCellValue().equals("")) {
                ++physicalMcount;
            }
            if (row.getCell(mentalHealthColNo).getStringCellValue().equals("") || !row.getCell(genderColNo).getStringCellValue().equals("Male") || !row.getCell(physicalHealthColNo).getStringCellValue().equals("")) continue;
            ++mentalMcount;
        }
        XSSFSheet sheet = this.graphsWorkbook.getNumberOfSheets() <= 2 ? this.graphsWorkbook.createSheet("2a. Health referrals") : this.graphsWorkbook.getSheetAt(2);
        XSSFRow titlesRow = sheet.createRow(0);
        XSSFRow fRow = sheet.createRow(1);
        XSSFRow mRow = sheet.createRow(2);
        Cell cell = titlesRow.createCell(1);
        cell.setCellValue("Mental");
        cell = titlesRow.createCell(2);
        cell.setCellValue("Physical");
        cell = titlesRow.createCell(3);
        cell.setCellValue("Both");
        Cell fcell = fRow.createCell(0);
        fcell.setCellValue("Female");
        fcell = fRow.createCell(1);
        fcell.setCellValue(mentalFcount);
        fcell = fRow.createCell(2);
        fcell.setCellValue(physicalFcount);
        fcell = fRow.createCell(3);
        fcell.setCellValue(bothFcount);
        Cell mcell = mRow.createCell(0);
        mcell.setCellValue("Male");
        mcell = mRow.createCell(1);
        mcell.setCellValue(mentalMcount);
        mcell = mRow.createCell(2);
        mcell.setCellValue(physicalMcount);
        mcell = mRow.createCell(3);
        mcell.setCellValue(bothMcount);
    }

    public void getGraph3() {
        int f1Count = 0;
        int ct1Count = 0;
        int st3Count = 0;
        int st6Count = 0;
        XSSFSheet refSheet = this.pswWorkbook.getSheetAt(this.pswWorkbook.getSheetIndex("Referrals"));
        int gradeColNo = PoiHelper.getCellColumnByString("Grade", refSheet);
        for (Row row : refSheet) {
            Cell gradeCell = CellUtil.getCell(row, gradeColNo);
            HashSet<String> f1grades = new HashSet<String>();
            f1grades.add("F1");
            f1grades.add("F2");
            f1grades.add("FY1");
            f1grades.add("FY2");
            HashSet<String> ct1grades = new HashSet<String>();
            ct1grades.add("CT1");
            ct1grades.add("CT2");
            ct1grades.add("ST1");
            ct1grades.add("ST2");
            HashSet<String> st3grades = new HashSet<String>();
            st3grades.add("ST3");
            st3grades.add("ST4");
            st3grades.add("ST5");
            st3grades.add("CT3");
            HashSet<String> st6grades = new HashSet<String>();
            st6grades.add("ST6");
            st6grades.add("ST7");
            st6grades.add("ST8");
            if (f1grades.contains(gradeCell.getStringCellValue())) {
                ++f1Count;
            }
            if (ct1grades.contains(gradeCell.getStringCellValue())) {
                ++ct1Count;
            }
            if (st3grades.contains(gradeCell.getStringCellValue())) {
                ++st3Count;
            }
            if (!st6grades.contains(gradeCell.getStringCellValue())) continue;
            ++st6Count;
        }
        XSSFSheet sheet = this.graphsWorkbook.getNumberOfSheets() <= 3 ? this.graphsWorkbook.createSheet("3. Referrals broken down by Stage") : this.graphsWorkbook.getSheetAt(3);
        double totalCount = f1Count + ct1Count + st3Count + st6Count;
        double f1perc = Math.round((double)f1Count / totalCount * 100.0);
        double ct1perc = Math.round((double)ct1Count / totalCount * 100.0);
        double st3perc = Math.round((double)st3Count / totalCount * 100.0);
        double st6perc = Math.round((double)st6Count / totalCount * 100.0);
        XSSFRow titlesRow = sheet.createRow(0);
        XSSFRow f1Row = sheet.createRow(1);
        XSSFRow ct1Row = sheet.createRow(2);
        XSSFRow st3Row = sheet.createRow(3);
        XSSFRow st6Row = sheet.createRow(4);
        Cell cell = titlesRow.createCell(1);
        cell.setCellValue("Total");
        Cell f1cell = f1Row.createCell(0);
        f1cell.setCellValue("F1 & F2 " + f1perc + "%");
        f1cell = f1Row.createCell(1);
        f1cell.setCellValue(f1Count);
        Cell ct1Cell = ct1Row.createCell(0);
        ct1Cell.setCellValue("CT1, CT2, ST1 & ST2 " + ct1perc + "%");
        ct1Cell = ct1Row.createCell(1);
        ct1Cell.setCellValue(ct1Count);
        Cell st3Cell = st3Row.createCell(0);
        st3Cell.setCellValue("ST3, ST4, ST5 & CT3 " + st3perc + "%");
        st3Cell = st3Row.createCell(1);
        st3Cell.setCellValue(st3Count);
        Cell st6Cell = st6Row.createCell(0);
        st6Cell.setCellValue("ST3, ST4, ST5 & CT3 " + st6perc + "%");
        st6Cell = st6Row.createCell(1);
        st6Cell.setCellValue(st6Count);
    }

    public void getGraph4() {
        int stCount = 0;
        int ftCount = 0;
        int gpCount = 0;
        int otherCount = 0;
        XSSFSheet refSheet = this.pswWorkbook.getSheetAt(this.pswWorkbook.getSheetIndex("Referrals"));
        int gradeColNo = PoiHelper.getCellColumnByString("Grade", refSheet);
        for (Row row : refSheet) {
            Cell cell = CellUtil.getCell(row, gradeColNo);
            if (cell.getStringCellValue().equals("FY1") || cell.getStringCellValue().equals("FY2")) {
                ++ftCount;
            }
            if (cell.getStringCellValue().equals("ST1") || cell.getStringCellValue().equals("ST2") || cell.getStringCellValue().equals("ST3") || cell.getStringCellValue().equals("ST4") || cell.getStringCellValue().equals("ST5") || cell.getStringCellValue().equals("ST6") || cell.getStringCellValue().equals("ST7") || cell.getStringCellValue().equals("ST8")) {
                ++stCount;
            }
            if (cell.getStringCellValue().equals("GPST1") || cell.getStringCellValue().equals("GPST2") || cell.getStringCellValue().equals("GPST3")) {
                ++gpCount;
            }
            if (!cell.getStringCellValue().equals("DCT1") && !cell.getStringCellValue().equals("DCT2") && !cell.getStringCellValue().equals("DF1") && !cell.getStringCellValue().equals("DF2")) continue;
            ++otherCount;
        }
        XSSFSheet sheet = this.graphsWorkbook.getNumberOfSheets() <= 4 ? this.graphsWorkbook.createSheet("4. Referrals broken down by programme grade") : this.graphsWorkbook.getSheetAt(4);
        XSSFRow titlesRow = sheet.createRow(0);
        XSSFRow ftRow = sheet.createRow(1);
        XSSFRow stRow = sheet.createRow(2);
        XSSFRow gpRow = sheet.createRow(3);
        XSSFRow otherRow = sheet.createRow(4);
        Cell cell = titlesRow.createCell(1);
        cell.setCellValue("Total");
        Cell ftcell = ftRow.createCell(0);
        ftcell.setCellValue("FT");
        ftcell = ftRow.createCell(1);
        ftcell.setCellValue(ftCount);
        Cell stCell = stRow.createCell(0);
        stCell.setCellValue("ST (inc Core)");
        stCell = stRow.createCell(1);
        stCell.setCellValue(stCount);
        Cell gpCell = gpRow.createCell(0);
        gpCell.setCellValue("GP");
        gpCell = gpRow.createCell(1);
        gpCell.setCellValue(gpCount);
        Cell otherCell = otherRow.createCell(0);
        otherCell.setCellValue("Other (inc Dental)");
        otherCell = otherRow.createCell(1);
        otherCell.setCellValue(otherCount);
    }

    public void getGraph5() {
        XSSFSheet ocSheet = this.pswWorkbook.getSheetAt(this.pswWorkbook.getSheetIndex("Open Cases history"));
        XSSFRow ocR0 = ocSheet.getRow(0);
        XSSFRow ocR1 = ocSheet.getRow(1);
        XSSFRow ocR2 = ocSheet.getRow(2);
        XSSFSheet sheet = this.graphsWorkbook.getNumberOfSheets() <= 5 ? this.graphsWorkbook.createSheet("5. Duration of PSW input for current open cases") : this.graphsWorkbook.getSheetAt(5);
        XSSFRow titlesRow = sheet.createRow(0);
        XSSFRow fRow = sheet.createRow(1);
        XSSFRow sRow = sheet.createRow(2);
        titlesRow.createCell(1).setCellValue(ocR0.getCell(1).getStringCellValue());
        titlesRow.createCell(2).setCellValue(ocR0.getCell(2).getStringCellValue());
        titlesRow.createCell(3).setCellValue(ocR0.getCell(3).getStringCellValue());
        titlesRow.createCell(4).setCellValue(ocR0.getCell(4).getStringCellValue());
        titlesRow.createCell(5).setCellValue(ocR0.getCell(5).getStringCellValue());
        XSSFCellStyle cellStyle = this.graphsWorkbook.createCellStyle();
        XSSFCreationHelper createHelper = this.graphsWorkbook.getCreationHelper();
        cellStyle.setDataFormat(createHelper.createDataFormat().getFormat("MMM-YY"));
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
    }

    public void getGraph6() {
        int anxiety = 0;
        int carreer = 0;
        int clinSkills = 0;
        int communication = 0;
        int conduct = 0;
        int cultural = 0;
        int exam = 0;
        int healthMental = 0;
        int healthPhysical = 0;
        int language = 0;
        int professionalism = 0;
        int adhd = 0;
        int asd = 0;
        int dyslexia = 0;
        int dyspraxia = 0;
        int srtt = 0;
        int team = 0;
        int time = 0;
        int capability = 0;
        int otherRefReason = 0;
        for (ReferralRecord record : this.recordList) {
            Date refDate = record.getRefDate();
            Calendar cal = Calendar.getInstance();
            cal.setTime(refDate);
            Date actualDate = Date.from(Instant.now());
            Calendar actualCal = Calendar.getInstance();
            actualCal.setTime(actualDate);
            actualCal.add(1, -2);
            if (!record.isCaseOpen() || !cal.before(actualCal)) continue;
            if (record.isAnxiety()) {
                ++anxiety;
            }
            if (record.isCapability()) {
                ++capability;
            }
            if (record.isCarreer()) {
                ++carreer;
            }
            if (record.isClinSkills()) {
                ++clinSkills;
            }
            if (record.isCommunication()) {
                ++communication;
            }
            if (record.isConduct()) {
                ++conduct;
            }
            if (record.isCultural()) {
                ++cultural;
            }
            if (record.isExam()) {
                ++exam;
            }
            if (record.isHealthMental()) {
                ++healthMental;
            }
            if (record.isHealthPhysical()) {
                ++healthPhysical;
            }
            if (record.isLanguage()) {
                ++language;
            }
            if (record.isProfessionalism()) {
                ++professionalism;
            }
            if (record.isAdhd()) {
                ++adhd;
            }
            if (record.isAsd()) {
                ++asd;
            }
            if (record.isDyslexia()) {
                ++dyslexia;
            }
            if (record.isDyspraxia()) {
                ++dyspraxia;
            }
            if (record.isSrtt()) {
                ++srtt;
            }
            if (record.isTeam()) {
                ++team;
            }
            if (record.isTime()) {
                ++time;
            }
            if (!record.isOtherReason()) continue;
            ++otherRefReason;
        }
        XSSFSheet sheet = this.graphsWorkbook.getNumberOfSheets() <= 6 ? this.graphsWorkbook.createSheet("6. Reasons for referrals open longer than 24 months") : this.graphsWorkbook.getSheetAt(6);
        sheet.createRow(0).createCell(1).setCellValue("Total");
        sheet.createRow(1).createCell(0).setCellValue("Anxiety / Stress");
        sheet.getRow(1).createCell(1).setCellValue(anxiety);
        sheet.createRow(2).createCell(0).setCellValue("Capability");
        sheet.getRow(2).createCell(1).setCellValue(capability);
        sheet.createRow(3).createCell(0).setCellValue("Career support");
        sheet.getRow(3).createCell(1).setCellValue(carreer);
        sheet.createRow(4).createCell(0).setCellValue("Clinical skills");
        sheet.getRow(4).createCell(1).setCellValue(clinSkills);
        sheet.createRow(5).createCell(0).setCellValue("Communication / Interpersonal skills");
        sheet.getRow(5).createCell(1).setCellValue(communication);
        sheet.createRow(6).createCell(0).setCellValue("Conduct");
        sheet.getRow(6).createCell(1).setCellValue(conduct);
        sheet.createRow(7).createCell(0).setCellValue("Cultural factors");
        sheet.getRow(7).createCell(1).setCellValue(cultural);
        sheet.createRow(8).createCell(0).setCellValue("Exam support");
        sheet.getRow(8).createCell(1).setCellValue(exam);
        sheet.createRow(9).createCell(0).setCellValue("Health Conditions (Mental)");
        sheet.getRow(9).createCell(1).setCellValue(healthMental);
        sheet.createRow(10).createCell(0).setCellValue("Health Conditions (Physical)");
        sheet.getRow(10).createCell(1).setCellValue(healthPhysical);
        sheet.createRow(11).createCell(0).setCellValue("Language support");
        sheet.getRow(11).createCell(1).setCellValue(language);
        sheet.createRow(12).createCell(0).setCellValue("Professionalism");
        sheet.getRow(12).createCell(1).setCellValue(professionalism);
        sheet.createRow(13).createCell(0).setCellValue("ADHD");
        sheet.getRow(13).createCell(1).setCellValue(adhd);
        sheet.createRow(14).createCell(0).setCellValue("ASD");
        sheet.getRow(14).createCell(1).setCellValue(asd);
        sheet.createRow(15).createCell(0).setCellValue("Dyslexia");
        sheet.getRow(15).createCell(1).setCellValue(dyslexia);
        sheet.createRow(16).createCell(0).setCellValue("Dyspraxia");
        sheet.getRow(16).createCell(1).setCellValue(dyspraxia);
        sheet.createRow(17).createCell(0).setCellValue("SRTT");
        sheet.getRow(17).createCell(1).setCellValue(srtt);
        sheet.createRow(18).createCell(0).setCellValue("Team working");
        sheet.getRow(18).createCell(1).setCellValue(team);
        sheet.createRow(19).createCell(0).setCellValue("Time / Workload Management");
        sheet.getRow(19).createCell(1).setCellValue(time);
        sheet.createRow(20).createCell(0).setCellValue("Other");
        sheet.getRow(20).createCell(1).setCellValue(otherRefReason);
    }

    public void getGraph7() {
        int noConcerns = 0;
        int onGoing = 0;
        int completed = 0;
        int released = 0;
        int resigned = 0;
        int other = 0;
        int death = 0;
        XSSFSheet ccSheet = this.pswWorkbook.getSheetAt(this.pswWorkbook.getSheetIndex("Closed Cases"));
        int totalsColumn = PoiHelper.getCellColumnByString("Outcome Key", ccSheet);
        noConcerns = Integer.parseInt(ccSheet.getRow(1).getCell(totalsColumn).getRawValue());
        onGoing = Integer.parseInt(ccSheet.getRow(2).getCell(totalsColumn).getRawValue());
        completed = Integer.parseInt(ccSheet.getRow(3).getCell(totalsColumn).getRawValue());
        released = Integer.parseInt(ccSheet.getRow(4).getCell(totalsColumn).getRawValue());
        resigned = Integer.parseInt(ccSheet.getRow(5).getCell(totalsColumn).getRawValue());
        other = Integer.parseInt(ccSheet.getRow(6).getCell(totalsColumn).getRawValue());
        death = Integer.parseInt(ccSheet.getRow(7).getCell(totalsColumn).getRawValue());
        XSSFSheet sheet = this.graphsWorkbook.getNumberOfSheets() <= 7 ? this.graphsWorkbook.createSheet("7. Rolling Analysis of case closures") : this.graphsWorkbook.getSheetAt(7);
        XSSFRow title = sheet.createRow(0);
        title.createCell(1).setCellValue("Total");
        XSSFRow noConcernsRow = sheet.createRow(1);
        noConcernsRow.createCell(0).setCellValue("Return to training no concerns");
        noConcernsRow.createCell(1).setCellValue(noConcerns);
        XSSFRow onGoingRow = sheet.createRow(2);
        onGoingRow.createCell(0).setCellValue("Return to training on going concerns");
        onGoingRow.createCell(1).setCellValue(onGoing);
        XSSFRow completedRow = sheet.createRow(3);
        completedRow.createCell(0).setCellValue("Completed training");
        completedRow.createCell(1).setCellValue(completed);
        XSSFRow releasedRow = sheet.createRow(4);
        releasedRow.createCell(0).setCellValue("Released from training");
        releasedRow.createCell(1).setCellValue(released);
        XSSFRow resignedRow = sheet.createRow(5);
        resignedRow.createCell(0).setCellValue("Resigned from training/Post");
        resignedRow.createCell(1).setCellValue(resigned);
        XSSFRow otherRow = sheet.createRow(6);
        otherRow.createCell(0).setCellValue("Other -Non engagement/break from training");
        otherRow.createCell(1).setCellValue(other);
        XSSFRow deathRow = sheet.createRow(7);
        deathRow.createCell(0).setCellValue("Suicide/death / Other (including trust Dr.)");
        deathRow.createCell(1).setCellValue(death);
    }

    public void getGraph8() {
        XSSFSheet refSheet = this.pswWorkbook.getSheetAt(this.pswWorkbook.getSheetIndex("Referrals"));
        XSSFSheet wssxSheet = this.pswWorkbook.getSheetAt(this.pswWorkbook.getSheetIndex("Wessex"));
        XSSFSheet ccSheet = this.pswWorkbook.getSheetAt(this.pswWorkbook.getSheetIndex("Closed Cases"));
        int dateOpnRefColumn = PoiHelper.getCellColumnByString("Date opened", ccSheet);
        int dateClosedRefColumn = PoiHelper.getCellColumnByString("Date Closed", ccSheet);
        int titlesRowNum = PoiHelper.getCellRowByString("Date Closed", ccSheet);
        int yr = DocHelper.getStartingYear();
        int yrp1 = yr + 1;
        int yrs1 = yr - 1;
        int startYrCount = (int)wssxSheet.getRow(3).getCell(1).getNumericCellValue();
        int aprYrs1Count = 0;
        int mayYrs1Count = 0;
        int juneYrs1Count = 0;
        int julyYrs1Count = 0;
        int augYrs1Count = 0;
        int septYrs1Count = 0;
        int octYrs1Count = 0;
        int novYrs1Count = 0;
        int decYrs1Count = 0;
        int janYrCount = 0;
        int febYrCount = 0;
        int marYrCount = 0;
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
        for (Row r : ccSheet) {
            if (r.getRowNum() <= titlesRowNum || PoiHelper.isRowEmpty(r)) continue;
            DataFormatter formatter = new DataFormatter(Locale.UK);
            Cell dateOpenCell = CellUtil.getCell(r, dateOpnRefColumn);
            Cell dateClosedCell = CellUtil.getCell(r, dateClosedRefColumn);
            formatter.formatCellValue(dateOpenCell);
            Date dateOpen = dateOpenCell.getDateCellValue();
            Calendar dateOpenCal = Calendar.getInstance();
            dateOpenCal.setTime(dateOpen);
            formatter.formatCellValue(dateClosedCell);
            Date dateClosed = dateClosedCell.getDateCellValue();
            Calendar dateClosedCal = Calendar.getInstance();
            dateClosedCal.setTime(dateClosed);
            if (dateOpenCal.get(1) == yr) {
                if (dateClosedCal.get(2) == 2) {
                    ++marYrCount;
                }
                if (dateClosedCal.get(2) == 1) {
                    ++febYrCount;
                }
                if (dateClosedCal.get(2) != 0) continue;
                ++janYrCount;
                continue;
            }
            if (dateOpenCal.get(1) != yrs1) continue;
            if (dateClosedCal.get(2) == 11) {
                ++decYrs1Count;
            }
            if (dateClosedCal.get(2) == 10) {
                ++novYrs1Count;
            }
            if (dateClosedCal.get(2) == 9) {
                ++octYrs1Count;
            }
            if (dateClosedCal.get(2) == 8) {
                ++septYrs1Count;
            }
            if (dateClosedCal.get(2) == 7) {
                ++augYrs1Count;
            }
            if (dateClosedCal.get(2) == 6) {
                ++julyYrs1Count;
            }
            if (dateClosedCal.get(2) == 5) {
                ++juneYrs1Count;
            }
            if (dateClosedCal.get(2) == 4) {
                ++mayYrs1Count;
            }
            if (dateClosedCal.get(2) != 3) continue;
            ++aprYrs1Count;
        }
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
        Date janYp1 = new GregorianCalendar(yrp1, 0, 1).getTime();
        Calendar janYp1Cal = Calendar.getInstance();
        janYp1Cal.setTime(janYp1);
        Date febYp1 = new GregorianCalendar(yrp1, 1, 1).getTime();
        Calendar febYp1Cal = Calendar.getInstance();
        febYp1Cal.setTime(febYp1);
        Date marYp1 = new GregorianCalendar(yrp1, 2, 1).getTime();
        Calendar marYp1Cal = Calendar.getInstance();
        marYp1Cal.setTime(marYp1);
        for (ReferralRecord record : this.recordList) {
            Calendar calendar = Calendar.getInstance();
            calendar.setTime(record.getRefDate());
            if (calendar.get(2) == aprYrCal.get(2)) {
                ++aprYrCount;
            }
            if (calendar.get(2) == mayYrCal.get(2)) {
                ++mayYrCount;
            }
            if (calendar.get(2) == juneYrCal.get(2)) {
                ++juneYrCount;
            }
            if (calendar.get(2) == julyYrCal.get(2)) {
                ++julyYrCount;
            }
            if (calendar.get(2) == augYrCal.get(2)) {
                ++augYrCount;
            }
            if (calendar.get(2) == septYrCal.get(2)) {
                ++septYrCount;
            }
            if (calendar.get(2) == octYrCal.get(2)) {
                ++octYrCount;
            }
            if (calendar.get(2) == novYrCal.get(2)) {
                ++novYrcount;
            }
            if (calendar.get(2) == decYrCal.get(2)) {
                ++decYrCount;
            }
            if (calendar.get(2) == janYp1Cal.get(2)) {
                ++janYrp1Count;
            }
            if (calendar.get(2) == febYp1Cal.get(2)) {
                ++febYrp1Count;
            }
            if (calendar.get(2) != marYp1Cal.get(2)) continue;
            ++marYrp1Count;
        }
        int aprYrs1Final = startYrCount - aprYrs1Count;
        int mayYrs1Final = aprYrs1Final - mayYrs1Count;
        int juneYrs1Final = mayYrs1Final - juneYrs1Count;
        int julyYrs1Final = juneYrs1Final - julyYrs1Count;
        int augYrs1Final = julyYrs1Final - augYrs1Count;
        int septYrs1Final = augYrs1Final - septYrs1Count;
        int octYrs1Final = septYrs1Final - octYrs1Count;
        int novYrs1Final = octYrs1Final - novYrs1Count;
        int decYrs1Final = novYrs1Final - decYrs1Count;
        int janYrFinal = decYrs1Final - janYrCount;
        int febYrFinal = janYrFinal - febYrCount;
        int marYrFinal = febYrFinal - marYrCount;
        int aprYrFinal = marYrp1Count + aprYrCount;
        int mayYrFinal = aprYrFinal + mayYrCount;
        int juneYrFinal = mayYrFinal + juneYrCount;
        int julyYrFinal = juneYrFinal + julyYrCount;
        int augYrFinal = julyYrFinal + augYrCount;
        int septYrFinal = augYrFinal + septYrCount;
        int octYrFinal = septYrFinal + octYrCount;
        int novYrFinal = octYrFinal + novYrcount;
        int decYrFinal = novYrFinal + decYrCount;
        int janYrp1Final = decYrFinal + janYrp1Count;
        int febYrp1Final = janYrp1Final + febYrp1Count;
        int marYrp1Final = febYrp1Final + marYrp1Count;
        XSSFSheet sheet = this.graphsWorkbook.getNumberOfSheets() <= 8 ? this.graphsWorkbook.createSheet("8. PSW open case load overall by month ") : this.graphsWorkbook.getSheetAt(8);
        sheet.createRow(0).createCell(1).setCellValue(yrs1 + "/" + yr);
        sheet.getRow(0).createCell(2).setCellValue(yr + "/" + yrp1);
        sheet.createRow(1).createCell(0).setCellValue("April");
        sheet.getRow(1).createCell(1).setCellValue(aprYrs1Final);
        sheet.getRow(1).createCell(2).setCellValue(aprYrFinal);
        sheet.createRow(2).createCell(0).setCellValue("May");
        sheet.getRow(2).createCell(1).setCellValue(mayYrs1Final);
        sheet.getRow(2).createCell(2).setCellValue(mayYrFinal);
        sheet.createRow(3).createCell(0).setCellValue("June");
        sheet.getRow(3).createCell(1).setCellValue(juneYrs1Final);
        sheet.getRow(3).createCell(2).setCellValue(juneYrFinal);
        sheet.createRow(4).createCell(0).setCellValue("July");
        sheet.getRow(4).createCell(1).setCellValue(julyYrs1Final);
        sheet.getRow(4).createCell(2).setCellValue(julyYrFinal);
        sheet.createRow(5).createCell(0).setCellValue("August");
        sheet.getRow(5).createCell(1).setCellValue(augYrs1Final);
        sheet.getRow(5).createCell(2).setCellValue(augYrFinal);
        sheet.createRow(6).createCell(0).setCellValue("September");
        sheet.getRow(6).createCell(1).setCellValue(septYrs1Final);
        sheet.getRow(6).createCell(2).setCellValue(septYrFinal);
        sheet.createRow(7).createCell(0).setCellValue("October");
        sheet.getRow(7).createCell(1).setCellValue(octYrs1Final);
        sheet.getRow(7).createCell(2).setCellValue(octYrFinal);
        sheet.createRow(8).createCell(0).setCellValue("November");
        sheet.getRow(8).createCell(1).setCellValue(novYrs1Final);
        sheet.getRow(8).createCell(2).setCellValue(novYrFinal);
        sheet.createRow(9).createCell(0).setCellValue("December");
        sheet.getRow(9).createCell(1).setCellValue(decYrs1Final);
        sheet.getRow(9).createCell(2).setCellValue(decYrFinal);
        sheet.createRow(10).createCell(0).setCellValue("January");
        sheet.getRow(10).createCell(1).setCellValue(janYrFinal);
        sheet.getRow(10).createCell(2).setCellValue(janYrp1Final);
        sheet.createRow(11).createCell(0).setCellValue("February");
        sheet.getRow(11).createCell(1).setCellValue(febYrFinal);
        sheet.getRow(11).createCell(2).setCellValue(febYrp1Final);
        sheet.createRow(12).createCell(0).setCellValue("March");
        sheet.getRow(12).createCell(1).setCellValue(marYrFinal);
        sheet.getRow(12).createCell(2).setCellValue(marYrp1Final);
    }

    public void getGraph9() {
        XSSFSheet graphSheet = this.pswWorkbook.getSheetAt(this.pswWorkbook.getSheetIndex("Graphs"));
        int titleColumnNo = PoiHelper.getCellColumnByString("Total Trainee Time", graphSheet);
        int titleRowNo = PoiHelper.getCellRowByString("Total Trainee Time", graphSheet);
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
        double timeApr = graphSheet.getRow(aprRow).getCell(titleColumnNo).getNumericCellValue() * 24.0;
        int hoursApr = (int)timeApr;
        int minutesApr = (int)(timeApr - (double)hoursApr) * 60;
        double timeMay = graphSheet.getRow(mayRow).getCell(titleColumnNo).getNumericCellValue() * 24.0;
        int hoursMay = (int)timeMay;
        int minutesMay = (int)(timeMay - (double)hoursMay) * 60;
        double timeJun = graphSheet.getRow(junRow).getCell(titleColumnNo).getNumericCellValue() * 24.0;
        int hoursJun = (int)timeJun;
        int minutesJun = (int)(timeJun - (double)hoursJun) * 60;
        double timeJul = graphSheet.getRow(julRow).getCell(titleColumnNo).getNumericCellValue() * 24.0;
        int hoursJul = (int)timeJul;
        int minutesJul = (int)(timeJul - (double)hoursJul) * 60;
        double timeAug = graphSheet.getRow(augRow).getCell(titleColumnNo).getNumericCellValue() * 24.0;
        int hoursAug = (int)timeAug;
        int minutesAug = (int)(timeAug - (double)hoursAug) * 60;
        double timeSep = graphSheet.getRow(sepRow).getCell(titleColumnNo).getNumericCellValue() * 24.0;
        int hoursSep = (int)timeSep;
        int minutesSep = (int)(timeSep - (double)hoursSep) * 60;
        double timeOct = graphSheet.getRow(octRow).getCell(titleColumnNo).getNumericCellValue() * 24.0;
        int hoursOct = (int)timeOct;
        int minutesOct = (int)(timeOct - (double)hoursOct) * 60;
        double timeNov = graphSheet.getRow(novRow).getCell(titleColumnNo).getNumericCellValue() * 24.0;
        int hoursNov = (int)timeNov;
        int minutesNov = (int)(timeNov - (double)hoursNov) * 60;
        double timeDec = graphSheet.getRow(decRow).getCell(titleColumnNo).getNumericCellValue() * 24.0;
        int hoursDec = (int)timeDec;
        int minutesDec = (int)(timeDec - (double)hoursDec) * 60;
        double timeJan = graphSheet.getRow(janRow).getCell(titleColumnNo).getNumericCellValue() * 24.0;
        int hoursJan = (int)timeJan;
        int minutesJan = (int)(timeJan - (double)hoursApr) * 60;
        double timeFeb = graphSheet.getRow(febRow).getCell(titleColumnNo).getNumericCellValue() * 24.0;
        int hoursFeb = (int)timeFeb;
        int minutesFeb = (int)(timeFeb - (double)hoursFeb) * 60;
        double timeMar = graphSheet.getRow(marRow).getCell(titleColumnNo).getNumericCellValue() * 24.0;
        int hoursMar = (int)timeMar;
        int minutesMar = (int)(timeMar - (double)hoursMar) * 60;
        XSSFSheet sheet = this.graphsWorkbook.getNumberOfSheets() <= 9 ? this.graphsWorkbook.createSheet("9. Total Hours for PSW CMs and SSG Experts ") : this.graphsWorkbook.getSheetAt(9);
        XSSFRow title = sheet.createRow(0);
        title.createCell(1).setCellValue("Total");
        XSSFRow aprGraphRow = sheet.createRow(1);
        aprGraphRow.createCell(0).setCellValue("April");
        aprGraphRow.createCell(1).setCellValue(hoursApr);
        XSSFRow mayGraphRow = sheet.createRow(2);
        mayGraphRow.createCell(0).setCellValue("May");
        mayGraphRow.createCell(1).setCellValue(hoursMay);
        XSSFRow junGraphRow = sheet.createRow(3);
        junGraphRow.createCell(0).setCellValue("June");
        junGraphRow.createCell(1).setCellValue(hoursJun);
        XSSFRow julGraphRow = sheet.createRow(4);
        julGraphRow.createCell(0).setCellValue("July");
        julGraphRow.createCell(1).setCellValue(hoursJul);
        XSSFRow augGraphRow = sheet.createRow(5);
        augGraphRow.createCell(0).setCellValue("August");
        augGraphRow.createCell(1).setCellValue(hoursAug);
        XSSFRow sepGraphRow = sheet.createRow(6);
        sepGraphRow.createCell(0).setCellValue("September");
        sepGraphRow.createCell(1).setCellValue(hoursSep);
        XSSFRow octGraphRow = sheet.createRow(7);
        octGraphRow.createCell(0).setCellValue("October");
        octGraphRow.createCell(1).setCellValue(hoursOct);
        XSSFRow novGraphRow = sheet.createRow(8);
        novGraphRow.createCell(0).setCellValue("November");
        novGraphRow.createCell(1).setCellValue(hoursNov);
        XSSFRow decGraphRow = sheet.createRow(9);
        decGraphRow.createCell(0).setCellValue("December");
        decGraphRow.createCell(1).setCellValue(hoursDec);
        XSSFRow janGraphRow = sheet.createRow(10);
        janGraphRow.createCell(0).setCellValue("January");
        janGraphRow.createCell(1).setCellValue(hoursJan);
        XSSFRow febGraphRow = sheet.createRow(11);
        febGraphRow.createCell(0).setCellValue("February");
        febGraphRow.createCell(1).setCellValue(hoursFeb);
        XSSFRow marGraphRow = sheet.createRow(12);
        marGraphRow.createCell(0).setCellValue("March");
        marGraphRow.createCell(1).setCellValue(hoursMar);
    }

    public void getGraph10() {
        XSSFSheet graphSheet = this.pswWorkbook.getSheetAt(this.pswWorkbook.getSheetIndex("Graphs"));
        int titleColumnNo = PoiHelper.getCellColumnByString("Total Trainee costs", graphSheet);
        int titleRowNo = PoiHelper.getCellRowByString("Total Trainee costs", graphSheet);
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
        XSSFSheet sheet = this.graphsWorkbook.getNumberOfSheets() <= 10 ? this.graphsWorkbook.createSheet("10. Monthly invoice trends for PSU CM and SSG Experts") : this.graphsWorkbook.getSheetAt(10);
        XSSFRow title = sheet.createRow(0);
        title.createCell(1).setCellValue("Total");
        XSSFRow aprGraphRow = sheet.createRow(1);
        aprGraphRow.createCell(0).setCellValue("April");
        aprGraphRow.createCell(1).setCellValue(aprCount);
        XSSFRow mayGraphRow = sheet.createRow(2);
        mayGraphRow.createCell(0).setCellValue("May");
        mayGraphRow.createCell(1).setCellValue(mayCount);
        XSSFRow junGraphRow = sheet.createRow(3);
        junGraphRow.createCell(0).setCellValue("June");
        junGraphRow.createCell(1).setCellValue(junCount);
        XSSFRow julGraphRow = sheet.createRow(4);
        julGraphRow.createCell(0).setCellValue("July");
        julGraphRow.createCell(1).setCellValue(julCount);
        XSSFRow augGraphRow = sheet.createRow(5);
        augGraphRow.createCell(0).setCellValue("August");
        augGraphRow.createCell(1).setCellValue(augCount);
        XSSFRow sepGraphRow = sheet.createRow(6);
        sepGraphRow.createCell(0).setCellValue("September");
        sepGraphRow.createCell(1).setCellValue(sepCount);
        XSSFRow octGraphRow = sheet.createRow(7);
        octGraphRow.createCell(0).setCellValue("October");
        octGraphRow.createCell(1).setCellValue(octCount);
        XSSFRow novGraphRow = sheet.createRow(8);
        novGraphRow.createCell(0).setCellValue("November");
        novGraphRow.createCell(1).setCellValue(novCount);
        XSSFRow decGraphRow = sheet.createRow(9);
        decGraphRow.createCell(0).setCellValue("December");
        decGraphRow.createCell(1).setCellValue(decCount);
        XSSFRow janGraphRow = sheet.createRow(10);
        janGraphRow.createCell(0).setCellValue("January");
        janGraphRow.createCell(1).setCellValue(janCount);
        XSSFRow febGraphRow = sheet.createRow(11);
        febGraphRow.createCell(0).setCellValue("February");
        febGraphRow.createCell(1).setCellValue(febCount);
        XSSFRow marGraphRow = sheet.createRow(12);
        marGraphRow.createCell(0).setCellValue("March");
        marGraphRow.createCell(1).setCellValue(marCount);
    }

    public void getGraph11() {
        XSSFSheet graphSheet = this.pswWorkbook.getSheetAt(this.pswWorkbook.getSheetIndex("Graphs"));
        int ssgColumnNo = PoiHelper.getCellColumnByString("SSG Costs", graphSheet);
        int cmColumnNo = PoiHelper.getCellColumnByString("CM Costs", graphSheet);
        int titleRowNo = PoiHelper.getCellRowByString("SSG Costs", graphSheet);
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
        double totalSsgCount = aprSsgCount + maySsgCount + junSsgCount + julSsgCount + augSsgCount + sepSsgCount + octSsgCount + novSsgCount + decSsgCount + janSsgCount + febSsgCount + marSsgCount;
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
        double totalCmCount = aprCmCount + mayCmCount + junCmCount + julCmCount + augCmCount + sepCmCount + octCmCount + novCmCount + decCmCount + janCmCount + febCmCount + marCmCount;
        XSSFSheet sheet = this.graphsWorkbook.getNumberOfSheets() <= 11 ? this.graphsWorkbook.createSheet("11. Monthly breakdown of CM and SSG invoices") : this.graphsWorkbook.getSheetAt(11);
        XSSFRow title = sheet.createRow(0);
        title.createCell(1).setCellValue("SSG Costs");
        title.createCell(2).setCellValue("CM Costs");
        XSSFRow aprGraphRow = sheet.createRow(1);
        aprGraphRow.createCell(0).setCellValue("April");
        aprGraphRow.createCell(1).setCellValue(aprSsgCount);
        aprGraphRow.createCell(2).setCellValue(aprCmCount);
        XSSFRow mayGraphRow = sheet.createRow(2);
        mayGraphRow.createCell(0).setCellValue("May");
        mayGraphRow.createCell(1).setCellValue(maySsgCount);
        mayGraphRow.createCell(2).setCellValue(mayCmCount);
        XSSFRow junGraphRow = sheet.createRow(3);
        junGraphRow.createCell(0).setCellValue("June");
        junGraphRow.createCell(1).setCellValue(junSsgCount);
        junGraphRow.createCell(2).setCellValue(junCmCount);
        XSSFRow julGraphRow = sheet.createRow(4);
        julGraphRow.createCell(0).setCellValue("July");
        julGraphRow.createCell(1).setCellValue(julSsgCount);
        julGraphRow.createCell(2).setCellValue(julCmCount);
        XSSFRow augGraphRow = sheet.createRow(5);
        augGraphRow.createCell(0).setCellValue("August");
        augGraphRow.createCell(1).setCellValue(augSsgCount);
        augGraphRow.createCell(2).setCellValue(augCmCount);
        XSSFRow sepGraphRow = sheet.createRow(6);
        sepGraphRow.createCell(0).setCellValue("September");
        sepGraphRow.createCell(1).setCellValue(sepSsgCount);
        sepGraphRow.createCell(2).setCellValue(sepCmCount);
        XSSFRow octGraphRow = sheet.createRow(7);
        octGraphRow.createCell(0).setCellValue("October");
        octGraphRow.createCell(1).setCellValue(octSsgCount);
        octGraphRow.createCell(2).setCellValue(octCmCount);
        XSSFRow novGraphRow = sheet.createRow(8);
        novGraphRow.createCell(0).setCellValue("November");
        novGraphRow.createCell(1).setCellValue(novSsgCount);
        novGraphRow.createCell(2).setCellValue(novCmCount);
        XSSFRow decGraphRow = sheet.createRow(9);
        decGraphRow.createCell(0).setCellValue("December");
        decGraphRow.createCell(1).setCellValue(decSsgCount);
        decGraphRow.createCell(2).setCellValue(decCmCount);
        XSSFRow janGraphRow = sheet.createRow(10);
        janGraphRow.createCell(0).setCellValue("January");
        janGraphRow.createCell(1).setCellValue(janSsgCount);
        janGraphRow.createCell(2).setCellValue(janCmCount);
        XSSFRow febGraphRow = sheet.createRow(11);
        febGraphRow.createCell(0).setCellValue("February");
        febGraphRow.createCell(1).setCellValue(febSsgCount);
        febGraphRow.createCell(2).setCellValue(febCmCount);
        XSSFRow marGraphRow = sheet.createRow(12);
        marGraphRow.createCell(0).setCellValue("March");
        marGraphRow.createCell(1).setCellValue(marSsgCount);
        marGraphRow.createCell(2).setCellValue(marCmCount);
        XSSFRow totalGraphRow = sheet.createRow(13);
        totalGraphRow.createCell(0).setCellValue("Total");
        totalGraphRow.createCell(1).setCellValue(totalSsgCount);
        totalGraphRow.createCell(2).setCellValue(totalCmCount);
    }

    public List<Integer> countTable0() {
        ArrayList<Integer> list = new ArrayList<Integer>();
        List<Integer> t1Integers = this.countT1Integers();
        int stCount = 0;
        int fCount = 0;
        int gpCount = 0;
        int otherCount = 0;
        int totalCount = 0;
        int casesClosed = t1Integers.get(0);
        int casesOClosed = t1Integers.get(1);
        XSSFSheet refSheet = this.pswWorkbook.getSheetAt(this.pswWorkbook.getSheetIndex("Referrals"));
        int gradeColNo = PoiHelper.getCellColumnByString("Grade", refSheet);
        for (Row row : refSheet) {
            if (PoiHelper.isRowEmpty(row) || PoiHelper.isCellEmpty(row.getCell(gradeColNo))) continue;
            switch (row.getCell(gradeColNo).getStringCellValue()) {
                case "FY1": 
                case "FY2": {
                    ++fCount;
                    break;
                }
                case "ST1": 
                case "ST2": 
                case "ST3": 
                case "ST4": 
                case "ST5": 
                case "ST6": 
                case "ST7": 
                case "ST8": 
                case "CT1": 
                case "CT2": 
                case "CT3": {
                    ++stCount;
                    break;
                }
                case "GPST1": 
                case "GPST2": 
                case "GPST3": {
                    ++gpCount;
                    break;
                }
                case "DF1": 
                case "DF2": 
                case "DCT1": 
                case "DCT2": 
                case "Pharmacy": {
                    ++otherCount;
                    break;
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

    public List<Integer> countTable1() {
        ArrayList<Integer> list = new ArrayList<Integer>();
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
        this.reLoadGraphWorkbook();
        XSSFSheet graphSheet = this.graphsWorkbook.getSheetAt(this.graphsWorkbook.getSheetIndex("2. Referral Reason"));
        int anxietyColNo = PoiHelper.getCellColumnByString("Anxiety / Stress", graphSheet);
        int capColNo = PoiHelper.getCellColumnByString("Capability", graphSheet);
        int carreerColNo = PoiHelper.getCellColumnByString("Career support", graphSheet);
        int clinSkillsColNo = PoiHelper.getCellColumnByString("Clinical skills", graphSheet);
        int communicationColNo = PoiHelper.getCellColumnByString("Communication / Interpersonal skills", graphSheet);
        int conductColNo = PoiHelper.getCellColumnByString("Conduct", graphSheet);
        int culturalColNo = PoiHelper.getCellColumnByString("Cultural factors", graphSheet);
        int examColNo = PoiHelper.getCellColumnByString("Exam support", graphSheet);
        int mentalHealthColNo = PoiHelper.getCellColumnByString("Health Conditions (Mental)", graphSheet);
        int physicalHealthColNo = PoiHelper.getCellColumnByString("Health Conditions (Physical)", graphSheet);
        int languageColNo = PoiHelper.getCellColumnByString("Language support", graphSheet);
        int professionalismColNo = PoiHelper.getCellColumnByString("Professionalism", graphSheet);
        int adhdColNo = PoiHelper.getCellColumnByString("ADHD", graphSheet);
        int asdColNo = PoiHelper.getCellColumnByString("ASD", graphSheet);
        int dyslexiaColNo = PoiHelper.getCellColumnByString("Dyslexia", graphSheet);
        int dyspraxiaColNo = PoiHelper.getCellColumnByString("Dyspraxia", graphSheet);
        int srttColNo = PoiHelper.getCellColumnByString("SRTT", graphSheet);
        int teamColNo = PoiHelper.getCellColumnByString("Team working", graphSheet);
        int timeColNo = PoiHelper.getCellColumnByString("Time / Workload Management", graphSheet);
        int otherColNo = PoiHelper.getCellColumnByString("Other", graphSheet);
        int femRowNo = PoiHelper.getCellRowByString("Female", graphSheet);
        int maleRowNo = PoiHelper.getCellRowByString("Male", graphSheet);
        anxietyFCount = (int)graphSheet.getRow(femRowNo).getCell(anxietyColNo).getNumericCellValue();
        capabilityFCount = (int)graphSheet.getRow(femRowNo).getCell(capColNo).getNumericCellValue();
        carreerFCount = (int)graphSheet.getRow(femRowNo).getCell(carreerColNo).getNumericCellValue();
        clinicalFCount = (int)graphSheet.getRow(femRowNo).getCell(clinSkillsColNo).getNumericCellValue();
        communicationFCount = (int)graphSheet.getRow(femRowNo).getCell(communicationColNo).getNumericCellValue();
        conductFCount = (int)graphSheet.getRow(femRowNo).getCell(conductColNo).getNumericCellValue();
        culturalFCount = (int)graphSheet.getRow(femRowNo).getCell(culturalColNo).getNumericCellValue();
        examFCount = (int)graphSheet.getRow(femRowNo).getCell(examColNo).getNumericCellValue();
        menHealthFCount = (int)graphSheet.getRow(femRowNo).getCell(mentalHealthColNo).getNumericCellValue();
        phHealthFCount = (int)graphSheet.getRow(femRowNo).getCell(physicalHealthColNo).getNumericCellValue();
        languageFCount = (int)graphSheet.getRow(femRowNo).getCell(languageColNo).getNumericCellValue();
        profFCount = (int)graphSheet.getRow(femRowNo).getCell(professionalismColNo).getNumericCellValue();
        adhdFCount = (int)graphSheet.getRow(femRowNo).getCell(adhdColNo).getNumericCellValue();
        asdFCount = (int)graphSheet.getRow(femRowNo).getCell(asdColNo).getNumericCellValue();
        dyslexiaFCount = (int)graphSheet.getRow(femRowNo).getCell(dyslexiaColNo).getNumericCellValue();
        dyspraxiaFCount = (int)graphSheet.getRow(femRowNo).getCell(dyspraxiaColNo).getNumericCellValue();
        srttFCount = (int)graphSheet.getRow(femRowNo).getCell(srttColNo).getNumericCellValue();
        teamFCount = (int)graphSheet.getRow(femRowNo).getCell(teamColNo).getNumericCellValue();
        timeFCount = (int)graphSheet.getRow(femRowNo).getCell(timeColNo).getNumericCellValue();
        otherFCount = (int)graphSheet.getRow(femRowNo).getCell(otherColNo).getNumericCellValue();
        anxietyMCount = (int)graphSheet.getRow(maleRowNo).getCell(anxietyColNo).getNumericCellValue();
        capabilityMCount = (int)graphSheet.getRow(maleRowNo).getCell(capColNo).getNumericCellValue();
        carreerMCount = (int)graphSheet.getRow(maleRowNo).getCell(carreerColNo).getNumericCellValue();
        clinicalMCount = (int)graphSheet.getRow(maleRowNo).getCell(clinSkillsColNo).getNumericCellValue();
        communicationMCount = (int)graphSheet.getRow(maleRowNo).getCell(communicationColNo).getNumericCellValue();
        conductMCount = (int)graphSheet.getRow(maleRowNo).getCell(conductColNo).getNumericCellValue();
        culturalMCount = (int)graphSheet.getRow(maleRowNo).getCell(culturalColNo).getNumericCellValue();
        examMCount = (int)graphSheet.getRow(maleRowNo).getCell(examColNo).getNumericCellValue();
        menHealthMCount = (int)graphSheet.getRow(maleRowNo).getCell(mentalHealthColNo).getNumericCellValue();
        phHealthMCount = (int)graphSheet.getRow(maleRowNo).getCell(physicalHealthColNo).getNumericCellValue();
        languageMCount = (int)graphSheet.getRow(maleRowNo).getCell(languageColNo).getNumericCellValue();
        profMCount = (int)graphSheet.getRow(maleRowNo).getCell(professionalismColNo).getNumericCellValue();
        adhdMCount = (int)graphSheet.getRow(maleRowNo).getCell(adhdColNo).getNumericCellValue();
        asdMCount = (int)graphSheet.getRow(maleRowNo).getCell(asdColNo).getNumericCellValue();
        dyslexiaMCount = (int)graphSheet.getRow(maleRowNo).getCell(dyslexiaColNo).getNumericCellValue();
        dyspraxiaMCount = (int)graphSheet.getRow(maleRowNo).getCell(dyspraxiaColNo).getNumericCellValue();
        srttMCount = (int)graphSheet.getRow(maleRowNo).getCell(srttColNo).getNumericCellValue();
        teamMCount = (int)graphSheet.getRow(maleRowNo).getCell(teamColNo).getNumericCellValue();
        timeMCount = (int)graphSheet.getRow(maleRowNo).getCell(timeColNo).getNumericCellValue();
        otherMCount = (int)graphSheet.getRow(maleRowNo).getCell(otherColNo).getNumericCellValue();
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
        ArrayList<Integer> list = new ArrayList<Integer>();
        int referredCount = this.countTotalReferrals();
        int f1TotalCount = 0;
        int f2TotalCount = 0;
        int f1ReferredCount = 0;
        int f2ReferredCount = 0;
        try {
            Row row;
            FileInputStream pswFileIn = new FileInputStream(this.psw);
            XSSFWorkbook pswWorkbook = new XSSFWorkbook(pswFileIn);
            XSSFSheet refSheet = pswWorkbook.getSheetAt(pswWorkbook.getSheetIndex("Referrals"));
            int gradeColNo = PoiHelper.getCellColumnByString("Grade", refSheet);
            Iterator<Row> rowIteratorRefSheet = refSheet.iterator();
            XSSFSheet foundationSheet = pswWorkbook.getSheetAt(pswWorkbook.getSheetIndex("Foundation"));
            int countColNo = PoiHelper.getCellColumnByString("Count of Trainees", foundationSheet);
            int f1CountRowNo = PoiHelper.getCellRowByString("Foundation Year 1", foundationSheet);
            int f2CountRowNo = PoiHelper.getCellRowByString("Foundation Year 2", foundationSheet);
            while (rowIteratorRefSheet.hasNext() && ((row = rowIteratorRefSheet.next()).getRowNum() <= 3 || row.getCell(gradeColNo) != null)) {
                if (row.getRowNum() > 3 && !row.getCell(gradeColNo).getStringCellValue().equals("") && row.getCell(gradeColNo).getStringCellValue().equals("FY1")) {
                    ++f1ReferredCount;
                }
                if (row.getRowNum() <= 3 || !row.getCell(gradeColNo).getStringCellValue().equals("FY2")) continue;
                ++f2ReferredCount;
            }
            f1TotalCount = (int)foundationSheet.getRow(f1CountRowNo).getCell(countColNo).getNumericCellValue();
            f2TotalCount = (int)foundationSheet.getRow(f2CountRowNo).getCell(countColNo).getNumericCellValue();
            list.add(referredCount);
            list.add(f1TotalCount);
            list.add(f2TotalCount);
            list.add(f1ReferredCount);
            list.add(f2ReferredCount);
        }
        catch (IOException ex) {
            Logger.getLogger(PoiHelper.class.getName()).log(Level.SEVERE, null, ex);
        }
        return list;
    }

    public List<Integer> countTable3() {
        Row row;
        ArrayList<Integer> list = new ArrayList<Integer>();
        int bournemouthRefNo = 0;
        int dorsetCountyRefNo = 0;
        int dorsetHealthRefNo = 0;
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
        int dorsetCountyTotal = 0;
        int dorsetHealthTotal = 0;
        int hhftTotal = 0;
        int iowTotal = 0;
        int jerseyTotal = 0;
        int pooleTotal = 0;
        int portsmouthTotal = 0;
        int salisburyTotal = 0;
        int solentTotal = 0;
        int southamptonTotal = 0;
        int southernTotal = 0;
        int totalWessex = 0;
        XSSFSheet refSheet = this.pswWorkbook.getSheetAt(this.pswWorkbook.getSheetIndex("Referrals"));
        int trustRColNo = PoiHelper.getCellColumnByString("Trust", refSheet);
        Iterator<Row> rowIteratorRefSheet = refSheet.iterator();
        XSSFSheet trustSheet = this.pswWorkbook.getSheetAt(this.pswWorkbook.getSheetIndex("Trust"));
        int traineeCountColNo = PoiHelper.getCellColumnByString("Count of Trainees", trustSheet);
        int trustColNo = PoiHelper.getCellColumnByString("Trust", trustSheet);
        Iterator<Row> rowIteratorTrustSheet = trustSheet.iterator();
        totalWessex = this.countTotalWessex();
        while (rowIteratorTrustSheet.hasNext()) {
            row = rowIteratorTrustSheet.next();
            switch (row.getCell(trustColNo).getStringCellValue()) {
                case "The Royal Bournemouth and Christchurch Hospitals NHS Foundation Trust": {
                    bournemouthTotal = (int)row.getCell(traineeCountColNo).getNumericCellValue();
                    break;
                }
                case "Dorset County Hospital NHS Foundation Trust": {
                    dorsetCountyTotal = (int)row.getCell(traineeCountColNo).getNumericCellValue();
                    break;
                }
                case "Dorset Healthcare University NHS Foundation Trust": {
                    dorsetHealthTotal = (int)row.getCell(traineeCountColNo).getNumericCellValue();
                    break;
                }
                case "Hampshire Hospitals NHS Foundation Trust": {
                    hhftTotal = (int)row.getCell(traineeCountColNo).getNumericCellValue();
                    break;
                }
                case "Isle of Wight NHS Trust": {
                    iowTotal = (int)row.getCell(traineeCountColNo).getNumericCellValue();
                    break;
                }
                case "States of Jersey": {
                    jerseyTotal = (int)row.getCell(traineeCountColNo).getNumericCellValue();
                    break;
                }
                case "Poole Hospital NHS Foundation Trust": {
                    pooleTotal = (int)row.getCell(traineeCountColNo).getNumericCellValue();
                    break;
                }
                case "Portsmouth Hospitals NHS Trust": {
                    portsmouthTotal = (int)row.getCell(traineeCountColNo).getNumericCellValue();
                    break;
                }
                case "Salisbury NHS Foundation Trust": {
                    salisburyTotal = (int)row.getCell(traineeCountColNo).getNumericCellValue();
                    break;
                }
                case "Solent NHS Trust": {
                    solentTotal = (int)row.getCell(traineeCountColNo).getNumericCellValue();
                    break;
                }
                case "University Hospital Southampton NHS Foundation Trust": {
                    southamptonTotal = (int)row.getCell(traineeCountColNo).getNumericCellValue();
                    break;
                }
                case "Southern Health NHS Foundation Trust": {
                    southernTotal = (int)row.getCell(traineeCountColNo).getNumericCellValue();
                }
            }
        }
        while (rowIteratorRefSheet.hasNext()) {
            row = rowIteratorRefSheet.next();
            if (PoiHelper.isRowEmpty(row) || PoiHelper.isCellEmpty(row.getCell(trustRColNo))) continue;
            if (row.getCell(trustRColNo).getStringCellValue().contains("The Royal Bournemouth and Christchurch Hospitals NHS Foundation Trust")) {
                ++bournemouthRefNo;
                continue;
            }
            if (row.getCell(trustRColNo).getStringCellValue().contains("Dorset County Hospital NHS Foundation Trust")) {
                ++dorsetCountyRefNo;
                continue;
            }
            if (row.getCell(trustRColNo).getStringCellValue().contains("Dorset Healthcare University NHS Foundation Trust")) {
                ++dorsetHealthRefNo;
                continue;
            }
            if (row.getCell(trustRColNo).getStringCellValue().contains("Hampshire Hospitals NHS Foundation Trust")) {
                ++hhftRefNo;
                continue;
            }
            if (row.getCell(trustRColNo).getStringCellValue().contains("Isle of Wight NHS Trust")) {
                ++iowRefNo;
                continue;
            }
            if (row.getCell(trustRColNo).getStringCellValue().contains("Jersey General Hospital, States of Jersey")) {
                ++jerseyRefNo;
                continue;
            }
            if (row.getCell(trustRColNo).getStringCellValue().contains("Poole Hospital NHS Foundation Trust")) {
                ++pooleRefNo;
                continue;
            }
            if (row.getCell(trustRColNo).getStringCellValue().contains("Portsmouth University Hospitals NHS Trust")) {
                ++portsmouthRefNo;
                continue;
            }
            if (row.getCell(trustRColNo).getStringCellValue().contains("Salisbury NHS Foundation Trust")) {
                ++salisburyRefNo;
                continue;
            }
            if (row.getCell(trustRColNo).getStringCellValue().contains("Solent NHS Trust")) {
                ++solentRefNo;
                continue;
            }
            if (row.getCell(trustRColNo).getStringCellValue().contains("University Hospital Southampton NHS Foundation Trust")) {
                ++southamptonRefNo;
                continue;
            }
            if (!row.getCell(trustRColNo).getStringCellValue().contains("Southern Health NHS Foundation Trust")) continue;
            ++southernRefNo;
        }
        list.add(bournemouthTotal);
        list.add(dorsetCountyTotal);
        list.add(dorsetHealthTotal);
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
        list.add(dorsetCountyRefNo);
        list.add(dorsetHealthRefNo);
        list.add(hhftRefNo);
        list.add(iowRefNo);
        list.add(jerseyRefNo);
        list.add(pooleRefNo);
        list.add(portsmouthRefNo);
        list.add(salisburyRefNo);
        list.add(solentRefNo);
        list.add(southamptonRefNo);
        list.add(southernRefNo);
        list.add(totalWessex);
        return list;
    }

    public List<Integer> countTable4() {
        Row row;
        ArrayList<Integer> list = new ArrayList<Integer>();
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
        double anaestheticsTotal = 0.0;
        double dentalTotal = 0.0;
        double emergTotal = 0.0;
        double foundationTotal = 0.0;
        double gpTotal = 0.0;
        double medicineTotal = 0.0;
        double obsTotal = 0.0;
        double occhealthTotal = 0.0;
        double paediatricsTotal = 0.0;
        double pathologyTotal = 0.0;
        double pharmacyTotal = 0.0;
        double psychTotal = 0.0;
        double pubhealthTotal = 0.0;
        double radioTotal = 0.0;
        double surgeryTotal = 0.0;
        int totalRefs = this.countTotalReferrals();
        int totalWssx = this.countTotalWessex();
        XSSFSheet refSheet = this.pswWorkbook.getSheetAt(this.pswWorkbook.getSheetIndex("Referrals"));
        Iterator<Row> rowIteratorRefSheet = refSheet.rowIterator();
        int specialtyRColNo = PoiHelper.getCellColumnByString("Specialty", refSheet);
        XSSFSheet specialtySheet = this.pswWorkbook.getSheetAt(this.pswWorkbook.getSheetIndex("Specialty"));
        Iterator<Row> rowIteratorSpcSheet = specialtySheet.rowIterator();
        int specialtyColNo = PoiHelper.getCellColumnByString("Specialty", specialtySheet);
        int traineeCountColNo = PoiHelper.getCellColumnByString("Count of Trainees", specialtySheet);
        while (rowIteratorRefSheet.hasNext()) {
            row = rowIteratorRefSheet.next();
            if (PoiHelper.isRowEmpty(row) || PoiHelper.isCellEmpty(row.getCell(specialtyRColNo))) continue;
            Cell c = row.getCell(specialtyRColNo);
            if (c.getStringCellValue().contains("Anaesthetics")) {
                ++anaestheticsRefNo;
                continue;
            }
            if (c.getStringCellValue().contains("Dental")) {
                ++dentalRefNo;
                continue;
            }
            if (c.getStringCellValue().contains("Emergency medicine")) {
                ++emergRefNo;
                continue;
            }
            if (c.getStringCellValue().contains("Foundation/Wessex")) {
                ++foundationRefNo;
                continue;
            }
            if (c.getStringCellValue().contains("GP")) {
                ++gpRefNo;
                continue;
            }
            if (c.getStringCellValue().contains("Medicine -")) {
                ++medicineRefNo;
                continue;
            }
            if (c.getStringCellValue().contains("Obstetrics")) {
                ++obsRefNo;
                continue;
            }
            if (c.getStringCellValue().contains("Occupational Health")) {
                ++occhealthRefNo;
                continue;
            }
            if (c.getStringCellValue().contains("Paediatrics")) {
                ++paediatricsRefNo;
                continue;
            }
            if (c.getStringCellValue().contains("Pathology")) {
                ++pathologyRefNo;
                continue;
            }
            if (c.getStringCellValue().contains("Pharmacy")) {
                ++pharmacyRefNo;
                continue;
            }
            if (c.getStringCellValue().contains("Psychiatry")) {
                ++psychRefNo;
                continue;
            }
            if (c.getStringCellValue().contains("Public Health")) {
                ++pubhealthRefNo;
                continue;
            }
            if (c.getStringCellValue().contains("Radiology")) {
                ++radioRefNo;
                continue;
            }
            if (!c.getStringCellValue().contains("Surgery -")) continue;
            ++surgeryRefNo;
        }
        while (rowIteratorSpcSheet.hasNext()) {
            row = rowIteratorSpcSheet.next();
            if (PoiHelper.isRowEmpty(row) || PoiHelper.isCellEmpty(row.getCell(specialtyColNo))) continue;
            switch (row.getCell(specialtyColNo).getStringCellValue()) {
                case "Anaesthetics/Wessex": 
                case "Core Anaesthetics Training/Wessex": {
                    anaestheticsTotal += row.getCell(traineeCountColNo).getNumericCellValue();
                    break;
                }
                case "Dental Core Training/Oxford/Wessex": 
                case "Dental Foundation/Oxford/Wessex": {
                    dentalTotal += row.getCell(traineeCountColNo).getNumericCellValue();
                    break;
                }
                case "Emergency medicine  (run through)/Wessex": 
                case "Emergency medicine/Wessex": {
                    emergTotal += row.getCell(traineeCountColNo).getNumericCellValue();
                    break;
                }
                case "Foundation/Wessex": {
                    foundationTotal = row.getCell(traineeCountColNo).getNumericCellValue();
                    break;
                }
                case "General Practice Basingstoke/Wessex": 
                case "General Practice Bournemouth/Wessex": 
                case "General Practice Dorchester/Wessex": 
                case "General Practice Isle of Wight/Wessex": 
                case "General Practice Poole/Wessex": 
                case "General Practice Portsmouth/Wessex": 
                case "General Practice Salisbury/Wessex": 
                case "General Practice Southampton/Wessex": 
                case "General Practice Winchester/Wessex": {
                    gpTotal += row.getCell(traineeCountColNo).getNumericCellValue();
                    break;
                }
                case "Acute Care Common Stem - Acute Medicine/Wessex": 
                case "Acute Internal Medicine/Wessex": 
                case "Genito-urinary medicine/Wessex": 
                case "Geriatric medicine/Wessex": 
                case "Intensive Care Medicine Single/Wessex": 
                case "Intensive care medicine/Wessex": 
                case "Internal Medicine Training Stage 1/Wessex": 
                case "Rehabilitation medicine/Wessex": 
                case "Renal medicine/Wessex": 
                case "Respiratory medicine/Wessex": {
                    medicineTotal += row.getCell(traineeCountColNo).getNumericCellValue();
                    break;
                }
                case "Obstetrics and gynaecology/Wessex": {
                    obsTotal = row.getCell(traineeCountColNo).getNumericCellValue();
                    break;
                }
                case "Occupational medicine/Wessex": {
                    occhealthTotal = row.getCell(traineeCountColNo).getNumericCellValue();
                    break;
                }
                case "Paediatric cardiology/Wessex": 
                case "Paediatrics/Wessex": {
                    paediatricsTotal += row.getCell(traineeCountColNo).getNumericCellValue();
                    break;
                }
                case "Chemical pathology/Wessex": 
                case "Histopathology/Wessex": {
                    pathologyTotal += row.getCell(traineeCountColNo).getNumericCellValue();
                    break;
                }
                case "Old age psychiatry/Wessex": 
                case "Psychiatry of learning disability/Wessex": 
                case "Child and adolescent psychiatry (run through)/Wessex": 
                case "Child and adolescent psychiatry/Wessex": 
                case "Core Psychiatry Training/Wessex": 
                case "Forensic psychiatry/Wessex": 
                case "General psychiatry/Wessex": {
                    psychTotal = psychTotal + pathologyTotal + row.getCell(traineeCountColNo).getNumericCellValue();
                    break;
                }
                case "Public health medicine/Wessex": {
                    pubhealthTotal = row.getCell(traineeCountColNo).getNumericCellValue();
                    break;
                }
                case "Clinical radiology/Wessex": {
                    radioTotal = row.getCell(traineeCountColNo).getNumericCellValue();
                    break;
                }
                case "Cardio-thoracic surgery (run through)/Wessex": 
                case "Core Surgical Training/Wessex": 
                case "General Surgery (run through)/Wessex": 
                case "General surgery/Wessex": 
                case "Neurosurgery/Wessex": 
                case "Oral and maxillo-facial surgery (run through)/Wessex": 
                case "Oral and maxillo-facial surgery/Wessex": 
                case "Oral Surgery/Oxford/Wessex": 
                case "Trauma and orthopaedic surgery/Wessex": {
                    surgeryTotal += row.getCell(traineeCountColNo).getNumericCellValue();
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
        return list;
    }

    public ArrayList<ReferralRecord> getTable5LineByTrust(String trst) {
        ArrayList<ReferralRecord> list = new ArrayList<ReferralRecord>();
        int i = 0;
        for (ReferralRecord record : this.recordList) {
            if (!record.getTrust().contains(trst) || !record.isExam()) continue;
            ++i;
            list.add(record);
        }
        return list;
    }

    public List<Double> countTable7() {
        this.reLoadGraphWorkbook();
        ArrayList<Double> list = new ArrayList<Double>();
        XSSFSheet costs = this.graphsWorkbook.getSheetAt(11);
        Double ssg = costs.getRow(13).getCell(1).getNumericCellValue();
        Double cm = costs.getRow(13).getCell(2).getNumericCellValue();
        Double total = ssg + cm;
        list.add(ssg);
        list.add(cm);
        list.add(total);
        return list;
    }

    public ArrayList<Table9Line> countTable8() {
        ArrayList<Table9Line> list = new ArrayList<Table9Line>();
        Table9Line anaesthetics = new Table9Line();
        Table9Line dental = new Table9Line();
        Table9Line dermatology = new Table9Line();
        Table9Line endocrinology = new Table9Line();
        Table9Line foundation = new Table9Line();
        Table9Line gastroenterology = new Table9Line();
        Table9Line gp = new Table9Line();
        Table9Line haematology = new Table9Line();
        Table9Line histopathology = new Table9Line();
        Table9Line emergMed = new Table9Line();
        Table9Line medicine = new Table9Line();
        Table9Line neurology = new Table9Line();
        Table9Line obs = new Table9Line();
        Table9Line occHealth = new Table9Line();
        Table9Line oncology = new Table9Line();
        Table9Line ophtalmology = new Table9Line();
        Table9Line paediatrics = new Table9Line();
        Table9Line pathology = new Table9Line();
        Table9Line pharmacy = new Table9Line();
        Table9Line psych = new Table9Line();
        Table9Line pubHealth = new Table9Line();
        Table9Line radiology = new Table9Line();
        Table9Line sexHealth = new Table9Line();
        Table9Line rheumathology = new Table9Line();
        Table9Line surgery = new Table9Line();
        for (ReferralRecord record : this.recordList) {
            if (record.getSpecialty().equals("Anaesthetics")) {
                anaesthetics = this.getLineBySpc(anaesthetics, record);
                continue;
            }
            if (record.getSpecialty().contains("Dental")) {
                dental = this.getLineBySpc(dental, record);
                continue;
            }
            if (record.getSpecialty().contains("Dermatology")) {
                dermatology = this.getLineBySpc(dermatology, record);
                continue;
            }
            if (record.getSpecialty().contains("Endocrinology & Diabetes")) {
                endocrinology = this.getLineBySpc(endocrinology, record);
                continue;
            }
            if (record.getSpecialty().contains("Foundation/Wessex")) {
                foundation = this.getLineBySpc(foundation, record);
                continue;
            }
            if (record.getSpecialty().contains("Gastroenterology")) {
                gastroenterology = this.getLineBySpc(gastroenterology, record);
                continue;
            }
            if (record.getSpecialty().contains("GP")) {
                gp = this.getLineBySpc(gp, record);
                continue;
            }
            if (record.getSpecialty().contains("Haematology")) {
                haematology = this.getLineBySpc(haematology, record);
                continue;
            }
            if (record.getSpecialty().contains("Histopathology")) {
                histopathology = this.getLineBySpc(histopathology, record);
                continue;
            }
            if (record.getSpecialty().contains("Emergency medicine")) {
                emergMed = this.getLineBySpc(emergMed, record);
                continue;
            }
            if (record.getSpecialty().contains("Medicine -")) {
                medicine = this.getLineBySpc(medicine, record);
                continue;
            }
            if (record.getSpecialty().contains("Neurology")) {
                neurology = this.getLineBySpc(neurology, record);
                continue;
            }
            if (record.getSpecialty().contains("Obstetrics")) {
                obs = this.getLineBySpc(obs, record);
                continue;
            }
            if (record.getSpecialty().contains("Occupational Health")) {
                occHealth = this.getLineBySpc(occHealth, record);
                continue;
            }
            if (record.getSpecialty().contains("Oncology")) {
                oncology = this.getLineBySpc(oncology, record);
                continue;
            }
            if (record.getSpecialty().contains("Ophtalmology")) {
                ophtalmology = this.getLineBySpc(ophtalmology, record);
                continue;
            }
            if (record.getSpecialty().contains("Paediatrics")) {
                paediatrics = this.getLineBySpc(paediatrics, record);
                continue;
            }
            if (record.getSpecialty().contains("Pathology")) {
                pathology = this.getLineBySpc(pathology, record);
                continue;
            }
            if (record.getSpecialty().contains("Pharmacy")) {
                pharmacy = this.getLineBySpc(pharmacy, record);
                continue;
            }
            if (record.getSpecialty().contains("Psychiatry")) {
                psych = this.getLineBySpc(psych, record);
                continue;
            }
            if (record.getSpecialty().contains("Public Health")) {
                pubHealth = this.getLineBySpc(pubHealth, record);
                continue;
            }
            if (record.getSpecialty().contains("Radiology")) {
                radiology = this.getLineBySpc(radiology, record);
                continue;
            }
            if (record.getSpecialty().contains("Sexual Health")) {
                sexHealth = this.getLineBySpc(sexHealth, record);
                continue;
            }
            if (record.getSpecialty().contains("Rheumatology")) {
                rheumathology = this.getLineBySpc(rheumathology, record);
                continue;
            }
            if (!record.getSpecialty().contains("Surgery -")) continue;
            surgery = this.getLineBySpc(surgery, record);
        }
        anaesthetics.setTitle("Anaesthetics");
        dental.setTitle("Dental");
        dermatology.setTitle("Dermatology");
        endocrinology.setTitle("Endocrinology & Diabetes");
        foundation.setTitle("Foundation/Wessex");
        gastroenterology.setTitle("Gastroenterology");
        gp.setTitle("GP");
        haematology.setTitle("Haematology");
        histopathology.setTitle("Histopathology");
        emergMed.setTitle("Emergency medicine");
        medicine.setTitle("Medicine");
        neurology.setTitle("Neurology");
        obs.setTitle("Obstetrics");
        occHealth.setTitle("Occupational Health");
        oncology.setTitle("Oncology");
        ophtalmology.setTitle("Ophtalmology");
        paediatrics.setTitle("Paediatrics");
        pathology.setTitle("Pathology");
        psych.setTitle("Psychiatry");
        pharmacy.setTitle("Pharmacy");
        pubHealth.setTitle("Public Health");
        radiology.setTitle("Radiology");
        sexHealth.setTitle("Sexual Health");
        rheumathology.setTitle("Rheumatology");
        surgery.setTitle("Surgery -");
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

    public ArrayList<Table9Line> countTable9() {
        ArrayList<Table9Line> list = new ArrayList<Table9Line>();
        Table9Line anaesthetics = new Table9Line();
        Table9Line dental = new Table9Line();
        Table9Line dermatology = new Table9Line();
        Table9Line endocrinology = new Table9Line();
        Table9Line foundation = new Table9Line();
        Table9Line gastroenterology = new Table9Line();
        Table9Line gp = new Table9Line();
        Table9Line haematology = new Table9Line();
        Table9Line histopathology = new Table9Line();
        Table9Line emergMed = new Table9Line();
        Table9Line medicine = new Table9Line();
        Table9Line neurology = new Table9Line();
        Table9Line obs = new Table9Line();
        Table9Line occHealth = new Table9Line();
        Table9Line oncology = new Table9Line();
        Table9Line ophtalmology = new Table9Line();
        Table9Line paediatrics = new Table9Line();
        Table9Line pathology = new Table9Line();
        Table9Line pharmacy = new Table9Line();
        Table9Line psych = new Table9Line();
        Table9Line pubHealth = new Table9Line();
        Table9Line radiology = new Table9Line();
        Table9Line sexHealth = new Table9Line();
        Table9Line rheumathology = new Table9Line();
        Table9Line surgery = new Table9Line();
        for (ReferralRecord record : this.recordList) {
            if (record.getSpecialty().equals("Anaesthetics")) {
                anaesthetics = this.getNonUkLineBySpc(anaesthetics, record);
                continue;
            }
            if (record.getSpecialty().contains("Dental")) {
                dental = this.getNonUkLineBySpc(dental, record);
                continue;
            }
            if (record.getSpecialty().contains("Dermatology")) {
                dermatology = this.getNonUkLineBySpc(dermatology, record);
                continue;
            }
            if (record.getSpecialty().contains("Endocrinology & Diabetes")) {
                endocrinology = this.getNonUkLineBySpc(endocrinology, record);
                continue;
            }
            if (record.getSpecialty().contains("Foundation/Wessex")) {
                foundation = this.getNonUkLineBySpc(foundation, record);
                continue;
            }
            if (record.getSpecialty().contains("Gastroenterology")) {
                gastroenterology = this.getNonUkLineBySpc(gastroenterology, record);
                continue;
            }
            if (record.getSpecialty().contains("GP")) {
                gp = this.getNonUkLineBySpc(gp, record);
                continue;
            }
            if (record.getSpecialty().contains("Haematology")) {
                haematology = this.getNonUkLineBySpc(haematology, record);
                continue;
            }
            if (record.getSpecialty().contains("Histopathology")) {
                histopathology = this.getNonUkLineBySpc(histopathology, record);
                continue;
            }
            if (record.getSpecialty().contains("Emergency medicine")) {
                emergMed = this.getNonUkLineBySpc(emergMed, record);
                continue;
            }
            if (record.getSpecialty().contains("Medicine -")) {
                medicine = this.getNonUkLineBySpc(medicine, record);
                continue;
            }
            if (record.getSpecialty().contains("Neurology")) {
                neurology = this.getNonUkLineBySpc(neurology, record);
                continue;
            }
            if (record.getSpecialty().contains("Obstetrics")) {
                obs = this.getNonUkLineBySpc(obs, record);
                continue;
            }
            if (record.getSpecialty().contains("Occupational Health")) {
                occHealth = this.getNonUkLineBySpc(occHealth, record);
                continue;
            }
            if (record.getSpecialty().contains("Oncology")) {
                oncology = this.getNonUkLineBySpc(oncology, record);
                continue;
            }
            if (record.getSpecialty().contains("Ophtalmology")) {
                ophtalmology = this.getNonUkLineBySpc(ophtalmology, record);
                continue;
            }
            if (record.getSpecialty().contains("Paediatrics")) {
                paediatrics = this.getNonUkLineBySpc(paediatrics, record);
                continue;
            }
            if (record.getSpecialty().contains("Pathology")) {
                pathology = this.getNonUkLineBySpc(pathology, record);
                continue;
            }
            if (record.getSpecialty().contains("Pharmacy")) {
                pharmacy = this.getNonUkLineBySpc(pharmacy, record);
                continue;
            }
            if (record.getSpecialty().contains("Psychiatry")) {
                psych = this.getNonUkLineBySpc(psych, record);
                continue;
            }
            if (record.getSpecialty().contains("Public Health")) {
                pubHealth = this.getNonUkLineBySpc(pubHealth, record);
                continue;
            }
            if (record.getSpecialty().contains("Radiology")) {
                radiology = this.getNonUkLineBySpc(radiology, record);
                continue;
            }
            if (record.getSpecialty().contains("Sexual Health")) {
                sexHealth = this.getNonUkLineBySpc(sexHealth, record);
                continue;
            }
            if (record.getSpecialty().contains("Rheumatology")) {
                rheumathology = this.getNonUkLineBySpc(rheumathology, record);
                continue;
            }
            if (!record.getSpecialty().contains("Surgery -")) continue;
            surgery = this.getNonUkLineBySpc(surgery, record);
        }
        anaesthetics.setTitle("Anaesthetics");
        dental.setTitle("Dental");
        dermatology.setTitle("Dermatology");
        endocrinology.setTitle("Endocrinology & Diabetes");
        foundation.setTitle("Foundation/Wessex");
        gastroenterology.setTitle("Gastroenterology");
        gp.setTitle("GP");
        haematology.setTitle("Haematology");
        histopathology.setTitle("Histopathology");
        emergMed.setTitle("Emergency medicine");
        medicine.setTitle("Medicine");
        neurology.setTitle("Neurology");
        obs.setTitle("Obstetrics");
        occHealth.setTitle("Occupational Health");
        oncology.setTitle("Oncology");
        ophtalmology.setTitle("Ophtalmology");
        paediatrics.setTitle("Paediatrics");
        pathology.setTitle("Pathology");
        psych.setTitle("Psychiatry");
        pharmacy.setTitle("Pharmacy");
        pubHealth.setTitle("Public Health");
        radiology.setTitle("Radiology");
        sexHealth.setTitle("Sexual Health");
        rheumathology.setTitle("Rheumatology");
        surgery.setTitle("Surgery -");
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

    private Table9Line getLineBySpc(Table9Line line, ReferralRecord record) {
        int male = line.getMale();
        int female = line.getFemale();
        int otherSex = line.getOtherSex();
        int uk = line.getUk();
        int nonUk = line.getNonUk();
        int age2329 = line.getAge2329();
        int age3035 = line.getAge3035();
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
        if (record.getGender().equals("Female")) {
            ++female;
        } else if (record.getGender().equals("Male")) {
            ++male;
        } else if (!record.getGender().equals("Male") || !record.getGender().equals("Female")) {
            ++otherSex;
        }
        if (record.getCountry().equals("UK")) {
            ++uk;
        } else if (!record.getCountry().equals("UK")) {
            ++nonUk;
        }
        if (record.getAge() >= 23 && record.getAge() <= 29) {
            ++age2329;
        } else if (record.getAge() >= 30 && record.getAge() <= 35) {
            ++age3035;
        } else if (record.getAge() >= 36 && record.getAge() <= 40) {
            ++age3540;
        } else if (record.getAge() > 40) {
            ++age40;
        }
        if (record.getEthnicity().equals("White British")) {
            ++whiteb;
        } else if (record.getEthnicity().equals("White Other")) {
            ++whiteo;
        } else if (record.getEthnicity().equals("Asian")) {
            ++asian;
        } else if (record.getEthnicity().equals("African")) {
            ++african;
        } else if (!record.getEthnicity().equals("")) {
            ++ethOther;
        }
        if (record.getReligion().equals("Christianity")) {
            ++christian;
        } else if (record.getReligion().equals("Islam")) {
            ++islam;
        } else if (record.getReligion().equals("Hinduism")) {
            ++hindu;
        } else if (record.getReligion().equals("Atheism")) {
            ++atheist;
        } else if (record.getReligion().equals("Sikhism")) {
            ++sikh;
        } else if (record.getReligion().equals("Judaism")) {
            ++judaism;
        } else if (record.getReligion().equals("Buddhism")) {
            ++buddhism;
        } else if (!record.getReligion().equals("")) {
            ++relOther;
        } else if (record.getReligion().equals("") || record.getReligion().equals("Prefer not to say")) {
            ++relPNS;
        }
        if (record.getDisability().equals("Yes")) {
            ++yes;
        } else if (record.getDisability().equals("No")) {
            ++no;
        }
        if (record.getSexOr().equals("Het")) {
            ++het;
        } else if (record.getSexOr().equals("Homosexual")) {
            ++homosexual;
        } else if (record.getSexOr().equals("Bi")) {
            ++bisexual;
        } else if (record.getSexOr().equals("") || record.getSexOr().equals("Prefer not to say")) {
            ++sexOrPNS;
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

    private Table9Line getNonUkLineBySpc(Table9Line line, ReferralRecord record) {
        int male = line.getMale();
        int female = line.getFemale();
        int otherSex = line.getOtherSex();
        int uk = line.getUk();
        int nonUk = line.getNonUk();
        int age2329 = line.getAge2329();
        int age3035 = line.getAge3035();
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
        if (!record.getCountry().equals("UK")) {
            if (record.getGender().equals("Female")) {
                ++female;
            } else if (record.getGender().equals("Male")) {
                ++male;
            } else if (!record.getGender().equals("Male") || !record.getGender().equals("Female")) {
                ++otherSex;
            }
            if (record.getCountry().equals("UK")) {
                ++uk;
            } else if (!record.getCountry().equals("UK")) {
                ++nonUk;
            }
            if (record.getAge() >= 23 && record.getAge() <= 29) {
                ++age2329;
            } else if (record.getAge() >= 30 && record.getAge() <= 35) {
                ++age3035;
            } else if (record.getAge() >= 36 && record.getAge() <= 40) {
                ++age3540;
            } else if (record.getAge() > 40) {
                ++age40;
            }
            if (record.getEthnicity().equals("White British")) {
                ++whiteb;
            } else if (record.getEthnicity().equals("White Other")) {
                ++whiteo;
            } else if (record.getEthnicity().equals("Asian")) {
                ++asian;
            } else if (record.getEthnicity().equals("African")) {
                ++african;
            } else if (!record.getEthnicity().equals("")) {
                ++ethOther;
            }
            if (record.getReligion().equals("Christianity")) {
                ++christian;
            } else if (record.getReligion().equals("Islam")) {
                ++islam;
            } else if (record.getReligion().equals("Hinduism")) {
                ++hindu;
            } else if (record.getReligion().equals("Atheism")) {
                ++atheist;
            } else if (record.getReligion().equals("Sikhism")) {
                ++sikh;
            } else if (record.getReligion().equals("Judaism")) {
                ++judaism;
            } else if (record.getReligion().equals("Buddhism")) {
                ++buddhism;
            } else if (!record.getReligion().equals("")) {
                ++relOther;
            } else if (record.getReligion().equals("") || record.getReligion().equals("Prefer not to say")) {
                ++relPNS;
            }
            if (record.getDisability().equals("Yes")) {
                ++yes;
            } else if (record.getDisability().equals("No")) {
                ++no;
            }
            if (record.getSexOr().equals("Het")) {
                ++het;
            } else if (record.getSexOr().equals("Homosexual")) {
                ++homosexual;
            } else if (record.getSexOr().equals("Bi")) {
                ++bisexual;
            } else if (record.getSexOr().equals("") || record.getSexOr().equals("Prefer not to say")) {
                ++sexOrPNS;
            }
        }
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
        ArrayList<Integer> list = new ArrayList<Integer>();
        int closedWithinPeriod = 0;
        int openedAndClosed = 0;
        XSSFSheet ccSheet = this.pswWorkbook.getSheet("Closed Cases");
        int dateOpenedColNo = PoiHelper.getCellColumnByString("Date opened", ccSheet);
        int dateClosedColNo = PoiHelper.getCellColumnByString("Date Closed", ccSheet);
        int titlesRow = PoiHelper.getCellRowByString("Date opened", ccSheet);
        for (Row r : ccSheet) {
            if (r.getRowNum() <= titlesRow || PoiHelper.isRowEmpty(r)) continue;
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
            if (calClosed.get(2) > 2 && calClosed.get(1) == DocHelper.getStartingYear() || calClosed.get(2) <= 2 && calClosed.get(1) == DocHelper.getStartingYear() + 1) {
                ++closedWithinPeriod;
            }
            if ((calClosed.get(2) <= 2 || calClosed.get(1) != DocHelper.getStartingYear()) && (calClosed.get(2) > 2 || calClosed.get(1) != DocHelper.getStartingYear() + 1) || (calOpened.get(2) <= 2 || calOpened.get(1) != DocHelper.getStartingYear()) && (calOpened.get(2) >= 2 || calOpened.get(1) != DocHelper.getStartingYear() + 1)) continue;
            ++openedAndClosed;
        }
        list.add(closedWithinPeriod);
        list.add(openedAndClosed);
        return list;
    }

    public Integer countTotalReferrals() {
        int totalRef = 0;
        XSSFSheet refSheet = this.pswWorkbook.getSheet("Referrals");
        Iterator<Row> rowIterator = refSheet.rowIterator();
        int firstRow = PoiHelper.getCellRowByString("ADHD", refSheet);
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            if (PoiHelper.isRowEmpty(row) || row.getRowNum() <= firstRow) continue;
            ++totalRef;
        }
        return totalRef;
    }

    private Integer countTotalWessex() {
        int totalRef = 0;
        XSSFSheet wessexSheet = this.pswWorkbook.getSheet("Wessex");
        int valueColNo = PoiHelper.getCellColumnByString("Value", wessexSheet);
        int countOfTraineesRowNo = PoiHelper.getCellRowByString("Count of Trainees", wessexSheet);
        totalRef = (int)wessexSheet.getRow(countOfTraineesRowNo).getCell(valueColNo).getNumericCellValue();
        return totalRef;
    }

    public static boolean isCellEmpty(Cell cell) {
        if (cell == null || cell.getCellType() == CellType.BLANK) {
            return true;
        }
        return cell.getCellType() == CellType.STRING && cell.getStringCellValue().isEmpty();
    }

    private static boolean isRowEmpty(Row row) {
        for (int c = row.getFirstCellNum(); c < row.getLastCellNum(); ++c) {
            Cell cell = row.getCell(c);
            if (cell == null || cell.getCellType() == CellType.BLANK) continue;
            return false;
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
                }
                catch (IllegalStateException illegalStateException) {
                    // empty catch block
                }
                if (!cellValueStr.equals(str)) continue;
                rowNumber = c.getRowIndex();
            }
        }
        return rowNumber;
    }

    public void genGraphFile() {
        try {
            this.graphs.createNewFile();
        }
        catch (IOException ex) {
            Logger.getLogger(PoiHelper.class.getName()).log(Level.SEVERE, null, ex);
        }
    }

    public void genMergedFormsFile() {
        try {
            this.mergedForms.createNewFile();
        }
        catch (IOException ex) {
            Logger.getLogger(PoiHelper.class.getName()).log(Level.SEVERE, null, ex);
        }
    }

    public ArrayList<TraineeFormRecord> getTraineeFormList() {
        ArrayList<TraineeFormRecord> list = new ArrayList<TraineeFormRecord>();
        XSSFSheet sheet = this.traineeFormWorkbook.getSheetAt(0);
        ArrayList<Integer> gmcList = new ArrayList<Integer>();
        int fNameCol = PoiHelper.getCellColumnByString("First name", sheet);
        int lNameCol = PoiHelper.getCellColumnByString("Last name", sheet);
        int traineeEmailCol = PoiHelper.getCellColumnByString("Preferred email", sheet);
        int akaCol = PoiHelper.getCellColumnByString("What would you like to be known as?", sheet);
        int gmcCol = PoiHelper.getCellColumnByString("GMC or GDC number", sheet);
        int genderCol = PoiHelper.getCellColumnByString("Gender", sheet);
        int ageCol = PoiHelper.getCellColumnByString("Age", sheet);
        int ethnicityCol = PoiHelper.getCellColumnByString("Ethnicity", sheet);
        int sexOrCol = PoiHelper.getCellColumnByString("Sexual orientation", sheet);
        int relCol = PoiHelper.getCellColumnByString("Religion", sheet);
        int disabilityCol = PoiHelper.getCellColumnByString("Disability", sheet);
        int learningDifcol = PoiHelper.getCellColumnByString("If so, please select which", sheet);
        int gradeCol = PoiHelper.getCellColumnByString("Position/Grade", sheet);
        int specialtyCol = PoiHelper.getCellColumnByString("Specialty", sheet);
        int lastARCPCol = PoiHelper.getCellColumnByString("Last ARCP outcome", sheet);
        int fullLFTCol = PoiHelper.getCellColumnByString("Full/LTFT", sheet);
        int trustCol = PoiHelper.getCellColumnByString("Trust", sheet);
        int countryCol = PoiHelper.getCellColumnByString("Country of Primary Qualification", sheet);
        int yearGradCol = PoiHelper.getCellColumnByString("Year of graduation", sheet);
        int ageGradCol = PoiHelper.getCellColumnByString("Age at graduation", sheet);
        int yearsCol = PoiHelper.getCellColumnByString("How many years have you been working in the UK?", sheet);
        int firstLangCol = PoiHelper.getCellColumnByString("First language", sheet);
        int extToTrainingCol = PoiHelper.getCellColumnByString("Has extension to training been given to you?", sheet);
        Iterator<Row> rowIterator = sheet.rowIterator();
        while (rowIterator.hasNext()) {
            Row nextRow = rowIterator.next();
            if (PoiHelper.isRowEmpty(nextRow) || nextRow.getRowNum() <= 0) continue;
            TraineeFormRecord traineeFormRecord = new TraineeFormRecord();
            traineeFormRecord.setfName(nextRow.getCell(fNameCol).getStringCellValue());
            if (!PoiHelper.isCellEmpty(nextRow.getCell(akaCol))) {
                traineeFormRecord.setAka(nextRow.getCell(akaCol).getStringCellValue());
            }
            traineeFormRecord.setEmail(nextRow.getCell(traineeEmailCol).getStringCellValue());
            traineeFormRecord.setlName(nextRow.getCell(lNameCol).getStringCellValue());
            traineeFormRecord.setGmc((int)nextRow.getCell(gmcCol).getNumericCellValue());
            gmcList.add((int)nextRow.getCell(gmcCol).getNumericCellValue());
            traineeFormRecord.setGender(nextRow.getCell(genderCol).getStringCellValue());
            traineeFormRecord.setAge((int)nextRow.getCell(ageCol).getNumericCellValue());
            traineeFormRecord.setEthnicity(nextRow.getCell(ethnicityCol).getStringCellValue());
            traineeFormRecord.setSexOr(nextRow.getCell(sexOrCol).getStringCellValue());
            traineeFormRecord.setReligion(nextRow.getCell(relCol).getStringCellValue());
            traineeFormRecord.setDisability(nextRow.getCell(disabilityCol).getStringCellValue());
            traineeFormRecord.setLearningDif(nextRow.getCell(learningDifcol).getStringCellValue());
            traineeFormRecord.setGrade(nextRow.getCell(gradeCol).getStringCellValue());
            traineeFormRecord.setSpecialty(nextRow.getCell(specialtyCol).getStringCellValue());
            traineeFormRecord.setLastARCP(nextRow.getCell(lastARCPCol).getStringCellValue());
            traineeFormRecord.setFullLTFT(nextRow.getCell(fullLFTCol).getStringCellValue());
            traineeFormRecord.setTrust(nextRow.getCell(trustCol).getStringCellValue());
            traineeFormRecord.setCountry(nextRow.getCell(countryCol).getStringCellValue());
            traineeFormRecord.setYearOfGrad((int)nextRow.getCell(yearGradCol).getNumericCellValue());
            traineeFormRecord.setAgeAtGrad((int)nextRow.getCell(ageGradCol).getNumericCellValue());
            traineeFormRecord.setYearsWorking((int)nextRow.getCell(yearsCol).getNumericCellValue());
            traineeFormRecord.setFirstLanguage(nextRow.getCell(firstLangCol).getStringCellValue());
            traineeFormRecord.setExtToTraining(nextRow.getCell(extToTrainingCol).getStringCellValue());
            list.add(traineeFormRecord);
        }
        HashSet<Integer> s = new HashSet<Integer>();
        for (Integer gmc : gmcList) {
            if (s.add(gmc)) continue;
            this.appendToLog("Duplicated GMC found in trainee form: " + gmc);
        }
        return list;
    }

    public ArrayList<ReferrerFormRecord> getReferrerFormList() {
        ArrayList<ReferrerFormRecord> list = new ArrayList<ReferrerFormRecord>();
        ArrayList<Integer> gmcList = new ArrayList<Integer>();
        XSSFSheet sheet = this.referrerFormWorkbook.getSheetAt(0);
        int refNameCol = PoiHelper.getCellColumnByString("Referrer Name", sheet);
        int refEmailCol = PoiHelper.getCellColumnByString("Referrer Email", sheet);
        int refJobCol = PoiHelper.getCellColumnByString("Referrer Job Role", sheet);
        int refDateCol = PoiHelper.getCellColumnByString("Referral Date", sheet);
        int titleCol = PoiHelper.getCellColumnByString("Title", sheet);
        int reasonsCol = PoiHelper.getCellColumnByString("Issues identified so far", sheet);
        int gmcCol = PoiHelper.getCellColumnByString("GMC or GDC number", sheet);
        Iterator<Row> rowIterator = sheet.rowIterator();
        while (rowIterator.hasNext()) {
            Row nextRow = rowIterator.next();
            if (PoiHelper.isRowEmpty(nextRow) || nextRow.getRowNum() <= 0) continue;
            ReferrerFormRecord referrerRecord = new ReferrerFormRecord();
            referrerRecord.setReferrerName(nextRow.getCell(refNameCol).getStringCellValue());
            referrerRecord.setReferrerEmail(nextRow.getCell(refEmailCol).getStringCellValue());
            referrerRecord.setReferrerJobRole(nextRow.getCell(refJobCol).getStringCellValue());
            referrerRecord.setReferralDate(nextRow.getCell(refDateCol).getDateCellValue());
            referrerRecord.setTitle(nextRow.getCell(titleCol).getStringCellValue());
            referrerRecord.setReasons(nextRow.getCell(reasonsCol).getStringCellValue());
            referrerRecord.setTraineeGMC((int)nextRow.getCell(gmcCol).getNumericCellValue());
            gmcList.add((int)nextRow.getCell(gmcCol).getNumericCellValue());
            list.add(referrerRecord);
        }
        HashSet<Integer> s = new HashSet<Integer>();
        for (Integer gmc : gmcList) {
            if (s.add(gmc)) continue;
            this.appendToLog("Duplicated GMC found in referrer form: " + gmc);
        }
        return list;
    }

    public void mergeForms() {
        ArrayList<TraineeFormRecord> traineeList = this.getTraineeFormList();
        ArrayList<ReferrerFormRecord> referrerList = this.getReferrerFormList();
        ArrayList<ReferralRecord> refRecordList = new ArrayList<ReferralRecord>();
        for (ReferrerFormRecord referrerRecord : referrerList) {
            int a = 0;
            for (TraineeFormRecord trainRecord : traineeList) {
                String learningDif;
                if (!Objects.equals(trainRecord.getGmc(), referrerRecord.getTraineeGMC())) continue;
                ++a;
                ReferralRecord record = new ReferralRecord();
                record.setReferrerName(referrerRecord.getReferrerName());
                record.setReferrerEmail(referrerRecord.getReferrerEmail());
                record.setReferrerJobRole(referrerRecord.getReferrerJobRole());
                record.setRefDate(referrerRecord.getReferralDate());
                record.setTraineeFName(trainRecord.getfName());
                record.setTraineeLName(trainRecord.getlName());
                record.setTraineeAka(trainRecord.getAka());
                record.setTraineeEmail(trainRecord.getEmail());
                record.setTitle(referrerRecord.getTitle());
                record.setTraineeGMC(referrerRecord.getTraineeGMC());
                record.setLastARCP(trainRecord.getLastARCP());
                record.setFullLTFT(trainRecord.getFullLTFT());
                record.setYearOfGrad(trainRecord.getYearOfGrad());
                record.setAgeAtGrad(trainRecord.getAgeAtGrad());
                record.setYearsWorking(trainRecord.getYearsWorking());
                record.setFirstLanguage(trainRecord.getFirstLanguage());
                record.setSpecialty(trainRecord.getSpecialty());
                record.setGender(trainRecord.getGender());
                record.setAge(trainRecord.getAge());
                record.setCountry(trainRecord.getCountry());
                record.setEthnicity(trainRecord.getEthnicity());
                record.setReligion(trainRecord.getReligion());
                record.setDisability(trainRecord.getDisability());
                record.setSexOr(trainRecord.getSexOr());
                record.setTrust(trainRecord.getTrust());
                record.setGrade(trainRecord.getGrade());
                String extTraining = trainRecord.getExtToTraining();
                if (extTraining.contains("ARCP")) {
                    record.setExtARCP(true);
                }
                if (extTraining.contains("Health")) {
                    record.setExtExam(true);
                }
                if (extTraining.contains("Exam")) {
                    record.setExtHealth(true);
                }
                if ((learningDif = trainRecord.getLearningDif()).contains("ADHD")) {
                    record.setAdhd(true);
                }
                if (learningDif.contains("ASD")) {
                    record.setAsd(true);
                }
                if (learningDif.contains("Dyslexia")) {
                    record.setDyslexia(true);
                }
                if (learningDif.contains("Dyspraxia")) {
                    record.setDyspraxia(true);
                }
                String reasons = referrerRecord.getReasons();
                String other = this.isOtherRefReason(reasons);
                if (reasons.contains("Anxiety / Stress")) {
                    record.setAnxiety(true);
                }
                if (reasons.contains("Career support")) {
                    record.setCarreer(true);
                }
                if (reasons.contains("Clinical skills")) {
                    record.setClinSkills(true);
                }
                if (reasons.contains("Communication")) {
                    record.setCommunication(true);
                }
                if (reasons.contains("Conduct")) {
                    record.setConduct(true);
                }
                if (reasons.contains("Cultural factors")) {
                    record.setCultural(true);
                }
                if (reasons.contains("Exam support")) {
                    record.setExam(true);
                }
                if (reasons.contains("Health Conditions (Mental)")) {
                    record.setHealthMental(true);
                }
                if (reasons.contains("Health Conditions (Physical)")) {
                    record.setHealthPhysical(true);
                }
                if (reasons.contains("Language support")) {
                    record.setLanguage(true);
                }
                if (reasons.contains("Professionalism")) {
                    record.setProfessionalism(true);
                }
                if (reasons.contains("ADHD")) {
                    record.setAdhd(true);
                }
                if (reasons.contains("ASD")) {
                    record.setAsd(true);
                }
                if (reasons.contains("Dyslexia")) {
                    record.setDyslexia(true);
                }
                if (reasons.contains("Dyspraxia")) {
                    record.setDyspraxia(true);
                }
                if (reasons.contains("SRTT")) {
                    record.setSrtt(true);
                }
                if (reasons.contains("Team working")) {
                    record.setTeam(true);
                }
                if (reasons.contains("Time / Workload Management")) {
                    record.setTime(true);
                }
                if (reasons.contains("Capability")) {
                    record.setCapability(true);
                }
                if (!other.equals("")) {
                    record.setOtherReason(true);
                    record.setOtherRefReason(this.isOtherRefReason(referrerRecord.getReasons()));
                }
                refRecordList.add(record);
            }
            if (a != 0) continue;
            this.appendToLog("No matching record found for GMC " + referrerRecord.getTraineeGMC());
        }
        XSSFSheet referrals = this.mergedFormsWorkbook.createSheet("Referrals");
        XSSFRow titlesRow = referrals.createRow(0);
        titlesRow.createCell(0).setCellValue("Referrer Name");
        titlesRow.createCell(1).setCellValue("Referrer Email");
        titlesRow.createCell(2).setCellValue("Referrer Job Role");
        titlesRow.createCell(3).setCellValue("Referral Date");
        titlesRow.createCell(4).setCellValue("Title");
        titlesRow.createCell(5).setCellValue("First name");
        titlesRow.createCell(6).setCellValue("What would you like to be known as?");
        titlesRow.createCell(7).setCellValue("Last name");
        titlesRow.createCell(8).setCellValue("Email");
        titlesRow.createCell(9).setCellValue("GMC or GDC number");
        titlesRow.createCell(10).setCellValue("Anxiety / Stress");
        titlesRow.createCell(11).setCellValue("Capability");
        titlesRow.createCell(12).setCellValue("Career support");
        titlesRow.createCell(13).setCellValue("Clinical skills");
        titlesRow.createCell(14).setCellValue("Communication / Interpersonal skills");
        titlesRow.createCell(15).setCellValue("Conduct");
        titlesRow.createCell(16).setCellValue("Cultural factors");
        titlesRow.createCell(17).setCellValue("Exam support");
        titlesRow.createCell(18).setCellValue("Health Conditions (Mental)");
        titlesRow.createCell(19).setCellValue("Health Conditions (Physical)");
        titlesRow.createCell(20).setCellValue("Language support");
        titlesRow.createCell(21).setCellValue("Professionalism");
        titlesRow.createCell(22).setCellValue("ADHD");
        titlesRow.createCell(23).setCellValue("ASD");
        titlesRow.createCell(24).setCellValue("Dyslexia");
        titlesRow.createCell(25).setCellValue("Dyspraxia");
        titlesRow.createCell(26).setCellValue("SRTT");
        titlesRow.createCell(27).setCellValue("Team working");
        titlesRow.createCell(28).setCellValue("Time / Workload Management");
        titlesRow.createCell(29).setCellValue("Other reasons");
        titlesRow.createCell(30).setCellValue("Gender");
        titlesRow.createCell(31).setCellValue("Age");
        titlesRow.createCell(32).setCellValue("Ethnicity");
        titlesRow.createCell(33).setCellValue("Sexual Orientation");
        titlesRow.createCell(34).setCellValue("Religion");
        titlesRow.createCell(35).setCellValue("Disability");
        titlesRow.createCell(36).setCellValue("ADHD");
        titlesRow.createCell(37).setCellValue("ASD");
        titlesRow.createCell(38).setCellValue("Dyslexia");
        titlesRow.createCell(39).setCellValue("Dyspraxia");
        titlesRow.createCell(40).setCellValue("Grade");
        titlesRow.createCell(41).setCellValue("Specialty");
        titlesRow.createCell(42).setCellValue("Last ARCP outcome");
        titlesRow.createCell(43).setCellValue("Full/LTFT");
        titlesRow.createCell(44).setCellValue("Trust");
        titlesRow.createCell(45).setCellValue("Country of Primary Qualification");
        titlesRow.createCell(46).setCellValue("Year of graduation");
        titlesRow.createCell(47).setCellValue("Age at graduation");
        titlesRow.createCell(48).setCellValue("How many years have you been working in the UK?");
        titlesRow.createCell(49).setCellValue("First language");
        titlesRow.createCell(50).setCellValue("Extension to training - ARCP Outcome");
        titlesRow.createCell(51).setCellValue("Extension to training - Health Factors");
        titlesRow.createCell(52).setCellValue("Extension to training - Exam Failure");
        int i = 0;
        for (ReferralRecord record : refRecordList) {
            XSSFRow newRow = referrals.createRow(++i);
            newRow.createCell(0).setCellValue(record.getReferrerName());
            newRow.createCell(1).setCellValue(record.getReferrerEmail());
            newRow.createCell(2).setCellValue(record.getReferrerJobRole());
            XSSFCellStyle cellStyle = this.mergedFormsWorkbook.createCellStyle();
            cellStyle.setDataFormat((short)14);
            Cell dateCell = newRow.createCell(3);
            dateCell.setCellStyle(cellStyle);
            dateCell.setCellValue(record.getRefDate());
            newRow.createCell(4).setCellValue(record.getTitle());
            newRow.createCell(5).setCellValue(record.getTraineeFName());
            newRow.createCell(6).setCellValue(record.getTraineeAka());
            newRow.createCell(7).setCellValue(record.getTraineeLName());
            newRow.createCell(8).setCellValue(record.getTraineeEmail());
            newRow.createCell(9).setCellValue(record.getTraineeGMC().intValue());
            if (record.isAnxiety()) {
                newRow.createCell(10).setCellValue("X");
            }
            if (record.isCapability()) {
                newRow.createCell(11).setCellValue("X");
            }
            if (record.isCarreer()) {
                newRow.createCell(12).setCellValue("X");
            }
            if (record.isClinSkills()) {
                newRow.createCell(13).setCellValue("X");
            }
            if (record.isCommunication()) {
                newRow.createCell(14).setCellValue("X");
            }
            if (record.isConduct()) {
                newRow.createCell(15).setCellValue("X");
            }
            if (record.isCultural()) {
                newRow.createCell(16).setCellValue("X");
            }
            if (record.isExam()) {
                newRow.createCell(17).setCellValue("X");
            }
            if (record.isHealthMental()) {
                newRow.createCell(18).setCellValue("X");
            }
            if (record.isHealthPhysical()) {
                newRow.createCell(19).setCellValue("X");
            }
            if (record.isLanguage()) {
                newRow.createCell(20).setCellValue("X");
            }
            if (record.isProfessionalism()) {
                newRow.createCell(21).setCellValue("X");
            }
            if (record.isAdhd()) {
                newRow.createCell(22).setCellValue("X");
            }
            if (record.isAsd()) {
                newRow.createCell(23).setCellValue("X");
            }
            if (record.isDyslexia()) {
                newRow.createCell(24).setCellValue("X");
            }
            if (record.isDyspraxia()) {
                newRow.createCell(25).setCellValue("X");
            }
            if (record.isSrtt()) {
                newRow.createCell(26).setCellValue("X");
            }
            if (record.isTeam()) {
                newRow.createCell(27).setCellValue("X");
            }
            if (record.isTime()) {
                newRow.createCell(28).setCellValue("X");
            }
            if (record.isOtherReason()) {
                newRow.createCell(29).setCellValue(record.getOtherRefReason());
            }
            newRow.createCell(30).setCellValue(record.getGender());
            newRow.createCell(31).setCellValue(record.getAge());
            newRow.createCell(32).setCellValue(record.getEthnicity());
            newRow.createCell(33).setCellValue(record.getSexOr());
            newRow.createCell(34).setCellValue(record.getReligion());
            newRow.createCell(35).setCellValue(record.getDisability());
            if (record.isAdhd()) {
                newRow.createCell(36).setCellValue("X");
            }
            if (record.isAsd()) {
                newRow.createCell(37).setCellValue("X");
            }
            if (record.isDyslexia()) {
                newRow.createCell(38).setCellValue("X");
            }
            if (record.isDyspraxia()) {
                newRow.createCell(39).setCellValue("X");
            }
            newRow.createCell(40).setCellValue(record.getGrade());
            newRow.createCell(41).setCellValue(record.getSpecialty());
            newRow.createCell(42).setCellValue(record.getLastARCP());
            newRow.createCell(43).setCellValue(record.getFullLTFT());
            newRow.createCell(44).setCellValue(record.getTrust());
            newRow.createCell(45).setCellValue(record.getCountry());
            newRow.createCell(46).setCellValue(record.getYearOfGrad().intValue());
            newRow.createCell(47).setCellValue(record.getAgeAtGrad().intValue());
            newRow.createCell(48).setCellValue(record.getYearsWorking().intValue());
            newRow.createCell(49).setCellValue(record.getFirstLanguage());
            if (record.isExtARCP()) {
                newRow.createCell(50).setCellValue("X");
            }
            if (record.isExtExam()) {
                newRow.createCell(51).setCellValue("X");
            }
            if (!record.isExtHealth()) continue;
            newRow.createCell(52).setCellValue("X");
        }
        this.writeToMergedForms();
    }

    public void mergeFullInfo() {
        XSSFSheet refSheet = this.referrerFormWorkbook.getSheetAt(0);
        XSSFSheet trainSheet = this.traineeFormWorkbook.getSheetAt(0);
        XSSFSheet newSheetRef = this.mergedFormsWorkbook.createSheet("Referrer");
        XSSFSheet newSheetTrain = this.mergedFormsWorkbook.createSheet("Trainee");
        int startColRefSheet = PoiHelper.getCellColumnByString("Referrer Name", refSheet);
        int startColTrainSheet = PoiHelper.getCellColumnByString("First name", trainSheet);
        for (Row row : refSheet) {
            int rowNum = row.getRowNum();
            XSSFRow newRow = newSheetRef.createRow(rowNum);
            if (PoiHelper.isRowEmpty(row)) continue;
            int i = -startColRefSheet - 1;
            for (int colNum = 0; colNum < row.getLastCellNum(); ++colNum) {
                Cell c = row.getCell(colNum, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                ++i;
                if (PoiHelper.isCellEmpty(c) || c.getColumnIndex() < startColRefSheet) continue;
                Cell cell = newRow.createCell(i);
                PoiHelper.copyCells(c, cell);
            }
        }
        for (Row row : trainSheet) {
            int autosumCol = PoiHelper.getCellColumnByString("How do you find quickly locating information in a document?", trainSheet);
            int rowNum = row.getRowNum();
            XSSFRow newRow = newSheetTrain.createRow(rowNum);
            if (PoiHelper.isRowEmpty(row)) continue;
            int i = -startColTrainSheet - 1;
            for (int colNum = 0; colNum < row.getLastCellNum(); ++colNum) {
                Cell c = row.getCell(colNum, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                ++i;
                if (PoiHelper.isCellEmpty(c) || c.getColumnIndex() < startColTrainSheet) continue;
                Cell cell = newRow.createCell(i);
                PoiHelper.copyCells(c, cell);
                if (c.getColumnIndex() != autosumCol) continue;
                if (rowNum == 0) {
                    newRow.createCell(++i).setCellValue("Scores Autosum");
                    continue;
                }
                int var1Col = i - 14;
                int var2Col = i - 13;
                int var3Col = i - 12;
                int var4Col = i - 11;
                int var5Col = i - 10;
                int var6Col = i - 9;
                int var7Col = i - 8;
                int var8Col = i - 7;
                int var9Col = i - 6;
                int var10Col = i - 5;
                int var11Col = i - 4;
                int var12Col = i - 3;
                int var13Col = i - 2;
                int var14Col = i - 1;
                int var15Col = i++;
                Double var1 = newRow.getCell(var1Col).getNumericCellValue();
                int var1Trans = PoiHelper.tranformScores(var1);
                Double var2 = newRow.getCell(var2Col).getNumericCellValue();
                int var2Trans = PoiHelper.tranformScores(var2);
                Double var3 = newRow.getCell(var3Col).getNumericCellValue();
                int var3Trans = PoiHelper.tranformScores(var3);
                Double var4 = newRow.getCell(var4Col).getNumericCellValue();
                int var4Trans = PoiHelper.tranformScores(var4);
                Double var5 = newRow.getCell(var5Col).getNumericCellValue();
                int var5Trans = PoiHelper.tranformScores(var5);
                Double var6 = newRow.getCell(var6Col).getNumericCellValue();
                int var6Trans = PoiHelper.tranformScores(var6);
                Double var7 = newRow.getCell(var7Col).getNumericCellValue();
                int var7Trans = PoiHelper.tranformScores(var7);
                Double var8 = newRow.getCell(var8Col).getNumericCellValue();
                int var8Trans = PoiHelper.tranformScores(var8);
                Double var9 = newRow.getCell(var9Col).getNumericCellValue();
                int var9Trans = PoiHelper.tranformScores(var9);
                Double var10 = newRow.getCell(var10Col).getNumericCellValue();
                int var10Trans = PoiHelper.tranformScores(var10);
                Double var11 = newRow.getCell(var11Col).getNumericCellValue();
                int var11Trans = PoiHelper.tranformScores(var11);
                Double var12 = newRow.getCell(var12Col).getNumericCellValue();
                int var12Trans = PoiHelper.tranformScores(var12);
                Double var13 = newRow.getCell(var13Col).getNumericCellValue();
                int var13Trans = PoiHelper.tranformScores(var13);
                Double var14 = newRow.getCell(var14Col).getNumericCellValue();
                int var14Trans = PoiHelper.tranformScores(var14);
                Double var15 = newRow.getCell(var15Col).getNumericCellValue();
                int var15Trans = PoiHelper.tranformScores(var15);
                int total = var1Trans + var2Trans + var3Trans + var4Trans + var5Trans + var6Trans + var7Trans + var8Trans + var9Trans + var10Trans + var11Trans + var12Trans + var13Trans + var14Trans + var15Trans;
                newRow.createCell(i).setCellValue(total);
            }
        }
        this.writeToMergedForms();
    }

    private static int tranformScores(Double var) {
        int tranformedScore = 0;
        if (var == 1.0) {
            tranformedScore = 3;
        } else if (var == 2.0) {
            tranformedScore = 6;
        } else if (var == 2.0) {
            tranformedScore = 9;
        } else if (var == 4.0) {
            tranformedScore = 12;
        }
        return tranformedScore;
    }

    private static boolean isNumeric(String strNum) {
        if (strNum == null) {
            return false;
        }
        try {
            double d = Double.parseDouble(strNum);
        }
        catch (NumberFormatException nfe) {
            return false;
        }
        return true;
    }

    public void deleteLastSheetGraphs() {
        this.reLoadGraphWorkbook();
        this.graphsWorkbook.removeSheetAt(this.graphsWorkbook.getSheetIndex("Evaluation Warning"));
        this.writeToGraphs();
    }

    public ArrayList<ReferralRecord> readReferralRecords() {
        ArrayList<ReferralRecord> records = new ArrayList<ReferralRecord>();
        XSSFSheet refSheet = this.pswWorkbook.getSheetAt(this.pswWorkbook.getSheetIndex("Referrals"));
        int referrerNameCol = PoiHelper.getCellColumnByString("Referrer Name", refSheet);
        int referrerJobCol = PoiHelper.getCellColumnByString("Referrer Job Role", refSheet);
        int referrerEmailCol = PoiHelper.getCellColumnByString("Referrer Email", refSheet);
        int traineeFNameCol = PoiHelper.getCellColumnByString("First Name", refSheet);
        int traineeLNameCol = PoiHelper.getCellColumnByString("Last Name", refSheet);
        int traineeGMCCol = PoiHelper.getCellColumnByString("GMC / GDC", refSheet);
        int traineeEmailCol = PoiHelper.getCellColumnByString("Email", refSheet);
        int spcCol = PoiHelper.getCellColumnByString("Specialty", refSheet);
        int genderCol = PoiHelper.getCellColumnByString("Gender", refSheet);
        int trainedCol = PoiHelper.getCellColumnByString("Country of Primary Qualification", refSheet);
        int ageCol = PoiHelper.getCellColumnByString("Age", refSheet);
        int ethCol = PoiHelper.getCellColumnByString("Ethnicity", refSheet);
        int relCol = PoiHelper.getCellColumnByString("Religion", refSheet);
        int disabilityCol = PoiHelper.getCellColumnByString("Disability", refSheet);
        int sexOrCol = PoiHelper.getCellColumnByString("Sexual Orientation", refSheet);
        int exSupportCol = PoiHelper.getCellColumnByString("Exam support", refSheet);
        int anxietyCol = PoiHelper.getCellColumnByString("Anxiety / Stress", refSheet);
        int capabilityCol = PoiHelper.getCellColumnByString("Capability", refSheet);
        int carreerCol = PoiHelper.getCellColumnByString("Career support", refSheet);
        int clinicalSkillsCol = PoiHelper.getCellColumnByString("Clinical skills", refSheet);
        int communicationCol = PoiHelper.getCellColumnByString("Communication / Interpersonal skills", refSheet);
        int conductCol = PoiHelper.getCellColumnByString("Conduct", refSheet);
        int culturalCol = PoiHelper.getCellColumnByString("Cultural factors", refSheet);
        int mentalHealthCol = PoiHelper.getCellColumnByString("Health Conditions (Mental)", refSheet);
        int physicalHealthCol = PoiHelper.getCellColumnByString("Health Conditions (Physical)", refSheet);
        int languageCol = PoiHelper.getCellColumnByString("Language support", refSheet);
        int profCol = PoiHelper.getCellColumnByString("Professionalism", refSheet);
        int adhdCol = PoiHelper.getCellColumnByString("ADHD", refSheet);
        int asdCol = PoiHelper.getCellColumnByString("ASD", refSheet);
        int dyslexiaCol = PoiHelper.getCellColumnByString("Dyslexia", refSheet);
        int dyspraxiaCol = PoiHelper.getCellColumnByString("Dyspraxia", refSheet);
        int srttCol = PoiHelper.getCellColumnByString("SRTT", refSheet);
        int teamCol = PoiHelper.getCellColumnByString("Team working", refSheet);
        int timeCol = PoiHelper.getCellColumnByString("Time / Workload Management", refSheet);
        int otherRefReasonCol = PoiHelper.getCellColumnByString("Other reasons", refSheet);
        int trustCol = PoiHelper.getCellColumnByString("Trust", refSheet);
        int gradeCol = PoiHelper.getCellColumnByString("Grade", refSheet);
        int dateCol = PoiHelper.getCellColumnByString("Referral Date", refSheet);
        int closedCol = PoiHelper.getCellColumnByString("Case open", refSheet);
        Iterator<Row> rowIterator = refSheet.iterator();
        boolean i = false;
        while (rowIterator.hasNext()) {
            Cell dateCell;
            Cell sexOrCell;
            Cell disabilityCell;
            Cell relCell;
            Cell ethCell;
            Cell ageCell;
            Cell trainedCell;
            Cell genderCell;
            Cell spCell;
            Cell traineeEmailCell;
            Cell traineGMCCell;
            Cell traineeLNameCell;
            Cell traineeFNameCell;
            Cell refEmailCell;
            Cell refJobCell;
            Row row = rowIterator.next();
            int currentRow = row.getRowNum();
            if (currentRow <= 3 || PoiHelper.isRowEmpty(row)) continue;
            Cell refNameCell = CellUtil.getCell(row, referrerNameCol);
            if (refNameCell.getStringCellValue().equals("")) {
                this.appendToLog("Empty value in Referrer Name column, row " + (currentRow + 1));
            }
            if ((refJobCell = CellUtil.getCell(row, referrerJobCol)).getStringCellValue().equals("")) {
                this.appendToLog("Empty value in Referrer Job Role column, row " + (currentRow + 1));
            }
            if ((refEmailCell = CellUtil.getCell(row, referrerEmailCol)).getStringCellValue().equals("")) {
                this.appendToLog("Empty value in Referrer Email column, row " + (currentRow + 1));
            }
            if ((traineeFNameCell = CellUtil.getCell(row, traineeFNameCol)).getStringCellValue().equals("")) {
                this.appendToLog("Empty value in Trainee First Name column, row " + (currentRow + 1));
            }
            if ((traineeLNameCell = CellUtil.getCell(row, traineeLNameCol)).getStringCellValue().equals("")) {
                this.appendToLog("Empty value in Trainee Last Name column, row " + (currentRow + 1));
            }
            if (PoiHelper.isCellEmpty(traineGMCCell = CellUtil.getCell(row, traineeGMCCol))) {
                this.appendToLog("Empty value in Trainee GMC column, row " + (currentRow + 1));
            }
            if ((traineeEmailCell = CellUtil.getCell(row, traineeEmailCol)).getStringCellValue().equals("")) {
                this.appendToLog("Empty value in Trainee Email column, row " + (currentRow + 1));
            }
            if ((spCell = CellUtil.getCell(row, spcCol)).getStringCellValue().equals("")) {
                this.appendToLog("Empty value in Specialty column, row " + (currentRow + 1));
            }
            if ((genderCell = CellUtil.getCell(row, genderCol)).getStringCellValue().equals("")) {
                this.appendToLog("Empty value in Gender column, row " + (currentRow + 1));
            }
            if ((trainedCell = CellUtil.getCell(row, trainedCol)).getStringCellValue().equals("")) {
                this.appendToLog("Empty value in Country of Primary Qualification column, row " + (currentRow + 1));
            }
            if (PoiHelper.isCellEmpty(ageCell = CellUtil.getCell(row, ageCol))) {
                this.appendToLog("Empty value in Age column, row " + (currentRow + 1));
            }
            if ((ethCell = CellUtil.getCell(row, ethCol)).getStringCellValue().equals("")) {
                this.appendToLog("Empty value in Ethnicity column, row " + (currentRow + 1));
            }
            if ((relCell = CellUtil.getCell(row, relCol)).getStringCellValue().equals("")) {
                this.appendToLog("Empty value in Religion column, row " + (currentRow + 1));
            }
            if ((disabilityCell = CellUtil.getCell(row, disabilityCol)).getStringCellValue().equals("")) {
                this.appendToLog("Empty value in Disablity column, row " + (currentRow + 1));
            }
            if ((sexOrCell = CellUtil.getCell(row, sexOrCol)).getStringCellValue().equals("")) {
                this.appendToLog("Empty value in Sexual Orientation column, row " + (currentRow + 1));
            }
            Cell anxietyCell = CellUtil.getCell(row, anxietyCol);
            Cell capabilityCell = CellUtil.getCell(row, capabilityCol);
            Cell carreerCell = CellUtil.getCell(row, carreerCol);
            Cell clinSkillsCell = CellUtil.getCell(row, clinicalSkillsCol);
            Cell communicationCell = CellUtil.getCell(row, communicationCol);
            Cell conductCell = CellUtil.getCell(row, conductCol);
            Cell culturalCell = CellUtil.getCell(row, culturalCol);
            Cell exSupportCell = CellUtil.getCell(row, exSupportCol);
            Cell mentalHealthCell = CellUtil.getCell(row, mentalHealthCol);
            Cell physicalHealthCell = CellUtil.getCell(row, physicalHealthCol);
            Cell languageCell = CellUtil.getCell(row, languageCol);
            Cell profCell = CellUtil.getCell(row, profCol);
            Cell adhdCell = CellUtil.getCell(row, adhdCol);
            Cell asdCell = CellUtil.getCell(row, asdCol);
            Cell dyslexiaCell = CellUtil.getCell(row, dyslexiaCol);
            Cell dyspraxiaCell = CellUtil.getCell(row, dyspraxiaCol);
            Cell srttCell = CellUtil.getCell(row, srttCol);
            Cell teamCell = CellUtil.getCell(row, teamCol);
            Cell timeCell = CellUtil.getCell(row, timeCol);
            Cell otherCell = CellUtil.getCell(row, otherRefReasonCol);
            Cell trustCell = CellUtil.getCell(row, trustCol);
            if (trustCell.getStringCellValue().equals("")) {
                this.appendToLog("Empty value in Trust column, row " + (currentRow + 1));
            }
            Cell gradeCell = CellUtil.getCell(row, gradeCol);
            if (disabilityCell.getStringCellValue().equals("")) {
                this.appendToLog("Empty value in Grade column, row " + (currentRow + 1));
            }
            if (PoiHelper.isCellEmpty(dateCell = CellUtil.getCell(row, dateCol))) {
                this.appendToLog("Empty value in Date column, row " + (currentRow + 1));
            }
            Cell closedCell = CellUtil.getCell(row, closedCol);
            ReferralRecord record = new ReferralRecord();
            record.setSpecialty(spCell.getStringCellValue());
            record.setGender(genderCell.getStringCellValue());
            record.setCountry(trainedCell.getStringCellValue());
            if (ageCell.getCellType() == CellType.NUMERIC) {
                record.setAge((int)ageCell.getNumericCellValue());
            } else {
                this.appendToLog("Unexpected value in Age column, row " + (currentRow + 1));
            }
            DataFormatter formatter = new DataFormatter(Locale.UK);
            formatter.formatCellValue(dateCell);
            try {
                record.setRefDate(dateCell.getDateCellValue());
            }
            catch (IllegalStateException e) {
                this.appendToLog("Unexpected value in Date column, row " + (currentRow + 1));
                break;
            }
            record.setEthnicity(ethCell.getStringCellValue());
            record.setReligion(relCell.getStringCellValue());
            record.setDisability(disabilityCell.getStringCellValue());
            record.setSexOr(sexOrCell.getStringCellValue());
            record.setTrust(trustCell.getStringCellValue());
            record.setGrade(gradeCell.getStringCellValue());
            if (!PoiHelper.isCellEmpty(anxietyCell)) {
                record.setAnxiety(true);
            }
            if (!PoiHelper.isCellEmpty(capabilityCell)) {
                record.setCapability(true);
            }
            if (!PoiHelper.isCellEmpty(carreerCell)) {
                record.setCarreer(true);
            }
            if (!PoiHelper.isCellEmpty(clinSkillsCell)) {
                record.setClinSkills(true);
            }
            if (!PoiHelper.isCellEmpty(communicationCell)) {
                record.setCommunication(true);
            }
            if (!PoiHelper.isCellEmpty(conductCell)) {
                record.setConduct(true);
            }
            if (!PoiHelper.isCellEmpty(culturalCell)) {
                record.setCultural(true);
            }
            if (!PoiHelper.isCellEmpty(exSupportCell)) {
                record.setExam(true);
            }
            if (!PoiHelper.isCellEmpty(mentalHealthCell)) {
                record.setHealthMental(true);
            }
            if (!PoiHelper.isCellEmpty(physicalHealthCell)) {
                record.setHealthPhysical(true);
            }
            if (!PoiHelper.isCellEmpty(languageCell)) {
                record.setLanguage(true);
            }
            if (!PoiHelper.isCellEmpty(profCell)) {
                record.setProfessionalism(true);
            }
            if (!PoiHelper.isCellEmpty(adhdCell)) {
                record.setAdhd(true);
            }
            if (!PoiHelper.isCellEmpty(asdCell)) {
                record.setAsd(true);
            }
            if (!PoiHelper.isCellEmpty(dyslexiaCell)) {
                record.setDyslexia(true);
            }
            if (!PoiHelper.isCellEmpty(dyspraxiaCell)) {
                record.setDyspraxia(true);
            }
            if (!PoiHelper.isCellEmpty(srttCell)) {
                record.setSrtt(true);
            }
            if (!PoiHelper.isCellEmpty(teamCell)) {
                record.setTeam(true);
            }
            if (!PoiHelper.isCellEmpty(timeCell)) {
                record.setTime(true);
            }
            if (!PoiHelper.isCellEmpty(otherCell)) {
                record.setOtherReason(true);
            }
            if (PoiHelper.isCellEmpty(closedCell)) {
                record.setCaseOpen(false);
            } else {
                record.setCaseOpen(true);
            }
            records.add(record);
        }
        this.recordList = records;
        return records;
    }

    public boolean isIntegrityCheck() {
        boolean isIntegrity = false;
        ArrayList<String> columns = new ArrayList<String>();
        columns.add("Referrer Name");
        columns.add("Referrer Job Role");
        columns.add("Referrer Email");
        columns.add("Referral Date");
        columns.add("Title");
        columns.add("First Name");
        columns.add("Last Name");
        columns.add("Email");
        columns.add("GMC / GDC");
        columns.add("Anxiety / Stress");
        columns.add("First Name");
        columns.add("Capability");
        columns.add("Career support");
        columns.add("Clinical skills");
        columns.add("Communication / Interpersonal skills");
        columns.add("Conduct");
        columns.add("Cultural factors");
        columns.add("Exam support");
        columns.add("Health Conditions (Physical)");
        columns.add("Health Conditions (Mental)");
        columns.add("Language support");
        columns.add("Professionalism");
        columns.add("ADHD");
        columns.add("ASD");
        columns.add("Dyslexia");
        columns.add("Dyspraxia");
        columns.add("Team working");
        columns.add("Time / Workload Management");
        columns.add("Other reasons");
        columns.add("Gender");
        columns.add("Age");
        columns.add("Ethnicity");
        columns.add("Sexual Orientation");
        columns.add("Religion");
        columns.add("Disability");
        columns.add("ADHD");
        columns.add("ASD");
        columns.add("Dyslexia");
        columns.add("Dyspraxia");
        columns.add("Grade");
        columns.add("Specialty");
        columns.add("Last ARCP outcome");
        columns.add("Full/LTFT");
        columns.add("Trust");
        columns.add("Country of Primary Qualification");
        columns.add("Year of graduation");
        columns.add("Age at graduation");
        columns.add("How many years working in the UK");
        columns.add("First Language");
        columns.add("Extension to training - ARCP Outcome");
        columns.add("Extension to training - Exam Failure");
        columns.add("Extension to training - Health Factors");
        columns.add("Case Manager");
        columns.add("Specialist Support Group");
        columns.add("Re-referral");
        columns.add("Case open");
        columns.add("Date case closed");
        columns.add("Outcome");
        XSSFSheet refSheet = this.pswWorkbook.getSheetAt(this.pswWorkbook.getSheetIndex("Referrals"));
        ArrayList<String> cellContents = new ArrayList<String>();
        for (Cell c : refSheet.getRow(3)) {
            if (PoiHelper.isCellEmpty(c)) continue;
            cellContents.add(c.getStringCellValue());
        }
        for (Cell c : refSheet.getRow(2)) {
            if (PoiHelper.isCellEmpty(c)) continue;
            cellContents.add(c.getStringCellValue());
        }
        ArrayList<String> missingColumns = new ArrayList<String>();
        for (int i = 0; i < columns.size(); ++i) {
            if (cellContents.contains(columns.get(i))) continue;
            missingColumns.add((String)columns.get(i));
        }
        if (!missingColumns.isEmpty()) {
            isIntegrity = true;
        }
        return isIntegrity;
    }

    public ArrayList<String> getMissingColumns() {
        ArrayList<String> columns = new ArrayList<String>();
        columns.add("Referrer Name");
        columns.add("Referrer Job Role");
        columns.add("Referrer Email");
        columns.add("Referral Date");
        columns.add("Title");
        columns.add("First Name");
        columns.add("Last Name");
        columns.add("Known as");
        columns.add("Email");
        columns.add("GMC / GDC");
        columns.add("Anxiety / Stress");
        columns.add("Capability");
        columns.add("Career support");
        columns.add("Clinical skills");
        columns.add("Communication / Interpersonal skills");
        columns.add("Conduct");
        columns.add("Cultural factors");
        columns.add("Exam support");
        columns.add("Health Conditions (Physical)");
        columns.add("Health Conditions (Mental)");
        columns.add("Language support");
        columns.add("Professionalism");
        columns.add("ADHD");
        columns.add("ASD");
        columns.add("Dyslexia");
        columns.add("Dyspraxia");
        columns.add("SRTT");
        columns.add("Team working");
        columns.add("Time / Workload Management");
        columns.add("Other reasons");
        columns.add("Gender");
        columns.add("Age");
        columns.add("Ethnicity");
        columns.add("Sexual Orientation");
        columns.add("Religion");
        columns.add("Disability");
        columns.add("ADHD");
        columns.add("ASD");
        columns.add("Dyslexia");
        columns.add("Dyspraxia");
        columns.add("Grade");
        columns.add("Specialty");
        columns.add("Last ARCP outcome");
        columns.add("Full/LTFT");
        columns.add("Trust");
        columns.add("Country of Primary Qualification");
        columns.add("Year of graduation");
        columns.add("Age at graduation");
        columns.add("How many years working in the UK");
        columns.add("First Language");
        columns.add("Extension to training - ARCP Outcome");
        columns.add("Extension to training - Exam Failure");
        columns.add("Extension to training - Health Factors");
        columns.add("Case Manager");
        columns.add("Specialist Support Group");
        columns.add("Re-referral");
        columns.add("Case open");
        columns.add("Date case closed");
        columns.add("Outcome");
        XSSFSheet refSheet = this.pswWorkbook.getSheetAt(this.pswWorkbook.getSheetIndex("Referrals"));
        ArrayList<String> cellContents = new ArrayList<String>();
        for (Cell c : refSheet.getRow(3)) {
            if (PoiHelper.isCellEmpty(c)) continue;
            cellContents.add(c.getStringCellValue());
        }
        for (Cell c : refSheet.getRow(2)) {
            if (PoiHelper.isCellEmpty(c)) continue;
            cellContents.add(c.getStringCellValue());
        }
        ArrayList<String> missingColumns = new ArrayList<String>();
        for (int i = 0; i < columns.size(); ++i) {
            if (cellContents.contains(columns.get(i))) continue;
            missingColumns.add((String)columns.get(i));
        }
        return missingColumns;
    }

    static Font copyFont(Font font1, Workbook wb2) {
        byte underline;
        short typeOffset;
        boolean isStrikeout;
        boolean isItalic;
        String fontName;
        short fontHeight;
        short color;
        boolean isBold = font1.getBold();
        Font font2 = wb2.findFont(isBold, color = font1.getColor(), fontHeight = font1.getFontHeight(), fontName = font1.getFontName(), isItalic = font1.getItalic(), isStrikeout = font1.getStrikeout(), typeOffset = font1.getTypeOffset(), underline = font1.getUnderline());
        if (font2 == null) {
            font2 = wb2.createFont();
            font2.setBold(isBold);
            font2.setColor(color);
            font2.setFontHeight(fontHeight);
            font2.setFontName(fontName);
            font2.setItalic(isItalic);
            font2.setStrikeout(isStrikeout);
            font2.setTypeOffset(typeOffset);
            font2.setUnderline(underline);
        }
        return font2;
    }

    static void copyStyles(Cell cell1, Cell cell2) {
        CellStyle style1 = cell1.getCellStyle();
        HashMap<String, Object> properties = new HashMap<String, Object>();
        short dataFormat1 = style1.getDataFormat();
        if (BuiltinFormats.getBuiltinFormat(dataFormat1) == null) {
            String formatString1 = style1.getDataFormatString();
            DataFormat format2 = cell2.getSheet().getWorkbook().createDataFormat();
            dataFormat1 = format2.getFormat(formatString1);
        }
        properties.put("dataFormat", dataFormat1);
        FillPatternType fillPattern = style1.getFillPattern();
        short fillForegroundColor = style1.getFillForegroundColor();
        properties.put("fillPattern", (Object)fillPattern);
        properties.put("fillForegroundColor", fillForegroundColor);
        Font font1 = cell1.getSheet().getWorkbook().getFontAt(style1.getFontIndexAsInt());
        Font font2 = PoiHelper.copyFont(font1, cell2.getSheet().getWorkbook());
        properties.put("font", font2.getIndexAsInt());
        CellUtil.setCellStyleProperties(cell2, properties);
    }

    static void copyCells(Cell cell1, Cell cell2) {
        switch (cell1.getCellType()) {
            case STRING: {
                String string1 = cell1.getStringCellValue();
                cell2.setCellValue(string1);
                break;
            }
            case NUMERIC: {
                if (DateUtil.isCellDateFormatted(cell1)) {
                    Date date1 = cell1.getDateCellValue();
                    cell2.setCellValue(date1);
                    break;
                }
                double cellValue1 = cell1.getNumericCellValue();
                cell2.setCellValue(cellValue1);
                break;
            }
            case BLANK: {
                cell2.setCellValue("");
            }
        }
        PoiHelper.copyStyles(cell1, cell2);
    }

    public ArrayList<File> getFiles() {
        ArrayList<File> list = new ArrayList<File>();
        list.add(this.psw);
        list.add(this.graphs);
        return list;
    }

    private String isOtherRefReason(String reasons) {
        String other = "";
        String[] split = reasons.split(";");
        boolean i = false;
        for (String str : split) {
            if (str.equals("Anxiety / Stress") || str.equals("Career support") || str.equals("Clinical skills") || str.equals("Communication / Interpersonal skills") || str.equals("Conduct") || str.equals("Cultural factors") || str.equals("Exam support") || str.equals("Health Conditions (Mental)") || str.equals("Health Conditions (Physical)") || str.equals("Language support") || str.equals("Professionalism") || str.equals("(SpLD) Specific Learning Difficulty/Difference - ADHD") || str.equals("(SpLD) Specific Learning Difficulty/Difference - ASD") || str.equals("(SpLD) Specific Learning Difficulty/Difference - Dyslexia") || str.equals("(SpLD) Specific Learning Difficulty/Difference - Dyspraxia") || str.equals("Supported Return to Training") || str.equals("Team working") || str.equals("Time / Workload Management") || str.equals("Capability")) continue;
            other = str;
        }
        return other;
    }

    public void writeToGraphs() {
        try {
            FileOutputStream fileOut = new FileOutputStream(this.graphs);
            this.graphsWorkbook.write(fileOut);
            fileOut.close();
        }
        catch (FileNotFoundException e) {
            MessageDialogs md = new MessageDialogs();
            md.showFileOpenDialog(this.graphs);
        }
        catch (IOException ex) {
            Logger.getLogger(PoiHelper.class.getName()).log(Level.SEVERE, null, ex);
        }
    }

    private void writeToMergedForms() {
        try {
            FileOutputStream fileOut = new FileOutputStream(this.mergedForms);
            this.mergedFormsWorkbook.write(fileOut);
            fileOut.close();
        }
        catch (FileNotFoundException e) {
            MessageDialogs md = new MessageDialogs();
            md.showFileOpenDialog(this.graphs);
        }
        catch (IOException ex) {
            Logger.getLogger(PoiHelper.class.getName()).log(Level.SEVERE, null, ex);
        }
    }

    public void appendToLog(String str) {
        try {
            FileWriter myWriter = new FileWriter(this.log, true);
            SimpleDateFormat formatter = new SimpleDateFormat("dd/MM/yyyy HH:mm:ss");
            Date date = new Date();
            myWriter.append(formatter.format(date) + "\n");
            myWriter.append(str + "\n\n");
            myWriter.close();
        }
        catch (FileNotFoundException myWriter) {
        }
        catch (IOException e) {
            e.printStackTrace();
        }
    }

    public boolean isFileOpen(File file) {
        boolean isFileOpen = false;
        try {
            FileOutputStream fileOut = new FileOutputStream(file);
            if (file.getName().contains("GRAPHS")) {
                this.graphsWorkbook.write(fileOut);
            } else if (this.mergedForms != null) {
                this.mergedFormsWorkbook.write(fileOut);
            }
            fileOut.close();
        }
        catch (FileNotFoundException e) {
            isFileOpen = true;
        }
        catch (IOException ex) {
            Logger.getLogger(PoiHelper.class.getName()).log(Level.SEVERE, null, ex);
        }
        return isFileOpen;
    }

    public void genLogFile() {
        try {
            this.log.createNewFile();
        }
        catch (IOException ex) {
            Logger.getLogger(PoiHelper.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
}
