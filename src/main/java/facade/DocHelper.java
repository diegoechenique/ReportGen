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
import java.io.InputStream;
import java.io.OutputStream;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.HashMap;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;
import static jdk.nashorn.internal.objects.NativeError.printStackTrace;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.wp.usermodel.Paragraph;
import org.apache.poi.xwpf.usermodel.BodyElementType;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.xmlbeans.XmlException;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRow;
import res.Strings;
import vo.RefRecord;
import vo.Table8Line;

/**
 *
 * @author diego
 */
public class DocHelper {

    private InputStream template = Thread.currentThread().getContextClassLoader().getResourceAsStream(Strings.DOC_PATH);
    private File outFile;
    private File psw;
    private File graphs;
    PoiHelper helper;

    int y = getStartingYear();
    int ys1 = getStartingYear() - 1;
    int ys2 = getStartingYear() - 2;
    int ys3 = getStartingYear() - 3;
    int yp1 = getStartingYear() + 1;

    String y2 = genSplitYearSeq(String.valueOf(yp1));

    int referredCount;
    int wessexCount;

    XWPFDocument doc;

    public DocHelper(PoiHelper poiHelper, File outFile) {
            this.outFile = outFile;   
            this.psw = poiHelper.getFiles().get(0);
            this.graphs = poiHelper.getFiles().get(1);
            helper = poiHelper;
            loadTemplate();
    }
    
    private void loadTemplate(){
        try {
            this.doc = new XWPFDocument(this.template);
        } catch (IOException ex) {
            //--Do something
        }
    }

    public void doTextUpdate() {
        for (XWPFParagraph p : doc.getParagraphs()) {
            replaceParagraph(p, genMap(Strings.DOC_YR2, y2));
            replaceParagraph(p, genMap(Strings.DOC_Y, String.valueOf(y)));
            replaceParagraph(p, genMap(Strings.DOC_YS1, String.valueOf(ys1)));
            replaceParagraph(p, genMap(Strings.DOC_YS2, String.valueOf(ys2)));
            replaceParagraph(p, genMap(Strings.DOC_YS3, String.valueOf(ys3)));
            replaceParagraph(p, genMap(Strings.DOC_YP1, String.valueOf(yp1)));
            replaceParagraph(p, genMap(Strings.DOC_TCR, String.valueOf(helper.countTotalReferrals())));
        }
        for (XWPFTable t : doc.getTables()) {
            replaceTable(t, genMap(Strings.DOC_Y, String.valueOf(y)));
            replaceTable(t, genMap(Strings.DOC_YR2, String.valueOf(y2)));
            replaceTable(t, genMap(Strings.DOC_YS1, String.valueOf(ys1)));
            replaceTable(t, genMap(Strings.DOC_YS2, String.valueOf(ys2)));
            replaceTable(t, genMap(Strings.DOC_YS3, String.valueOf(ys3)));
            replaceTable(t, genMap(Strings.DOC_YP1, String.valueOf(yp1)));
        }
    }

    public void getTable0() {
        List<Integer> t0list = helper.countTable0();
        int stCount = t0list.get(0);
        int fCount = t0list.get(1);
        int gpCount = t0list.get(2);
        int otherCount = t0list.get(3);
        int total = t0list.get(4);
        int casesClosed = t0list.get(5);
        int casesOClosed = t0list.get(6);
        XWPFTable table0 = doc.getTables().get(0);
        replaceTable(table0, genMap(Strings.TABLE_A, Integer.toString(stCount)));
        replaceTable(table0, genMap(Strings.TABLE_B, Integer.toString(fCount)));
        replaceTable(table0, genMap(Strings.TABLE_C, Integer.toString(gpCount)));
        replaceTable(table0, genMap(Strings.TABLE_D, Integer.toString(otherCount)));
        replaceTable(table0, genMap(Strings.TABLE_E, Integer.toString(total)));
        replaceTable(table0, genMap(Strings.TABLE_F, Integer.toString(casesClosed)));
        replaceTable(table0, genMap(Strings.TABLE_G, Integer.toString(casesOClosed)));
    }

    public void getTable1() {
        
        List<Integer> list = helper.countTable1();
        XWPFTable table1 = doc.getTables().get(1);

 
        int anxietyMCount = list.get(0);
        int anxietyFCount = list.get(1);
        int capabilityMCount = list.get(2);
        int capabilityFCount = list.get(3);
        int carreerMCount = list.get(4);
        int carreerFCount = list.get(5);
        int clinicalMCount = list.get(6);
        int clinicalFCount = list.get(7);
        int communicationMCount = list.get(8);
        int communicationFCount = list.get(9);
        int conductMCount = list.get(10);
        int conductFCount = list.get(11);
        int culturalMCount = list.get(12);
        int culturalFCount = list.get(13);
        int examMCount = list.get(14);
        int examFCount = list.get(15);
        int phHealthMCount = list.get(16);
        int phHealthFCount = list.get(17);
        int menHealthMCount = list.get(18);
        int menHealthFCount = list.get(19);
        int languageMCount = list.get(20);
        int languageFCount = list.get(21);
        int profMCount = list.get(22);
        int profFCount = list.get(23);
        int adhdMCount = list.get(24);
        int adhdFCount = list.get(25);
        int asdMCount = list.get(26);
        int asdFCount = list.get(27);
        int dyslexiaMCount = list.get(28);
        int dyslexiaFCount = list.get(29);
        int dyspraxiaMCount = list.get(30);
        int dyspraxiaFCount = list.get(31);
        int srttMCount = list.get(32);
        int srttFCount = list.get(33);
        int teamMCount = list.get(34);
        int teamFCount = list.get(35);
        int timeMCount = list.get(36);
        int timeFCount = list.get(37);
        int otherMCount = list.get(38);
        int otherFCount = list.get(39);
        
        writeSmallCell(table1.getRow(2).getCell(4), String.valueOf(anxietyMCount));
        writeSmallCell(table1.getRow(3).getCell(4), String.valueOf(anxietyFCount));
        writeSmallCell(table1.getRow(4).getCell(4), String.valueOf(capabilityMCount));
        writeSmallCell(table1.getRow(5).getCell(4), String.valueOf(capabilityFCount));
        writeSmallCell(table1.getRow(6).getCell(4), String.valueOf(carreerMCount));
        writeSmallCell(table1.getRow(7).getCell(4), String.valueOf(carreerFCount));
        writeSmallCell(table1.getRow(8).getCell(4), String.valueOf(clinicalMCount));
        writeSmallCell(table1.getRow(9).getCell(4), String.valueOf(clinicalFCount));
        writeSmallCell(table1.getRow(10).getCell(4), String.valueOf(communicationMCount));
        writeSmallCell(table1.getRow(11).getCell(4), String.valueOf(communicationFCount));
        writeSmallCell(table1.getRow(12).getCell(4), String.valueOf(conductMCount));
        writeSmallCell(table1.getRow(13).getCell(4), String.valueOf(conductFCount));
        writeSmallCell(table1.getRow(14).getCell(4), String.valueOf(culturalMCount));
        writeSmallCell(table1.getRow(15).getCell(4), String.valueOf(culturalFCount));
        writeSmallCell(table1.getRow(16).getCell(4), String.valueOf(examMCount));
        writeSmallCell(table1.getRow(17).getCell(4), String.valueOf(examFCount));
        writeSmallCell(table1.getRow(18).getCell(4), String.valueOf(menHealthMCount));
        writeSmallCell(table1.getRow(19).getCell(4), String.valueOf(menHealthFCount));
        writeSmallCell(table1.getRow(20).getCell(4), String.valueOf(phHealthMCount));
        writeSmallCell(table1.getRow(21).getCell(4), String.valueOf(phHealthFCount));
        writeSmallCell(table1.getRow(22).getCell(4), String.valueOf(languageMCount));
        writeSmallCell(table1.getRow(23).getCell(4), String.valueOf(languageFCount));
        writeSmallCell(table1.getRow(24).getCell(4), String.valueOf(profMCount));
        writeSmallCell(table1.getRow(25).getCell(4), String.valueOf(profFCount));
        writeSmallCell(table1.getRow(26).getCell(4), String.valueOf(adhdMCount));
        writeSmallCell(table1.getRow(27).getCell(4), String.valueOf(adhdFCount));
        writeSmallCell(table1.getRow(28).getCell(4), String.valueOf(asdMCount));
        writeSmallCell(table1.getRow(29).getCell(4), String.valueOf(asdFCount));
        writeSmallCell(table1.getRow(30).getCell(4), String.valueOf(dyslexiaMCount));
        writeSmallCell(table1.getRow(31).getCell(4), String.valueOf(dyslexiaFCount));
        writeSmallCell(table1.getRow(32).getCell(4), String.valueOf(dyspraxiaMCount));
        writeSmallCell(table1.getRow(33).getCell(4), String.valueOf(dyspraxiaFCount));
        writeSmallCell(table1.getRow(34).getCell(4), String.valueOf(srttMCount));
        writeSmallCell(table1.getRow(35).getCell(4), String.valueOf(srttFCount));
        writeSmallCell(table1.getRow(36).getCell(4), String.valueOf(teamMCount));
        writeSmallCell(table1.getRow(37).getCell(4), String.valueOf(teamFCount));
        writeSmallCell(table1.getRow(38).getCell(4), String.valueOf(timeMCount));
        writeSmallCell(table1.getRow(39).getCell(4), String.valueOf(timeFCount));
        writeSmallCell(table1.getRow(40).getCell(4), String.valueOf(otherMCount));
        writeSmallCell(table1.getRow(41).getCell(4), String.valueOf(otherFCount));

    }

    public void getTable2() {
        List<Integer> t2list = helper.countTable2();
        double f1TotalCount = t2list.get(1);
        double f2TotalCount = t2list.get(2);
        double f1RefCount = t2list.get(3);
        double f2RefCount = t2list.get(4);
        double totalFCount = f1TotalCount + f2TotalCount;
        double totalFRef = f1RefCount + f2RefCount;
        double f1Perc = (f1RefCount / f1TotalCount) * 100;
        double f2Perc = (f2RefCount / f2TotalCount) * 100;
        double fPerc = ((totalFRef) / totalFCount) * 100;
        double fRefPerc = ((totalFRef) / referredCount) * 100;
        DecimalFormat format = new DecimalFormat("#.#");
        XWPFTable table2 = doc.getTables().get(2);
        replaceTable(table2, genMap(Strings.TABLE_B, String.valueOf((int) totalFRef) + " (" + format.format(fPerc) + "%)"));
        replaceTable(table2, genMap(Strings.TABLE_C, "F1 = " + (int) f1RefCount + " F2 = " + (int) f2RefCount));
        replaceTable(table2, genMap(Strings.TABLE_D, String.valueOf(format.format(fRefPerc) + "%")));
        replaceTable(table2, genMap(Strings.TABLE_E, "F1 = " + (int) f1TotalCount + " F2 = " + (int) f2TotalCount));
        replaceTable(table2, genMap(Strings.TABLE_F, String.valueOf(format.format(f1Perc) + "%")));
        replaceTable(table2, genMap(Strings.TABLE_G, String.valueOf(format.format(f2Perc) + "%")));
    }

    public void getTable3() {
        List<Integer> t3list = helper.countTable3();
        double bournemouthTotal = t3list.get(0);
        double dorchesterTotal = t3list.get(1);
        double dorsetTotal = t3list.get(2);
        double hhftTotal = t3list.get(3);
        double iowTotal = t3list.get(4);
        double jerseyTotal = t3list.get(5);
        double pooleTotal = t3list.get(6);
        double portsmouthTotal = t3list.get(7);
        double salisburyTotal = t3list.get(8);
        double solentTotal = t3list.get(9);
        double southamptonTotal = t3list.get(10);
        double southernTotal = t3list.get(11);
        double bournemouthRefNo = t3list.get(12);
        double dorchesterRefNo = t3list.get(13);
        double dorsetRefNo = t3list.get(14);
        double hhftRefNo = t3list.get(15);
        double iowRefNo = t3list.get(16);
        double jerseyRefNo = t3list.get(17);
        double pooleRefNo = t3list.get(18);
        double portsmouthRefNo = t3list.get(19);
        double salisburyRefNo = t3list.get(20);
        double solentRefNo = t3list.get(21);
        double southamptonRefNo = t3list.get(22);
        double southernRefNo = t3list.get(23);
        wessexCount = t3list.get(25);
        double bournemouthTrInTrust = Math.round((bournemouthRefNo / bournemouthTotal) * 100);
        double dorchesterTrInTrust = Math.round((dorchesterRefNo / dorchesterTotal) * 100);
        double dorsetTrInTrust = Math.round((dorsetRefNo / dorsetTotal) * 100);
        double hhtfTrInTrust = Math.round((hhftRefNo / hhftTotal) * 100);
        double iowTrInTrust = Math.round((iowRefNo / iowTotal) * 100);
        double jerseyTrInTrust = Math.round((jerseyRefNo / jerseyTotal) * 100);
        double pooleTrInTrust = Math.round((pooleRefNo / pooleTotal) * 100);
        double portsmouthTrInTrust = Math.round((portsmouthRefNo / portsmouthTotal) * 100);
        double salisburyTrInTrust = Math.round((salisburyRefNo / salisburyTotal) * 100);
        double solentTrInTrust = Math.round((solentRefNo / solentTotal) * 100);
        double southamptonTrInTrust = Math.round((southamptonRefNo / southamptonTotal) * 100);
        double southernTrInTrust = Math.round((southernRefNo / southernTotal) * 100);
        double bournemouthOfWssx = Math.round((bournemouthTotal / wessexCount) * 100);
        double dorchesterOfWssx = Math.round((dorchesterTotal / wessexCount) * 100);
        double dorsetOfWssx = Math.round((dorsetTotal / wessexCount) * 100);
        double hhtfOfWssx = Math.round((hhftTotal / wessexCount) * 100);
        double iowOfWssx = Math.round((iowTotal / wessexCount) * 100);
        double jerseyOfWssx = Math.round((jerseyTotal / wessexCount) * 100);
        double pooleOfWssx = Math.round((pooleTotal / wessexCount) * 100);
        double portsmouthOfWssx = Math.round((portsmouthTotal / wessexCount) * 100);
        double salisburyOfWssx = Math.round((salisburyTotal / wessexCount) * 100);
        double solentOfWssx = Math.round((solentTotal / wessexCount) * 100);
        double southamptonOfWssx = Math.round((southamptonTotal / wessexCount) * 100);
        double southernOfWssx = Math.round((southernTotal / wessexCount) * 100);
        double bournemouthOfPSU = Math.round((bournemouthRefNo / referredCount) * 100);
        double dorchesterOfPSU = Math.round((dorchesterRefNo / referredCount) * 100);
        double dorsetOfPSU = Math.round((dorsetRefNo / referredCount) * 100);
        double hhtfOfPSU = Math.round((hhftRefNo / referredCount) * 100);
        double iowOfPSU = Math.round((iowRefNo / referredCount) * 100);
        double jerseyOfPSU = Math.round((jerseyRefNo / referredCount) * 100);
        double pooleOfPSU = Math.round((pooleRefNo / referredCount) * 100);
        double portsmouthOfPSU = Math.round((portsmouthRefNo / referredCount) * 100);
        double salisburyOfPSU = Math.round((salisburyRefNo / referredCount) * 100);
        double solentOfPSU = Math.round((solentRefNo / referredCount) * 100);
        double southamptonOfPSU = Math.round((southamptonRefNo / referredCount) * 100);
        double southernOfPSU = Math.round((southernRefNo / referredCount) * 100);
        XWPFTable table3 = doc.getTables().get(3);
        replaceTable(table3, genMap(Strings.TABLE_A, String.valueOf((int) bournemouthRefNo)));
        replaceTable(table3, genMap(Strings.TABLE_B, String.valueOf((int) dorchesterRefNo)));
        replaceTable(table3, genMap(Strings.TABLE_C, String.valueOf((int) dorsetRefNo)));
        replaceTable(table3, genMap(Strings.TABLE_D, String.valueOf((int) hhftRefNo)));
        replaceTable(table3, genMap(Strings.TABLE_E, String.valueOf((int) iowRefNo)));
        replaceTable(table3, genMap(Strings.TABLE_F, String.valueOf((int) jerseyRefNo)));
        replaceTable(table3, genMap(Strings.TABLE_G, String.valueOf((int) pooleRefNo)));
        replaceTable(table3, genMap(Strings.TABLE_H, String.valueOf((int) portsmouthRefNo)));
        replaceTable(table3, genMap(Strings.TABLE_I, String.valueOf((int) salisburyRefNo)));
        replaceTable(table3, genMap(Strings.TABLE_J, String.valueOf((int) solentRefNo)));
        replaceTable(table3, genMap(Strings.TABLE_K, String.valueOf((int) southamptonRefNo)));
        replaceTable(table3, genMap(Strings.TABLE_L, String.valueOf((int) southernRefNo)));
        replaceTable(table3, genMap(Strings.TABLE_M, String.valueOf((int) bournemouthTotal)));
        replaceTable(table3, genMap(Strings.TABLE_N, String.valueOf((int) dorchesterTotal)));
        replaceTable(table3, genMap(Strings.TABLE_O, String.valueOf((int) dorsetTotal)));
        replaceTable(table3, genMap(Strings.TABLE_P, String.valueOf((int) hhftTotal)));
        replaceTable(table3, genMap(Strings.TABLE_Q, String.valueOf((int) iowTotal)));
        replaceTable(table3, genMap(Strings.TABLE_R, String.valueOf((int) jerseyTotal)));
        replaceTable(table3, genMap(Strings.TABLE_S, String.valueOf((int) pooleTotal)));
        replaceTable(table3, genMap(Strings.TABLE_T, String.valueOf((int) portsmouthTotal)));
        replaceTable(table3, genMap(Strings.TABLE_U, String.valueOf((int) salisburyTotal)));
        replaceTable(table3, genMap(Strings.TABLE_V, String.valueOf((int) solentTotal)));
        replaceTable(table3, genMap(Strings.TABLE_W, String.valueOf((int) southamptonTotal)));
        replaceTable(table3, genMap(Strings.TABLE_X, String.valueOf((int) southernTotal)));
        replaceTable(table3, genMap(Strings.TABLE_Y, String.valueOf((int) bournemouthTrInTrust + "%")));
        replaceTable(table3, genMap(Strings.TABLE_Z, String.valueOf((int) dorchesterTrInTrust + "%")));
        replaceTable(table3, genMap(Strings.TABLE_AA, String.valueOf((int) dorsetTrInTrust + "%")));
        replaceTable(table3, genMap(Strings.TABLE_AB, String.valueOf((int) hhtfTrInTrust + "%")));
        replaceTable(table3, genMap(Strings.TABLE_AC, String.valueOf((int) iowTrInTrust + "%")));
        replaceTable(table3, genMap(Strings.TABLE_AD, String.valueOf((int) jerseyTrInTrust + "%")));
        replaceTable(table3, genMap(Strings.TABLE_AE, String.valueOf((int) pooleTrInTrust + "%")));
        replaceTable(table3, genMap(Strings.TABLE_AF, String.valueOf((int) portsmouthTrInTrust + "%")));
        replaceTable(table3, genMap(Strings.TABLE_AG, String.valueOf((int) salisburyTrInTrust + "%")));
        replaceTable(table3, genMap(Strings.TABLE_AH, String.valueOf((int) solentTrInTrust + "%")));
        replaceTable(table3, genMap(Strings.TABLE_AI, String.valueOf((int) southamptonTrInTrust + "%")));
        replaceTable(table3, genMap(Strings.TABLE_AJ, String.valueOf((int) southernTrInTrust + "%")));
        replaceTable(table3, genMap(Strings.TABLE_AK, String.valueOf((int) bournemouthOfWssx + "%")));
        replaceTable(table3, genMap(Strings.TABLE_AL, String.valueOf((int) dorchesterOfWssx + "%")));
        replaceTable(table3, genMap(Strings.TABLE_AM, String.valueOf((int) dorsetOfWssx + "%")));
        replaceTable(table3, genMap(Strings.TABLE_AN, String.valueOf((int) hhtfOfWssx + "%")));
        replaceTable(table3, genMap(Strings.TABLE_AO, String.valueOf((int) iowOfWssx + "%")));
        replaceTable(table3, genMap(Strings.TABLE_AP, String.valueOf((int) jerseyOfWssx + "%")));
        replaceTable(table3, genMap(Strings.TABLE_AQ, String.valueOf((int) pooleOfWssx + "%")));
        replaceTable(table3, genMap(Strings.TABLE_AR, String.valueOf((int) portsmouthOfWssx + "%")));
        replaceTable(table3, genMap(Strings.TABLE_AS, String.valueOf((int) salisburyOfWssx + "%")));
        replaceTable(table3, genMap(Strings.TABLE_AT, String.valueOf((int) solentOfWssx + "%")));
        replaceTable(table3, genMap(Strings.TABLE_AU, String.valueOf((int) southamptonOfWssx + "%")));
        replaceTable(table3, genMap(Strings.TABLE_AV, String.valueOf((int) solentOfWssx + "%")));
        replaceTable(table3, genMap(Strings.TABLE_AW, String.valueOf((int) bournemouthOfPSU + "%")));
        replaceTable(table3, genMap(Strings.TABLE_AX, String.valueOf((int) dorchesterOfPSU + "%")));
        replaceTable(table3, genMap(Strings.TABLE_AY, String.valueOf((int) dorsetOfPSU + "%")));
        replaceTable(table3, genMap(Strings.TABLE_AZ, String.valueOf((int) hhtfOfPSU + "%")));
        replaceTable(table3, genMap(Strings.TABLE_BA, String.valueOf((int) iowOfPSU + "%")));
        replaceTable(table3, genMap(Strings.TABLE_BB, String.valueOf((int) jerseyOfPSU + "%")));
        replaceTable(table3, genMap(Strings.TABLE_BC, String.valueOf((int) pooleOfPSU + "%")));
        replaceTable(table3, genMap(Strings.TABLE_BD, String.valueOf((int) portsmouthOfPSU + "%")));
        replaceTable(table3, genMap(Strings.TABLE_BE, String.valueOf((int) salisburyOfPSU + "%")));
        replaceTable(table3, genMap(Strings.TABLE_BF, String.valueOf((int) solentOfPSU + "%")));
        replaceTable(table3, genMap(Strings.TABLE_BG, String.valueOf((int) southamptonOfPSU + "%")));
        replaceTable(table3, genMap(Strings.TABLE_BH, String.valueOf((int) southernOfPSU + "%")));
    }

    public void getTable4() {
        List<Integer> t4list = helper.countTable4();
        double anaestheticsRefNo = t4list.get(0);
        double dentalRefNo = t4list.get(1);
        double emergRefNo = t4list.get(2);
        double foundationRefNo = t4list.get(3);
        double gpRefNo = t4list.get(4);
        double medicineRefNo = t4list.get(5);
        double obsRefNo = t4list.get(6);
        double occhealthRefNo = t4list.get(7);
        double paediatricsRefNo = t4list.get(8);
        double pathologyRefNo = t4list.get(9);
        double pharmacyRefNo = t4list.get(10);
        double psychRefNo = t4list.get(11);
        double pubhealthRefNo = t4list.get(12);
        double radioRefNo = t4list.get(13);
        double surgeryRefNo = t4list.get(14);
        double anaestheticsTotal = t4list.get(15);
        double dentalTotal = t4list.get(16);
        double emergTotal = t4list.get(17);
        double foundationTotal = t4list.get(18);
        double gpTotal = t4list.get(19);
        double medicineTotal = t4list.get(20);
        double obsTotal = t4list.get(21);
        double occhealthTotal = t4list.get(22);
        double paediatricsTotal = t4list.get(23);
        double pathologyTotal = t4list.get(24);
        double pharmacyTotal = t4list.get(25);
        double psychTotal = t4list.get(26);
        double pubhealthTotal = t4list.get(27);
        double radioTotal = t4list.get(28);
        double surgeryTotal = t4list.get(29);
        double anaestheticsInSpc = Math.round((anaestheticsRefNo / anaestheticsTotal) * 100);
        double dentalInSpc = Math.round((dentalRefNo / dentalTotal) * 100);
        double emergInSpc = Math.round((emergRefNo / emergTotal) * 100);
        double foundationInSpc = Math.round((foundationRefNo / foundationTotal) * 100);
        double gpInSpc = Math.round((gpRefNo / gpTotal) * 100);
        double medicineInSpc = Math.round((medicineRefNo / medicineTotal) * 100);
        double obsInSpc = Math.round((obsRefNo / obsTotal) * 100);
        double occhealthInSpc = Math.round((occhealthRefNo / occhealthTotal) * 100);
        double paediatricsInSpc = Math.round((paediatricsRefNo / paediatricsTotal) * 100);
        double pathologyInSpc = Math.round((pathologyRefNo / pathologyTotal) * 100);
        double pharmacyInSpc = Math.round((pharmacyRefNo / pharmacyTotal) * 100);
        double psychInSpc = Math.round((psychRefNo / psychTotal) * 100);
        double pubhealthInSpc = Math.round((pubhealthRefNo / pubhealthTotal) * 100);
        double radioInSpc = Math.round((radioRefNo / radioTotal) * 100);
        double surgeryInSpc = Math.round((surgeryRefNo / surgeryTotal) * 100);
        double anaestheticsOfWssx = Math.round((anaestheticsTotal / wessexCount) * 100);
        double dentalOfWssx = Math.round((dentalTotal / wessexCount) * 100);
        double emergOfWssx = Math.round((emergTotal / wessexCount) * 100);
        double foundationOfWssx = Math.round((foundationTotal / wessexCount) * 100);
        double gpOfWssx = Math.round((gpTotal / wessexCount) * 100);
        double medicineOfWssx = Math.round((medicineTotal / wessexCount) * 100);
        double obsOfWssx = Math.round((obsTotal / wessexCount) * 100);
        double occhealthOfWssx = Math.round((occhealthTotal / wessexCount) * 100);
        double paediatricsOfWssx = Math.round((paediatricsTotal / wessexCount) * 100);
        double pathologyOfWssx = Math.round((pathologyTotal / wessexCount) * 100);
        double pharmacyOfWssx = Math.round((pharmacyTotal / wessexCount) * 100);
        double psychOfWssx = Math.round((psychTotal / wessexCount) * 100);
        double pubhealthOfWssx = Math.round((pubhealthTotal / wessexCount) * 100);
        double radioOfWssx = Math.round((radioTotal / wessexCount) * 100);
        double surgeryOfWssx = Math.round((surgeryTotal / wessexCount) * 100);
        double anaestheticsOfPSU = Math.round((anaestheticsRefNo / referredCount) * 100);
        double dentalOfPSU = Math.round((dentalRefNo / referredCount) * 100);
        double emergOfPSU = Math.round((emergRefNo / referredCount) * 100);
        double foundationOfPSU = Math.round((foundationRefNo / referredCount) * 100);
        double gpOfPSU = Math.round((gpRefNo / referredCount) * 100);
        double medicineOfPSU = Math.round((medicineRefNo / referredCount) * 100);
        double obsOfPSU = Math.round((obsRefNo / referredCount) * 100);
        double occhealthOfPSU = Math.round((occhealthRefNo / referredCount) * 100);
        double paediatricsOfPSU = Math.round((paediatricsRefNo / referredCount) * 100);
        double pathologyOfPSU = Math.round((pathologyRefNo / referredCount) * 100);
        double pharmacyOfPSU = Math.round((pharmacyRefNo / referredCount) * 100);
        double psychOfPSU = Math.round((psychRefNo / referredCount) * 100);
        double pubhealthOfPSU = Math.round((pubhealthRefNo / referredCount) * 100);
        double radioOfPSU = Math.round((radioRefNo / referredCount) * 100);
        double surgeryOfPSU = Math.round((surgeryRefNo / referredCount) * 100);
        XWPFTable table4 = doc.getTables().get(4);
        replaceTable(table4, genMap(Strings.TABLE_A, String.valueOf((int) anaestheticsRefNo)));
        replaceTable(table4, genMap(Strings.TABLE_B, String.valueOf((int) dentalRefNo)));
        replaceTable(table4, genMap(Strings.TABLE_C, String.valueOf((int) emergRefNo)));
        replaceTable(table4, genMap(Strings.TABLE_D, String.valueOf((int) foundationRefNo)));
        replaceTable(table4, genMap(Strings.TABLE_E, String.valueOf((int) gpRefNo)));
        replaceTable(table4, genMap(Strings.TABLE_F, String.valueOf((int) medicineRefNo)));
        replaceTable(table4, genMap(Strings.TABLE_G, String.valueOf((int) obsRefNo)));
        replaceTable(table4, genMap(Strings.TABLE_H, String.valueOf((int) occhealthRefNo)));
        replaceTable(table4, genMap(Strings.TABLE_I, String.valueOf((int) paediatricsRefNo)));
        replaceTable(table4, genMap(Strings.TABLE_J, String.valueOf((int) pathologyRefNo)));
        replaceTable(table4, genMap(Strings.TABLE_K, String.valueOf((int) pharmacyRefNo)));
        replaceTable(table4, genMap(Strings.TABLE_L, String.valueOf((int) psychRefNo)));
        replaceTable(table4, genMap(Strings.TABLE_M, String.valueOf((int) pubhealthRefNo)));
        replaceTable(table4, genMap(Strings.TABLE_N, String.valueOf((int) radioRefNo)));
        replaceTable(table4, genMap(Strings.TABLE_O, String.valueOf((int) surgeryRefNo)));
        replaceTable(table4, genMap(Strings.TABLE_P, String.valueOf((int) anaestheticsTotal)));
        replaceTable(table4, genMap(Strings.TABLE_Q, String.valueOf((int) dentalTotal)));
        replaceTable(table4, genMap(Strings.TABLE_R, String.valueOf((int) emergTotal)));
        replaceTable(table4, genMap(Strings.TABLE_S, String.valueOf((int) foundationTotal)));
        replaceTable(table4, genMap(Strings.TABLE_T, String.valueOf((int) gpTotal)));
        replaceTable(table4, genMap(Strings.TABLE_U, String.valueOf((int) medicineTotal)));
        replaceTable(table4, genMap(Strings.TABLE_V, String.valueOf((int) obsTotal)));
        replaceTable(table4, genMap(Strings.TABLE_W, String.valueOf((int) occhealthTotal)));
        replaceTable(table4, genMap(Strings.TABLE_X, String.valueOf((int) paediatricsTotal)));
        replaceTable(table4, genMap(Strings.TABLE_Y, String.valueOf((int) pathologyTotal)));
        replaceTable(table4, genMap(Strings.TABLE_Z, String.valueOf((int) pharmacyTotal)));
        replaceTable(table4, genMap(Strings.TABLE_AA, String.valueOf((int) psychTotal)));
        replaceTable(table4, genMap(Strings.TABLE_AB, String.valueOf((int) pubhealthTotal)));
        replaceTable(table4, genMap(Strings.TABLE_AC, String.valueOf((int) radioTotal)));
        replaceTable(table4, genMap(Strings.TABLE_AD, String.valueOf((int) surgeryTotal)));
        replaceTable(table4, genMap(Strings.TABLE_AE, String.valueOf((int) anaestheticsInSpc) + "%"));
        replaceTable(table4, genMap(Strings.TABLE_AF, String.valueOf((int) dentalInSpc) + "%"));
        replaceTable(table4, genMap(Strings.TABLE_AG, String.valueOf((int) emergInSpc) + "%"));
        replaceTable(table4, genMap(Strings.TABLE_AH, String.valueOf((int) foundationInSpc) + "%"));
        replaceTable(table4, genMap(Strings.TABLE_AI, String.valueOf((int) gpInSpc) + "%"));
        replaceTable(table4, genMap(Strings.TABLE_AJ, String.valueOf((int) medicineInSpc) + "%"));
        replaceTable(table4, genMap(Strings.TABLE_AK, String.valueOf((int) obsInSpc) + "%"));
        replaceTable(table4, genMap(Strings.TABLE_AL, String.valueOf((int) occhealthInSpc) + "%"));
        replaceTable(table4, genMap(Strings.TABLE_AM, String.valueOf((int) paediatricsInSpc) + "%"));
        replaceTable(table4, genMap(Strings.TABLE_AN, String.valueOf((int) pathologyInSpc) + "%"));
        replaceTable(table4, genMap(Strings.TABLE_AO, String.valueOf((int) pharmacyInSpc) + "%"));
        replaceTable(table4, genMap(Strings.TABLE_AP, String.valueOf((int) psychInSpc) + "%"));
        replaceTable(table4, genMap(Strings.TABLE_AQ, String.valueOf((int) pubhealthInSpc) + "%"));
        replaceTable(table4, genMap(Strings.TABLE_AR, String.valueOf((int) radioInSpc) + "%"));
        replaceTable(table4, genMap(Strings.TABLE_AS, String.valueOf((int) surgeryInSpc) + "%"));
        replaceTable(table4, genMap(Strings.TABLE_AT, String.valueOf((int) anaestheticsOfWssx) + "%"));
        replaceTable(table4, genMap(Strings.TABLE_AU, String.valueOf((int) dentalOfWssx) + "%"));
        replaceTable(table4, genMap(Strings.TABLE_AV, String.valueOf((int) emergOfWssx) + "%"));
        replaceTable(table4, genMap(Strings.TABLE_AW, String.valueOf((int) foundationOfWssx) + "%"));
        replaceTable(table4, genMap(Strings.TABLE_AX, String.valueOf((int) gpOfWssx) + "%"));
        replaceTable(table4, genMap(Strings.TABLE_AY, String.valueOf((int) medicineOfWssx) + "%"));
        replaceTable(table4, genMap(Strings.TABLE_AZ, String.valueOf((int) obsOfWssx) + "%"));
        replaceTable(table4, genMap(Strings.TABLE_BA, String.valueOf((int) occhealthOfWssx) + "%"));
        replaceTable(table4, genMap(Strings.TABLE_BB, String.valueOf((int) paediatricsOfWssx) + "%"));
        replaceTable(table4, genMap(Strings.TABLE_BC, String.valueOf((int) pathologyOfWssx) + "%"));
        replaceTable(table4, genMap(Strings.TABLE_BD, String.valueOf((int) pharmacyOfWssx) + "%"));
        replaceTable(table4, genMap(Strings.TABLE_BE, String.valueOf((int) psychOfWssx) + "%"));
        replaceTable(table4, genMap(Strings.TABLE_BF, String.valueOf((int) pubhealthOfWssx) + "%"));
        replaceTable(table4, genMap(Strings.TABLE_BG, String.valueOf((int) radioOfWssx) + "%"));
        replaceTable(table4, genMap(Strings.TABLE_BH, String.valueOf((int) surgeryOfWssx) + "%"));
        replaceTable(table4, genMap(Strings.TABLE_BI, String.valueOf((int) anaestheticsOfPSU) + "%"));
        replaceTable(table4, genMap(Strings.TABLE_BJ, String.valueOf((int) dentalOfPSU) + "%"));
        replaceTable(table4, genMap(Strings.TABLE_BK, String.valueOf((int) emergOfPSU) + "%"));
        replaceTable(table4, genMap(Strings.TABLE_BL, String.valueOf((int) foundationOfPSU) + "%"));
        replaceTable(table4, genMap(Strings.TABLE_BM, String.valueOf((int) gpOfPSU) + "%"));
        replaceTable(table4, genMap(Strings.TABLE_BN, String.valueOf((int) medicineOfPSU) + "%"));
        replaceTable(table4, genMap(Strings.TABLE_BO, String.valueOf((int) obsOfPSU) + "%"));
        replaceTable(table4, genMap(Strings.TABLE_BP, String.valueOf((int) occhealthOfPSU) + "%"));
        replaceTable(table4, genMap(Strings.TABLE_BQ, String.valueOf((int) paediatricsOfPSU) + "%"));
        replaceTable(table4, genMap(Strings.TABLE_BR, String.valueOf((int) pathologyOfPSU) + "%"));
        replaceTable(table4, genMap(Strings.TABLE_BS, String.valueOf((int) pharmacyOfPSU) + "%"));
        replaceTable(table4, genMap(Strings.TABLE_BT, String.valueOf((int) psychOfPSU) + "%"));
        replaceTable(table4, genMap(Strings.TABLE_BU, String.valueOf((int) pubhealthOfPSU) + "%"));
        replaceTable(table4, genMap(Strings.TABLE_BV, String.valueOf((int) radioOfPSU) + "%"));
        replaceTable(table4, genMap(Strings.TABLE_BW, String.valueOf((int) surgeryOfPSU) + "%"));
    }

    public void getTable5() {
        XWPFTable table5 = doc.getTables().get(5);
        ArrayList<ArrayList<RefRecord>> list = new ArrayList<>();
        ArrayList<RefRecord> t5bournemouthList = helper.getTable5LineByTrust(Strings.PSW_TRUST_BOURNEMOUTH);
        ArrayList<RefRecord> t5dorchesterList = helper.getTable5LineByTrust(Strings.PSW_TRUST_DORCHESTER);
        ArrayList<RefRecord> t5dorsetList = helper.getTable5LineByTrust(Strings.PSW_TRUST_DORSET);
        ArrayList<RefRecord> t5hhftList = helper.getTable5LineByTrust(Strings.PSW_TRUST_HHFT);
        ArrayList<RefRecord> t5iowList = helper.getTable5LineByTrust(Strings.PSW_TRUST_IOW);
        ArrayList<RefRecord> t5jerseyList = helper.getTable5LineByTrust(Strings.PSW_TRUST_JERSEY);
        ArrayList<RefRecord> t5pooleList = helper.getTable5LineByTrust(Strings.PSW_TRUST_POOLE);
        ArrayList<RefRecord> t5portsmouthList = helper.getTable5LineByTrust(Strings.PSW_TRUST_PORTSMOUTH);
        ArrayList<RefRecord> t5salisburyList = helper.getTable5LineByTrust(Strings.PSW_TRUST_SALISBURY);
        ArrayList<RefRecord> t5solentList = helper.getTable5LineByTrust(Strings.PSW_TRUST_SOLENT);
        ArrayList<RefRecord> t5southamptonList = helper.getTable5LineByTrust(Strings.PSW_TRUST_SOUTHAMPTON);
        ArrayList<RefRecord> t5southernList = helper.getTable5LineByTrust(Strings.PSW_TRUST_SOUTHERN);

        list.add(t5bournemouthList);
        list.add(t5dorchesterList);
        list.add(t5dorsetList);
        list.add(t5hhftList);
        list.add(t5iowList);
        list.add(t5jerseyList);
        list.add(t5pooleList);
        list.add(t5portsmouthList);
        list.add(t5salisburyList);
        list.add(t5solentList);
        list.add(t5southamptonList);
        list.add(t5southernList);
       
        for (ArrayList<RefRecord> rList : list) { 
            if(!rList.isEmpty()){
                for (int i=0; i<rList.size();i++) {
                    int cursor = i;
                    writeT5Line(table5, rList.get(i), cursor);
                }
            }
        }

        table5.removeRow(1);
        table5.removeRow(1);
    }
    
    

    public void getTable7() {
        XWPFTable table7 = doc.getTables().get(7);
        List<Double> list = helper.countTable7();
        Double ssg = list.get(0);
        Double cm = list.get(1);
        Double total = list.get(2);
        replaceTable(table7, genMap(Strings.TABLE_A, String.valueOf(ssg)));
        replaceTable(table7, genMap(Strings.TABLE_B, String.valueOf(cm)));
        replaceTable(table7, genMap(Strings.TABLE_C, String.valueOf(total)));
    }

    public void getTable8() {
        ArrayList<Table8Line> linesArray = helper.countTable8();

        Table8Line anaesthetics = linesArray.get(0);
        Table8Line dental = linesArray.get(1);
        Table8Line dermatology = linesArray.get(2);
        Table8Line endocrinology = linesArray.get(3);
        Table8Line foundation = linesArray.get(4);
        Table8Line gastroenterology = linesArray.get(5);
        Table8Line gp = linesArray.get(6);
        Table8Line haematology = linesArray.get(7);
        Table8Line histopathology = linesArray.get(8);
        Table8Line emergMed = linesArray.get(9);
        Table8Line medicine = linesArray.get(10);
        Table8Line neurology = linesArray.get(11);
        Table8Line obs = linesArray.get(12);
        Table8Line occHealth = linesArray.get(13);
        Table8Line oncology = linesArray.get(14);
        Table8Line ophtalmology = linesArray.get(15);
        Table8Line paediatrics = linesArray.get(16);
        Table8Line pathology = linesArray.get(17);
        Table8Line pharmacy = linesArray.get(18);
        Table8Line psych = linesArray.get(19);
        Table8Line pubHealth = linesArray.get(20);
        Table8Line radiology = linesArray.get(21);
        Table8Line sexHealth = linesArray.get(22);
        Table8Line rheumathology = linesArray.get(23);
        Table8Line surgery = linesArray.get(24);

        writeT8Line(anaesthetics);
        writeT8Line(dental);
        writeT8Line(dermatology);
        writeT8Line(endocrinology);
        writeT8Line(foundation);
        writeT8Line(gastroenterology);
        writeT8Line(gp);
        writeT8Line(haematology);
        writeT8Line(histopathology);
        writeT8Line(emergMed);
        writeT8Line(medicine);
        writeT8Line(neurology);
        writeT8Line(obs);
        writeT8Line(occHealth);
        writeT8Line(oncology);
        writeT8Line(ophtalmology);
        writeT8Line(paediatrics);
        writeT8Line(pathology);
        writeT8Line(pharmacy);
        writeT8Line(psych);
        writeT8Line(pubHealth);
        writeT8Line(radiology);
        writeT8Line(sexHealth);
        writeT8Line(rheumathology);
        writeT8Line(surgery);

        doc.getTables().get(8).removeRow(2);
    }

    private void writeT8Line(Table8Line line) {

        XWPFTable table8 = doc.getTables().get(8);
        XWPFTableRow oldRow = table8.getRow(2);

        if (!line.isEmpty()) {

            CTRow firstCTRow = null;
            try {
                firstCTRow = CTRow.Factory.parse(oldRow.getCtRow().newInputStream());
            } catch (XmlException | IOException | NullPointerException e) {
                printStackTrace(e);
            }
            XWPFTableRow firstRow = new XWPFTableRow(firstCTRow, table8);
            int rowNum = table8.getRows().size() - 1;
            String title = line.getTitle() + "(" + String.valueOf(line.getTotalCount()) + ")";
            writeSmallCell(firstRow.getCell(0), title);
            if (line.getMale() != 0) {
                writeSmallCell(firstRow.getCell(1), String.valueOf(line.getMale()));
            }
            if (line.getFemale() != 0) {
                writeSmallCell(firstRow.getCell(2), String.valueOf(line.getFemale()));
            }
            if (line.getUk() != 0) {
                writeSmallCell(firstRow.getCell(3), String.valueOf(line.getUk()));
            }
            if (line.getNonUk() != 0) {
                writeSmallCell(firstRow.getCell(4), String.valueOf(line.getNonUk()));
            }
            if (line.getAge2329() != 0) {
                writeSmallCell(firstRow.getCell(5), String.valueOf(line.getAge2329()));
            }
            if (line.getAge3035() != 0) {
                writeSmallCell(firstRow.getCell(6), String.valueOf(line.getAge3035()));
            }
            if (line.getAge3540() != 0) {
                writeSmallCell(firstRow.getCell(7), String.valueOf(line.getAge3540()));
            }
            if (line.getAge40() != 0) {
                writeSmallCell(firstRow.getCell(8), String.valueOf(line.getAge40()));
            }
            if (line.getWhiteb() != 0) {
                writeSmallCell(firstRow.getCell(9), String.valueOf(line.getWhiteb()));
            }
            if (line.getWhiteo() != 0) {
                writeSmallCell(firstRow.getCell(10), String.valueOf(line.getWhiteo()));
            }
            if (line.getAsian() != 0) {
                writeSmallCell(firstRow.getCell(11), String.valueOf(line.getAsian()));
            }
            if (line.getAfrican() != 0) {
                writeSmallCell(firstRow.getCell(12), String.valueOf(line.getAfrican()));
            }
            if (line.getEthOther() != 0) {
                writeSmallCell(firstRow.getCell(13), String.valueOf(line.getEthOther()));
            }
            if (line.getChristian() != 0) {
                writeSmallCell(firstRow.getCell(14), String.valueOf(line.getChristian()));
            }
            if (line.getIslam() != 0) {
                writeSmallCell(firstRow.getCell(15), String.valueOf(line.getIslam()));
            }
            if (line.getHindu() != 0) {
                writeSmallCell(firstRow.getCell(16), String.valueOf(line.getHindu()));
            }
            if (line.getAtheist() != 0) {
                writeSmallCell(firstRow.getCell(17), String.valueOf(line.getAtheist()));
            }
            if (line.getSikh() != 0) {
                writeSmallCell(firstRow.getCell(18), String.valueOf(line.getSikh()));
            }
            if (line.getJudaism() != 0) {
                writeSmallCell(firstRow.getCell(19), String.valueOf(line.getJudaism()));
            }
            if (line.getBuddhism() != 0) {
                writeSmallCell(firstRow.getCell(20), String.valueOf(line.getBuddhism()));
            }
            if (line.getRelOther() != 0) {
                writeSmallCell(firstRow.getCell(21), String.valueOf(line.getRelOther()));
            }
            if (line.getRelPNS() != 0) {
                writeSmallCell(firstRow.getCell(22), String.valueOf(line.getRelPNS()));
            }
            if (line.getYes() != 0) {
                writeSmallCell(firstRow.getCell(23), String.valueOf(line.getYes()));
            }
            if (line.getNo() != 0) {
                writeSmallCell(firstRow.getCell(24), String.valueOf(line.getNo()));
            }
            if (line.getHet() != 0) {
                writeSmallCell(firstRow.getCell(25), String.valueOf(line.getHet()));
            }
            if (line.getBisexual() != 0) {
                writeSmallCell(firstRow.getCell(26), String.valueOf(line.getBisexual()));
            }
            if (line.getHomosexual() != 0) {
                writeSmallCell(firstRow.getCell(27), String.valueOf(line.getHomosexual()));
            }
            if (line.getSexOrPNS() != 0) {
                writeSmallCell(firstRow.getCell(28), String.valueOf(line.getSexOrPNS()));
            }
            table8.addRow(firstRow);
        }
    }

    public void getTable9() {
        ArrayList<Table8Line> linesArray = helper.countTable9();

        Table8Line anaesthetics = linesArray.get(0);
        Table8Line dental = linesArray.get(1);
        Table8Line dermatology = linesArray.get(2);
        Table8Line endocrinology = linesArray.get(3);
        Table8Line foundation = linesArray.get(4);
        Table8Line gastroenterology = linesArray.get(5);
        Table8Line gp = linesArray.get(6);
        Table8Line haematology = linesArray.get(7);
        Table8Line histopathology = linesArray.get(8);
        Table8Line emergMed = linesArray.get(9);
        Table8Line medicine = linesArray.get(10);
        Table8Line neurology = linesArray.get(11);
        Table8Line obs = linesArray.get(12);
        Table8Line occHealth = linesArray.get(13);
        Table8Line oncology = linesArray.get(14);
        Table8Line ophtalmology = linesArray.get(15);
        Table8Line paediatrics = linesArray.get(16);
        Table8Line pathology = linesArray.get(17);
        Table8Line pharmacy = linesArray.get(18);
        Table8Line psych = linesArray.get(19);
        Table8Line pubHealth = linesArray.get(20);
        Table8Line radiology = linesArray.get(21);
        Table8Line sexHealth = linesArray.get(22);
        Table8Line rheumathology = linesArray.get(23);
        Table8Line surgery = linesArray.get(24);

        writeT9Line(anaesthetics);
        writeT9Line(dental);
        writeT9Line(dermatology);
        writeT9Line(endocrinology);
        writeT9Line(foundation);
        writeT9Line(gastroenterology);
        writeT9Line(gp);
        writeT9Line(haematology);
        writeT9Line(histopathology);
        writeT9Line(emergMed);
        writeT9Line(medicine);
        writeT9Line(neurology);
        writeT9Line(obs);
        writeT9Line(occHealth);
        writeT9Line(oncology);
        writeT9Line(ophtalmology);
        writeT9Line(paediatrics);
        writeT9Line(pathology);
        writeT9Line(pharmacy);
        writeT9Line(psych);
        writeT9Line(pubHealth);
        writeT9Line(radiology);
        writeT9Line(sexHealth);
        writeT9Line(rheumathology);
        writeT9Line(surgery);

        doc.getTables().get(9).removeRow(2);
    }

    private void writeT9Line(Table8Line line) {
        XWPFTable table9 = doc.getTables().get(9);
        XWPFTableRow oldRow = table9.getRow(2);

        if (!line.isEmpty()) {

            CTRow firstCTRow = null;
            try {
                firstCTRow = CTRow.Factory.parse(oldRow.getCtRow().newInputStream());
            } catch (XmlException | IOException | NullPointerException e) {
                printStackTrace(e);
            }
            XWPFTableRow firstRow = new XWPFTableRow(firstCTRow, table9);
            int rowNum = table9.getRows().size() - 1;

            String title = line.getTitle() + "(" + String.valueOf(line.getTotalCount()) + ")";
            writeSmallCell(firstRow.getCell(0), title);
            if (line.getMale() != 0) {
                writeSmallCell(firstRow.getCell(1), String.valueOf(line.getMale()));
            }
            if (line.getFemale() != 0) {
                writeSmallCell(firstRow.getCell(2), String.valueOf(line.getFemale()));
            }
            if (line.getAge2329() != 0) {
                writeSmallCell(firstRow.getCell(3), String.valueOf(line.getAge2329()));
            }
            if (line.getAge3035() != 0) {
                writeSmallCell(firstRow.getCell(4), String.valueOf(line.getAge3035()));
            }
            if (line.getAge3540() != 0) {
                writeSmallCell(firstRow.getCell(5), String.valueOf(line.getAge3540()));
            }
            if (line.getAge40() != 0) {
                writeSmallCell(firstRow.getCell(6), String.valueOf(line.getAge40()));
            }
            if (line.getWhiteb() != 0) {
                writeSmallCell(firstRow.getCell(7), String.valueOf(line.getWhiteb()));
            }
            if (line.getWhiteo() != 0) {
                writeSmallCell(firstRow.getCell(8), String.valueOf(line.getWhiteo()));
            }
            if (line.getAsian() != 0) {
                writeSmallCell(firstRow.getCell(9), String.valueOf(line.getAsian()));
            }
            if (line.getAfrican() != 0) {
                writeSmallCell(firstRow.getCell(10), String.valueOf(line.getAfrican()));
            }
            if (line.getEthOther() != 0) {
                writeSmallCell(firstRow.getCell(11), String.valueOf(line.getEthOther()));
            }
            if (line.getChristian() != 0) {
                writeSmallCell(firstRow.getCell(12), String.valueOf(line.getChristian()));
            }
            if (line.getIslam() != 0) {
                writeSmallCell(firstRow.getCell(13), String.valueOf(line.getIslam()));
            }
            if (line.getHindu() != 0) {
                writeSmallCell(firstRow.getCell(14), String.valueOf(line.getHindu()));
            }
            if (line.getAtheist() != 0) {
                writeSmallCell(firstRow.getCell(15), String.valueOf(line.getAtheist()));
            }
            if (line.getSikh() != 0) {
                writeSmallCell(firstRow.getCell(16), String.valueOf(line.getSikh()));
            }
            if (line.getJudaism() != 0) {
                writeSmallCell(firstRow.getCell(17), String.valueOf(line.getJudaism()));
            }
            if (line.getBuddhism() != 0) {
                writeSmallCell(firstRow.getCell(18), String.valueOf(line.getBuddhism()));
            }
            if (line.getRelOther() != 0) {
                writeSmallCell(firstRow.getCell(19), String.valueOf(line.getRelOther()));
            }
            if (line.getRelPNS() != 0) {
                writeSmallCell(firstRow.getCell(20), String.valueOf(line.getRelPNS()));
            }
            if (line.getHet() != 0) {
                writeSmallCell(firstRow.getCell(21), String.valueOf(line.getHet()));
            }
            if (line.getBisexual() != 0) {
                writeSmallCell(firstRow.getCell(22), String.valueOf(line.getBisexual()));
            }
            if (line.getHomosexual() != 0) {
                writeSmallCell(firstRow.getCell(23), String.valueOf(line.getHomosexual()));
            }
            if (line.getSexOrPNS() != 0) {
                writeSmallCell(firstRow.getCell(24), String.valueOf(line.getSexOrPNS()));
            }
            if (line.getYes() != 0) {
                writeSmallCell(firstRow.getCell(25), String.valueOf(line.getYes()));
            }
            if (line.getNo() != 0) {
                writeSmallCell(firstRow.getCell(26), String.valueOf(line.getNo()));
            }
            table9.addRow(firstRow);
        }
    }

    public void saveDoc() {

        FileOutputStream fos = null;
        try {
            fos = new FileOutputStream(outFile);
            doc.write(fos);
            fos.close();
            doc.close();
        } catch (IOException ex) {
            Logger.getLogger(DocHelper.class.getName()).log(Level.SEVERE, null, ex);
        }

    }

    private static void writeSmallCell(XWPFTableCell cell, String str) {
        XWPFRun run = cell.getParagraphs().get(0).createRun();
        run.setFontSize(8);
        run.setText(str);
    }

    /**
     * Generates the output file based on the template
     *
     * @param inFile
     * @param outFile
     */
    public static void genDocOut(File inFile, File outFile) {

        try {
            XWPFDocument docOut = new XWPFDocument(new FileInputStream(inFile));
            FileOutputStream fos = new FileOutputStream(outFile);
            docOut.write(fos);
            docOut.close();
            fos.close();
        } catch (FileNotFoundException ex) {
            Logger.getLogger(DocHelper.class.getName()).log(Level.SEVERE, null, ex);
        } catch (IOException ex) {
            Logger.getLogger(DocHelper.class.getName()).log(Level.SEVERE, null, ex);
        }
    }

    /**
     * Saves the document
     *
     * @param inFile
     * @param outFile
     */
    public void genDocFile() {
        try {
            outFile.createNewFile();
        } catch (IOException ex) {
            Logger.getLogger(DocHelper.class.getName()).log(Level.SEVERE, null, ex);
        }
    }

    /**
     * Replaces the text of a run within a cell
     *
     * @param table
     * @param m
     */
    private static void replaceTable(XWPFTable table, HashMap m) {
        for (XWPFTableRow row : table.getRows()) {
            for (XWPFTableCell cell : row.getTableCells()) {
                for (XWPFParagraph p : cell.getParagraphs()) {
                    replaceParagraphInCell(p, m);
                }
            }
        }
    }

    private static void replaceTableSmall(XWPFTable table, HashMap m) {
        for (XWPFTableRow row : table.getRows()) {
            for (XWPFTableCell cell : row.getTableCells()) {
                for (XWPFParagraph p : cell.getParagraphs()) {
                    replaceParagraphInCellSmall(p, m);
                }
            }
        }
    }

    /**
     * Replaces the text found within the runs of a paragraph
     *
     * @param paragraph
     * @param m
     */
    private static void replaceParagraphInCell(XWPFParagraph paragraph, HashMap m) {
        String text = paragraph.getText();
        int size = paragraph.getRuns().size();
        if (text != null && text.contains(m.get("a").toString())) {
            for (int i = 0; i < size; i++) {
                paragraph.removeRun(0);
            }
            text = text.replace(m.get("a").toString(), m.get("b").toString());
            XWPFRun run = paragraph.createRun();
            run.setText(text, 0);
        }
    }

    private static void replaceParagraphInCellSmall(XWPFParagraph paragraph, HashMap m) {
        String text = paragraph.getText();
        int size = paragraph.getRuns().size();
        if (text != null && text.contains(m.get("a").toString())) {
            for (int i = 0; i < size; i++) {
                paragraph.removeRun(0);
            }
            text = text.replace(m.get("a").toString(), m.get("b").toString());
            XWPFRun run = paragraph.createRun();
            run.setFontSize(8);
            run.setText(text, 0);
        }
    }

    private static void replaceParagraph(XWPFParagraph paragraph, HashMap m) {
        for (XWPFRun r : paragraph.getRuns()) {
            String text = r.getText(r.getTextPosition());
            if (text != null && text.contains(m.get("a").toString())) {
                text = text.replace(m.get("a").toString(), m.get("b").toString());
                r.setText(text, 0);
            }
        }
    }

    /**
     * Finds a given value within the doc and replaces it
     *
     * @param m
     * @param outFile
     */
    public static void findAndReplace(HashMap m, File outFile) {

        try {

            XWPFDocument xdoc = new XWPFDocument(new FileInputStream(outFile));
            for (XWPFParagraph p : xdoc.getParagraphs()) {
                replaceParagraph(p, m);
            }
        } catch (FileNotFoundException ex) {
            Logger.getLogger(DocHelper.class.getName()).log(Level.SEVERE, null, ex);
        } catch (IOException ex) {
            Logger.getLogger(DocHelper.class.getName()).log(Level.SEVERE, null, ex);
        }
    }

    private void writeT5Line(XWPFTable table, RefRecord record,int cursor) {

        
        
        String trust = record.getTrust();
        String gender = record.getGender();
        String grade = record.getGrade();
        String school = record.getSchool();
        List<String> addRefList = record.getAddRef();
        String addRef = "";
        for(int i=0; i<addRefList.size();i++){
            if(i==0){
                if(addRefList.size()>1){
                    addRef = addRefList.get(0)+"; ";
                }
                else{
                    addRef = addRefList.get(0);
                }
            }
            else{
                addRef = addRef + addRefList.get(i) + "; ";
            }
        }
        String country = record.getCountry();
        int age = record.getAge();
        String ethnicity = record.getEthnicity();
        String sexOr = record.getSexOr();
        String religion = record.getReligion();
        String disability = record.getDisability();

        int tempateRowFId = 1;
        XWPFTableRow rowTemplateF = table.getRow(tempateRowFId);
        int tempateRowMId = 2;
        XWPFTableRow rowTemplateM = table.getRow(tempateRowMId);
        
        if(cursor==0){
            if (gender.equals("F")) {
                
                XWPFTableRow oldRow = rowTemplateF;
                CTRow firstCTRow = null;
                try {
                    firstCTRow = CTRow.Factory.parse(oldRow.getCtRow().newInputStream());
                } catch (XmlException | IOException e) {
                    
                }
                
                XWPFTableRow firstRow = new XWPFTableRow(firstCTRow, table);
                XWPFTableCell cell0 = firstRow.getCell(0);
                writeSmallCell(cell0, trust);
                XWPFTableCell cell1 = firstRow.getCell(1);
                writeSmallCell(cell1, gender);
                XWPFTableCell cell2 = firstRow.getCell(2);
                writeSmallCell(cell2, grade);
                XWPFTableCell cell3 = firstRow.getCell(3);
                writeSmallCell(cell3, school);
                XWPFTableCell cell4 = firstRow.getCell(4);
                writeSmallCell(cell4, addRef);
                XWPFTableCell cell5 = firstRow.getCell(5);
                writeSmallCell(cell5, country);
                XWPFTableCell cell6 = firstRow.getCell(6);
                writeSmallCell(cell6, String.valueOf(age));
                XWPFTableCell cell7 = firstRow.getCell(7);
                writeSmallCell(cell7, ethnicity);
                XWPFTableCell cell8 = firstRow.getCell(8);
                writeSmallCell(cell8, sexOr);
                XWPFTableCell cell9 = firstRow.getCell(9);
                writeSmallCell(cell9, religion);
                XWPFTableCell cell10 = firstRow.getCell(10);
                writeSmallCell(cell10, disability);
                
                table.addRow(firstRow);
            } else if (gender.equals("M")) {
                
                XWPFTableRow oldRow = rowTemplateM;
                CTRow firstCTRow = null;
                try {
                    firstCTRow = CTRow.Factory.parse(oldRow.getCtRow().newInputStream());
                } catch (XmlException | IOException e) {
                    e.printStackTrace();
                }
                
                XWPFTableRow firstRow = new XWPFTableRow(firstCTRow, table);
                XWPFTableCell cell0 = firstRow.getCell(0);
                writeSmallCell(cell0, trust);
                XWPFTableCell cell1 = firstRow.getCell(1);
                writeSmallCell(cell1, gender);
                XWPFTableCell cell2 = firstRow.getCell(2);
                writeSmallCell(cell2, grade);
                XWPFTableCell cell3 = firstRow.getCell(3);
                writeSmallCell(cell3, school);
                XWPFTableCell cell4 = firstRow.getCell(4);
                writeSmallCell(cell4, addRef);
                XWPFTableCell cell5 = firstRow.getCell(5);
                writeSmallCell(cell5, country);
                XWPFTableCell cell6 = firstRow.getCell(6);
                writeSmallCell(cell6, String.valueOf(age));
                XWPFTableCell cell7 = firstRow.getCell(7);
                writeSmallCell(cell7, ethnicity);
                XWPFTableCell cell8 = firstRow.getCell(8);
                writeSmallCell(cell8, sexOr);
                XWPFTableCell cell9 = firstRow.getCell(9);
                writeSmallCell(cell9, religion);
                XWPFTableCell cell10 = firstRow.getCell(10);
                writeSmallCell(cell10, disability);
                
                table.addRow(firstRow);
            }
        }
        else{
            if (gender.equals("F")) {
                
                XWPFTableRow oldRow = rowTemplateF;
                CTRow firstCTRow = null;
                try {
                    firstCTRow = CTRow.Factory.parse(oldRow.getCtRow().newInputStream());
                } catch (XmlException | IOException e) {
                    
                }
                
                XWPFTableRow firstRow = new XWPFTableRow(firstCTRow, table);
                XWPFTableCell cell1 = firstRow.getCell(1);
                writeSmallCell(cell1, gender);
                XWPFTableCell cell2 = firstRow.getCell(2);
                writeSmallCell(cell2, grade);
                XWPFTableCell cell3 = firstRow.getCell(3);
                writeSmallCell(cell3, school);
                XWPFTableCell cell4 = firstRow.getCell(4);
                writeSmallCell(cell4, addRef);
                XWPFTableCell cell5 = firstRow.getCell(5);
                writeSmallCell(cell5, country);
                XWPFTableCell cell6 = firstRow.getCell(6);
                writeSmallCell(cell6, String.valueOf(age));
                XWPFTableCell cell7 = firstRow.getCell(7);
                writeSmallCell(cell7, ethnicity);
                XWPFTableCell cell8 = firstRow.getCell(8);
                writeSmallCell(cell8, sexOr);
                XWPFTableCell cell9 = firstRow.getCell(9);
                writeSmallCell(cell9, religion);
                XWPFTableCell cell10 = firstRow.getCell(10);
                writeSmallCell(cell10, disability);
                
                table.addRow(firstRow);
            } else if (gender.equals("M")) {
                
                XWPFTableRow oldRow = rowTemplateM;
                CTRow firstCTRow = null;
                try {
                    firstCTRow = CTRow.Factory.parse(oldRow.getCtRow().newInputStream());
                } catch (XmlException | IOException e) {
                    e.printStackTrace();
                }
                
                XWPFTableRow firstRow = new XWPFTableRow(firstCTRow, table);
                XWPFTableCell cell1 = firstRow.getCell(1);
                writeSmallCell(cell1, gender);
                XWPFTableCell cell2 = firstRow.getCell(2);
                writeSmallCell(cell2, grade);
                XWPFTableCell cell3 = firstRow.getCell(3);
                writeSmallCell(cell3, school);
                XWPFTableCell cell4 = firstRow.getCell(4);
                writeSmallCell(cell4, addRef);
                XWPFTableCell cell5 = firstRow.getCell(5);
                writeSmallCell(cell5, country);
                XWPFTableCell cell6 = firstRow.getCell(6);
                writeSmallCell(cell6, String.valueOf(age));
                XWPFTableCell cell7 = firstRow.getCell(7);
                writeSmallCell(cell7, ethnicity);
                XWPFTableCell cell8 = firstRow.getCell(8);
                writeSmallCell(cell8, sexOr);
                XWPFTableCell cell9 = firstRow.getCell(9);
                writeSmallCell(cell9, religion);
                XWPFTableCell cell10 = firstRow.getCell(10);
                writeSmallCell(cell10, disability);
                
                table.addRow(firstRow);
            }   
        }
    }

    /**
     * Generates the map which will be used to find and replace
     *
     * @param a
     * @param b
     * @return
     */
    public static HashMap<String, String> genMap(String a, String b) {
        HashMap<String, String> hash = new HashMap<>();
        hash.put("a", a);
        hash.put("b", b);
        return hash;

    }

    /**
     * Returns the actual year
     *
     * @return
     */
    public static int getStartingYear() {
        int year = Calendar.getInstance().get(Calendar.YEAR)-1;
        return year;
    }

    public static String genSplitYearSeq(String str) {
        List<String> strList = new ArrayList<>();

        for (String s : str.split("")) {
            strList.add(s);
        }

        String finalStr = strList.get(2) + strList.get(3);
        return finalStr;
    }
}
