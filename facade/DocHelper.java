/*
 * Decompiled with CFR 0.152.
 * 
 * Could not load the following classes:
 *  jdk.nashorn.internal.objects.NativeError
 */
package facade;

import facade.PoiHelper;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.HashMap;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;
import jdk.nashorn.internal.objects.NativeError;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.xmlbeans.XmlException;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRow;
import vo.ReferralRecord;
import vo.Table9Line;

public class DocHelper {
    private InputStream template = Thread.currentThread().getContextClassLoader().getResourceAsStream("QGReportSample.docx");
    private final File doc;
    private final File psw;
    private final File graphs;
    PoiHelper helper;
    int y = DocHelper.getStartingYear();
    int ys1 = DocHelper.getStartingYear() - 1;
    int ys2 = DocHelper.getStartingYear() - 2;
    int ys3 = DocHelper.getStartingYear() - 3;
    int yp1 = DocHelper.getStartingYear() + 1;
    String y2 = DocHelper.genSplitYearSeq(String.valueOf(this.yp1));
    int referredCount;
    int wessexCount;
    XWPFDocument xdoc;

    public DocHelper(File psw, File graphs, File doc, ArrayList<ReferralRecord> records) {
        this.doc = doc;
        this.psw = psw;
        this.graphs = graphs;
        this.helper = new PoiHelper(psw, graphs, records);
        this.referredCount = this.helper.countTotalReferrals();
        this.loadTemplate();
    }

    private void loadTemplate() {
        try {
            this.xdoc = new XWPFDocument(this.template);
        }
        catch (IOException iOException) {
            // empty catch block
        }
    }

    public void doTextUpdate() {
        for (XWPFParagraph p : this.xdoc.getParagraphs()) {
            DocHelper.replaceParagraph(p, DocHelper.genMap("YR2", this.y2));
            DocHelper.replaceParagraph(p, DocHelper.genMap("YR", String.valueOf(this.y)));
            DocHelper.replaceParagraph(p, DocHelper.genMap("YS1", String.valueOf(this.ys1)));
            DocHelper.replaceParagraph(p, DocHelper.genMap("YS2", String.valueOf(this.ys2)));
            DocHelper.replaceParagraph(p, DocHelper.genMap("YS3", String.valueOf(this.ys3)));
            DocHelper.replaceParagraph(p, DocHelper.genMap("YP1", String.valueOf(this.yp1)));
            DocHelper.replaceParagraph(p, DocHelper.genMap("TCR", String.valueOf(this.helper.countTotalReferrals())));
        }
        for (XWPFTable t : this.xdoc.getTables()) {
            DocHelper.replaceTable(t, DocHelper.genMap("YR", String.valueOf(this.y)));
            DocHelper.replaceTable(t, DocHelper.genMap("YR2", String.valueOf(this.y2)));
            DocHelper.replaceTable(t, DocHelper.genMap("YS1", String.valueOf(this.ys1)));
            DocHelper.replaceTable(t, DocHelper.genMap("YS2", String.valueOf(this.ys2)));
            DocHelper.replaceTable(t, DocHelper.genMap("YS3", String.valueOf(this.ys3)));
            DocHelper.replaceTable(t, DocHelper.genMap("YP1", String.valueOf(this.yp1)));
            DocHelper.replaceTable(t, DocHelper.genMap("TCR", String.valueOf(this.helper.countTotalReferrals())));
        }
    }

    public void getTable0() {
        List<Integer> t0list = this.helper.countTable0();
        int stCount = t0list.get(0);
        int fCount = t0list.get(1);
        int gpCount = t0list.get(2);
        int otherCount = t0list.get(3);
        int total = t0list.get(4);
        int casesClosed = t0list.get(5);
        int casesOClosed = t0list.get(6);
        XWPFTable table0 = this.xdoc.getTables().get(0);
        DocHelper.replaceTable(table0, DocHelper.genMap("TA", Integer.toString(stCount)));
        DocHelper.replaceTable(table0, DocHelper.genMap("TB", Integer.toString(fCount)));
        DocHelper.replaceTable(table0, DocHelper.genMap("TC", Integer.toString(gpCount)));
        DocHelper.replaceTable(table0, DocHelper.genMap("TD", Integer.toString(otherCount)));
        DocHelper.replaceTable(table0, DocHelper.genMap("TE", Integer.toString(total)));
        DocHelper.replaceTable(table0, DocHelper.genMap("TF", Integer.toString(casesClosed)));
        DocHelper.replaceTable(table0, DocHelper.genMap("TG", Integer.toString(casesOClosed)));
    }

    public void getTable1() {
        List<Integer> list = this.helper.countTable1();
        XWPFTable table1 = this.xdoc.getTables().get(1);
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
        DocHelper.writeSmallCell(table1.getRow(2).getCell(4), String.valueOf(anxietyMCount));
        DocHelper.writeSmallCell(table1.getRow(3).getCell(4), String.valueOf(anxietyFCount));
        DocHelper.writeSmallCell(table1.getRow(4).getCell(4), String.valueOf(capabilityMCount));
        DocHelper.writeSmallCell(table1.getRow(5).getCell(4), String.valueOf(capabilityFCount));
        DocHelper.writeSmallCell(table1.getRow(6).getCell(4), String.valueOf(carreerMCount));
        DocHelper.writeSmallCell(table1.getRow(7).getCell(4), String.valueOf(carreerFCount));
        DocHelper.writeSmallCell(table1.getRow(8).getCell(4), String.valueOf(clinicalMCount));
        DocHelper.writeSmallCell(table1.getRow(9).getCell(4), String.valueOf(clinicalFCount));
        DocHelper.writeSmallCell(table1.getRow(10).getCell(4), String.valueOf(communicationMCount));
        DocHelper.writeSmallCell(table1.getRow(11).getCell(4), String.valueOf(communicationFCount));
        DocHelper.writeSmallCell(table1.getRow(12).getCell(4), String.valueOf(conductMCount));
        DocHelper.writeSmallCell(table1.getRow(13).getCell(4), String.valueOf(conductFCount));
        DocHelper.writeSmallCell(table1.getRow(14).getCell(4), String.valueOf(culturalMCount));
        DocHelper.writeSmallCell(table1.getRow(15).getCell(4), String.valueOf(culturalFCount));
        DocHelper.writeSmallCell(table1.getRow(16).getCell(4), String.valueOf(examMCount));
        DocHelper.writeSmallCell(table1.getRow(17).getCell(4), String.valueOf(examFCount));
        DocHelper.writeSmallCell(table1.getRow(18).getCell(4), String.valueOf(menHealthMCount));
        DocHelper.writeSmallCell(table1.getRow(19).getCell(4), String.valueOf(menHealthFCount));
        DocHelper.writeSmallCell(table1.getRow(20).getCell(4), String.valueOf(phHealthMCount));
        DocHelper.writeSmallCell(table1.getRow(21).getCell(4), String.valueOf(phHealthFCount));
        DocHelper.writeSmallCell(table1.getRow(22).getCell(4), String.valueOf(languageMCount));
        DocHelper.writeSmallCell(table1.getRow(23).getCell(4), String.valueOf(languageFCount));
        DocHelper.writeSmallCell(table1.getRow(24).getCell(4), String.valueOf(profMCount));
        DocHelper.writeSmallCell(table1.getRow(25).getCell(4), String.valueOf(profFCount));
        DocHelper.writeSmallCell(table1.getRow(26).getCell(4), String.valueOf(adhdMCount));
        DocHelper.writeSmallCell(table1.getRow(27).getCell(4), String.valueOf(adhdFCount));
        DocHelper.writeSmallCell(table1.getRow(28).getCell(4), String.valueOf(asdMCount));
        DocHelper.writeSmallCell(table1.getRow(29).getCell(4), String.valueOf(asdFCount));
        DocHelper.writeSmallCell(table1.getRow(30).getCell(4), String.valueOf(dyslexiaMCount));
        DocHelper.writeSmallCell(table1.getRow(31).getCell(4), String.valueOf(dyslexiaFCount));
        DocHelper.writeSmallCell(table1.getRow(32).getCell(4), String.valueOf(dyspraxiaMCount));
        DocHelper.writeSmallCell(table1.getRow(33).getCell(4), String.valueOf(dyspraxiaFCount));
        DocHelper.writeSmallCell(table1.getRow(34).getCell(4), String.valueOf(srttMCount));
        DocHelper.writeSmallCell(table1.getRow(35).getCell(4), String.valueOf(srttFCount));
        DocHelper.writeSmallCell(table1.getRow(36).getCell(4), String.valueOf(teamMCount));
        DocHelper.writeSmallCell(table1.getRow(37).getCell(4), String.valueOf(teamFCount));
        DocHelper.writeSmallCell(table1.getRow(38).getCell(4), String.valueOf(timeMCount));
        DocHelper.writeSmallCell(table1.getRow(39).getCell(4), String.valueOf(timeFCount));
        DocHelper.writeSmallCell(table1.getRow(40).getCell(4), String.valueOf(otherMCount));
        DocHelper.writeSmallCell(table1.getRow(41).getCell(4), String.valueOf(otherFCount));
    }

    public void getTable2() {
        List<Integer> t2list = this.helper.countTable2();
        double f1TotalCount = t2list.get(1).intValue();
        double f2TotalCount = t2list.get(2).intValue();
        double f1RefCount = t2list.get(3).intValue();
        double f2RefCount = t2list.get(4).intValue();
        double totalFCount = f1TotalCount + f2TotalCount;
        double totalFRef = f1RefCount + f2RefCount;
        double f1Perc = f1RefCount / f1TotalCount * 100.0;
        double f2Perc = f2RefCount / f2TotalCount * 100.0;
        double fPerc = totalFRef / totalFCount * 100.0;
        double fRefPerc = totalFRef / (double)this.referredCount * 100.0;
        DecimalFormat format = new DecimalFormat("#.#");
        XWPFTable table2 = this.xdoc.getTables().get(2);
        DocHelper.replaceTable(table2, DocHelper.genMap("TB", String.valueOf((int)totalFRef) + " (" + format.format(fPerc) + "%)"));
        DocHelper.replaceTable(table2, DocHelper.genMap("TC", "F1 = " + (int)f1RefCount + " F2 = " + (int)f2RefCount));
        DocHelper.replaceTable(table2, DocHelper.genMap("TD", String.valueOf(format.format(fRefPerc) + "%")));
        DocHelper.replaceTable(table2, DocHelper.genMap("TE", "F1 = " + (int)f1TotalCount + " F2 = " + (int)f2TotalCount));
        DocHelper.replaceTable(table2, DocHelper.genMap("TF", String.valueOf(format.format(f1Perc) + "%")));
        DocHelper.replaceTable(table2, DocHelper.genMap("TG", String.valueOf(format.format(f2Perc) + "%")));
    }

    public void getTable3() {
        List<Integer> t3list = this.helper.countTable3();
        double bournemouthTotal = t3list.get(0).intValue();
        double dorsetCountyTotal = t3list.get(1).intValue();
        double dorsetHealthTotal = t3list.get(2).intValue();
        double hhftTotal = t3list.get(3).intValue();
        double iowTotal = t3list.get(4).intValue();
        double jerseyTotal = t3list.get(5).intValue();
        double pooleTotal = t3list.get(6).intValue();
        double portsmouthTotal = t3list.get(7).intValue();
        double salisburyTotal = t3list.get(8).intValue();
        double solentTotal = t3list.get(9).intValue();
        double southamptonTotal = t3list.get(10).intValue();
        double southernTotal = t3list.get(11).intValue();
        double bournemouthRefNo = t3list.get(12).intValue();
        double dorsetCountyRefNo = t3list.get(13).intValue();
        double dorsetHealthRefNo = t3list.get(14).intValue();
        double hhftRefNo = t3list.get(15).intValue();
        double iowRefNo = t3list.get(16).intValue();
        double jerseyRefNo = t3list.get(17).intValue();
        double pooleRefNo = t3list.get(18).intValue();
        double portsmouthRefNo = t3list.get(19).intValue();
        double salisburyRefNo = t3list.get(20).intValue();
        double solentRefNo = t3list.get(21).intValue();
        double southamptonRefNo = t3list.get(22).intValue();
        double southernRefNo = t3list.get(23).intValue();
        this.wessexCount = t3list.get(24);
        double bournemouthTrInTrust = Math.round(bournemouthRefNo / bournemouthTotal * 100.0);
        double dorsetCountyTrInTrust = Math.round(dorsetCountyRefNo / dorsetCountyTotal * 100.0);
        double dorsetHealthTrInTrust = Math.round(dorsetHealthRefNo / dorsetHealthTotal * 100.0);
        double hhtfTrInTrust = Math.round(hhftRefNo / hhftTotal * 100.0);
        double iowTrInTrust = Math.round(iowRefNo / iowTotal * 100.0);
        double jerseyTrInTrust = Math.round(jerseyRefNo / jerseyTotal * 100.0);
        double pooleTrInTrust = Math.round(pooleRefNo / pooleTotal * 100.0);
        double portsmouthTrInTrust = Math.round(portsmouthRefNo / portsmouthTotal * 100.0);
        double salisburyTrInTrust = Math.round(salisburyRefNo / salisburyTotal * 100.0);
        double solentTrInTrust = Math.round(solentRefNo / solentTotal * 100.0);
        double southamptonTrInTrust = Math.round(southamptonRefNo / southamptonTotal * 100.0);
        double southernTrInTrust = Math.round(southernRefNo / southernTotal * 100.0);
        double bournemouthOfWssx = Math.round(bournemouthTotal / (double)this.wessexCount * 100.0);
        double dorsetCountyOfWssx = Math.round(dorsetCountyTotal / (double)this.wessexCount * 100.0);
        double dorsetHealthOfWssx = Math.round(dorsetHealthTotal / (double)this.wessexCount * 100.0);
        double hhtfOfWssx = Math.round(hhftTotal / (double)this.wessexCount * 100.0);
        double iowOfWssx = Math.round(iowTotal / (double)this.wessexCount * 100.0);
        double jerseyOfWssx = Math.round(jerseyTotal / (double)this.wessexCount * 100.0);
        double pooleOfWssx = Math.round(pooleTotal / (double)this.wessexCount * 100.0);
        double portsmouthOfWssx = Math.round(portsmouthTotal / (double)this.wessexCount * 100.0);
        double salisburyOfWssx = Math.round(salisburyTotal / (double)this.wessexCount * 100.0);
        double solentOfWssx = Math.round(solentTotal / (double)this.wessexCount * 100.0);
        double southamptonOfWssx = Math.round(southamptonTotal / (double)this.wessexCount * 100.0);
        double southernOfWssx = Math.round(southernTotal / (double)this.wessexCount * 100.0);
        double bournemouthOfPSW = Math.round(bournemouthRefNo / (double)this.referredCount * 100.0);
        double dorsetCountyOfPSW = Math.round(dorsetCountyRefNo / (double)this.referredCount * 100.0);
        double dorsetHealthOfPSW = Math.round(dorsetHealthRefNo / (double)this.referredCount * 100.0);
        double hhtfOfPSW = Math.round(hhftRefNo / (double)this.referredCount * 100.0);
        double iowOfPSW = Math.round(iowRefNo / (double)this.referredCount * 100.0);
        double jerseyOfPSW = Math.round(jerseyRefNo / (double)this.referredCount * 100.0);
        double pooleOfPSW = Math.round(pooleRefNo / (double)this.referredCount * 100.0);
        double portsmouthOfPSW = Math.round(portsmouthRefNo / (double)this.referredCount * 100.0);
        double salisburyOfPSW = Math.round(salisburyRefNo / (double)this.referredCount * 100.0);
        double solentOfPSW = Math.round(solentRefNo / (double)this.referredCount * 100.0);
        double southamptonOfPSW = Math.round(southamptonRefNo / (double)this.referredCount * 100.0);
        double southernOfPSW = Math.round(southernRefNo / (double)this.referredCount * 100.0);
        XWPFTable table3 = this.xdoc.getTables().get(3);
        DocHelper.replaceTable(table3, DocHelper.genMap("TA", String.valueOf((int)bournemouthRefNo)));
        DocHelper.replaceTable(table3, DocHelper.genMap("TB", String.valueOf((int)dorsetCountyRefNo)));
        DocHelper.replaceTable(table3, DocHelper.genMap("TC", String.valueOf((int)dorsetHealthRefNo)));
        DocHelper.replaceTable(table3, DocHelper.genMap("TD", String.valueOf((int)hhftRefNo)));
        DocHelper.replaceTable(table3, DocHelper.genMap("TE", String.valueOf((int)iowRefNo)));
        DocHelper.replaceTable(table3, DocHelper.genMap("TF", String.valueOf((int)jerseyRefNo)));
        DocHelper.replaceTable(table3, DocHelper.genMap("TG", String.valueOf((int)pooleRefNo)));
        DocHelper.replaceTable(table3, DocHelper.genMap("TH", String.valueOf((int)portsmouthRefNo)));
        DocHelper.replaceTable(table3, DocHelper.genMap("TI", String.valueOf((int)salisburyRefNo)));
        DocHelper.replaceTable(table3, DocHelper.genMap("TJ", String.valueOf((int)solentRefNo)));
        DocHelper.replaceTable(table3, DocHelper.genMap("TK", String.valueOf((int)southamptonRefNo)));
        DocHelper.replaceTable(table3, DocHelper.genMap("TL", String.valueOf((int)southernRefNo)));
        DocHelper.replaceTable(table3, DocHelper.genMap("TM", String.valueOf((int)bournemouthTotal)));
        DocHelper.replaceTable(table3, DocHelper.genMap("TN", String.valueOf((int)dorsetCountyTotal)));
        DocHelper.replaceTable(table3, DocHelper.genMap("TO", String.valueOf((int)dorsetHealthTotal)));
        DocHelper.replaceTable(table3, DocHelper.genMap("TP", String.valueOf((int)hhftTotal)));
        DocHelper.replaceTable(table3, DocHelper.genMap("TQ", String.valueOf((int)iowTotal)));
        DocHelper.replaceTable(table3, DocHelper.genMap("TR", String.valueOf((int)jerseyTotal)));
        DocHelper.replaceTable(table3, DocHelper.genMap("TS", String.valueOf((int)pooleTotal)));
        DocHelper.replaceTable(table3, DocHelper.genMap("TT", String.valueOf((int)portsmouthTotal)));
        DocHelper.replaceTable(table3, DocHelper.genMap("TU", String.valueOf((int)salisburyTotal)));
        DocHelper.replaceTable(table3, DocHelper.genMap("TV", String.valueOf((int)solentTotal)));
        DocHelper.replaceTable(table3, DocHelper.genMap("TW", String.valueOf((int)southamptonTotal)));
        DocHelper.replaceTable(table3, DocHelper.genMap("TX", String.valueOf((int)southernTotal)));
        DocHelper.replaceTable(table3, DocHelper.genMap("TY", String.valueOf((int)bournemouthTrInTrust + "%")));
        DocHelper.replaceTable(table3, DocHelper.genMap("TZ", String.valueOf((int)dorsetCountyTrInTrust + "%")));
        DocHelper.replaceTable(table3, DocHelper.genMap("T1A", String.valueOf((int)dorsetHealthTrInTrust + "%")));
        DocHelper.replaceTable(table3, DocHelper.genMap("T1B", String.valueOf((int)hhtfTrInTrust + "%")));
        DocHelper.replaceTable(table3, DocHelper.genMap("T1C", String.valueOf((int)iowTrInTrust + "%")));
        DocHelper.replaceTable(table3, DocHelper.genMap("T1D", String.valueOf((int)jerseyTrInTrust + "%")));
        DocHelper.replaceTable(table3, DocHelper.genMap("T1E", String.valueOf((int)pooleTrInTrust + "%")));
        DocHelper.replaceTable(table3, DocHelper.genMap("T1F", String.valueOf((int)portsmouthTrInTrust + "%")));
        DocHelper.replaceTable(table3, DocHelper.genMap("T1G", String.valueOf((int)salisburyTrInTrust + "%")));
        DocHelper.replaceTable(table3, DocHelper.genMap("T1H", String.valueOf((int)solentTrInTrust + "%")));
        DocHelper.replaceTable(table3, DocHelper.genMap("T1I", String.valueOf((int)southamptonTrInTrust + "%")));
        DocHelper.replaceTable(table3, DocHelper.genMap("T1J", String.valueOf((int)southernTrInTrust + "%")));
        DocHelper.replaceTable(table3, DocHelper.genMap("T1K", String.valueOf((int)bournemouthOfWssx + "%")));
        DocHelper.replaceTable(table3, DocHelper.genMap("T1L", String.valueOf((int)dorsetCountyOfWssx + "%")));
        DocHelper.replaceTable(table3, DocHelper.genMap("T1M", String.valueOf((int)dorsetHealthOfWssx + "%")));
        DocHelper.replaceTable(table3, DocHelper.genMap("T1N", String.valueOf((int)hhtfOfWssx + "%")));
        DocHelper.replaceTable(table3, DocHelper.genMap("T1O", String.valueOf((int)iowOfWssx + "%")));
        DocHelper.replaceTable(table3, DocHelper.genMap("T1P", String.valueOf((int)jerseyOfWssx + "%")));
        DocHelper.replaceTable(table3, DocHelper.genMap("T1Q", String.valueOf((int)pooleOfWssx + "%")));
        DocHelper.replaceTable(table3, DocHelper.genMap("T1R", String.valueOf((int)portsmouthOfWssx + "%")));
        DocHelper.replaceTable(table3, DocHelper.genMap("T1S", String.valueOf((int)salisburyOfWssx + "%")));
        DocHelper.replaceTable(table3, DocHelper.genMap("T1T", String.valueOf((int)solentOfWssx + "%")));
        DocHelper.replaceTable(table3, DocHelper.genMap("T1U", String.valueOf((int)southamptonOfWssx + "%")));
        DocHelper.replaceTable(table3, DocHelper.genMap("T1V", String.valueOf((int)southernOfWssx + "%")));
        DocHelper.replaceTable(table3, DocHelper.genMap("T1W", String.valueOf((int)bournemouthOfPSW + "%")));
        DocHelper.replaceTable(table3, DocHelper.genMap("T1X", String.valueOf((int)dorsetCountyOfPSW + "%")));
        DocHelper.replaceTable(table3, DocHelper.genMap("T1Y", String.valueOf((int)dorsetHealthOfPSW + "%")));
        DocHelper.replaceTable(table3, DocHelper.genMap("T1Z", String.valueOf((int)hhtfOfPSW + "%")));
        DocHelper.replaceTable(table3, DocHelper.genMap("T2A", String.valueOf((int)iowOfPSW + "%")));
        DocHelper.replaceTable(table3, DocHelper.genMap("T2B", String.valueOf((int)jerseyOfPSW + "%")));
        DocHelper.replaceTable(table3, DocHelper.genMap("T2C", String.valueOf((int)pooleOfPSW + "%")));
        DocHelper.replaceTable(table3, DocHelper.genMap("T2D", String.valueOf((int)portsmouthOfPSW + "%")));
        DocHelper.replaceTable(table3, DocHelper.genMap("T2E", String.valueOf((int)salisburyOfPSW + "%")));
        DocHelper.replaceTable(table3, DocHelper.genMap("T2F", String.valueOf((int)solentOfPSW + "%")));
        DocHelper.replaceTable(table3, DocHelper.genMap("T2G", String.valueOf((int)southamptonOfPSW + "%")));
        DocHelper.replaceTable(table3, DocHelper.genMap("T2H", String.valueOf((int)southernOfPSW + "%")));
    }

    public void getTable4() {
        List<Integer> t4list = this.helper.countTable4();
        double anaestheticsRefNo = t4list.get(0).intValue();
        double dentalRefNo = t4list.get(1).intValue();
        double emergRefNo = t4list.get(2).intValue();
        double foundationRefNo = t4list.get(3).intValue();
        double gpRefNo = t4list.get(4).intValue();
        double medicineRefNo = t4list.get(5).intValue();
        double obsRefNo = t4list.get(6).intValue();
        double occhealthRefNo = t4list.get(7).intValue();
        double paediatricsRefNo = t4list.get(8).intValue();
        double pathologyRefNo = t4list.get(9).intValue();
        double pharmacyRefNo = t4list.get(10).intValue();
        double psychRefNo = t4list.get(11).intValue();
        double pubhealthRefNo = t4list.get(12).intValue();
        double radioRefNo = t4list.get(13).intValue();
        double surgeryRefNo = t4list.get(14).intValue();
        double anaestheticsTotal = t4list.get(15).intValue();
        double dentalTotal = t4list.get(16).intValue();
        double emergTotal = t4list.get(17).intValue();
        double foundationTotal = t4list.get(18).intValue();
        double gpTotal = t4list.get(19).intValue();
        double medicineTotal = t4list.get(20).intValue();
        double obsTotal = t4list.get(21).intValue();
        double occhealthTotal = t4list.get(22).intValue();
        double paediatricsTotal = t4list.get(23).intValue();
        double pathologyTotal = t4list.get(24).intValue();
        double pharmacyTotal = t4list.get(25).intValue();
        double psychTotal = t4list.get(26).intValue();
        double pubhealthTotal = t4list.get(27).intValue();
        double radioTotal = t4list.get(28).intValue();
        double surgeryTotal = t4list.get(29).intValue();
        double anaestheticsInSpc = Math.round(anaestheticsRefNo / anaestheticsTotal * 100.0);
        double dentalInSpc = Math.round(dentalRefNo / dentalTotal * 100.0);
        double emergInSpc = Math.round(emergRefNo / emergTotal * 100.0);
        double foundationInSpc = Math.round(foundationRefNo / foundationTotal * 100.0);
        double gpInSpc = Math.round(gpRefNo / gpTotal * 100.0);
        double medicineInSpc = Math.round(medicineRefNo / medicineTotal * 100.0);
        double obsInSpc = Math.round(obsRefNo / obsTotal * 100.0);
        double occhealthInSpc = Math.round(occhealthRefNo / occhealthTotal * 100.0);
        double paediatricsInSpc = Math.round(paediatricsRefNo / paediatricsTotal * 100.0);
        double pathologyInSpc = Math.round(pathologyRefNo / pathologyTotal * 100.0);
        double pharmacyInSpc = Math.round(pharmacyRefNo / pharmacyTotal * 100.0);
        double psychInSpc = Math.round(psychRefNo / psychTotal * 100.0);
        double pubhealthInSpc = Math.round(pubhealthRefNo / pubhealthTotal * 100.0);
        double radioInSpc = Math.round(radioRefNo / radioTotal * 100.0);
        double surgeryInSpc = Math.round(surgeryRefNo / surgeryTotal * 100.0);
        double anaestheticsOfWssx = Math.round(anaestheticsTotal / (double)this.wessexCount * 100.0);
        double dentalOfWssx = Math.round(dentalTotal / (double)this.wessexCount * 100.0);
        double emergOfWssx = Math.round(emergTotal / (double)this.wessexCount * 100.0);
        double foundationOfWssx = Math.round(foundationTotal / (double)this.wessexCount * 100.0);
        double gpOfWssx = Math.round(gpTotal / (double)this.wessexCount * 100.0);
        double medicineOfWssx = Math.round(medicineTotal / (double)this.wessexCount * 100.0);
        double obsOfWssx = Math.round(obsTotal / (double)this.wessexCount * 100.0);
        double occhealthOfWssx = Math.round(occhealthTotal / (double)this.wessexCount * 100.0);
        double paediatricsOfWssx = Math.round(paediatricsTotal / (double)this.wessexCount * 100.0);
        double pathologyOfWssx = Math.round(pathologyTotal / (double)this.wessexCount * 100.0);
        double pharmacyOfWssx = Math.round(pharmacyTotal / (double)this.wessexCount * 100.0);
        double psychOfWssx = Math.round(psychTotal / (double)this.wessexCount * 100.0);
        double pubhealthOfWssx = Math.round(pubhealthTotal / (double)this.wessexCount * 100.0);
        double radioOfWssx = Math.round(radioTotal / (double)this.wessexCount * 100.0);
        double surgeryOfWssx = Math.round(surgeryTotal / (double)this.wessexCount * 100.0);
        double anaestheticsOfPSW = Math.round(anaestheticsRefNo / (double)this.referredCount * 100.0);
        double dentalOfPSW = Math.round(dentalRefNo / (double)this.referredCount * 100.0);
        double emergOfPSW = Math.round(emergRefNo / (double)this.referredCount * 100.0);
        double foundationOfPSW = Math.round(foundationRefNo / (double)this.referredCount * 100.0);
        double gpOfPSW = Math.round(gpRefNo / (double)this.referredCount * 100.0);
        double medicineOfPSW = Math.round(medicineRefNo / (double)this.referredCount * 100.0);
        double obsOfPSW = Math.round(obsRefNo / (double)this.referredCount * 100.0);
        double occhealthOfPSW = Math.round(occhealthRefNo / (double)this.referredCount * 100.0);
        double paediatricsOfPSW = Math.round(paediatricsRefNo / (double)this.referredCount * 100.0);
        double pathologyOfPSW = Math.round(pathologyRefNo / (double)this.referredCount * 100.0);
        double pharmacyOfPSW = Math.round(pharmacyRefNo / (double)this.referredCount * 100.0);
        double psychOfPSW = Math.round(psychRefNo / (double)this.referredCount * 100.0);
        double pubhealthOfPSW = Math.round(pubhealthRefNo / (double)this.referredCount * 100.0);
        double radioOfPSW = Math.round(radioRefNo / (double)this.referredCount * 100.0);
        double surgeryOfPSW = Math.round(surgeryRefNo / (double)this.referredCount * 100.0);
        XWPFTable table4 = this.xdoc.getTables().get(4);
        DocHelper.replaceTable(table4, DocHelper.genMap("TA", String.valueOf((int)anaestheticsRefNo)));
        DocHelper.replaceTable(table4, DocHelper.genMap("TB", String.valueOf((int)dentalRefNo)));
        DocHelper.replaceTable(table4, DocHelper.genMap("TC", String.valueOf((int)emergRefNo)));
        DocHelper.replaceTable(table4, DocHelper.genMap("TD", String.valueOf((int)foundationRefNo)));
        DocHelper.replaceTable(table4, DocHelper.genMap("TE", String.valueOf((int)gpRefNo)));
        DocHelper.replaceTable(table4, DocHelper.genMap("TF", String.valueOf((int)medicineRefNo)));
        DocHelper.replaceTable(table4, DocHelper.genMap("TG", String.valueOf((int)obsRefNo)));
        DocHelper.replaceTable(table4, DocHelper.genMap("TH", String.valueOf((int)occhealthRefNo)));
        DocHelper.replaceTable(table4, DocHelper.genMap("TI", String.valueOf((int)paediatricsRefNo)));
        DocHelper.replaceTable(table4, DocHelper.genMap("TJ", String.valueOf((int)pathologyRefNo)));
        DocHelper.replaceTable(table4, DocHelper.genMap("TK", String.valueOf((int)pharmacyRefNo)));
        DocHelper.replaceTable(table4, DocHelper.genMap("TL", String.valueOf((int)psychRefNo)));
        DocHelper.replaceTable(table4, DocHelper.genMap("TM", String.valueOf((int)pubhealthRefNo)));
        DocHelper.replaceTable(table4, DocHelper.genMap("TN", String.valueOf((int)radioRefNo)));
        DocHelper.replaceTable(table4, DocHelper.genMap("TO", String.valueOf((int)surgeryRefNo)));
        DocHelper.replaceTable(table4, DocHelper.genMap("TP", String.valueOf((int)anaestheticsTotal)));
        DocHelper.replaceTable(table4, DocHelper.genMap("TQ", String.valueOf((int)dentalTotal)));
        DocHelper.replaceTable(table4, DocHelper.genMap("TR", String.valueOf((int)emergTotal)));
        DocHelper.replaceTable(table4, DocHelper.genMap("TS", String.valueOf((int)foundationTotal)));
        DocHelper.replaceTable(table4, DocHelper.genMap("TT", String.valueOf((int)gpTotal)));
        DocHelper.replaceTable(table4, DocHelper.genMap("TU", String.valueOf((int)medicineTotal)));
        DocHelper.replaceTable(table4, DocHelper.genMap("TV", String.valueOf((int)obsTotal)));
        DocHelper.replaceTable(table4, DocHelper.genMap("TW", String.valueOf((int)occhealthTotal)));
        DocHelper.replaceTable(table4, DocHelper.genMap("TX", String.valueOf((int)paediatricsTotal)));
        DocHelper.replaceTable(table4, DocHelper.genMap("TY", String.valueOf((int)pathologyTotal)));
        DocHelper.replaceTable(table4, DocHelper.genMap("TZ", String.valueOf((int)pharmacyTotal)));
        DocHelper.replaceTable(table4, DocHelper.genMap("T1A", String.valueOf((int)psychTotal)));
        DocHelper.replaceTable(table4, DocHelper.genMap("T1B", String.valueOf((int)pubhealthTotal)));
        DocHelper.replaceTable(table4, DocHelper.genMap("T1C", String.valueOf((int)radioTotal)));
        DocHelper.replaceTable(table4, DocHelper.genMap("T1D", String.valueOf((int)surgeryTotal)));
        DocHelper.replaceTable(table4, DocHelper.genMap("T1E", String.valueOf((int)anaestheticsInSpc) + "%"));
        DocHelper.replaceTable(table4, DocHelper.genMap("T1F", String.valueOf((int)dentalInSpc) + "%"));
        DocHelper.replaceTable(table4, DocHelper.genMap("T1G", String.valueOf((int)emergInSpc) + "%"));
        DocHelper.replaceTable(table4, DocHelper.genMap("T1H", String.valueOf((int)foundationInSpc) + "%"));
        DocHelper.replaceTable(table4, DocHelper.genMap("T1I", String.valueOf((int)gpInSpc) + "%"));
        DocHelper.replaceTable(table4, DocHelper.genMap("T1J", String.valueOf((int)medicineInSpc) + "%"));
        DocHelper.replaceTable(table4, DocHelper.genMap("T1K", String.valueOf((int)obsInSpc) + "%"));
        DocHelper.replaceTable(table4, DocHelper.genMap("T1L", String.valueOf((int)occhealthInSpc) + "%"));
        DocHelper.replaceTable(table4, DocHelper.genMap("T1M", String.valueOf((int)paediatricsInSpc) + "%"));
        DocHelper.replaceTable(table4, DocHelper.genMap("T1N", String.valueOf((int)pathologyInSpc) + "%"));
        DocHelper.replaceTable(table4, DocHelper.genMap("T1O", String.valueOf((int)pharmacyInSpc) + "%"));
        DocHelper.replaceTable(table4, DocHelper.genMap("T1P", String.valueOf((int)psychInSpc) + "%"));
        DocHelper.replaceTable(table4, DocHelper.genMap("T1Q", String.valueOf((int)pubhealthInSpc) + "%"));
        DocHelper.replaceTable(table4, DocHelper.genMap("T1R", String.valueOf((int)radioInSpc) + "%"));
        DocHelper.replaceTable(table4, DocHelper.genMap("T1S", String.valueOf((int)surgeryInSpc) + "%"));
        DocHelper.replaceTable(table4, DocHelper.genMap("T1T", String.valueOf((int)anaestheticsOfWssx) + "%"));
        DocHelper.replaceTable(table4, DocHelper.genMap("T1U", String.valueOf((int)dentalOfWssx) + "%"));
        DocHelper.replaceTable(table4, DocHelper.genMap("T1V", String.valueOf((int)emergOfWssx) + "%"));
        DocHelper.replaceTable(table4, DocHelper.genMap("T1W", String.valueOf((int)foundationOfWssx) + "%"));
        DocHelper.replaceTable(table4, DocHelper.genMap("T1X", String.valueOf((int)gpOfWssx) + "%"));
        DocHelper.replaceTable(table4, DocHelper.genMap("T1Y", String.valueOf((int)medicineOfWssx) + "%"));
        DocHelper.replaceTable(table4, DocHelper.genMap("T1Z", String.valueOf((int)obsOfWssx) + "%"));
        DocHelper.replaceTable(table4, DocHelper.genMap("T2A", String.valueOf((int)occhealthOfWssx) + "%"));
        DocHelper.replaceTable(table4, DocHelper.genMap("T2B", String.valueOf((int)paediatricsOfWssx) + "%"));
        DocHelper.replaceTable(table4, DocHelper.genMap("T2C", String.valueOf((int)pathologyOfWssx) + "%"));
        DocHelper.replaceTable(table4, DocHelper.genMap("T2D", String.valueOf((int)pharmacyOfWssx) + "%"));
        DocHelper.replaceTable(table4, DocHelper.genMap("T2E", String.valueOf((int)psychOfWssx) + "%"));
        DocHelper.replaceTable(table4, DocHelper.genMap("T2F", String.valueOf((int)pubhealthOfWssx) + "%"));
        DocHelper.replaceTable(table4, DocHelper.genMap("T2G", String.valueOf((int)radioOfWssx) + "%"));
        DocHelper.replaceTable(table4, DocHelper.genMap("T2H", String.valueOf((int)surgeryOfWssx) + "%"));
        DocHelper.replaceTable(table4, DocHelper.genMap("T2I", String.valueOf((int)anaestheticsOfPSW) + "%"));
        DocHelper.replaceTable(table4, DocHelper.genMap("T2J", String.valueOf((int)dentalOfPSW) + "%"));
        DocHelper.replaceTable(table4, DocHelper.genMap("T2K", String.valueOf((int)emergOfPSW) + "%"));
        DocHelper.replaceTable(table4, DocHelper.genMap("T2L", String.valueOf((int)foundationOfPSW) + "%"));
        DocHelper.replaceTable(table4, DocHelper.genMap("T2M", String.valueOf((int)gpOfPSW) + "%"));
        DocHelper.replaceTable(table4, DocHelper.genMap("T2N", String.valueOf((int)medicineOfPSW) + "%"));
        DocHelper.replaceTable(table4, DocHelper.genMap("T2O", String.valueOf((int)obsOfPSW) + "%"));
        DocHelper.replaceTable(table4, DocHelper.genMap("T2P", String.valueOf((int)occhealthOfPSW) + "%"));
        DocHelper.replaceTable(table4, DocHelper.genMap("T2Q", String.valueOf((int)paediatricsOfPSW) + "%"));
        DocHelper.replaceTable(table4, DocHelper.genMap("T2R", String.valueOf((int)pathologyOfPSW) + "%"));
        DocHelper.replaceTable(table4, DocHelper.genMap("T2S", String.valueOf((int)pharmacyOfPSW) + "%"));
        DocHelper.replaceTable(table4, DocHelper.genMap("T2T", String.valueOf((int)psychOfPSW) + "%"));
        DocHelper.replaceTable(table4, DocHelper.genMap("T2U", String.valueOf((int)pubhealthOfPSW) + "%"));
        DocHelper.replaceTable(table4, DocHelper.genMap("T2V", String.valueOf((int)radioOfPSW) + "%"));
        DocHelper.replaceTable(table4, DocHelper.genMap("T2W", String.valueOf((int)surgeryOfPSW) + "%"));
    }

    public void getTable5() {
        XWPFTable table5 = this.xdoc.getTables().get(5);
        ArrayList<ArrayList<ReferralRecord>> list = new ArrayList<ArrayList<ReferralRecord>>();
        ArrayList<ReferralRecord> t5bournemouthList = this.helper.getTable5LineByTrust("The Royal Bournemouth and Christchurch Hospitals NHS Foundation Trust");
        ArrayList<ReferralRecord> t5dorchesterList = this.helper.getTable5LineByTrust("Dorchester");
        ArrayList<ReferralRecord> t5dorsetCountyList = this.helper.getTable5LineByTrust("Dorset County Hospital NHS Foundation Trust");
        ArrayList<ReferralRecord> t5dorsetHealthList = this.helper.getTable5LineByTrust("Dorset Healthcare University NHS Foundation Trust");
        ArrayList<ReferralRecord> t5hhftList = this.helper.getTable5LineByTrust("Hampshire Hospitals NHS Foundation Trust");
        ArrayList<ReferralRecord> t5iowList = this.helper.getTable5LineByTrust("Isle of Wight NHS Trust");
        ArrayList<ReferralRecord> t5jerseyList = this.helper.getTable5LineByTrust("Jersey General Hospital, States of Jersey");
        ArrayList<ReferralRecord> t5pooleList = this.helper.getTable5LineByTrust("Poole Hospital NHS Foundation Trust");
        ArrayList<ReferralRecord> t5portsmouthList = this.helper.getTable5LineByTrust("Portsmouth University Hospitals NHS Trust");
        ArrayList<ReferralRecord> t5salisburyList = this.helper.getTable5LineByTrust("Salisbury NHS Foundation Trust");
        ArrayList<ReferralRecord> t5solentList = this.helper.getTable5LineByTrust("Solent NHS Trust");
        ArrayList<ReferralRecord> t5southamptonList = this.helper.getTable5LineByTrust("University Hospital Southampton NHS Foundation Trust");
        ArrayList<ReferralRecord> t5southernList = this.helper.getTable5LineByTrust("Southern Health NHS Foundation Trust");
        ArrayList<ReferralRecord> t5gpPlacementList = this.helper.getTable5LineByTrust("GP Placement");
        list.add(t5bournemouthList);
        list.add(t5dorchesterList);
        list.add(t5dorsetCountyList);
        list.add(t5dorsetHealthList);
        list.add(t5hhftList);
        list.add(t5iowList);
        list.add(t5jerseyList);
        list.add(t5pooleList);
        list.add(t5portsmouthList);
        list.add(t5salisburyList);
        list.add(t5solentList);
        list.add(t5southamptonList);
        list.add(t5southernList);
        list.add(t5gpPlacementList);
        for (ArrayList arrayList : list) {
            if (arrayList.isEmpty()) continue;
            for (int i = 0; i < arrayList.size(); ++i) {
                int cursor = i;
                this.writeT5Line(table5, (ReferralRecord)arrayList.get(i), cursor);
            }
        }
        table5.removeRow(1);
        table5.removeRow(1);
    }

    public void getTable7() {
        XWPFTable table7 = this.xdoc.getTables().get(7);
        List<Double> list = this.helper.countTable7();
        Double ssg = list.get(0);
        Double cm = list.get(1);
        Double total = list.get(2);
        DocHelper.replaceTable(table7, DocHelper.genMap("TA", String.valueOf(ssg)));
        DocHelper.replaceTable(table7, DocHelper.genMap("TB", String.valueOf(cm)));
        DocHelper.replaceTable(table7, DocHelper.genMap("TC", String.valueOf(total)));
    }

    public void getTable8() {
        ArrayList<Table9Line> linesArray = this.helper.countTable8();
        Table9Line anaesthetics = linesArray.get(0);
        Table9Line dental = linesArray.get(1);
        Table9Line dermatology = linesArray.get(2);
        Table9Line enxdocrinology = linesArray.get(3);
        Table9Line foundation = linesArray.get(4);
        Table9Line gastroenterology = linesArray.get(5);
        Table9Line gp = linesArray.get(6);
        Table9Line haematology = linesArray.get(7);
        Table9Line histopathology = linesArray.get(8);
        Table9Line emergMed = linesArray.get(9);
        Table9Line medicine = linesArray.get(10);
        Table9Line neurology = linesArray.get(11);
        Table9Line obs = linesArray.get(12);
        Table9Line occHealth = linesArray.get(13);
        Table9Line oncology = linesArray.get(14);
        Table9Line ophtalmology = linesArray.get(15);
        Table9Line paediatrics = linesArray.get(16);
        Table9Line pathology = linesArray.get(17);
        Table9Line pharmacy = linesArray.get(18);
        Table9Line psych = linesArray.get(19);
        Table9Line pubHealth = linesArray.get(20);
        Table9Line radiology = linesArray.get(21);
        Table9Line sexHealth = linesArray.get(22);
        Table9Line rheumathology = linesArray.get(23);
        Table9Line surgery = linesArray.get(24);
        this.writeT8Line(anaesthetics);
        this.writeT8Line(dental);
        this.writeT8Line(dermatology);
        this.writeT8Line(enxdocrinology);
        this.writeT8Line(foundation);
        this.writeT8Line(gastroenterology);
        this.writeT8Line(gp);
        this.writeT8Line(haematology);
        this.writeT8Line(histopathology);
        this.writeT8Line(emergMed);
        this.writeT8Line(medicine);
        this.writeT8Line(neurology);
        this.writeT8Line(obs);
        this.writeT8Line(occHealth);
        this.writeT8Line(oncology);
        this.writeT8Line(ophtalmology);
        this.writeT8Line(paediatrics);
        this.writeT8Line(pathology);
        this.writeT8Line(pharmacy);
        this.writeT8Line(psych);
        this.writeT8Line(pubHealth);
        this.writeT8Line(radiology);
        this.writeT8Line(sexHealth);
        this.writeT8Line(rheumathology);
        this.writeT8Line(surgery);
        this.xdoc.getTables().get(8).removeRow(2);
    }

    private void writeT8Line(Table9Line line) {
        XWPFTable table8 = this.xdoc.getTables().get(8);
        XWPFTableRow oldRow = table8.getRow(2);
        line.countTotal();
        if (!line.isEmpty()) {
            CTRow firstCTRow = null;
            try {
                firstCTRow = CTRow.Factory.parse(oldRow.getCtRow().newInputStream());
            }
            catch (IOException | NullPointerException | XmlException e) {
                NativeError.printStackTrace((Object)e);
            }
            XWPFTableRow firstRow = new XWPFTableRow(firstCTRow, table8);
            int rowNum = table8.getRows().size() - 1;
            String title = line.getTitle() + "(" + String.valueOf(line.getTotalCount()) + ")";
            DocHelper.writeSmallCell(firstRow.getCell(0), title);
            if (line.getMale() != 0) {
                DocHelper.writeSmallCell(firstRow.getCell(1), String.valueOf(line.getMale()));
            }
            if (line.getFemale() != 0) {
                DocHelper.writeSmallCell(firstRow.getCell(2), String.valueOf(line.getFemale()));
            }
            if (line.getUk() != 0) {
                DocHelper.writeSmallCell(firstRow.getCell(3), String.valueOf(line.getUk()));
            }
            if (line.getNonUk() != 0) {
                DocHelper.writeSmallCell(firstRow.getCell(4), String.valueOf(line.getNonUk()));
            }
            if (line.getAge2329() != 0) {
                DocHelper.writeSmallCell(firstRow.getCell(5), String.valueOf(line.getAge2329()));
            }
            if (line.getAge3035() != 0) {
                DocHelper.writeSmallCell(firstRow.getCell(6), String.valueOf(line.getAge3035()));
            }
            if (line.getAge3540() != 0) {
                DocHelper.writeSmallCell(firstRow.getCell(7), String.valueOf(line.getAge3540()));
            }
            if (line.getAge40() != 0) {
                DocHelper.writeSmallCell(firstRow.getCell(8), String.valueOf(line.getAge40()));
            }
            if (line.getWhiteb() != 0) {
                DocHelper.writeSmallCell(firstRow.getCell(9), String.valueOf(line.getWhiteb()));
            }
            if (line.getWhiteo() != 0) {
                DocHelper.writeSmallCell(firstRow.getCell(10), String.valueOf(line.getWhiteo()));
            }
            if (line.getAsian() != 0) {
                DocHelper.writeSmallCell(firstRow.getCell(11), String.valueOf(line.getAsian()));
            }
            if (line.getAfrican() != 0) {
                DocHelper.writeSmallCell(firstRow.getCell(12), String.valueOf(line.getAfrican()));
            }
            if (line.getEthOther() != 0) {
                DocHelper.writeSmallCell(firstRow.getCell(13), String.valueOf(line.getEthOther()));
            }
            if (line.getChristian() != 0) {
                DocHelper.writeSmallCell(firstRow.getCell(14), String.valueOf(line.getChristian()));
            }
            if (line.getIslam() != 0) {
                DocHelper.writeSmallCell(firstRow.getCell(15), String.valueOf(line.getIslam()));
            }
            if (line.getHindu() != 0) {
                DocHelper.writeSmallCell(firstRow.getCell(16), String.valueOf(line.getHindu()));
            }
            if (line.getAtheist() != 0) {
                DocHelper.writeSmallCell(firstRow.getCell(17), String.valueOf(line.getAtheist()));
            }
            if (line.getSikh() != 0) {
                DocHelper.writeSmallCell(firstRow.getCell(18), String.valueOf(line.getSikh()));
            }
            if (line.getJudaism() != 0) {
                DocHelper.writeSmallCell(firstRow.getCell(19), String.valueOf(line.getJudaism()));
            }
            if (line.getBuddhism() != 0) {
                DocHelper.writeSmallCell(firstRow.getCell(20), String.valueOf(line.getBuddhism()));
            }
            if (line.getRelOther() != 0) {
                DocHelper.writeSmallCell(firstRow.getCell(21), String.valueOf(line.getRelOther()));
            }
            if (line.getRelPNS() != 0) {
                DocHelper.writeSmallCell(firstRow.getCell(22), String.valueOf(line.getRelPNS()));
            }
            if (line.getYes() != 0) {
                DocHelper.writeSmallCell(firstRow.getCell(23), String.valueOf(line.getYes()));
            }
            if (line.getNo() != 0) {
                DocHelper.writeSmallCell(firstRow.getCell(24), String.valueOf(line.getNo()));
            }
            if (line.getHet() != 0) {
                DocHelper.writeSmallCell(firstRow.getCell(25), String.valueOf(line.getHet()));
            }
            if (line.getBisexual() != 0) {
                DocHelper.writeSmallCell(firstRow.getCell(26), String.valueOf(line.getBisexual()));
            }
            if (line.getHomosexual() != 0) {
                DocHelper.writeSmallCell(firstRow.getCell(27), String.valueOf(line.getHomosexual()));
            }
            if (line.getSexOrPNS() != 0) {
                DocHelper.writeSmallCell(firstRow.getCell(28), String.valueOf(line.getSexOrPNS()));
            }
            table8.addRow(firstRow);
        }
    }

    public void getTable9() {
        ArrayList<Table9Line> linesArray = this.helper.countTable9();
        Table9Line anaesthetics = linesArray.get(0);
        Table9Line dental = linesArray.get(1);
        Table9Line dermatology = linesArray.get(2);
        Table9Line enxdocrinology = linesArray.get(3);
        Table9Line foundation = linesArray.get(4);
        Table9Line gastroenterology = linesArray.get(5);
        Table9Line gp = linesArray.get(6);
        Table9Line haematology = linesArray.get(7);
        Table9Line histopathology = linesArray.get(8);
        Table9Line emergMed = linesArray.get(9);
        Table9Line medicine = linesArray.get(10);
        Table9Line neurology = linesArray.get(11);
        Table9Line obs = linesArray.get(12);
        Table9Line occHealth = linesArray.get(13);
        Table9Line oncology = linesArray.get(14);
        Table9Line ophtalmology = linesArray.get(15);
        Table9Line paediatrics = linesArray.get(16);
        Table9Line pathology = linesArray.get(17);
        Table9Line pharmacy = linesArray.get(18);
        Table9Line psych = linesArray.get(19);
        Table9Line pubHealth = linesArray.get(20);
        Table9Line radiology = linesArray.get(21);
        Table9Line sexHealth = linesArray.get(22);
        Table9Line rheumathology = linesArray.get(23);
        Table9Line surgery = linesArray.get(24);
        this.writeT9Line(anaesthetics);
        this.writeT9Line(dental);
        this.writeT9Line(dermatology);
        this.writeT9Line(enxdocrinology);
        this.writeT9Line(foundation);
        this.writeT9Line(gastroenterology);
        this.writeT9Line(gp);
        this.writeT9Line(haematology);
        this.writeT9Line(histopathology);
        this.writeT9Line(emergMed);
        this.writeT9Line(medicine);
        this.writeT9Line(neurology);
        this.writeT9Line(obs);
        this.writeT9Line(occHealth);
        this.writeT9Line(oncology);
        this.writeT9Line(ophtalmology);
        this.writeT9Line(paediatrics);
        this.writeT9Line(pathology);
        this.writeT9Line(pharmacy);
        this.writeT9Line(psych);
        this.writeT9Line(pubHealth);
        this.writeT9Line(radiology);
        this.writeT9Line(sexHealth);
        this.writeT9Line(rheumathology);
        this.writeT9Line(surgery);
        this.xdoc.getTables().get(9).removeRow(2);
    }

    private void writeT9Line(Table9Line line) {
        XWPFTable table9 = this.xdoc.getTables().get(9);
        XWPFTableRow oldRow = table9.getRow(2);
        line.countTotal();
        if (!line.isEmpty()) {
            CTRow firstCTRow = null;
            try {
                firstCTRow = CTRow.Factory.parse(oldRow.getCtRow().newInputStream());
            }
            catch (IOException | NullPointerException | XmlException e) {
                NativeError.printStackTrace((Object)e);
            }
            XWPFTableRow firstRow = new XWPFTableRow(firstCTRow, table9);
            int rowNum = table9.getRows().size() - 1;
            String title = line.getTitle() + "(" + String.valueOf(line.getTotalCount()) + ")";
            DocHelper.writeSmallCell(firstRow.getCell(0), title);
            if (line.getMale() != 0) {
                DocHelper.writeSmallCell(firstRow.getCell(1), String.valueOf(line.getMale()));
            }
            if (line.getFemale() != 0) {
                DocHelper.writeSmallCell(firstRow.getCell(2), String.valueOf(line.getFemale()));
            }
            if (line.getAge2329() != 0) {
                DocHelper.writeSmallCell(firstRow.getCell(3), String.valueOf(line.getAge2329()));
            }
            if (line.getAge3035() != 0) {
                DocHelper.writeSmallCell(firstRow.getCell(4), String.valueOf(line.getAge3035()));
            }
            if (line.getAge3540() != 0) {
                DocHelper.writeSmallCell(firstRow.getCell(5), String.valueOf(line.getAge3540()));
            }
            if (line.getAge40() != 0) {
                DocHelper.writeSmallCell(firstRow.getCell(6), String.valueOf(line.getAge40()));
            }
            if (line.getWhiteb() != 0) {
                DocHelper.writeSmallCell(firstRow.getCell(7), String.valueOf(line.getWhiteb()));
            }
            if (line.getWhiteo() != 0) {
                DocHelper.writeSmallCell(firstRow.getCell(8), String.valueOf(line.getWhiteo()));
            }
            if (line.getAsian() != 0) {
                DocHelper.writeSmallCell(firstRow.getCell(9), String.valueOf(line.getAsian()));
            }
            if (line.getAfrican() != 0) {
                DocHelper.writeSmallCell(firstRow.getCell(10), String.valueOf(line.getAfrican()));
            }
            if (line.getEthOther() != 0) {
                DocHelper.writeSmallCell(firstRow.getCell(11), String.valueOf(line.getEthOther()));
            }
            if (line.getChristian() != 0) {
                DocHelper.writeSmallCell(firstRow.getCell(12), String.valueOf(line.getChristian()));
            }
            if (line.getIslam() != 0) {
                DocHelper.writeSmallCell(firstRow.getCell(13), String.valueOf(line.getIslam()));
            }
            if (line.getHindu() != 0) {
                DocHelper.writeSmallCell(firstRow.getCell(14), String.valueOf(line.getHindu()));
            }
            if (line.getAtheist() != 0) {
                DocHelper.writeSmallCell(firstRow.getCell(15), String.valueOf(line.getAtheist()));
            }
            if (line.getSikh() != 0) {
                DocHelper.writeSmallCell(firstRow.getCell(16), String.valueOf(line.getSikh()));
            }
            if (line.getJudaism() != 0) {
                DocHelper.writeSmallCell(firstRow.getCell(17), String.valueOf(line.getJudaism()));
            }
            if (line.getBuddhism() != 0) {
                DocHelper.writeSmallCell(firstRow.getCell(18), String.valueOf(line.getBuddhism()));
            }
            if (line.getRelOther() != 0) {
                DocHelper.writeSmallCell(firstRow.getCell(19), String.valueOf(line.getRelOther()));
            }
            if (line.getRelPNS() != 0) {
                DocHelper.writeSmallCell(firstRow.getCell(20), String.valueOf(line.getRelPNS()));
            }
            if (line.getHet() != 0) {
                DocHelper.writeSmallCell(firstRow.getCell(21), String.valueOf(line.getHet()));
            }
            if (line.getBisexual() != 0) {
                DocHelper.writeSmallCell(firstRow.getCell(22), String.valueOf(line.getBisexual()));
            }
            if (line.getHomosexual() != 0) {
                DocHelper.writeSmallCell(firstRow.getCell(23), String.valueOf(line.getHomosexual()));
            }
            if (line.getSexOrPNS() != 0) {
                DocHelper.writeSmallCell(firstRow.getCell(24), String.valueOf(line.getSexOrPNS()));
            }
            if (line.getYes() != 0) {
                DocHelper.writeSmallCell(firstRow.getCell(25), String.valueOf(line.getYes()));
            }
            if (line.getNo() != 0) {
                DocHelper.writeSmallCell(firstRow.getCell(26), String.valueOf(line.getNo()));
            }
            table9.addRow(firstRow);
        }
    }

    public void saveDoc() {
        FileOutputStream fos = null;
        try {
            fos = new FileOutputStream(this.doc);
            this.xdoc.write(fos);
            fos.close();
            this.xdoc.close();
        }
        catch (IOException ex) {
            Logger.getLogger(DocHelper.class.getName()).log(Level.SEVERE, null, ex);
        }
    }

    private static void writeSmallCell(XWPFTableCell cell, String str) {
        XWPFRun run = cell.getParagraphs().get(0).createRun();
        run.setFontSize(8);
        run.setText(str);
    }

    public static void genDocOut(File inFile, File doc) {
        try {
            XWPFDocument xdocOut = new XWPFDocument(new FileInputStream(inFile));
            FileOutputStream fos = new FileOutputStream(doc);
            xdocOut.write(fos);
            xdocOut.close();
            fos.close();
        }
        catch (FileNotFoundException ex) {
            Logger.getLogger(DocHelper.class.getName()).log(Level.SEVERE, null, ex);
        }
        catch (IOException ex) {
            Logger.getLogger(DocHelper.class.getName()).log(Level.SEVERE, null, ex);
        }
    }

    public void genDocFile() {
        try {
            this.doc.createNewFile();
        }
        catch (IOException ex) {
            Logger.getLogger(DocHelper.class.getName()).log(Level.SEVERE, null, ex);
        }
    }

    private static void replaceTable(XWPFTable table, HashMap m) {
        for (XWPFTableRow row : table.getRows()) {
            for (XWPFTableCell cell : row.getTableCells()) {
                for (XWPFParagraph p : cell.getParagraphs()) {
                    DocHelper.replaceParagraphInCell(p, m);
                }
            }
        }
    }

    private static void replaceTableSmall(XWPFTable table, HashMap m) {
        for (XWPFTableRow row : table.getRows()) {
            for (XWPFTableCell cell : row.getTableCells()) {
                for (XWPFParagraph p : cell.getParagraphs()) {
                    DocHelper.replaceParagraphInCellSmall(p, m);
                }
            }
        }
    }

    private static void replaceParagraphInCell(XWPFParagraph paragraph, HashMap m) {
        String text = paragraph.getText();
        int size = paragraph.getRuns().size();
        if (text != null && text.contains(m.get("a").toString())) {
            for (int i = 0; i < size; ++i) {
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
            for (int i = 0; i < size; ++i) {
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
            if (text == null || !text.contains(m.get("a").toString())) continue;
            text = text.replace(m.get("a").toString(), m.get("b").toString());
            r.setText(text, 0);
        }
    }

    private void writeT5Line(XWPFTable table, ReferralRecord record, int cursor) {
        String trust = record.getTrust();
        String gender = record.getGender();
        String grade = record.getGrade();
        String specialty = record.getSpecialty();
        List<String> addRefList = record.getAddRef();
        String addRef = "";
        for (int i = 0; i < addRefList.size(); ++i) {
            if (i == 0) {
                if (addRefList.size() > 1) {
                    addRef = addRefList.get(0) + "; ";
                    continue;
                }
                addRef = addRefList.get(0);
                continue;
            }
            addRef = addRef + addRefList.get(i) + "; ";
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
        if (cursor == 0) {
            if (gender.equals("Female")) {
                XWPFTableRow oldRow = rowTemplateF;
                CTRow firstCTRow = null;
                try {
                    firstCTRow = CTRow.Factory.parse(oldRow.getCtRow().newInputStream());
                }
                catch (IOException | XmlException exception) {
                    // empty catch block
                }
                XWPFTableRow firstRow = new XWPFTableRow(firstCTRow, table);
                XWPFTableCell cell0 = firstRow.getCell(0);
                DocHelper.writeSmallCell(cell0, trust);
                XWPFTableCell cell1 = firstRow.getCell(1);
                DocHelper.writeSmallCell(cell1, gender);
                XWPFTableCell cell2 = firstRow.getCell(2);
                DocHelper.writeSmallCell(cell2, grade);
                XWPFTableCell cell3 = firstRow.getCell(3);
                DocHelper.writeSmallCell(cell3, specialty);
                XWPFTableCell cell4 = firstRow.getCell(4);
                DocHelper.writeSmallCell(cell4, addRef);
                XWPFTableCell cell5 = firstRow.getCell(5);
                DocHelper.writeSmallCell(cell5, country);
                XWPFTableCell cell6 = firstRow.getCell(6);
                DocHelper.writeSmallCell(cell6, String.valueOf(age));
                XWPFTableCell cell7 = firstRow.getCell(7);
                DocHelper.writeSmallCell(cell7, ethnicity);
                XWPFTableCell cell8 = firstRow.getCell(8);
                DocHelper.writeSmallCell(cell8, sexOr);
                XWPFTableCell cell9 = firstRow.getCell(9);
                DocHelper.writeSmallCell(cell9, religion);
                XWPFTableCell cell10 = firstRow.getCell(10);
                DocHelper.writeSmallCell(cell10, disability);
                table.addRow(firstRow);
            } else if (gender.equals("Male")) {
                XWPFTableRow oldRow = rowTemplateM;
                CTRow firstCTRow = null;
                try {
                    firstCTRow = CTRow.Factory.parse(oldRow.getCtRow().newInputStream());
                }
                catch (IOException | XmlException e) {
                    e.printStackTrace();
                }
                XWPFTableRow firstRow = new XWPFTableRow(firstCTRow, table);
                XWPFTableCell cell0 = firstRow.getCell(0);
                DocHelper.writeSmallCell(cell0, trust);
                XWPFTableCell cell1 = firstRow.getCell(1);
                DocHelper.writeSmallCell(cell1, gender);
                XWPFTableCell cell2 = firstRow.getCell(2);
                DocHelper.writeSmallCell(cell2, grade);
                XWPFTableCell cell3 = firstRow.getCell(3);
                DocHelper.writeSmallCell(cell3, specialty);
                XWPFTableCell cell4 = firstRow.getCell(4);
                DocHelper.writeSmallCell(cell4, addRef);
                XWPFTableCell cell5 = firstRow.getCell(5);
                DocHelper.writeSmallCell(cell5, country);
                XWPFTableCell cell6 = firstRow.getCell(6);
                DocHelper.writeSmallCell(cell6, String.valueOf(age));
                XWPFTableCell cell7 = firstRow.getCell(7);
                DocHelper.writeSmallCell(cell7, ethnicity);
                XWPFTableCell cell8 = firstRow.getCell(8);
                DocHelper.writeSmallCell(cell8, sexOr);
                XWPFTableCell cell9 = firstRow.getCell(9);
                DocHelper.writeSmallCell(cell9, religion);
                XWPFTableCell cell10 = firstRow.getCell(10);
                DocHelper.writeSmallCell(cell10, disability);
                table.addRow(firstRow);
            }
        } else if (gender.equals("Female")) {
            XWPFTableRow oldRow = rowTemplateF;
            CTRow firstCTRow = null;
            try {
                firstCTRow = CTRow.Factory.parse(oldRow.getCtRow().newInputStream());
            }
            catch (IOException | XmlException firstRow) {
                // empty catch block
            }
            XWPFTableRow firstRow = new XWPFTableRow(firstCTRow, table);
            XWPFTableCell cell1 = firstRow.getCell(1);
            DocHelper.writeSmallCell(cell1, gender);
            XWPFTableCell cell2 = firstRow.getCell(2);
            DocHelper.writeSmallCell(cell2, grade);
            XWPFTableCell cell3 = firstRow.getCell(3);
            DocHelper.writeSmallCell(cell3, specialty);
            XWPFTableCell cell4 = firstRow.getCell(4);
            DocHelper.writeSmallCell(cell4, addRef);
            XWPFTableCell cell5 = firstRow.getCell(5);
            DocHelper.writeSmallCell(cell5, country);
            XWPFTableCell cell6 = firstRow.getCell(6);
            DocHelper.writeSmallCell(cell6, String.valueOf(age));
            XWPFTableCell cell7 = firstRow.getCell(7);
            DocHelper.writeSmallCell(cell7, ethnicity);
            XWPFTableCell cell8 = firstRow.getCell(8);
            DocHelper.writeSmallCell(cell8, sexOr);
            XWPFTableCell cell9 = firstRow.getCell(9);
            DocHelper.writeSmallCell(cell9, religion);
            XWPFTableCell cell10 = firstRow.getCell(10);
            DocHelper.writeSmallCell(cell10, disability);
            table.addRow(firstRow);
        } else if (gender.equals("Male")) {
            XWPFTableRow oldRow = rowTemplateM;
            CTRow firstCTRow = null;
            try {
                firstCTRow = CTRow.Factory.parse(oldRow.getCtRow().newInputStream());
            }
            catch (IOException | XmlException e) {
                e.printStackTrace();
            }
            XWPFTableRow firstRow = new XWPFTableRow(firstCTRow, table);
            XWPFTableCell cell1 = firstRow.getCell(1);
            DocHelper.writeSmallCell(cell1, gender);
            XWPFTableCell cell2 = firstRow.getCell(2);
            DocHelper.writeSmallCell(cell2, grade);
            XWPFTableCell cell3 = firstRow.getCell(3);
            DocHelper.writeSmallCell(cell3, specialty);
            XWPFTableCell cell4 = firstRow.getCell(4);
            DocHelper.writeSmallCell(cell4, addRef);
            XWPFTableCell cell5 = firstRow.getCell(5);
            DocHelper.writeSmallCell(cell5, country);
            XWPFTableCell cell6 = firstRow.getCell(6);
            DocHelper.writeSmallCell(cell6, String.valueOf(age));
            XWPFTableCell cell7 = firstRow.getCell(7);
            DocHelper.writeSmallCell(cell7, ethnicity);
            XWPFTableCell cell8 = firstRow.getCell(8);
            DocHelper.writeSmallCell(cell8, sexOr);
            XWPFTableCell cell9 = firstRow.getCell(9);
            DocHelper.writeSmallCell(cell9, religion);
            XWPFTableCell cell10 = firstRow.getCell(10);
            DocHelper.writeSmallCell(cell10, disability);
            table.addRow(firstRow);
        }
    }

    public static HashMap<String, String> genMap(String a, String b) {
        HashMap<String, String> hash = new HashMap<String, String>();
        hash.put("a", a);
        hash.put("b", b);
        return hash;
    }

    public static int getStartingYear() {
        int year = Calendar.getInstance().get(1) - 1;
        return year;
    }

    public static String genSplitYearSeq(String str) {
        ArrayList<String> strList = new ArrayList<String>();
        for (String s : str.split("")) {
            strList.add(s);
        }
        String finalStr = (String)strList.get(2) + (String)strList.get(3);
        return finalStr;
    }
}
