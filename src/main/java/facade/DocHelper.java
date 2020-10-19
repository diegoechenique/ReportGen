/*
* To change this license header, choose License Headers in Project Properties.
* To change this template file, choose Tools | Templates
* and open the template in the editor.
*/
package facade;

import com.sun.org.apache.xalan.internal.xsltc.compiler.Template;
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

/**
 *
 * @author diego
 */
public class DocHelper {
    
    
    private InputStream template = Thread.currentThread().getContextClassLoader().getResourceAsStream(Strings.DOC_PATH);
    private File outFile;
    private File psw;
    private File graphs;
    
    int twoYrsAgo = getYear()-2;
    int prevYear = getYear()-1;
    int thisYear = getYear();
    String yrPastCurrentYYYY = prevYear+"/"+thisYear;
    
    
    public DocHelper(File outFile,File psw, File graphs){
        this.outFile=outFile;
        this.psw = psw;
        this.graphs = graphs;
    }
    
    
    /**
     * Generates the first table in the document
     * @param docOut
     */
    public void getTable0(){
        PoiHelper helper = new PoiHelper(psw, graphs);
        List<Integer> list = helper.countTable0();
        
        int casesClosed = 0;
        int casesOClosed = 0;
        
        try {
            XWPFDocument doc = new XWPFDocument(new FileInputStream(outFile));

            
            XWPFTable table0 = doc.getTables().get(0);
            
        } catch (FileNotFoundException ex) {
            Logger.getLogger(DocHelper.class.getName()).log(Level.SEVERE, null, ex);
        } catch (IOException ex) {
            Logger.getLogger(DocHelper.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
    
    public void getTables(){
        
        PoiHelper helper = new PoiHelper(psw, graphs);
        
        //table 0
        try {
//            
            XWPFDocument doc = new XWPFDocument(this.template);
//            
//            List<Integer> t0list = helper.countTable0();
//            int stCount = t0list.get(0);
//            int fCount = t0list.get(1);
//            int gpCount = t0list.get(2);
//            int otherCount = t0list.get(3);
//            int total = t0list.get(4);
//            int casesClosed = 0;
//            int casesOClosed = 0;
//            
//            XWPFTable table0 = doc.getTables().get(0);
//            
//            replaceTable(table0,genMap(Strings.TABLE_A, Integer.toString(stCount)));
//            replaceTable(table0,genMap(Strings.TABLE_B, Integer.toString(fCount)));
//            replaceTable(table0,genMap(Strings.TABLE_C, Integer.toString(gpCount)));
//            replaceTable(table0,genMap(Strings.TABLE_D, Integer.toString(otherCount)));
//            replaceTable(table0,genMap(Strings.TABLE_E, Integer.toString(total)));
//            replaceTable(table0,genMap(Strings.TABLE_F, Integer.toString(casesClosed)));
//            replaceTable(table0,genMap(Strings.TABLE_G, Integer.toString(casesOClosed)));
//            
//            //table 1
//            List<Integer> t1list = helper.countTable1();
//            
//            int capFCount = t1list.get(0);
//            int capMCount = t1list.get(1);
//            int exFailFCount = t1list.get(2);
//            int exFailMCount = t1list.get(3);
//            int healthFCount = t1list.get(4);
//            int healthMCount = t1list.get(5);
//            int carreerFCount = t1list.get(6);
//            int carreerMCount = t1list.get(7);
//            int conductFCount = t1list.get(8);
//            int conductMCount = t1list.get(9);
//            int otherFCount = t1list.get(10);
//            int otherMCount = t1list.get(11);
//            int referredCount = t1list.get(12);
//            
//            XWPFTable table1 = doc.getTables().get(1);
//            replaceTable(table1,genMap(Strings.TABLE_A, Integer.toString(referredCount)));
//            replaceTable(table1,genMap(Strings.TABLE_B, genSplitYearSeq(String.valueOf(prevYear))+"/"+genSplitYearSeq(String.valueOf(thisYear))));
//            replaceTable(table1,genMap(Strings.TABLE_C, Integer.toString(capMCount)));
//            replaceTable(table1,genMap(Strings.TABLE_D, Integer.toString(capFCount)));
//            replaceTable(table1,genMap(Strings.TABLE_E, Integer.toString(exFailMCount)));
//            replaceTable(table1,genMap(Strings.TABLE_F, Integer.toString(exFailFCount)));
//            replaceTable(table1,genMap(Strings.TABLE_G, Integer.toString(healthMCount)));
//            replaceTable(table1,genMap(Strings.TABLE_H, Integer.toString(healthFCount)));
//            replaceTable(table1,genMap(Strings.TABLE_I, Integer.toString(carreerMCount)));
//            replaceTable(table1,genMap(Strings.TABLE_J, Integer.toString(carreerFCount)));
//            replaceTable(table1,genMap(Strings.TABLE_K, Integer.toString(conductMCount)));
//            replaceTable(table1,genMap(Strings.TABLE_L, Integer.toString(conductFCount)));
//            replaceTable(table1,genMap(Strings.TABLE_M, Integer.toString(otherMCount)));
//            replaceTable(table1,genMap(Strings.TABLE_N, Integer.toString(otherFCount)));
//            
//            
//            //table 2
//            List<Integer> t2list = helper.countTable2();
//            
//            double f1TotalCount = t2list.get(1);
//            double f2TotalCount = t2list.get(2);
//            double f1RefCount = t2list.get(3);
//            double f2RefCount = t2list.get(4);
//            double totalFCount = f1TotalCount+f2TotalCount;
//            double totalFRef = f1RefCount+f2RefCount;
//            double f1Perc = (f1RefCount/f1TotalCount)*100;
//            double f2Perc = (f2RefCount/f2TotalCount)*100;
//            
//            double fPerc = ((totalFRef)/totalFCount)*100;
//            double fRefPerc = ((totalFRef)/referredCount)*100;
//            DecimalFormat format = new DecimalFormat("#.#");
//            
//            XWPFTable table2 = doc.getTables().get(2);
//            replaceTable(table2,genMap(Strings.TABLE_A, String.valueOf(prevYear)+"/"+genSplitYearSeq(String.valueOf(thisYear))));
//            replaceTable(table2,genMap(Strings.TABLE_B, String.valueOf((int)totalFRef)+" ("+format.format(fPerc)+"%)"));
//            replaceTable(table2,genMap(Strings.TABLE_C, "F1 = "+(int)f1RefCount+" F2 = "+(int)f2RefCount));
//            replaceTable(table2,genMap(Strings.TABLE_D, String.valueOf(format.format(fRefPerc)+"%")));
//            replaceTable(table2,genMap(Strings.TABLE_E, "F1 = "+(int)f1TotalCount+" F2 = "+(int)f2TotalCount));
//            replaceTable(table2,genMap(Strings.TABLE_F, String.valueOf(format.format(f1Perc)+"%")));
//            replaceTable(table2,genMap(Strings.TABLE_G, String.valueOf(format.format(f2Perc)+"%")));
//            
//            //table 3
//            List<Integer> t3list = helper.countTable3();
//            
//            double bournemouthTotal = t3list.get(0);
//            double dorchesterTotal = t3list.get(1);
//            double dorsetTotal = t3list.get(2);
//            double hhftTotal = t3list.get(3);
//            double iowTotal = t3list.get(4);
//            double jerseyTotal = t3list.get(5);
//            double pooleTotal = t3list.get(6);
//            double portsmouthTotal = t3list.get(7);
//            double salisburyTotal = t3list.get(8);
//            double solentTotal = t3list.get(9);
//            double southamptonTotal = t3list.get(10);
//            double southernTotal = t3list.get(11);
//            
//            double bournemouthRefNo = t3list.get(12);
//            double dorchesterRefNo = t3list.get(13);
//            double dorsetRefNo = t3list.get(14);
//            double hhftRefNo = t3list.get(15);
//            double iowRefNo = t3list.get(16);
//            double jerseyRefNo = t3list.get(17);
//            double pooleRefNo = t3list.get(18);
//            double portsmouthRefNo = t3list.get(19);
//            double salisburyRefNo = t3list.get(20);
//            double solentRefNo = t3list.get(21);
//            double southamptonRefNo = t3list.get(22);
//            double southernRefNo = t3list.get(23);
//            
//            double totalRefs = t3list.get(24);
//            double totalWessex = t3list.get(25);
//            
//            double bournemouthTrInTrust = Math.round((bournemouthRefNo/bournemouthTotal)*100);
//            double dorchesterTrInTrust = Math.round((dorchesterRefNo/dorchesterTotal)*100);
//            double dorsetTrInTrust = Math.round((dorsetRefNo/dorsetTotal)*100);
//            double hhtfTrInTrust = Math.round((hhftRefNo/hhftTotal)*100);
//            double iowTrInTrust = Math.round((iowRefNo/iowTotal)*100);
//            double jerseyTrInTrust = Math.round((jerseyRefNo/jerseyTotal)*100);
//            double pooleTrInTrust = Math.round((pooleRefNo/pooleTotal)*100);
//            double portsmouthTrInTrust = Math.round((portsmouthRefNo/portsmouthTotal)*100);
//            double salisburyTrInTrust = Math.round((salisburyRefNo/salisburyTotal)*100);
//            double solentTrInTrust = Math.round((solentRefNo/solentTotal)*100);
//            double southamptonTrInTrust = Math.round((southamptonRefNo/southamptonTotal)*100);
//            double southernTrInTrust = Math.round((southernRefNo/southernTotal)*100);
//            
//            double bournemouthOfWssx = Math.round((bournemouthTotal/totalWessex)*100);
//            double dorchesterOfWssx = Math.round((dorchesterTotal/totalWessex)*100);
//            double dorsetOfWssx = Math.round((dorsetTotal/totalWessex)*100);
//            double hhtfOfWssx = Math.round((hhftTotal/totalWessex)*100);
//            double iowOfWssx = Math.round((iowTotal/totalWessex)*100);
//            double jerseyOfWssx = Math.round((jerseyTotal/totalWessex)*100);
//            double pooleOfWssx = Math.round((pooleTotal/totalWessex)*100);
//            double portsmouthOfWssx = Math.round((portsmouthTotal/totalWessex)*100);
//            double salisburyOfWssx = Math.round((salisburyTotal/totalWessex)*100);
//            double solentOfWssx = Math.round((solentTotal/totalWessex)*100);
//            double southamptonOfWssx = Math.round((southamptonTotal/totalWessex)*100);
//            double southernOfWssx = Math.round((southernTotal/totalWessex)*100);
//            
//            double bournemouthOfPSU = Math.round((bournemouthRefNo/totalRefs)*100);
//            double dorchesterOfPSU = Math.round((dorchesterRefNo/totalRefs)*100);
//            double dorsetOfPSU = Math.round((dorsetRefNo/totalRefs)*100);
//            double hhtfOfPSU = Math.round((hhftRefNo/totalRefs)*100);
//            double iowOfPSU = Math.round((iowRefNo/totalRefs)*100);
//            double jerseyOfPSU = Math.round((jerseyRefNo/totalRefs)*100);
//            double pooleOfPSU = Math.round((pooleRefNo/totalRefs)*100);
//            double portsmouthOfPSU = Math.round((portsmouthRefNo/totalRefs)*100);
//            double salisburyOfPSU = Math.round((salisburyRefNo/totalRefs)*100);
//            double solentOfPSU = Math.round((solentRefNo/totalRefs)*100);
//            double southamptonOfPSU = Math.round((southamptonRefNo/totalRefs)*100);
//            double southernOfPSU = Math.round((southernRefNo/totalRefs)*100);
//            
//            
//            
//            
//            XWPFTable table3 = doc.getTables().get(3);
//            
//            replaceTable(table3, genMap(Strings.TABLE_A, String.valueOf((int)bournemouthRefNo)));
//            replaceTable(table3, genMap(Strings.TABLE_B, String.valueOf((int)dorchesterRefNo)));
//            replaceTable(table3, genMap(Strings.TABLE_C, String.valueOf((int)dorsetRefNo)));
//            replaceTable(table3, genMap(Strings.TABLE_D, String.valueOf((int)hhftRefNo)));
//            replaceTable(table3, genMap(Strings.TABLE_E, String.valueOf((int)iowRefNo)));
//            replaceTable(table3, genMap(Strings.TABLE_F, String.valueOf((int)jerseyRefNo)));
//            replaceTable(table3, genMap(Strings.TABLE_G, String.valueOf((int)pooleRefNo)));
//            replaceTable(table3, genMap(Strings.TABLE_H, String.valueOf((int)portsmouthRefNo)));
//            replaceTable(table3, genMap(Strings.TABLE_I, String.valueOf((int)salisburyRefNo)));
//            replaceTable(table3, genMap(Strings.TABLE_J, String.valueOf((int)solentRefNo)));
//            replaceTable(table3, genMap(Strings.TABLE_K, String.valueOf((int)southamptonRefNo)));
//            replaceTable(table3, genMap(Strings.TABLE_L, String.valueOf((int)southernRefNo)));
//            
//            replaceTable(table3, genMap(Strings.TABLE_M, String.valueOf((int)bournemouthTotal)));
//            replaceTable(table3, genMap(Strings.TABLE_N, String.valueOf((int)dorchesterTotal)));
//            replaceTable(table3, genMap(Strings.TABLE_O, String.valueOf((int)dorsetTotal)));
//            replaceTable(table3, genMap(Strings.TABLE_P, String.valueOf((int)hhftTotal)));
//            replaceTable(table3, genMap(Strings.TABLE_Q, String.valueOf((int)iowTotal)));
//            replaceTable(table3, genMap(Strings.TABLE_R, String.valueOf((int)jerseyTotal)));
//            replaceTable(table3, genMap(Strings.TABLE_S, String.valueOf((int)pooleTotal)));
//            replaceTable(table3, genMap(Strings.TABLE_T, String.valueOf((int)portsmouthTotal)));
//            replaceTable(table3, genMap(Strings.TABLE_U, String.valueOf((int)salisburyTotal)));
//            replaceTable(table3, genMap(Strings.TABLE_V, String.valueOf((int)solentTotal)));
//            replaceTable(table3, genMap(Strings.TABLE_W, String.valueOf((int)southamptonTotal)));
//            replaceTable(table3, genMap(Strings.TABLE_X, String.valueOf((int)southernTotal)));
//            
//            replaceTable(table3, genMap(Strings.TABLE_Y, String.valueOf((int) bournemouthTrInTrust+"%")));
//            replaceTable(table3, genMap(Strings.TABLE_Z, String.valueOf((int)dorchesterTrInTrust+"%")));
//            replaceTable(table3, genMap(Strings.TABLE_AA, String.valueOf((int)dorsetTrInTrust+"%")));
//            replaceTable(table3, genMap(Strings.TABLE_AB, String.valueOf((int)hhtfTrInTrust+"%")));
//            replaceTable(table3, genMap(Strings.TABLE_AC, String.valueOf((int)iowTrInTrust+"%")));
//            replaceTable(table3, genMap(Strings.TABLE_AD, String.valueOf((int)jerseyTrInTrust+"%")));
//            replaceTable(table3, genMap(Strings.TABLE_AE, String.valueOf((int)pooleTrInTrust+"%")));
//            replaceTable(table3, genMap(Strings.TABLE_AF, String.valueOf((int)portsmouthTrInTrust+"%")));
//            replaceTable(table3, genMap(Strings.TABLE_AG, String.valueOf((int)salisburyTrInTrust+"%")));
//            replaceTable(table3, genMap(Strings.TABLE_AH, String.valueOf((int)solentTrInTrust+"%")));
//            replaceTable(table3, genMap(Strings.TABLE_AI, String.valueOf((int)southamptonTrInTrust+"%")));
//            replaceTable(table3, genMap(Strings.TABLE_AJ, String.valueOf((int)southernTrInTrust+"%")));
//            
//            replaceTable(table3, genMap(Strings.TABLE_AK, String.valueOf((int)bournemouthOfWssx+"%")));
//            replaceTable(table3, genMap(Strings.TABLE_AL, String.valueOf((int)dorchesterOfWssx+"%")));
//            replaceTable(table3, genMap(Strings.TABLE_AM, String.valueOf((int)dorsetOfWssx+"%")));
//            replaceTable(table3, genMap(Strings.TABLE_AN, String.valueOf((int)hhtfOfWssx+"%")));
//            replaceTable(table3, genMap(Strings.TABLE_AO, String.valueOf((int)iowOfWssx+"%")));
//            replaceTable(table3, genMap(Strings.TABLE_AP, String.valueOf((int)jerseyOfWssx+"%")));
//            replaceTable(table3, genMap(Strings.TABLE_AQ, String.valueOf((int)pooleOfWssx+"%")));
//            replaceTable(table3, genMap(Strings.TABLE_AR, String.valueOf((int)portsmouthOfWssx+"%")));
//            replaceTable(table3, genMap(Strings.TABLE_AS, String.valueOf((int)salisburyOfWssx+"%")));
//            replaceTable(table3, genMap(Strings.TABLE_AT, String.valueOf((int)solentOfWssx+"%")));
//            replaceTable(table3, genMap(Strings.TABLE_AU, String.valueOf((int)southamptonOfWssx+"%")));
//            replaceTable(table3, genMap(Strings.TABLE_AV, String.valueOf((int)solentOfWssx+"%")));
//            
//            replaceTable(table3, genMap(Strings.TABLE_AW, String.valueOf((int)bournemouthOfPSU+"%")));
//            replaceTable(table3, genMap(Strings.TABLE_AX, String.valueOf((int)dorchesterOfPSU+"%")));
//            replaceTable(table3, genMap(Strings.TABLE_AY, String.valueOf((int)dorsetOfPSU+"%")));
//            replaceTable(table3, genMap(Strings.TABLE_AZ, String.valueOf((int)hhtfOfPSU+"%")));
//            replaceTable(table3, genMap(Strings.TABLE_BA, String.valueOf((int)iowOfPSU+"%")));
//            replaceTable(table3, genMap(Strings.TABLE_BB, String.valueOf((int)jerseyOfPSU+"%")));
//            replaceTable(table3, genMap(Strings.TABLE_BC, String.valueOf((int)pooleOfPSU+"%")));
//            replaceTable(table3, genMap(Strings.TABLE_BD, String.valueOf((int)portsmouthOfPSU+"%")));
//            replaceTable(table3, genMap(Strings.TABLE_BE, String.valueOf((int)salisburyOfPSU+"%")));
//            replaceTable(table3, genMap(Strings.TABLE_BF, String.valueOf((int)solentOfPSU+"%")));
//            replaceTable(table3, genMap(Strings.TABLE_BG, String.valueOf((int)southamptonOfPSU+"%")));
//            replaceTable(table3, genMap(Strings.TABLE_BH, String.valueOf((int)southernOfPSU+"%")));
//            
//            //table 4
//            List<Integer> t4list = helper.countTable4();
//            
//            double anaestheticsRefNo = t4list.get(0);
//            double dentalRefNo = t4list.get(1);
//            double emergRefNo = t4list.get(2);
//            double foundationRefNo = t4list.get(3);
//            double gpRefNo = t4list.get(4);
//            double medicineRefNo = t4list.get(5);
//            double obsRefNo = t4list.get(6);
//            double occhealthRefNo = t4list.get(7);
//            double paediatricsRefNo = t4list.get(8);
//            double pathologyRefNo = t4list.get(9);
//            double pharmacyRefNo = t4list.get(10);
//            double psychRefNo = t4list.get(11);
//            double pubhealthRefNo = t4list.get(12);
//            double radioRefNo = t4list.get(13);
//            double surgeryRefNo = t4list.get(14);
//            
//            double anaestheticsTotal = t4list.get(15);
//            double dentalTotal = t4list.get(16);
//            double emergTotal = t4list.get(17);
//            double foundationTotal = t4list.get(18);
//            double gpTotal = t4list.get(19);
//            double medicineTotal = t4list.get(20);
//            double obsTotal = t4list.get(21);
//            double occhealthTotal = t4list.get(22);
//            double paediatricsTotal = t4list.get(23);
//            double pathologyTotal = t4list.get(24);
//            double pharmacyTotal = t4list.get(25);
//            double psychTotal = t4list.get(26);
//            double pubhealthTotal = t4list.get(27);
//            double radioTotal = t4list.get(28);
//            double surgeryTotal = t4list.get(29);
//            
//            double anaestheticsInSpc = Math.round((anaestheticsRefNo/anaestheticsTotal)*100);
//            double dentalInSpc = Math.round((dentalRefNo/dentalTotal)*100);
//            double emergInSpc = Math.round((emergRefNo/emergTotal)*100);
//            double foundationInSpc = Math.round((foundationRefNo/foundationTotal)*100);
//            double gpInSpc = Math.round((gpRefNo/gpTotal)*100);
//            double medicineInSpc = Math.round((medicineRefNo/medicineTotal)*100);
//            double obsInSpc = Math.round((obsRefNo/obsTotal)*100);
//            double occhealthInSpc = Math.round((occhealthRefNo/occhealthTotal)*100);
//            double paediatricsInSpc = Math.round((paediatricsRefNo/paediatricsTotal)*100);
//            double pathologyInSpc = Math.round((pathologyRefNo/pathologyTotal)*100);
//            double pharmacyInSpc = Math.round((pharmacyRefNo/pharmacyTotal)*100);
//            double psychInSpc = Math.round((psychRefNo/psychTotal)*100);
//            double pubhealthInSpc = Math.round((pubhealthRefNo/pubhealthTotal)*100);
//            double radioInSpc = Math.round((radioRefNo/radioTotal)*100);
//            double surgeryInSpc = Math.round((surgeryRefNo/surgeryTotal)*100);
//            
//            double anaestheticsOfWssx = Math.round((anaestheticsTotal/totalWessex)*100);
//            double dentalOfWssx = Math.round((dentalTotal/totalWessex)*100);
//            double emergOfWssx = Math.round((emergTotal/totalWessex)*100);
//            double foundationOfWssx = Math.round((foundationTotal/totalWessex)*100);
//            double gpOfWssx = Math.round((gpTotal/totalWessex)*100);
//            double medicineOfWssx = Math.round((medicineTotal/totalWessex)*100);
//            double obsOfWssx = Math.round((obsTotal/totalWessex)*100);
//            double occhealthOfWssx = Math.round((occhealthTotal/totalWessex)*100);
//            double paediatricsOfWssx = Math.round((paediatricsTotal/totalWessex)*100);
//            double pathologyOfWssx = Math.round((pathologyTotal/totalWessex)*100);
//            double pharmacyOfWssx = Math.round((pharmacyTotal/totalWessex)*100);
//            double psychOfWssx = Math.round((psychTotal/totalWessex)*100);
//            double pubhealthOfWssx = Math.round((pubhealthTotal/totalWessex)*100);
//            double radioOfWssx = Math.round((radioTotal/totalWessex)*100);
//            double surgeryOfWssx = Math.round((surgeryTotal/totalWessex)*100);
//            
//            double anaestheticsOfPSU = Math.round((anaestheticsRefNo/totalRefs)*100);
//            double dentalOfPSU = Math.round((dentalRefNo/totalRefs)*100);
//            double emergOfPSU = Math.round((emergRefNo/totalRefs)*100);
//            double foundationOfPSU = Math.round((foundationRefNo/totalRefs)*100);
//            double gpOfPSU = Math.round((gpRefNo/totalRefs)*100);
//            double medicineOfPSU = Math.round((medicineRefNo/totalRefs)*100);
//            double obsOfPSU = Math.round((obsRefNo/totalRefs)*100);
//            double occhealthOfPSU = Math.round((occhealthRefNo/totalRefs)*100);
//            double paediatricsOfPSU = Math.round((paediatricsRefNo/totalRefs)*100);
//            double pathologyOfPSU = Math.round((pathologyRefNo/totalRefs)*100);
//            double pharmacyOfPSU = Math.round((pharmacyRefNo/totalRefs)*100);
//            double psychOfPSU = Math.round((psychRefNo/totalRefs)*100);
//            double pubhealthOfPSU = Math.round((pubhealthRefNo/totalRefs)*100);
//            double radioOfPSU = Math.round((radioRefNo/totalRefs)*100);
//            double surgeryOfPSU = Math.round((surgeryRefNo/totalRefs)*100);
//            
//            
//            XWPFTable table4 = doc.getTables().get(4);
//            
//            replaceTable(table4, genMap(Strings.TABLE_A, String.valueOf((int)anaestheticsRefNo)));
//            replaceTable(table4, genMap(Strings.TABLE_B, String.valueOf((int)dentalRefNo)));
//            replaceTable(table4, genMap(Strings.TABLE_C, String.valueOf((int)emergRefNo)));
//            replaceTable(table4, genMap(Strings.TABLE_D, String.valueOf((int)foundationRefNo)));
//            replaceTable(table4, genMap(Strings.TABLE_E, String.valueOf((int)gpRefNo)));
//            replaceTable(table4, genMap(Strings.TABLE_F, String.valueOf((int)medicineRefNo)));
//            replaceTable(table4, genMap(Strings.TABLE_G, String.valueOf((int)obsRefNo)));
//            replaceTable(table4, genMap(Strings.TABLE_H, String.valueOf((int)occhealthRefNo)));
//            replaceTable(table4, genMap(Strings.TABLE_I, String.valueOf((int)paediatricsRefNo)));
//            replaceTable(table4, genMap(Strings.TABLE_J, String.valueOf((int)pathologyRefNo)));
//            replaceTable(table4, genMap(Strings.TABLE_K, String.valueOf((int)pharmacyRefNo)));
//            replaceTable(table4, genMap(Strings.TABLE_L, String.valueOf((int)psychRefNo)));
//            replaceTable(table4, genMap(Strings.TABLE_M, String.valueOf((int)pubhealthRefNo)));
//            replaceTable(table4, genMap(Strings.TABLE_N, String.valueOf((int)radioRefNo)));
//            replaceTable(table4, genMap(Strings.TABLE_O, String.valueOf((int)surgeryRefNo)));
//            
//            replaceTable(table4, genMap(Strings.TABLE_P, String.valueOf((int)anaestheticsTotal)));
//            replaceTable(table4, genMap(Strings.TABLE_Q, String.valueOf((int)dentalTotal)));
//            replaceTable(table4, genMap(Strings.TABLE_R, String.valueOf((int)emergTotal)));
//            replaceTable(table4, genMap(Strings.TABLE_S, String.valueOf((int)foundationTotal)));
//            replaceTable(table4, genMap(Strings.TABLE_T, String.valueOf((int)gpTotal)));
//            replaceTable(table4, genMap(Strings.TABLE_U, String.valueOf((int)medicineTotal)));
//            replaceTable(table4, genMap(Strings.TABLE_V, String.valueOf((int)obsTotal)));
//            replaceTable(table4, genMap(Strings.TABLE_W, String.valueOf((int)occhealthTotal)));
//            replaceTable(table4, genMap(Strings.TABLE_X, String.valueOf((int)paediatricsTotal)));
//            replaceTable(table4, genMap(Strings.TABLE_Y, String.valueOf((int)pathologyTotal)));
//            replaceTable(table4, genMap(Strings.TABLE_Z, String.valueOf((int)pharmacyTotal)));
//            replaceTable(table4, genMap(Strings.TABLE_AA, String.valueOf((int)psychTotal)));
//            replaceTable(table4, genMap(Strings.TABLE_AB, String.valueOf((int)pubhealthTotal)));
//            replaceTable(table4, genMap(Strings.TABLE_AC, String.valueOf((int)radioTotal)));
//            replaceTable(table4, genMap(Strings.TABLE_AD, String.valueOf((int)surgeryTotal)));
//            
//            replaceTable(table4, genMap(Strings.TABLE_AE, String.valueOf((int)anaestheticsInSpc)+"%"));
//            replaceTable(table4, genMap(Strings.TABLE_AF, String.valueOf((int)dentalInSpc)+"%"));
//            replaceTable(table4, genMap(Strings.TABLE_AG, String.valueOf((int)emergInSpc)+"%"));
//            replaceTable(table4, genMap(Strings.TABLE_AH, String.valueOf((int)foundationInSpc)+"%"));
//            replaceTable(table4, genMap(Strings.TABLE_AI, String.valueOf((int)gpInSpc)+"%"));
//            replaceTable(table4, genMap(Strings.TABLE_AJ, String.valueOf((int)medicineInSpc)+"%"));
//            replaceTable(table4, genMap(Strings.TABLE_AK, String.valueOf((int)obsInSpc)+"%"));
//            replaceTable(table4, genMap(Strings.TABLE_AL, String.valueOf((int)occhealthInSpc)+"%"));
//            replaceTable(table4, genMap(Strings.TABLE_AM, String.valueOf((int)paediatricsInSpc)+"%"));
//            replaceTable(table4, genMap(Strings.TABLE_AN, String.valueOf((int)pathologyInSpc)+"%"));
//            replaceTable(table4, genMap(Strings.TABLE_AO, String.valueOf((int)pharmacyInSpc)+"%"));
//            replaceTable(table4, genMap(Strings.TABLE_AP, String.valueOf((int)psychInSpc)+"%"));
//            replaceTable(table4, genMap(Strings.TABLE_AQ, String.valueOf((int)pubhealthInSpc)+"%"));
//            replaceTable(table4, genMap(Strings.TABLE_AR, String.valueOf((int)radioInSpc)+"%"));
//            replaceTable(table4, genMap(Strings.TABLE_AS, String.valueOf((int)surgeryInSpc)+"%"));
//            
//            replaceTable(table4, genMap(Strings.TABLE_AT, String.valueOf((int)anaestheticsOfWssx)+"%"));
//            replaceTable(table4, genMap(Strings.TABLE_AU, String.valueOf((int)dentalOfWssx)+"%"));
//            replaceTable(table4, genMap(Strings.TABLE_AV, String.valueOf((int)emergOfWssx)+"%"));
//            replaceTable(table4, genMap(Strings.TABLE_AW, String.valueOf((int)foundationOfWssx)+"%"));
//            replaceTable(table4, genMap(Strings.TABLE_AX, String.valueOf((int)gpOfWssx)+"%"));
//            replaceTable(table4, genMap(Strings.TABLE_AY, String.valueOf((int)medicineOfWssx)+"%"));
//            replaceTable(table4, genMap(Strings.TABLE_AZ, String.valueOf((int)obsOfWssx)+"%"));
//            replaceTable(table4, genMap(Strings.TABLE_BA, String.valueOf((int)occhealthOfWssx)+"%"));
//            replaceTable(table4, genMap(Strings.TABLE_BB, String.valueOf((int)paediatricsOfWssx)+"%"));
//            replaceTable(table4, genMap(Strings.TABLE_BC, String.valueOf((int)pathologyOfWssx)+"%"));
//            replaceTable(table4, genMap(Strings.TABLE_BD, String.valueOf((int)pharmacyOfWssx)+"%"));
//            replaceTable(table4, genMap(Strings.TABLE_BE, String.valueOf((int)psychOfWssx)+"%"));
//            replaceTable(table4, genMap(Strings.TABLE_BF, String.valueOf((int)pubhealthOfWssx)+"%"));
//            replaceTable(table4, genMap(Strings.TABLE_BG, String.valueOf((int)radioOfWssx)+"%"));
//            replaceTable(table4, genMap(Strings.TABLE_BH, String.valueOf((int)surgeryOfWssx)+"%"));
//            
//            replaceTable(table4, genMap(Strings.TABLE_BI, String.valueOf((int)anaestheticsOfPSU)+"%"));
//            replaceTable(table4, genMap(Strings.TABLE_BJ, String.valueOf((int)dentalOfPSU)+"%"));
//            replaceTable(table4, genMap(Strings.TABLE_BK, String.valueOf((int)emergOfPSU)+"%"));
//            replaceTable(table4, genMap(Strings.TABLE_BL, String.valueOf((int)foundationOfPSU)+"%"));
//            replaceTable(table4, genMap(Strings.TABLE_BM, String.valueOf((int)gpOfPSU)+"%"));
//            replaceTable(table4, genMap(Strings.TABLE_BN, String.valueOf((int)medicineOfPSU)+"%"));
//            replaceTable(table4, genMap(Strings.TABLE_BO, String.valueOf((int)obsOfPSU)+"%"));
//            replaceTable(table4, genMap(Strings.TABLE_BP, String.valueOf((int)occhealthOfPSU)+"%"));
//            replaceTable(table4, genMap(Strings.TABLE_BQ, String.valueOf((int)paediatricsOfPSU)+"%"));
//            replaceTable(table4, genMap(Strings.TABLE_BR, String.valueOf((int)pathologyOfPSU)+"%"));
//            replaceTable(table4, genMap(Strings.TABLE_BS, String.valueOf((int)pharmacyOfPSU)+"%"));
//            replaceTable(table4, genMap(Strings.TABLE_BT, String.valueOf((int)psychOfPSU)+"%"));
//            replaceTable(table4, genMap(Strings.TABLE_BU, String.valueOf((int)pubhealthOfPSU)+"%"));
//            replaceTable(table4, genMap(Strings.TABLE_BV, String.valueOf((int)radioOfPSU)+"%"));
//            replaceTable(table4, genMap(Strings.TABLE_BW, String.valueOf((int)surgeryOfPSU)+"%"));
//            
//            
//            //table 5
//            String lastLn = "0";
//            List<String> t5bournemouthList;
//            List<String> t5dorchesterList;
//            List<String> t5dorsetList;
//            List<String> t5hhftList;
//            List<String> t5iowList;
//            List<String> t5jerseyList;
//            List<String> t5pooleList;
//            List<String> t5portsmouthList;
//            List<String> t5salisburyList;
//            List<String> t5southamptonList;
//            List<String> t5solentList;
//            List<String> t5southernList;
//            
//            XWPFTable table5 = doc.getTables().get(5);
//            
//            t5bournemouthList  = helper.getTable5Line(Strings.PSW_TRUST_BOURNEMOUTH, lastLn);
//            insertLineT5(table5, t5bournemouthList, Strings.PSW_TRUST_BOURNEMOUTH);
//            
//            t5dorchesterList = helper.getTable5Line(Strings.PSW_TRUST_DORCHESTER, lastLn);
//            insertLineT5(table5, t5dorchesterList, Strings.PSW_TRUST_DORCHESTER);
//            
//            t5dorsetList = helper.getTable5Line(Strings.PSW_TRUST_DORSET, lastLn);
//            insertLineT5(table5, t5dorsetList, Strings.PSW_TRUST_DORSET);
//            
//            t5hhftList = helper.getTable5Line(Strings.PSW_TRUST_HHFT, lastLn);
//            insertLineT5(table5, t5hhftList, Strings.PSW_TRUST_HHFT);
//            
//            t5iowList = helper.getTable5Line(Strings.PSW_TRUST_IOW, lastLn);
//            insertLineT5(table5, t5iowList, Strings.PSW_TRUST_IOW);
//            
//            t5jerseyList = helper.getTable5Line(Strings.PSW_TRUST_JERSEY, lastLn);
//            insertLineT5(table5, t5jerseyList, Strings.PSW_TRUST_JERSEY);
//            
//            t5pooleList = helper.getTable5Line(Strings.PSW_TRUST_POOLE, lastLn);
//            insertLineT5(table5, t5pooleList, Strings.PSW_TRUST_POOLE);
//            
//            t5portsmouthList = helper.getTable5Line(Strings.PSW_TRUST_PORTSMOUTH, lastLn);
//            insertLineT5(table5, t5portsmouthList, Strings.PSW_TRUST_PORTSMOUTH);
//            
//            t5salisburyList = helper.getTable5Line(Strings.PSW_TRUST_SALISBURY, lastLn);
//            insertLineT5(table5, t5salisburyList, Strings.PSW_TRUST_SALISBURY);
//            
//            t5solentList = helper.getTable5Line(Strings.PSW_TRUST_SOLENT, lastLn);
//            insertLineT5(table5, t5solentList, Strings.PSW_TRUST_SOLENT);
//            
//            t5southamptonList = helper.getTable5Line(Strings.PSW_TRUST_SOUTHAMPTON, lastLn);
//            insertLineT5(table5, t5southamptonList, Strings.PSW_TRUST_SOUTHAMPTON);
//            
//            t5southernList = helper.getTable5Line(Strings.PSW_TRUST_SOUTHERN, lastLn);
//            insertLineT5(table5, t5dorchesterList, Strings.PSW_TRUST_SOUTHERN);
//            
//            table5.removeRow(1);
//            table5.removeRow(1);
//            

            
            //table 6
            
            XWPFTable table6 = doc.getTables().get(6);
            
            replaceTable(table6, genMap(Strings.TABLE_A, String.valueOf(yrPastCurrentYYYY)));
            
            doc.write(new FileOutputStream(outFile));
            
            
        } catch (FileNotFoundException ex) {
            Logger.getLogger(DocHelper.class.getName()).log(Level.SEVERE, null, ex);
        } catch (IOException ex) {
            Logger.getLogger(DocHelper.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
    
    private static void writeSmallCell(XWPFTableCell cell, String str){
        XWPFRun run = cell.getParagraphs().get(0).getRuns().get(0);
        run.setFontSize(8);
        run.setText(str);
    }
    
    /**
     * Generates the output file based on the template
     * @param inFile
     * @param outFile
     */
    public static void genDocOut(File inFile, File outFile){
        
        try {
            XWPFDocument docOut = new XWPFDocument(new FileInputStream(inFile));
            docOut.write(new FileOutputStream(outFile));
        } catch (FileNotFoundException ex) {
            Logger.getLogger(DocHelper.class.getName()).log(Level.SEVERE, null, ex);
        } catch (IOException ex) {
            Logger.getLogger(DocHelper.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
    
    
    
    
    /**
     * Saves the document
     * @param inFile
     * @param outFile
     */
    public void saveDoc (){
        
        
        try {
            outFile.createNewFile();
        } catch (IOException ex) {
            Logger.getLogger(DocHelper.class.getName()).log(Level.SEVERE, null, ex);
        }
        
        
    }
    
    /**
     * Replaces the text of a run within a cell
     * @param table
     * @param m
     */
    private static void replaceTable(XWPFTable table,HashMap m) {
        for (XWPFTableRow row : table.getRows()) {
            for (XWPFTableCell cell : row.getTableCells()) {
                for(XWPFParagraph p:cell.getParagraphs()){
                    replaceParagraphInCell(p, m);
                }
            }
        }
    }
    
    
    private static void replaceTableSmall(XWPFTable table,HashMap m) {
        for (XWPFTableRow row : table.getRows()) {
            for (XWPFTableCell cell : row.getTableCells()) {
                for(XWPFParagraph p:cell.getParagraphs()){
                    replaceParagraphInCellSmall(p, m);
                }
            }
        }
    }
    
    /**
     * Replaces the text found within the runs of a paragraph
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
    
    private static void replaceParagraph(XWPFParagraph paragraph, HashMap m){
        for (XWPFRun r : paragraph.getRuns()) {
            String text = r.getText(r.getTextPosition());
            if (text != null && text.contains(m.get("a").toString())) {
                text = text.replace(m.get("a").toString(), m.get("b").toString());
                r.setText(text,0);
            }
        }
    }
    
    /**
     * Finds a given value within the doc and replaces it
     * @param m
     * @param outFile
     */
    public static void findAndReplace(HashMap m,File outFile){
        
        try {
            
            XWPFDocument xdoc = new XWPFDocument(new FileInputStream(outFile));
            for (XWPFParagraph p : xdoc.getParagraphs()) {
                replaceParagraph(p, m);
            }
            xdoc.write(new FileOutputStream(outFile));
        } catch (FileNotFoundException ex) {
            Logger.getLogger(DocHelper.class.getName()).log(Level.SEVERE, null, ex);
        } catch (IOException ex) {
            Logger.getLogger(DocHelper.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
    
    private void insertLineT5(XWPFTable table, List<String> list, String trust){
        
        PoiHelper helper = new PoiHelper(psw, graphs);
        
        String gender = list.get(0);
        String grade = list.get(1);
        String school = list.get(2);
        String addRef = list.get(3);
        String country = list.get(4);
        String age = list.get(5);
        String ethnicity = list.get(6);
        String sexOr = list.get(7);
        String religion = list.get(8);
        String disability = list.get(9);
        String lastLine = list.get(10);
        
        int tempateRowFId = 1;
        XWPFTableRow rowTemplateF = table.getRow(tempateRowFId);
        int tempateRowMId = 2;
        XWPFTableRow rowTemplateM = table.getRow(tempateRowMId);
        
        if(!gender.equals("")){
            if(gender.equals("F")){
                
                XWPFTableRow oldRow = rowTemplateF;
                CTRow firstCTRow = null;
                try {
                    firstCTRow = CTRow.Factory.parse(oldRow.getCtRow().newInputStream());
                } catch (XmlException | IOException e) {e.printStackTrace();
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
                writeSmallCell(cell6, age);
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
            
            else if(gender.equals("M")){
                
                XWPFTableRow oldRow = rowTemplateM;
                CTRow firstCTRow = null;
                try {
                    firstCTRow = CTRow.Factory.parse(oldRow.getCtRow().newInputStream());
                } catch (XmlException | IOException e) {e.printStackTrace();
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
                writeSmallCell(cell6, age);
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
            
            while(!gender.equals("")){
                
                list = helper.getTable5Line(trust, lastLine);
                
                gender = list.get(0);
                grade = list.get(1);
                school = list.get(2);
                addRef = list.get(3);
                country = list.get(4);
                age = list.get(5);
                ethnicity = list.get(6);
                sexOr = list.get(7);
                religion = list.get(8);
                disability = list.get(9);
                lastLine = list.get(10);
                
                if(gender.equals("")){
                    break;
                }
                
                if(gender.equals("F")){
                    
                    CTRow newCTRow = null;
                    XWPFTableRow oldRow = rowTemplateF;
                    try {
                        newCTRow = CTRow.Factory.parse(oldRow.getCtRow().newInputStream());
                    } catch (XmlException | IOException e) {e.printStackTrace();
                    }
                    XWPFTableRow newRow = new XWPFTableRow(newCTRow, table);
                    
                    XWPFTableCell newcell1 = newRow.getCell(1);
                    writeSmallCell(newcell1, gender);
                    XWPFTableCell newcell2 = newRow.getCell(2);
                    writeSmallCell(newcell2, grade);
                    XWPFTableCell newcell3 = newRow.getCell(3);
                    writeSmallCell(newcell3, school);
                    XWPFTableCell newcell4 = newRow.getCell(4);
                    writeSmallCell(newcell4, addRef);
                    XWPFTableCell newcell5 = newRow.getCell(5);
                    writeSmallCell(newcell5, country);
                    XWPFTableCell newcell6 = newRow.getCell(6);
                    writeSmallCell(newcell6, age);
                    XWPFTableCell newcell7 = newRow.getCell(7);
                    writeSmallCell(newcell7, ethnicity);
                    XWPFTableCell newcell8 = newRow.getCell(8);
                    writeSmallCell(newcell8, sexOr);
                    XWPFTableCell newcell9 = newRow.getCell(9);
                    writeSmallCell(newcell9, religion);
                    XWPFTableCell newcell10 = newRow.getCell(10);
                    writeSmallCell(newcell10, disability);
                    
                    table.addRow(newRow);
                }
                
                else if(gender.equals("M")){
                    
                    CTRow newCTRow = null;
                    XWPFTableRow oldRow = rowTemplateM;
                    try {
                        newCTRow = CTRow.Factory.parse(oldRow.getCtRow().newInputStream());
                    } catch (XmlException | IOException e) {e.printStackTrace();
                    }
                    XWPFTableRow newRow = new XWPFTableRow(newCTRow, table);
                    
                    XWPFTableCell newcell1 = newRow.getCell(1);
                    writeSmallCell(newcell1, gender);
                    XWPFTableCell newcell2 = newRow.getCell(2);
                    writeSmallCell(newcell2, grade);
                    XWPFTableCell newcell3 = newRow.getCell(3);
                    writeSmallCell(newcell3, school);
                    XWPFTableCell newcell4 = newRow.getCell(4);
                    writeSmallCell(newcell4, addRef);
                    XWPFTableCell newcell5 = newRow.getCell(5);
                    writeSmallCell(newcell5, country);
                    XWPFTableCell newcell6 = newRow.getCell(6);
                    writeSmallCell(newcell6, age);
                    XWPFTableCell newcell7 = newRow.getCell(7);
                    writeSmallCell(newcell7, ethnicity);
                    XWPFTableCell newcell8 = newRow.getCell(8);
                    writeSmallCell(newcell8, sexOr);
                    XWPFTableCell newcell9 = newRow.getCell(9);
                    writeSmallCell(newcell9, religion);
                    XWPFTableCell newcell10 = newRow.getCell(10);
                    writeSmallCell(newcell10, disability);
                    
                    table.addRow(newRow);
                }
            }
        }
    }
    
    /**
     * Generates the map which will be used to find and replace
     * @param a
     * @param b
     * @return
     */
    public static HashMap<String, String> genMap(String a, String b){
        
        HashMap<String, String> hash = new HashMap<>();
        hash.put("a", a);
        hash.put("b", b);
        return hash;
        
    }
    
    /**
     * Returns the actual year
     * @return
     */
    public static int getYear(){
        int year = Calendar.getInstance().get(Calendar.YEAR);
        return year;
    }
    
    public static String genSplitYearSeq(String str){
        List<String> strList = new ArrayList<>();
        
        
        for(String s:str.split("")){
            strList.add(s);
        }
        
        String finalStr = strList.get(2)+strList.get(3);
        return finalStr;
    }
}
