/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package main;

import com.grapecity.documents.excel.Workbook;
import facade.DocHelper;
import facade.GcExcelHelper;
import facade.PoiHelper;
import gui.MainJFrame;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Calendar;
import java.util.HashMap;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.JFrame;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import res.Strings;

/**
 *
 * @author diego
 */
public class Main {
    
    public static void main(String[] args) {
        
        MainJFrame frame = new MainJFrame();
        frame.setVisible(true);
        frame.setLocationRelativeTo(null);
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
//
//            int year = DocHelper.getYear();
//            int nextYear = year+1;
////            
//
////            
////            
////            
//////            PoiHelper.countOpenedAndClosed(in);
////            
//            HashMap m = DocHelper.genMap("currentYearSlash", year+"/"+nextYear);
//            DocHelper.findAndReplace(m, docOut);
////            DocHelper.getTable0();
//            PoiHelper.getGraph1();
//            PoiHelper.getGraph2();
//            PoiHelper.getGraph3();
//            PoiHelper.getGraph4();
//            PoiHelper.getGraph5();
//            PoiHelper.getGraph6();
//            PoiHelper.getGraph7();
//            PoiHelper.getGraph8();
//            PoiHelper.getGraph9();
//            PoiHelper.getGraph10();
//            PoiHelper.getGraph11();
//            GcExcelHelper.genGraphs();
//            DocHelper.getTables();

            
            
            
//        DocHelper dh = new DocHelper();
//        DocHelper.genOutput(doc, docOut);
//        dh.getTable0(docOut);
       
    }
    
}
