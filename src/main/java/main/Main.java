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
import java.util.ArrayList;
import java.util.Calendar;
import java.util.HashMap;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.JFrame;
import javax.swing.SwingUtilities;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import res.Strings;
import vo.RefRecord;

/**
 *
 * @author diego
 */
public class Main {
    
    public static void main(String[] args) {
        
//        SwingUtilities.invokeLater(new Runnable(){
//            @Override
//            public void run() {
//                MainJFrame frame = new MainJFrame();
//                frame.setVisible(true);
//                frame.setLocationRelativeTo(null);
//                frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);            }
//            
//        });

    File psw = new File(Strings.PSW_PATH);
    File doc = new File (Strings.DOC_OUT_PATH);
    File graphs = new File(Strings.GRAPH_OUT_PATH);
    
    PoiHelper poiHelper = new PoiHelper(psw, graphs);
    DocHelper docHelper = new DocHelper(poiHelper, doc);
    GcExcelHelper gchelper = new GcExcelHelper(graphs);
  
    docHelper.doTextUpdate();
//    poiHelper.getRefRecords();
//    poiHelper.getGraph7();
//    docHelper.getTable0();
    docHelper.saveDoc();
////    
    

    }
}
