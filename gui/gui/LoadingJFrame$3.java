/*
 * Decompiled with CFR 0.152.
 */
package gui;

import facade.DocHelper;
import facade.GcExcelHelper;
import facade.PoiHelper;
import gui.LoadingJFrame;
import java.util.ArrayList;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.JOptionPane;
import javax.swing.SwingWorker;
import vo.ReferralRecord;

class LoadingJFrame.3
extends SwingWorker<Void, String> {
    LoadingJFrame.3() {
    }

    @Override
    protected Void doInBackground() {
        try {
            Thread.sleep(100L);
            this.publish("Reading database...");
            LoadingJFrame.this.poiHelper = new PoiHelper(LoadingJFrame.this.psw, LoadingJFrame.this.graphs);
            LoadingJFrame.this.poiHelper.genGraphFile();
            ArrayList<ReferralRecord> records = LoadingJFrame.this.poiHelper.readReferralRecords();
            LoadingJFrame.this.poiHelper.setRecordList(records);
            LoadingJFrame.this.docHelper = new DocHelper(LoadingJFrame.this.psw, LoadingJFrame.this.graphs, LoadingJFrame.this.doc, records);
            LoadingJFrame.this.gcExcelHelper = new GcExcelHelper(LoadingJFrame.this.graphs);
            LoadingJFrame.this.poiHelper.genLogFile();
            this.publish("Generating Graphs.xlsx output");
            LoadingJFrame.this.poiHelper.genGraphFile();
            this.publish("Generating sheet for graph 1");
            LoadingJFrame.this.poiHelper.getGraph0();
            this.publish("Generating sheet for graph 2");
            LoadingJFrame.this.poiHelper.getGraph1();
            this.publish("Generating sheet for graph 3");
            LoadingJFrame.this.poiHelper.getGraph2();
            this.publish("Generating sheet for graph 4");
            LoadingJFrame.this.poiHelper.getGraph3();
            this.publish("Generating sheet for graph 5");
            LoadingJFrame.this.poiHelper.getGraph4();
            this.publish("Generating sheet for graph 6");
            LoadingJFrame.this.poiHelper.getGraph5();
            this.publish("Generating sheet for graph 7");
            LoadingJFrame.this.poiHelper.getGraph6();
            this.publish("Generating sheet for graph 8");
            LoadingJFrame.this.poiHelper.getGraph7();
            this.publish("Generating sheet for graph 9");
            LoadingJFrame.this.poiHelper.getGraph8();
            this.publish("Generating sheet for graph 10");
            LoadingJFrame.this.poiHelper.getGraph9();
            this.publish("Generating sheet for graph 11");
            LoadingJFrame.this.poiHelper.getGraph10();
            this.publish("Generating sheet for graph 12");
            LoadingJFrame.this.poiHelper.getGraph11();
            this.publish("OverWriting Graphs.xlsx");
            LoadingJFrame.this.poiHelper.writeToGraphs();
            this.publish("Creating charts");
            LoadingJFrame.this.gcExcelHelper.genGraphs();
            LoadingJFrame.this.poiHelper.deleteLastSheetGraphs();
            this.publish("Generating Output.docx output");
            LoadingJFrame.this.docHelper.genDocFile();
            this.publish("Generating table 1");
            LoadingJFrame.this.docHelper.getTable0();
            this.publish("Generating table 2");
            LoadingJFrame.this.docHelper.getTable1();
            this.publish("Generating table 3");
            LoadingJFrame.this.docHelper.getTable2();
            this.publish("Generating table 4");
            LoadingJFrame.this.docHelper.getTable3();
            this.publish("Generating table 5");
            LoadingJFrame.this.docHelper.getTable4();
            this.publish("Generating table 6");
            LoadingJFrame.this.docHelper.getTable5();
            this.publish("Generating table 8");
            LoadingJFrame.this.docHelper.getTable7();
            this.publish("Generating table 9");
            LoadingJFrame.this.docHelper.getTable8();
            this.publish("Generating table 10");
            LoadingJFrame.this.docHelper.getTable9();
            this.publish("Updating text...");
            LoadingJFrame.this.docHelper.doTextUpdate();
            this.publish("OverWriting Output.docx");
            LoadingJFrame.this.docHelper.saveDoc();
            this.publish("Program finished - Result OK");
        }
        catch (Exception e) {
            Logger.getLogger(LoadingJFrame.class.getName()).log(Level.SEVERE, null, e);
        }
        return null;
    }

    @Override
    protected void process(List<String> chunks) {
        for (String text : chunks) {
            LoadingJFrame.this.jTextArea1.append(text + "\n");
        }
    }

    @Override
    protected void done() {
        LoadingJFrame.this.jButton1.setEnabled(true);
        JOptionPane.showMessageDialog(null, "Output saved", "NHS Education - Report generation tool", 1);
    }
}
