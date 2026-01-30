/*
 * Decompiled with CFR 0.152.
 */
package gui;

import facade.DocHelper;
import facade.GcExcelHelper;
import facade.PoiHelper;
import gui.MainJFrame;
import gui.MessageDialogs;
import java.awt.EventQueue;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;
import java.io.File;
import java.util.ArrayList;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.GroupLayout;
import javax.swing.JButton;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JScrollPane;
import javax.swing.JTextArea;
import javax.swing.LayoutStyle;
import javax.swing.SwingWorker;
import javax.swing.UIManager;
import javax.swing.UnsupportedLookAndFeelException;
import vo.ReferralRecord;

public class LoadingJFrame
extends JFrame {
    JTextArea jta;
    PoiHelper poiHelper;
    GcExcelHelper gcExcelHelper;
    MainJFrame main;
    File psw;
    File graphs;
    File doc;
    DocHelper docHelper;
    private JButton jButton1;
    private JLabel jLabel1;
    private JScrollPane jScrollPane1;
    private JTextArea jTextArea1;

    public LoadingJFrame() {
        try {
            for (UIManager.LookAndFeelInfo info : UIManager.getInstalledLookAndFeels()) {
                if (!"Windows".equals(info.getName())) continue;
                UIManager.setLookAndFeel(info.getClassName());
                break;
            }
            this.addWindowListener(new WindowAdapter(){

                @Override
                public void windowClosing(WindowEvent e) {
                    LoadingJFrame.this.pleaseHold();
                }
            });
            this.initComponents();
        }
        catch (ClassNotFoundException | IllegalAccessException | InstantiationException | UnsupportedLookAndFeelException ex) {
            Logger.getLogger(LoadingJFrame.class.getName()).log(Level.SEVERE, null, ex);
        }
    }

    public void run(MainJFrame main) {
        this.main = main;
        ArrayList<File> fileList = main.getFileList();
        this.doc = fileList.get(0);
        this.graphs = fileList.get(1);
        this.psw = fileList.get(2);
        this.gcExcelHelper = new GcExcelHelper(this.graphs);
        this.genReport();
    }

    private void initComponents() {
        this.jLabel1 = new JLabel();
        this.jScrollPane1 = new JScrollPane();
        this.jTextArea1 = new JTextArea();
        this.jButton1 = new JButton();
        this.setDefaultCloseOperation(0);
        this.setTitle("NHS Education - Loading...");
        this.setResizable(false);
        this.jLabel1.setText("Loading, please hold...");
        this.jTextArea1.setEditable(false);
        this.jTextArea1.setColumns(20);
        this.jTextArea1.setRows(5);
        this.jScrollPane1.setViewportView(this.jTextArea1);
        this.jButton1.setText("Finish");
        this.jButton1.setEnabled(false);
        this.jButton1.addMouseListener(new MouseAdapter(){

            @Override
            public void mouseClicked(MouseEvent evt) {
                LoadingJFrame.this.jButton1MouseClicked(evt);
            }
        });
        GroupLayout layout = new GroupLayout(this.getContentPane());
        this.getContentPane().setLayout(layout);
        layout.setHorizontalGroup(layout.createParallelGroup(GroupLayout.Alignment.LEADING).addGroup(layout.createSequentialGroup().addContainerGap().addGroup(layout.createParallelGroup(GroupLayout.Alignment.LEADING).addComponent(this.jScrollPane1, -1, 380, Short.MAX_VALUE).addGroup(layout.createSequentialGroup().addComponent(this.jLabel1).addGap(0, 0, Short.MAX_VALUE)).addGroup(GroupLayout.Alignment.TRAILING, layout.createSequentialGroup().addGap(0, 0, Short.MAX_VALUE).addComponent(this.jButton1))).addContainerGap()));
        layout.setVerticalGroup(layout.createParallelGroup(GroupLayout.Alignment.LEADING).addGroup(layout.createSequentialGroup().addContainerGap().addComponent(this.jLabel1).addPreferredGap(LayoutStyle.ComponentPlacement.RELATED).addComponent(this.jScrollPane1, -2, 120, -2).addPreferredGap(LayoutStyle.ComponentPlacement.UNRELATED).addComponent(this.jButton1).addContainerGap(-1, Short.MAX_VALUE)));
        this.pack();
    }

    private void jButton1MouseClicked(MouseEvent evt) {
        this.setVisible(false);
        this.main.setEnabled(true);
        this.main.setVisible(true);
    }

    private void genReport() {
        SwingWorker<Void, String> worker = new SwingWorker<Void, String>(){

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
        };
        worker.execute();
    }

    public static void main(String[] args) {
        try {
            for (UIManager.LookAndFeelInfo info : UIManager.getInstalledLookAndFeels()) {
                if (!"Nimbus".equals(info.getName())) continue;
                UIManager.setLookAndFeel(info.getClassName());
                break;
            }
        }
        catch (ClassNotFoundException ex) {
            Logger.getLogger(LoadingJFrame.class.getName()).log(Level.SEVERE, null, ex);
        }
        catch (InstantiationException ex) {
            Logger.getLogger(LoadingJFrame.class.getName()).log(Level.SEVERE, null, ex);
        }
        catch (IllegalAccessException ex) {
            Logger.getLogger(LoadingJFrame.class.getName()).log(Level.SEVERE, null, ex);
        }
        catch (UnsupportedLookAndFeelException ex) {
            Logger.getLogger(LoadingJFrame.class.getName()).log(Level.SEVERE, null, ex);
        }
        EventQueue.invokeLater(new Runnable(){

            @Override
            public void run() {
                new LoadingJFrame().setVisible(true);
            }
        });
    }

    private void pleaseHold() {
        if (this.jButton1.isEnabled()) {
            this.dispose();
        } else {
            MessageDialogs md = new MessageDialogs();
            md.showPleaseHold();
        }
    }
}
