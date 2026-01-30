/*
 * Decompiled with CFR 0.152.
 */
package gui;

import facade.DocHelper;
import facade.GcExcelHelper;
import facade.PoiHelper;
import gui.ExcelJFileChooser;
import gui.LoadingJFrame;
import gui.MessageDialogs;
import java.awt.Color;
import java.awt.Font;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.imageio.ImageIO;
import javax.swing.GroupLayout;
import javax.swing.ImageIcon;
import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JLayeredPane;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JTextField;
import javax.swing.LayoutStyle;
import javax.swing.OverlayLayout;
import javax.swing.UIManager;
import javax.swing.UnsupportedLookAndFeelException;
import javax.swing.filechooser.FileNameExtensionFilter;

public class MainJFrame
extends JFrame {
    private InputStream logoIn;
    private ImageIcon logoIcon;
    private InputStream folderIn;
    private ImageIcon folderIcon;
    private File pswFile;
    private File traineeFile;
    private File referrerFile;
    private File mergedFile;
    private String docPath;
    private String graphPath;
    private File doc;
    private File graphs;
    private PoiHelper poiHelper;
    private DocHelper docHelper;
    private GcExcelHelper gcHelper;
    private JButton jButton1;
    private JButton jButton2;
    private JButton jButton3;
    private JButton jButton4;
    private JButton jButton5;
    private JLabel jLabel2;
    private JLabel jLabel4;
    private JLabel jLabel5;
    private JLabel jLabel6;
    private JLabel jLabel7;
    private JLabel jLabel8;
    private JLabel jLabel9;
    private JLayeredPane jLayeredPane1;
    private JPanel jPanel1;
    private JTextField jtf1;
    private JTextField jtf2;
    private JTextField jtf3;

    public MainJFrame() {
        try {
            this.logoIn = this.getClass().getResourceAsStream("/res/nhs_logo.png");
            this.logoIcon = new ImageIcon(ImageIO.read(this.logoIn));
            this.folderIn = this.getClass().getResourceAsStream("/res/folder.png");
            this.folderIcon = new ImageIcon(ImageIO.read(this.folderIn));
            try {
                for (UIManager.LookAndFeelInfo info : UIManager.getInstalledLookAndFeels()) {
                    if (!"Windows".equals(info.getName())) continue;
                    UIManager.setLookAndFeel(info.getClassName());
                    break;
                }
            }
            catch (ClassNotFoundException | IllegalAccessException | InstantiationException | UnsupportedLookAndFeelException ex) {
                Logger.getLogger(MainJFrame.class.getName()).log(Level.SEVERE, null, ex);
            }
        }
        catch (IOException ex) {
            Logger.getLogger(MainJFrame.class.getName()).log(Level.SEVERE, null, ex);
        }
        this.initComponents();
    }

    private void initComponents() {
        this.jPanel1 = new JPanel();
        this.jLayeredPane1 = new JLayeredPane();
        this.jLabel4 = new JLabel();
        this.jLabel9 = new JLabel();
        this.jButton5 = new JButton();
        this.jButton4 = new JButton();
        this.jLabel8 = new JLabel();
        this.jLabel7 = new JLabel();
        this.jLabel2 = new JLabel();
        this.jButton3 = new JButton();
        this.jButton1 = new JButton();
        this.jtf1 = new JTextField();
        this.jtf2 = new JTextField();
        this.jButton2 = new JButton();
        this.jtf3 = new JTextField();
        this.jLabel6 = new JLabel();
        this.jLabel5 = new JLabel();
        this.setDefaultCloseOperation(3);
        this.setTitle("NHS Education - Report generation tool");
        this.setBackground(new Color(255, 255, 255));
        this.setResizable(false);
        this.getContentPane().setLayout(new OverlayLayout(this.getContentPane()));
        this.jPanel1.setBackground(new Color(255, 255, 255));
        this.jLayeredPane1.setBackground(new Color(0, 94, 184));
        this.jLayeredPane1.setForeground(new Color(0, 94, 184));
        this.jLayeredPane1.setOpaque(true);
        this.jLabel4.setBackground(new Color(255, 255, 255));
        this.jLabel4.setFont(new Font("Arial", 1, 18));
        this.jLabel4.setForeground(new Color(240, 240, 240));
        this.jLabel4.setText("<html>Wessex Professional Support<br/>and Wellbeing Unit</html>");
        this.jLabel4.setInheritsPopupMenu(false);
        this.jLabel9.setIcon(this.logoIcon);
        this.jLabel9.setToolTipText("");
        this.jLayeredPane1.setLayer(this.jLabel4, JLayeredPane.POPUP_LAYER);
        this.jLayeredPane1.setLayer(this.jLabel9, JLayeredPane.DEFAULT_LAYER);
        GroupLayout jLayeredPane1Layout = new GroupLayout(this.jLayeredPane1);
        this.jLayeredPane1.setLayout(jLayeredPane1Layout);
        jLayeredPane1Layout.setHorizontalGroup(jLayeredPane1Layout.createParallelGroup(GroupLayout.Alignment.LEADING).addGroup(jLayeredPane1Layout.createSequentialGroup().addContainerGap().addComponent(this.jLabel4, -2, -1, -2).addPreferredGap(LayoutStyle.ComponentPlacement.RELATED, -1, Short.MAX_VALUE).addComponent(this.jLabel9)));
        jLayeredPane1Layout.setVerticalGroup(jLayeredPane1Layout.createParallelGroup(GroupLayout.Alignment.LEADING).addComponent(this.jLabel4, -1, 45, Short.MAX_VALUE).addGroup(jLayeredPane1Layout.createSequentialGroup().addComponent(this.jLabel9).addGap(0, 0, Short.MAX_VALUE)));
        this.jButton5.setBackground(new Color(0, 102, 204));
        this.jButton5.setFont(new Font("Arial", 0, 14));
        this.jButton5.setForeground(new Color(255, 255, 255));
        this.jButton5.setText("Generate report");
        this.jButton5.setBorderPainted(false);
        this.jButton5.setContentAreaFilled(false);
        this.jButton5.setOpaque(true);
        this.jButton5.addMouseListener(new MouseAdapter(){

            @Override
            public void mouseClicked(MouseEvent evt) {
                MainJFrame.this.jButton5MouseClicked(evt);
            }
        });
        this.jButton4.setBackground(new Color(0, 102, 204));
        this.jButton4.setFont(new Font("Arial", 0, 14));
        this.jButton4.setForeground(new Color(255, 255, 255));
        this.jButton4.setText("Merge forms");
        this.jButton4.setBorderPainted(false);
        this.jButton4.setContentAreaFilled(false);
        this.jButton4.setOpaque(true);
        this.jButton4.addMouseListener(new MouseAdapter(){

            @Override
            public void mouseClicked(MouseEvent evt) {
                MainJFrame.this.jButton4MouseClicked(evt);
            }
        });
        this.jLabel8.setText("Please select the referrer form");
        this.jLabel7.setText("Please select the trainee form");
        this.jLabel2.setText("Please select the PSW data spreadsheet:");
        this.jButton3.setIcon(this.folderIcon);
        this.jButton3.addMouseListener(new MouseAdapter(){

            @Override
            public void mouseClicked(MouseEvent evt) {
                MainJFrame.this.jButton3MouseClicked(evt);
            }
        });
        this.jButton1.setIcon(this.folderIcon);
        this.jButton1.setFocusPainted(false);
        this.jButton1.addMouseListener(new MouseAdapter(){

            @Override
            public void mouseClicked(MouseEvent evt) {
                MainJFrame.this.jButton1MouseClicked(evt);
            }
        });
        this.jtf1.setEditable(false);
        this.jtf1.setEnabled(false);
        this.jtf2.setEditable(false);
        this.jtf2.setEnabled(false);
        this.jButton2.setIcon(this.folderIcon);
        this.jButton2.addMouseListener(new MouseAdapter(){

            @Override
            public void mouseClicked(MouseEvent evt) {
                MainJFrame.this.jButton2MouseClicked(evt);
            }
        });
        this.jtf3.setEditable(false);
        this.jtf3.setEnabled(false);
        this.jLabel6.setBackground(new Color(255, 255, 255));
        this.jLabel6.setFont(new Font("Tahoma", 1, 11));
        this.jLabel6.setForeground(new Color(0, 102, 204));
        this.jLabel6.setText("Merge forms");
        this.jLabel5.setBackground(new Color(255, 255, 255));
        this.jLabel5.setFont(new Font("Tahoma", 1, 11));
        this.jLabel5.setForeground(new Color(0, 102, 204));
        this.jLabel5.setText("Generate report");
        GroupLayout jPanel1Layout = new GroupLayout(this.jPanel1);
        this.jPanel1.setLayout(jPanel1Layout);
        jPanel1Layout.setHorizontalGroup(jPanel1Layout.createParallelGroup(GroupLayout.Alignment.LEADING).addGroup(jPanel1Layout.createSequentialGroup().addContainerGap().addGroup(jPanel1Layout.createParallelGroup(GroupLayout.Alignment.LEADING).addGroup(jPanel1Layout.createParallelGroup(GroupLayout.Alignment.TRAILING).addComponent(this.jButton4).addGroup(jPanel1Layout.createParallelGroup(GroupLayout.Alignment.LEADING).addComponent(this.jLabel8).addGroup(jPanel1Layout.createParallelGroup(GroupLayout.Alignment.TRAILING).addGroup(jPanel1Layout.createSequentialGroup().addComponent(this.jtf2, -2, 212, -2).addPreferredGap(LayoutStyle.ComponentPlacement.RELATED).addComponent(this.jButton2, -2, 20, -2)).addGroup(jPanel1Layout.createSequentialGroup().addComponent(this.jtf1, -2, 212, -2).addPreferredGap(LayoutStyle.ComponentPlacement.RELATED).addComponent(this.jButton1, -2, 20, -2))))).addComponent(this.jLabel7).addComponent(this.jLabel6)).addPreferredGap(LayoutStyle.ComponentPlacement.RELATED, 6, Short.MAX_VALUE).addGroup(jPanel1Layout.createParallelGroup(GroupLayout.Alignment.LEADING).addGroup(jPanel1Layout.createSequentialGroup().addGroup(jPanel1Layout.createParallelGroup(GroupLayout.Alignment.LEADING).addGroup(GroupLayout.Alignment.TRAILING, jPanel1Layout.createParallelGroup(GroupLayout.Alignment.LEADING).addGroup(jPanel1Layout.createSequentialGroup().addGap(4, 4, 4).addComponent(this.jtf3, -2, 212, -2).addPreferredGap(LayoutStyle.ComponentPlacement.RELATED).addComponent(this.jButton3, -2, 20, -2)).addComponent(this.jLabel2)).addComponent(this.jButton5, GroupLayout.Alignment.TRAILING)).addContainerGap(-1, Short.MAX_VALUE)).addGroup(jPanel1Layout.createSequentialGroup().addComponent(this.jLabel5).addGap(0, 0, Short.MAX_VALUE)))).addGroup(jPanel1Layout.createSequentialGroup().addComponent(this.jLayeredPane1).addContainerGap()));
        jPanel1Layout.setVerticalGroup(jPanel1Layout.createParallelGroup(GroupLayout.Alignment.LEADING).addGroup(jPanel1Layout.createSequentialGroup().addComponent(this.jLayeredPane1, -2, -1, -2).addPreferredGap(LayoutStyle.ComponentPlacement.RELATED).addGroup(jPanel1Layout.createParallelGroup(GroupLayout.Alignment.BASELINE).addComponent(this.jLabel5).addComponent(this.jLabel6)).addPreferredGap(LayoutStyle.ComponentPlacement.UNRELATED).addGroup(jPanel1Layout.createParallelGroup(GroupLayout.Alignment.LEADING).addGroup(GroupLayout.Alignment.TRAILING, jPanel1Layout.createSequentialGroup().addComponent(this.jLabel2).addPreferredGap(LayoutStyle.ComponentPlacement.RELATED).addComponent(this.jtf3, -2, -1, -2)).addGroup(GroupLayout.Alignment.TRAILING, jPanel1Layout.createParallelGroup(GroupLayout.Alignment.LEADING).addGroup(jPanel1Layout.createSequentialGroup().addComponent(this.jLabel8).addGap(8, 8, 8).addGroup(jPanel1Layout.createParallelGroup(GroupLayout.Alignment.LEADING, false).addComponent(this.jtf1).addComponent(this.jButton1, -2, 20, -2))).addComponent(this.jButton3, GroupLayout.Alignment.TRAILING, -2, 20, -2))).addGap(13, 13, 13).addComponent(this.jLabel7).addPreferredGap(LayoutStyle.ComponentPlacement.UNRELATED).addGroup(jPanel1Layout.createParallelGroup(GroupLayout.Alignment.LEADING, false).addComponent(this.jButton2, -2, 20, -2).addComponent(this.jtf2)).addGap(18, 18, 18).addGroup(jPanel1Layout.createParallelGroup(GroupLayout.Alignment.BASELINE).addComponent(this.jButton5).addComponent(this.jButton4)).addContainerGap(-1, Short.MAX_VALUE)));
        this.getContentPane().add(this.jPanel1);
        this.pack();
    }

    private void jButton1MouseClicked(MouseEvent evt) {
        ExcelJFileChooser chooser = new ExcelJFileChooser();
        this.referrerFile = chooser.run();
        this.jtf1.setText(this.referrerFile.getName());
        this.jtf1.setEnabled(true);
        this.referrerFile = new File(this.referrerFile.getAbsolutePath());
    }

    private void jButton2MouseClicked(MouseEvent evt) {
        ExcelJFileChooser chooser = new ExcelJFileChooser();
        this.traineeFile = chooser.run();
        this.jtf2.setText(this.traineeFile.getName());
        this.jtf2.setEnabled(true);
        this.traineeFile = new File(this.traineeFile.getAbsolutePath());
    }

    private void jButton4MouseClicked(MouseEvent evt) {
        if (this.jtf1.getText().equals("") || this.jtf2.getText().equals("")) {
            MessageDialogs md = new MessageDialogs();
            md.showNoFile();
        } else {
            JFileChooser fileChooser = new JFileChooser(){

                @Override
                public void approveSelection() {
                    String file = this.getSelectedFile() + ".xlsx";
                    File f = new File(file);
                    if (f.exists() && this.getDialogType() == 1) {
                        int result = JOptionPane.showConfirmDialog(this, "The file exists, overwrite?", "Existing file", 0);
                        switch (result) {
                            case 0: {
                                super.approveSelection();
                                return;
                            }
                            case 1: {
                                return;
                            }
                            case -1: {
                                return;
                            }
                        }
                    }
                    super.approveSelection();
                }
            };
            FileNameExtensionFilter filter = new FileNameExtensionFilter("Microsoft Excel Open XML", "xlsx");
            fileChooser.setFileFilter(filter);
            fileChooser.setSelectedFile(new File("Merged Report"));
            int option = fileChooser.showSaveDialog(this);
            if (option == 0) {
                String mergedFilePath = fileChooser.getSelectedFile().getAbsolutePath();
                this.mergedFile = new File(mergedFilePath + ".xlsx");
                PoiHelper helper = new PoiHelper(this.referrerFile, this.traineeFile, this.mergedFile);
                helper.genMergedFormsFile();
                if (helper.isFileOpen(this.mergedFile)) {
                    MessageDialogs md = new MessageDialogs();
                    md.showFileOpenDialog(this.mergedFile);
                } else {
                    helper.mergeForms();
                    helper.genLogFile();
                    helper.mergeFullInfo();
                    JOptionPane.getRootFrame().dispose();
                    JOptionPane.showMessageDialog(this, "Forms merged successfully");
                }
            }
        }
    }

    private void jButton3MouseClicked(MouseEvent evt) {
        ExcelJFileChooser chooser = new ExcelJFileChooser();
        this.pswFile = chooser.run();
        this.jtf3.setText(this.pswFile.getName());
        this.jtf3.setEnabled(true);
        this.pswFile = new File(this.pswFile.getAbsolutePath());
    }

    private void jButton5MouseClicked(MouseEvent evt) {
        if (this.jtf3.getText().equals("")) {
            MessageDialogs md = new MessageDialogs();
            md.showNoFile();
        } else {
            JFileChooser fileChooser = new JFileChooser(){

                @Override
                public void approveSelection() {
                    String file = this.getSelectedFile() + ".docx";
                    File f = new File(file);
                    if (f.exists() && this.getDialogType() == 1) {
                        int result = JOptionPane.showConfirmDialog(this, "The file exists, overwrite?", "Existing file", 0);
                        switch (result) {
                            case 0: {
                                super.approveSelection();
                                return;
                            }
                            case 1: {
                                return;
                            }
                            case -1: {
                                return;
                            }
                        }
                    }
                    super.approveSelection();
                }
            };
            FileNameExtensionFilter filter = new FileNameExtensionFilter("Microsoft Word Open XML", "docx");
            fileChooser.setFileFilter(filter);
            fileChooser.setSelectedFile(new File("Q&G Report"));
            int option = fileChooser.showSaveDialog(this);
            if (option == 0) {
                String docPath = fileChooser.getSelectedFile().getAbsolutePath();
                this.graphs = new File(docPath + " - GRAPHS.xlsx");
                this.doc = new File(docPath + ".docx");
                this.poiHelper = new PoiHelper(this.pswFile, this.graphs);
                if (this.poiHelper.isFileOpen(this.graphs)) {
                    MessageDialogs md = new MessageDialogs();
                    md.showFileOpenDialog(this.graphs);
                } else if (this.poiHelper.isFileOpen(this.doc)) {
                    MessageDialogs md = new MessageDialogs();
                    md.showFileOpenDialog(this.doc);
                } else if (this.poiHelper.isIntegrityCheck()) {
                    MessageDialogs md = new MessageDialogs();
                    md.showColumnIntegrity(this.poiHelper.getMissingColumns());
                } else {
                    LoadingJFrame ljf = new LoadingJFrame();
                    ljf.setVisible(true);
                    this.setEnabled(false);
                    ljf.setLocationRelativeTo(null);
                    ljf.setEnabled(true);
                    ljf.addWindowListener(new WindowAdapter(){

                        @Override
                        public void windowClosed(WindowEvent e) {
                            MainJFrame.this.setEnabled(true);
                            MainJFrame.this.setVisible(true);
                        }
                    });
                    ljf.run(this);
                }
            }
        }
    }

    public ArrayList<File> getFileList() {
        ArrayList<File> list = new ArrayList<File>();
        list.add(this.doc);
        list.add(this.graphs);
        list.add(this.pswFile);
        list.add(this.referrerFile);
        list.add(this.traineeFile);
        return list;
    }
}
