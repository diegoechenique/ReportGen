/*
 * Decompiled with CFR 0.152.
 */
package gui;

import java.io.File;
import javax.swing.JFileChooser;
import javax.swing.JOptionPane;

class MainJFrame.6
extends JFileChooser {
    MainJFrame.6() {
    }

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
}
