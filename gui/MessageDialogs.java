/*
 * Decompiled with CFR 0.152.
 */
package gui;

import java.io.File;
import java.util.ArrayList;
import javax.swing.JOptionPane;

public class MessageDialogs {
    public void showFileOpenDialog(File file) {
        JOptionPane.showMessageDialog(null, "The program cannot access " + file.getName() + " because it is being used by another program, please close it and try again.", "Warning", 1);
    }

    public void showColumnIntegrity(ArrayList<String> strList) {
        String strStrings = "";
        for (String str : strList) {
            strStrings = strStrings + str + "\n";
        }
        JOptionPane.showMessageDialog(null, "Error: Could not find column(s): " + strStrings + "Please check the format of the spreadsheet and try again.", "Warning", 1);
    }

    public void showNoFile() {
        JOptionPane.showMessageDialog(null, "No file selected! Please specify a File and try again", "Warning", 1);
    }

    public void showPleaseSelect() {
        JOptionPane.showMessageDialog(null, "Please select a destination folder", "Warning", 1);
    }

    public void showPleaseHold() {
        JOptionPane.showMessageDialog(null, "Please wait for the program to finish", "Warning", 1);
    }
}
