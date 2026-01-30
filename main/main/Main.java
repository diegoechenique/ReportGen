/*
 * Decompiled with CFR 0.152.
 */
package main;

import gui.MainJFrame;
import javax.swing.SwingUtilities;

public class Main {
    public static void main(String[] args) {
        SwingUtilities.invokeLater(new Runnable(){

            @Override
            public void run() {
                MainJFrame frame = new MainJFrame();
                frame.setVisible(true);
                frame.setLocationRelativeTo(null);
                frame.setDefaultCloseOperation(3);
            }
        });
    }
}
