/*
 * Decompiled with CFR 0.152.
 */
package gui;

import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;

class MainJFrame.8
extends WindowAdapter {
    MainJFrame.8() {
    }

    @Override
    public void windowClosed(WindowEvent e) {
        MainJFrame.this.setEnabled(true);
        MainJFrame.this.setVisible(true);
    }
}
