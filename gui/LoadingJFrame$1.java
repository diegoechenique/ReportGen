/*
 * Decompiled with CFR 0.152.
 */
package gui;

import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;

class LoadingJFrame.1
extends WindowAdapter {
    LoadingJFrame.1() {
    }

    @Override
    public void windowClosing(WindowEvent e) {
        LoadingJFrame.this.pleaseHold();
    }
}
