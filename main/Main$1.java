/*
 * Decompiled with CFR 0.152.
 */
package main;

import gui.MainJFrame;

static class Main.1
implements Runnable {
    Main.1() {
    }

    @Override
    public void run() {
        MainJFrame frame = new MainJFrame();
        frame.setVisible(true);
        frame.setLocationRelativeTo(null);
        frame.setDefaultCloseOperation(3);
    }
}
