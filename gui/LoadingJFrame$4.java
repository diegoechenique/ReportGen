/*
 * Decompiled with CFR 0.152.
 */
package gui;

import gui.LoadingJFrame;

static class LoadingJFrame.4
implements Runnable {
    LoadingJFrame.4() {
    }

    @Override
    public void run() {
        new LoadingJFrame().setVisible(true);
    }
}
