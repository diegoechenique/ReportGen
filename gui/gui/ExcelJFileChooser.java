/*
 * Decompiled with CFR 0.152.
 */
package gui;

import java.io.File;
import javax.swing.JFileChooser;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.filechooser.FileSystemView;

public class ExcelJFileChooser {
    public File run() {
        JFileChooser jfc = new JFileChooser(FileSystemView.getFileSystemView().getHomeDirectory());
        FileNameExtensionFilter filter = new FileNameExtensionFilter("Microsoft Excel Open XML", "xlsx");
        jfc.setFileFilter(filter);
        int returnValue = jfc.showOpenDialog(null);
        File selectedFile = null;
        if (returnValue == 0) {
            selectedFile = jfc.getSelectedFile();
        }
        return selectedFile;
    }
}
