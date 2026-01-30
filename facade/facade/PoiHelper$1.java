/*
 * Decompiled with CFR 0.152.
 */
package facade;

import org.apache.poi.ss.usermodel.CellType;

static class PoiHelper.1 {
    static final /* synthetic */ int[] $SwitchMap$org$apache$poi$ss$usermodel$CellType;

    static {
        $SwitchMap$org$apache$poi$ss$usermodel$CellType = new int[CellType.values().length];
        try {
            PoiHelper.1.$SwitchMap$org$apache$poi$ss$usermodel$CellType[CellType.STRING.ordinal()] = 1;
        }
        catch (NoSuchFieldError noSuchFieldError) {
            // empty catch block
        }
        try {
            PoiHelper.1.$SwitchMap$org$apache$poi$ss$usermodel$CellType[CellType.NUMERIC.ordinal()] = 2;
        }
        catch (NoSuchFieldError noSuchFieldError) {
            // empty catch block
        }
        try {
            PoiHelper.1.$SwitchMap$org$apache$poi$ss$usermodel$CellType[CellType.BLANK.ordinal()] = 3;
        }
        catch (NoSuchFieldError noSuchFieldError) {
            // empty catch block
        }
    }
}
