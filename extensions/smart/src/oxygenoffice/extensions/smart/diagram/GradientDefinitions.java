package oxygenoffice.extensions.smart.diagram;

public class GradientDefinitions {

    protected static final int[]    aBLUE_COLORS     = { 0x000077, 0x0099ff };
    protected static final int[]    aAQUA_COLORS     = { 65535, 0x0099ff };
    protected static final int[]    aRED_COLORS      = { 0xff0000, 0x660000 };
    protected static final int[]    aFIRE_COLORS     = { 0xffff00, 0xe00000 };
    protected static final int[]    aSUN_COLORS      = { 0xfffd45, 0xff8000 };
    protected static final int[]    aGREEN_COLORS    = { 65280, 0x0e4b00 };
    protected static final int[]    aOLIVE_COLORS    = { 14677829, 0x3c4900 };
    protected static final int[]    aPURPLE_COLORS   = { 16711935, 0x5e1b5d };
    protected static final int[]    aPINK_COLORS     = { 0xffdeec, 0xd60084 };
    protected static final int[]    aINDIAN_COLORS   = { 0xffeeee, 0xcd5c5c };
    protected static final int[]    aMAROON_COLORS   = { 0xffaa88, 0x800000 };
    protected static final int[]    aBROWN_COLORS    = { 0xecba74, 0x5c2d0a };

    protected static final int[][]  aGradients = {  aBLUE_COLORS, aAQUA_COLORS, aRED_COLORS, aFIRE_COLORS,
                                                    aSUN_COLORS, aGREEN_COLORS, aOLIVE_COLORS, aPURPLE_COLORS,
                                                    aPINK_COLORS, aINDIAN_COLORS, aMAROON_COLORS, aBROWN_COLORS };

    public final static short       BLUE_GRADIENTS      = 4;
    public final static short       AQUA_GRADIENTS      = 5;
    public final static short       RED_GRADIENTS       = 6;
    public final static short       FIRE_GRADIENTS      = 7;
    public final static short       SUN_GRADIENTS       = 8;
    public final static short       GREEN_GRADIENTS     = 9;
    public final static short       OLIVE_GRADIENTS     = 10;
    public final static short       PURPLE_GRADIENTS    = 11;
    public final static short       PINK_GRADIENTS      = 12;
    public final static short       INDIAN_GRADIENTS    = 13;
    public final static short       MAROON_GRADIENTS    = 14;
    public final static short       BROWN_GRADIENTS     = 15;

    public final static short       FIRST_GRAD_VALUE    = 4;
    public final static short       NUM_OF_GRADIENTS    = 12;


    public static boolean isPreDefinedGradient(short style){
        return  style == BLUE_GRADIENTS || style == AQUA_GRADIENTS ||
                style == RED_GRADIENTS || style == FIRE_GRADIENTS ||
                style == SUN_GRADIENTS || style == GREEN_GRADIENTS ||
                style == OLIVE_GRADIENTS || style == PURPLE_GRADIENTS ||
                style == PINK_GRADIENTS || style == INDIAN_GRADIENTS ||
                style == MAROON_GRADIENTS || style == BROWN_GRADIENTS;
    }

    public static int getPreDefinedGradient(int style, int index, int iSteps){
        int colorCode = style - FIRST_GRAD_VALUE;
        if(colorCode >= 0 && colorCode < NUM_OF_GRADIENTS)
            return GradientDefinitions.getGradientColor(aGradients[colorCode][0], aGradients[colorCode][1], index, iSteps);
        return 0x0099ff;
    }

    // return the interpolated value between pBegin and pEnd
    public static int interpolate(int iBegin, int iEnd, int iStep, int iMax) {
        if (iBegin < iEnd) {
            double diff = iEnd - iBegin;
            return (int)(diff * ((double)iStep / (double)iMax))+ iBegin;

        } else {
            double diff = iBegin - iEnd;
            return (int)((diff * (1 - ((double)iStep / (double)iMax))) + iEnd);
        }
    }

    public static int getGradientColor(int iBeginColor, int iEndColor, int index, int iSteps){
        int RBegin      = (iBeginColor & 0xff0000) >> 16;
        int GBegin      = (iBeginColor & 0x00ff00) >> 8;
        int BBegin      = iBeginColor & 0x0000ff;
        int REnd        = (iEndColor & 0xff0000) >> 16;
        int GEnd        = (iEndColor & 0x00ff00) >> 8;
        int BEnd        = iEndColor & 0x0000ff;

        int R = interpolate(RBegin, REnd, index, iSteps);
        int G = interpolate(GBegin, GEnd, index, iSteps);
        int B = interpolate(BBegin, BEnd, index, iSteps);
        return ((((R << 8) | G) << 8) | B);
    }

}
