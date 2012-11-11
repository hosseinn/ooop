package oxygenoffice.extensions.smart.diagram;

public class SchemeDefinitions {
    
    protected static final int[]    aBLUE_SCHEME     = { 0x0099ff, 0x000077 };
    protected static final int[]    aAQUA_SCHEME     = { 65535, 0x0099ff };
    protected static final int[]    aRED_SCHEME      = { 0xff0000, 0x660000 };
    protected static final int[]    aFIRE_SCHEME     = { 0xffff00, 0xe00000 };
    protected static final int[]    aSUN_SCHEME      = { 0xfffd45, 0xff8000 };
    protected static final int[]    aGREEN_SCHEME    = { 65280, 0x0e4b00 };
    protected static final int[]    aOLIVE_SCHEME    = { 14677829, 0x3c4900 };
    protected static final int[]    aPURPLE_SCHEME   = { 16711935, 0x5e1b5d };
    protected static final int[]    aPINK_SCHEME     = { 0xffdeec, 0xd60084 };
    protected static final int[]    aINDIAN_SCHEME   = { 0xffeeee, 0xcd5c5c };
    protected static final int[]    aMAROON_SCHEME   = { 0xF9CFB5, 0xA33E03 };
    protected static final int[]    aBROWN_SCHEME    = { 0xecba74, 0x5c2d0a };

    public static final int[][]  aColorSchemes = {  aBLUE_SCHEME, aAQUA_SCHEME, aRED_SCHEME, aFIRE_SCHEME,
                                                    aSUN_SCHEME, aGREEN_SCHEME, aOLIVE_SCHEME, aPURPLE_SCHEME,
                                                    aPINK_SCHEME, aINDIAN_SCHEME, aMAROON_SCHEME, aBROWN_SCHEME };

    protected static final int      NUM_OF_SCHEMES = 12;

    
    public static int getGradientColor(int colorCode, int index, int iSteps){
        if(colorCode >= 0 && colorCode < NUM_OF_SCHEMES)
            return SchemeDefinitions.getGradientColor(aColorSchemes[colorCode][0], aColorSchemes[colorCode][1], index, iSteps);
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
