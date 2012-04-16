package oxygenoffice.extensions.smart.diagram;

public class GradientInterpolate {

    // return the interpolated value between pBegin and pEnd
    public int interpolate(int iBegin, int iEnd, int iStep, int iMax) {
        if (iBegin < iEnd) {
            return ((iEnd - iBegin) * (iStep / iMax)) + iBegin;
        } else {
            return ((iBegin - iEnd) * (1 - (iStep / iMax))) + iEnd;
        }
    }

    public int[] getGradientColors(int iBeginColor, int iEndColor, int iSteps){
        int[] aValues   = new int[iSteps+1];
        int RBegin      = (iBeginColor & 0xff0000) >> 16;
        int GBegin      = (iBeginColor & 0x00ff00) >> 8;
        int BBegin      = iBeginColor & 0x0000ff;
        int REnd        = (iEndColor & 0xff0000) >> 16;
        int GEnd        = (iEndColor & 0x00ff00) >> 8;
        int BEnd        = iEndColor & 0x0000ff;
        int R, G, B;

        for(int i=0; i<=iSteps; i++){
            R = interpolate(RBegin, REnd, i, iSteps);
            G = interpolate(GBegin, GEnd, i, iSteps);
            B = interpolate(BBegin, BEnd, i, iSteps);
            aValues[i] = (((R << 8) | G) << 8) | B;
        }
        return aValues;
    }
}