package oxygenoffice.extensions.smart.diagram;


public abstract class DiagramProperties {

    protected   short                   m_Style             = 0;

    //---------------- SIZE ---------------------------------------------------------
    public final static short           UD_SIZE             = 0;
    public final static short           FULL_SIZE           = 1;
    
    protected short                     m_sSize             = UD_SIZE;
    protected int                       m_iUD_width;
    protected int                       m_iUD_height;
    protected int                       m_iUD_XPos;
    protected int                       m_iUD_YPos;
    //---------------- SELECTION ----------------------------------------------------
    public final static short           SELECTED_SHAPES     = 0;
    public final static short           SIBLING_SHAPES      = 1;
    public final static short           BRANCH_SHAPES       = 2;

    protected boolean                   m_IsSelectAllShape  = true;
    protected short                     m_sSelectedArea     = SELECTED_SHAPES;
    //-------------------------------------------------------------------------------


    //---------------- COLORS -------------------------------------------------------
    //public final int                  BLUECOLOR                       = 255;
    protected int                       m_iColor                        = 0x18a303;
    protected int                       m_iStartColor                   = 0x43c330;
    protected int                       m_iEndColor                     = 0x18a303;
                                                             
    public final static int[]           aLOGREENS                       = { 0x106802, 0x18a303, 0x43c330, 0x92e285, 0xccf4c6 };
    public final static int[]           aLOBLUES                        = { 0x023F62, 0x0369A3, 0x1C99E0, 0x63BBEE, 0xAADCF7 };
    public final static int[]           aLOPURPLES                      = { 0x530260, 0x8E03A3, 0xC254D2, 0xDC85E9, 0xF2CBF8 };
    public final static int[]           aLOORANGES                      = { 0x622502, 0xA33E03, 0xD36118, 0xF09E6F, 0xF9CFB5 };
    public final static int[]           aLOYELLOWS                      = { 0x876900, 0xC99C00, 0xE9B913, 0xF5CE53, 0xFDE9A9 };
    public final static int             GreenAccent                     = 0x2CEE0E;
    public final static int             BlueAccent                      = 0x00A0FC;
    public final static int             OrangeAccent                    = 0xFC5C00;
    public final static int             PurpleAccent                    = 0xE327FF;
    public final static int             YellowAccent                    = 0xFFD74C;
                                                                        //old colors: 65280, 255, 16711680, 16776960, 9699435, 16737843, 47359, 12076800 } 
    public final static int[]           aORGCHARTCOLORS                 = { 0xc5000b, OrangeAccent, aLOGREENS[1], 0xffff00, aLOPURPLES[1], 0xFF0000, aLOGREENS[1], 0x2300dc };
    public final static int[]           aBASECOLORS                     = { GreenAccent, 0x0000ff, 0xff0000, 0xffff00, PurpleAccent, OrangeAccent, aLOBLUES[3], aLOORANGES[2]};                                                                    // 0x80f 0x00f 0x08f 0x0ff 0x0f8 0x0f0 0x8f0 0xff0 0xf80 0xf00 0xf08 0xf0f
    public final static int[]           aRAINBOWCOLORS                  = { aLOPURPLES[1], 0x2323dc, aLOGREENS[1], 0x88ff00, 0xffff00, 0xff8800, 0xff0000, 0xff00ff };
    public static int[]                 _aBaseColors                    = { GreenAccent, 0x0000ff, 0xff0000, 0xffff00, PurpleAccent, OrangeAccent, aLOBLUES[3], aLOORANGES[2]};

    public final static int[]           aGREENDARK                      = { aLOGREENS[2], aLOGREENS[1] };
    public final static int[]           aGREENBRIGHT                    = { aLOGREENS[4], aLOGREENS[3] };
    public final static int[]           aBLUEDARK                       = { aLOBLUES[2], aLOBLUES[1] };
    public final static int[]           aBLUEBRIGHT                     = { aLOBLUES[4], aLOBLUES[3] };
    public final static int[]           aPURPLEDARK                     = { aLOPURPLES[2], aLOPURPLES[1] };
    public final static int[]           aPURPLEBRIGHT                   = { aLOPURPLES[4], aLOPURPLES[3] };
    public final static int[]           aORANGEDARK                     = { aLOORANGES[2], aLOORANGES[1] };
    public final static int[]           aORANGEBRIGHT                   = { aLOORANGES[4], aLOORANGES[3] };
    public final static int[]           aYELLOWDARK                     = { aLOYELLOWS[2], aLOYELLOWS[1] };
    public final static int[]           aYELLOWBRIGHT                   = { aLOYELLOWS[4], aLOYELLOWS[3] };
    
    public final static int[][]         aLOCOLORS                       = { aGREENDARK, aGREENBRIGHT, aBLUEDARK, aBLUEBRIGHT, aPURPLEDARK, aPURPLEBRIGHT, aORANGEDARK, aORANGEBRIGHT, aYELLOWDARK, aYELLOWBRIGHT };
    public final static int[][]         aLOCOLORS2                      = { aLOGREENS, aLOBLUES, aLOPURPLES, aLOORANGES, aLOYELLOWS };
    protected short                     m_sColorMode                    = SIMPLE_COLOR_MODE;
    
    public final static short           SIMPLE_COLOR_MODE               = 0;
    public final static short           BASE_COLORS_MODE                = 1;
    public final static short           BASE_COLORS_WITH_GRADIENT_MODE  = 2;
    public final static short           GRADIENT_COLOR_MODE             = 3;

    public final static short           GREEN_DARK_GRADIENT_COLOR_MODE    = 4;
    public final static short           GREEN_BRIGHT_GRADIENT_COLOR_MODE  = 5;
    public final static short           BLUE_DARK_GRADIENT_COLOR_MODE     = 6;
    public final static short           BLUE_BRIGHT_GRADIENT_COLOR_MODE   = 7;
    public final static short           PURPLE_DARK_GRADIENT_COLOR_MODE   = 8;
    public final static short           PURPLE_BRIGHT_GRADIENT_COLOR_MODE = 9;
    public final static short           ORANGE_DARK_GRADIENT_COLOR_MODE   = 10;
    public final static short           ORANGE_BRIGHT_GRADIENT_COLOR_MODE = 11;
    public final static short           YELLOW_DARK_GRADIENT_COLOR_MODE   = 12;
    public final static short           YELLOW_BRIGHT_GRADIENT_COLOR_MODE = 13;

    public final static short           BLUE_SCHEME_COLOR_MODE          = 14;
    public final static short           AQUA_SCHEME_COLOR_MODE          = 15;
    public final static short           RED_SCHEME_COLOR_MODE           = 16;
    public final static short           FIRE_SCHEME_COLOR_MODE          = 17;
    public final static short           SUN_SCHEME_COLOR_MODE           = 18;
    public final static short           GREEN_SCHEME_COLOR_MODE         = 19;
    public final static short           OLIVE_SCHEME_COLOR_MODE         = 20;
    public final static short           PURPLE_SCHEME_COLOR_MODE        = 21;
    public final static short           PINK_SCHEME_COLOR_MODE          = 22;
    public final static short           INDIAN_SCHEME_COLOR_MODE        = 23;
    public final static short           MAROON_SCHEME_COLOR_MODE        = 24;
    public final static short           BROWN_SCHEME_COLOR_MODE         = 25;
    
    public final static short           GREEN_DARK_COLOR_MODE     = 26;
    public final static short           GREEN_BRIGHT_COLOR_MODE   = 27;
    public final static short           BLUE_DARK_COLOR_MODE      = 28;
    public final static short           BLUE_BRIGHT_COLOR_MODE    = 29;
    public final static short           PURPLE_DARK_COLOR_MODE    = 30;
    public final static short           PURPLE_BRIGHT_COLOR_MODE  = 31;
    public final static short           ORANGE_DARK_COLOR_MODE   = 32;
    public final static short           ORANGE_BRIGHT_COLOR_MODE = 33;
    public final static short           YELLOW_DARK_COLOR_MODE   = 34;
    public final static short           YELLOW_BRIGHT_COLOR_MODE = 35;

    public final static short           FIRST_COLORTHEMEGRADIENT_MODE_VALUE = 4;
    public final static short           LAST_COLORTHEMEGRADIENT_MODE_VALUE  = 13;
    public final static short           FIRST_COLORSCHEME_MODE_VALUE        = 14;
    public final static short           LAST_COLORSCHEME_MODE_VALUE         = 25;
    public final static short           FIRST_COLORTHEME_MODE_VALUE         = 26;
    public final static short           LAST_COLORTHEME_MODE_VALUE          = 35;
    
    protected short                     m_sGradientDirection            = VERTICAL;
    public final static short           VERTICAL                        = 0;
    public final static short           HORIZONTAL                      = 1;
    protected boolean                   m_IsModifyColors                = false;
    
    public final static int             DEFAULT_TEXT_COLOR              = 0x000000;
    //-------------------------------------------------------------------------------

    //----------------- LINE --------------------------------------------------------
    protected int                       m_iShapesLineWidth      = LINE_WIDTH100;
    public final static int             LINE_WIDTH000           = 0;
    public final static int             LINE_WIDTH100           = 100;
    public final static int             LINE_WIDTH200           = 200;
    public final static int             LINE_WIDTH300           = 300;
    public final static int             LINE_WIDTH400           = 400;
    public final static int             LINE_WIDTH500           = 500;

    protected int                       m_iLineColor            = 0x000000;
    //-------------------------------------------------------------------------------

    //----------------- ROUNDED -----------------------------------------------------
    protected short                     m_sRounded          = MEDIUM_ROUNDED;
    public final static short           NULL_ROUNDED        = 0;
    public final static short           MEDIUM_ROUNDED      = 1;
    public final static short           EXTRA_ROUNDED       = 2;

    protected final static int          CORNER_RADIUS1      = 300;
    protected final static int          CORNER_RADIUS2      = 500;
    protected final static int          CORNER_RADIUS3      = 700;
    //-------------------------------------------------------------------------------


    //----------------- TRANSPARENCY ------------------------------------------------
    protected short                     m_sTransparency     = MEDIUM_TRANSP;
    public final static short           NULL_TRANSP         = 0;
    public final static short           MEDIUM_TRANSP       = 1;
    public final static short           EXTRA_TRANSP        = 2;

    public final static int             TRANSP_VALUE1       = 35;
    public final static int             TRANSP_VALUE2       = 50;
    public final static int             TRANSP_VALUE3       = 65;
    //-------------------------------------------------------------------------------


    //----------------- FONTS -------------------------------------------------------
    //text properties with default starting values for every diagram
    protected boolean                   m_IsTextToFitToSize = true;
    protected float                     m_fFontSize         = (float)32.0;
    //values have defined in FontSize class
    protected boolean                   m_IsTextColorChange = false;
    protected int                       m_iTextColor        = 0;
    //-------------------------------------------------------------------------------


    //----------------- ADDITIONAL BOOLEAN PROPS ------------------------------------
    protected boolean                   m_IsOutline         = true;
    protected boolean                   m_IsFrame           = true;
    protected boolean                   m_IsRoundedFrame    = true;

    protected boolean                   m_IsShadow          = false;
    public final static int             SHADOW_DIST1        = 250;
    public final static int             SHADOW_DIST2        = 400;
    public final static int             SHADOW_TRANSP       = 30;
    //-------------------------------------------------------------------------------

    //----------------- CONNECTORS --------------------------------------------------
    protected boolean                   m_IsShownConnectors = true;
    protected int                       m_iConnectorsLineWidth  = LINE_WIDTH100;
    protected short                     m_sConnType         = CONN_STANDARD;
    public final static short           CONN_STANDARD       = 0;
    public final static short           CONN_LINE           = 1;
    public final static short           CONN_STRAIGHT       = 2;
    public final static short           CONN_CURVED         = 3;

    protected boolean                   m_IsConnStartArrow  = false;
    protected boolean                   m_IsConnEndArrow    = false;
    
    public final static int             DEFAULT_CONN_COLOR  = 0x000000;
    protected int                       m_iConnColor        = DEFAULT_CONN_COLOR;

    public void setDefaultProps(){
        m_Style                 = 0;
        m_sSize                 = UD_SIZE;
        m_IsSelectAllShape      = true;
        m_sSelectedArea         = SELECTED_SHAPES;
        m_iColor                = 0x18a303;
        m_iStartColor           = 0x43c330;
        m_iEndColor             = 0x18a303;
        setDefaultBaseColors();
        m_sColorMode            = SIMPLE_COLOR_MODE;
        m_sGradientDirection    = VERTICAL;
        m_IsModifyColors        = false;
        m_iShapesLineWidth      = LINE_WIDTH100;
        m_iConnectorsLineWidth  = LINE_WIDTH100;
        m_iLineColor            = 0x000000;
        m_sRounded              = MEDIUM_ROUNDED;
        m_sTransparency         = MEDIUM_TRANSP;
        m_IsTextToFitToSize     = true;
        m_fFontSize             = (float)32.0;
        m_IsTextColorChange     = false;
        m_iTextColor            = 0;
        m_IsOutline             = true;
        m_IsFrame               = true;
        m_IsRoundedFrame        = true;
        m_IsShadow              = false;
        setShownConnectorsProp(true);
        m_sConnType             = CONN_STANDARD;
        m_IsConnStartArrow      = false;
        m_IsConnEndArrow        = false;
        m_iConnColor            = DEFAULT_CONN_COLOR;
    }
    //-------------------------------------------------------------------------------

    public void setStyleProp(short selected) {
        m_Style = selected;
    }

    public short getStyleProp() {
        return m_Style;
    }
    
    //---------------- SIZE ---------------------------------------------------------
    public void setSizeProp(short n){
        m_sSize = n;
    }

    public short getSizeProp(){
        return m_sSize;
    }
    
    public void increaseSizeProp(){
        ++m_sSize;
        if(m_sSize > Diagram.FULL_SIZE)
            m_sSize = 0;
    }
    
    public void setUDWidthProp(int s){
        m_iUD_width = s;
    }

    public int getUDWidthProp(){
        return m_iUD_width;
    }
    
    public void setUDHeightProp(int s){
        m_iUD_height = s;
    }

    public int getUDHeightProp(){
        return m_iUD_height;
    }
    
    public void setUDXPosProp(int pos){
        m_iUD_XPos = pos;
    }

    public int getUDXPosProp(){
        return m_iUD_XPos;
    }
    
    public void setUDYPosProp(int pos){
        m_iUD_YPos = pos;
    }

    public int getUDYPosProp(){
        return m_iUD_YPos;
    }
    
    //---------------- SELECTION ----------------------------------------------------
    public void setSelectedAllShapesProp(boolean bool){
        m_IsSelectAllShape = bool;
    }

    public boolean isSelectedAllShapesProp(){
        return m_IsSelectAllShape;
    }

    public void setSelectedAreaProp(short n){
        m_sSelectedArea = n;
    }

    public short getSeletctedAreaProp(){
        return m_sSelectedArea;
    }
    //-------------------------------------------------------------------------------


    //---------------- COLORS -------------------------------------------------------   
    public void setColorProp(int color){
//        writeColor(color);
        m_iColor = color;
    }
    
    public void writeColor(int color){
        int r = (color & 0xff0000) >> 16;
        int g = (color & 0x00ff00) >> 8;
        int b = color & 0x0000ff;
        String str = color + " hexa: " + Integer.toHexString(color) +  " rgb: " + r + " " + g + " " + b;
        System.out.println(str);
    }

    public int getColorProp(){
        return m_iColor;
    }


    public void setStartColorProp(int color){
        m_iStartColor = color;
    }

    public int getStartColorProp(){
        return m_iStartColor;
    }

    public void setEndColorProp(int color){
        m_iEndColor = color;
    }

    public int getEndColorProp(){
        return m_iEndColor;
    }
    
    public int getArrowColor() {
        if(isColorThemeGradientMode()){
            return getEndColorProp();
        }else if(isColorSchemeMode()){
            return SchemeDefinitions.aColorSchemes[getColorModeProp() - FIRST_COLORSCHEME_MODE_VALUE][1];
        }else{ 
            return _aBaseColors[0];
        }
    }

    public void setColorThemeGradientColors(){
        short s = (short)(getColorModeProp() - FIRST_COLORTHEMEGRADIENT_MODE_VALUE);
        if(s >= 0 && s < 10){
            setStartColorProp(aLOCOLORS[s][0]);
            setEndColorProp(aLOCOLORS[s][1]);
            setLineColorProp(aLOCOLORS[s][1]);
            setColorProp(aLOCOLORS[s][0]);
        }
    }
    
    public void setColorThemeColors(){
        short s = (short)(getColorModeProp() - FIRST_COLORTHEME_MODE_VALUE);
        if(s >= 0 && s < 10){
            setColorProp(aLOCOLORS[s][0]);
            setLineColorProp(aLOCOLORS[s][1]);
        }
    }

    public boolean isColorThemeGradient(short s){
        if(getStartColorProp() == aLOCOLORS[s][0] && getEndColorProp() == aLOCOLORS[s][1])
            return true;
        return false;
    }
    
    public short getColorThemeGradientNum(){
        for(short i = 0; i < 10; i++)
            if(getStartColorProp() == aLOCOLORS[i][0] && getEndColorProp() == aLOCOLORS[i][1])
                return i;
        return -1;
    }
    
    public short getColorThemeGradientNum(int iStartColor, int iEndColor){
        for(short i = 0; i < 10; i++)
            if(iStartColor == aLOCOLORS[i][0] && iEndColor == aLOCOLORS[i][1])
                return i;
        return -1;
    }

    public boolean isModifyColorsProp(){
        return m_IsModifyColors;
    }

    public void setModifyColorsProp(boolean bool){
        m_IsModifyColors = bool;
    }

    public void setColorModeProp(short colorMode){
        m_sColorMode = colorMode;
    }

    public short getColorModeProp(){
        return m_sColorMode;
    }

    public void setGradientDirectionProp(short n){
        m_sGradientDirection = n;
    }

    public short getGradientDirectionProp(){
        return m_sGradientDirection;
    }

    public boolean isSimpleColorMode(){
        if(m_sColorMode == Diagram.SIMPLE_COLOR_MODE)
            return true;
        return false;
    }

    public boolean isBaseColorsMode(){
        if(m_sColorMode == Diagram.BASE_COLORS_MODE)
            return true;
        return false;
    }

    public boolean isBaseColorsWithGradientMode(){
        if(m_sColorMode == Diagram.BASE_COLORS_WITH_GRADIENT_MODE)
            return true;
        return false;
    }

    public boolean isGradientColorMode(){
        if(m_sColorMode == Diagram.GRADIENT_COLOR_MODE)
            return true;
        return false;
    }

    public boolean isColorThemeGradientMode(){
        if(m_sColorMode >= Diagram.FIRST_COLORTHEMEGRADIENT_MODE_VALUE && m_sColorMode <= Diagram.LAST_COLORTHEMEGRADIENT_MODE_VALUE)
            return true;
        return false;
    }

    public boolean isColorSchemeMode(){
        if(m_sColorMode >= Diagram.FIRST_COLORSCHEME_MODE_VALUE && m_sColorMode <= Diagram.LAST_COLORSCHEME_MODE_VALUE)
            return true;
        return false;
    }
    
    public boolean isColorThemeMode(){
        if(m_sColorMode >= Diagram.FIRST_COLORTHEME_MODE_VALUE && m_sColorMode <= Diagram.LAST_COLORTHEME_MODE_VALUE)
            return true;
        return false;
    }
    
    public boolean isAnyGradientColorMode(){
        if(isGradientColorMode() || isColorThemeGradientMode() || isColorSchemeMode())
            return true;
        return false;
    }
    
    public abstract void setDefaultBaseColors();
    
    public static void _setDefaultBaseColors(String sMode){
        if(sMode.equals("BaseColors"))
            for(int i = 0; i < _aBaseColors.length; i++)
                _aBaseColors[i] = aBASECOLORS[i];
        if(sMode.equals("RainbowColors"))
            for(int i = 0; i < _aBaseColors.length; i++)
                _aBaseColors[i] = aRAINBOWCOLORS[i];
    }
    //-------------------------------------------------------------------------------

    //----------------- LINE --------------------------------------------------------
    public int getShapesLineWidhtProp(){
        return m_iShapesLineWidth;
    }
    public void setShapesLineWidthProp(int lineWidth){
        m_iShapesLineWidth = lineWidth;
    }
    /*
    public void setDefaultLineWidthProp(){
        m_iShapesLineWidth = LINE_WIDTH100;
    }
     */
    public int getConnectorsLineWidhtProp(){
        return m_iConnectorsLineWidth;
    }
    public void setConnectorsLineWidthProp(int lineWidth){
        m_iConnectorsLineWidth = lineWidth;
    }

    public int getLineWidthValue(short s) {
        if(s == 0)
            return LINE_WIDTH000;
        if(s == 1)
            return LINE_WIDTH100;
        if(s == 2)
            return LINE_WIDTH200;
        if(s == 3)
            return LINE_WIDTH300;
        if(s == 4)
            return LINE_WIDTH400;
        if(s == 5)
            return LINE_WIDTH500;
        return 0;
    }

    public void setLineColorProp(int color){
        m_iLineColor = color;
    }

    public int getLineColorProp(){
        return m_iLineColor;
    }
/*
    public void setDefaultLineColor(){
        m_iLineColor = 0x000000;
    }
  */
    public int getDefaultLineColor(){
        return 0x000000;
    }
    //-------------------------------------------------------------------------------

    //----------------- ROUNDED -----------------------------------------------------
    public void setRoundedProp(short type){
        m_sRounded = type;
    }
    
    public short getRoundedProp(){
        return m_sRounded;
    }
    //-------------------------------------------------------------------------------


    //----------------- TRANSPARENCY ------------------------------------------------
    public void setTransparencyProp(short type){
        m_sTransparency = type;
    }

    public short getTransparencyProp(){
        return m_sTransparency;
    }
    //-------------------------------------------------------------------------------


    //----------------- FONTS -------------------------------------------------------
    public void setTextColorProp(int color){
        m_iTextColor = color;
    }

    public int getTextColorProp(){
        return m_iTextColor;
    }

    public void setTextColorChange(boolean bool){
        m_IsTextColorChange = bool;
    }

    public boolean isTextColorChange(){
        return m_IsTextColorChange;
    }

    public void setTextFitProp(boolean bool){
        m_IsTextToFitToSize = bool;
    }

    public boolean isTextFitProp(){
        return m_IsTextToFitToSize;
    }

    public void setFontSizeProp(float fontSize){
        m_fFontSize = fontSize;
    }

    public float getFontSizeProp(){
        return m_fFontSize;
    }
    //-------------------------------------------------------------------------------


    //----------------- ADDITIONAL BOOLEAN PROPS ------------------------------------
    public void setOutlineProp(boolean bool){
        m_IsOutline = bool;
    }

    public boolean isOutlineProp(){
        return m_IsOutline;
    }

    public void setFrameProp(boolean bool){
        m_IsFrame = bool;
    }

    public boolean isFrameProp(){
        return m_IsFrame;
    }

    public void setRoundedFrameProp(boolean bool){
        m_IsRoundedFrame = bool;
    }

    public boolean isRoundedFrameProp(){
        return m_IsRoundedFrame;
    }

    public void setShadowProp(boolean bool){
        m_IsShadow = bool;
    }

    public boolean isShadowProp(){
        return m_IsShadow;
    }
    //-------------------------------------------------------------------------------


    //----------------- CONNECTORS --------------------------------------------------
    public boolean isShownConnectorsProp(){
        return m_IsShownConnectors;
    }

    public void setShownConnectorsProp(boolean bool){
        m_IsShownConnectors = bool;
    }

    public void setConnectorTypeProp(short type){
        m_sConnType = type;
    }

    public short getConnectorTypeProp(){
        return m_sConnType;
    }

    public void setConnectorStartArrowProp(boolean bool){
        m_IsConnStartArrow = bool;
    }

    public boolean isConnectorStartArrowProp(){
        return m_IsConnStartArrow;
    }

    public void setConnectorEndArrowProp(boolean bool){
        m_IsConnEndArrow = bool;
    }

    public boolean isConnectorEndArrowProp(){
        return m_IsConnEndArrow;
    }
    
    public int getConnectorColorProp(){
        return m_iConnColor;
    }
    
    public void setConnectorColorProp(int color){
        m_iConnColor = color;
    }
    
    public int getDefalutConnectorColor(){
        return DEFAULT_CONN_COLOR;
    }
    //-------------------------------------------------------------------------------

}
