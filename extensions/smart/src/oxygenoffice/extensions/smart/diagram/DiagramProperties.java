package oxygenoffice.extensions.smart.diagram;

public abstract class DiagramProperties {

    protected   short                   m_Style;

    public final static short           SELECTED_SHAPES     = 0;
    public final static short           SIBLING_SHAPES      = 1;
    public final static short           BRANCH_SHAPES       = 2;

    public final static short           VERTICAL            = 0;
    public final static short           HORIZONTAL          = 1;

    public final static short           NULL_ROUNDED        = 0;
    public final static short           MEDIUM_ROUNDED      = 1;
    public final static short           EXTRA_ROUNDED       = 2;

    public final static short           NULL_TRANSP         = 0;
    public final static short           MEDIUM_TRANSP       = 1;
    public final static short           EXTRA_TRANSP        = 2;


    //properties of style of diagram
    protected boolean                   m_IsSelectAllShape;
    protected short                     m_sSelectedArea;
    //SELECTED_SHAPES = 0, SIBLING_SHAPES = 1, BRANCH_SHAPES = 2;

    protected boolean                   m_IsModifyColors;
    protected boolean                   m_IsColor;
    protected boolean                   m_IsBaseColors              = true;
    protected boolean                   m_IsBaseColorsWithGradients = false;
    protected boolean                   m_IsGradients               = false;
    protected short                     m_sGradientDirection;
    // VERTICAL, HORIZONTAL

    protected boolean                   m_IsPreDefinedGradients     = false;
    //protected boolean                   m_IsBlueGradients           = false;
    //protected boolean                   m_IsRedGradients            = false;

    protected int                       m_iColor            = 255;
    protected int                       m_iStartColor       = 16711680;
    protected int                       m_iEndColor         = 8388608;

    protected short                     m_sRounded;
    //NULL_ROUNDED = 0, MEDIUM_ROUNDED = 1, EXTRA_ROUNDED = 2

    protected short                     m_sTransparency;
    //NULL_TRANSP = 0, MEDIUM_TRANSP = 1, EXTRA_TRANSP = 2;

    protected boolean                   m_IsOutline;
    protected boolean                   m_IsFrame;
    protected boolean                   m_IsRoundedFrame;
    protected boolean                   m_IsShadow;
    protected boolean                   m_IsAction;

    //text properties with default starting values for every diagram
    protected boolean                   m_IsTextToFitToSize = true;
    protected float                     m_fFontSize = (float)32.0;
    //values have defined in FontSize class
    protected boolean                   m_IsTextColorChage = false;
    protected int                       m_iTextColor            = 0;

    public void setStyle(short selected) {
        m_Style = selected;
    }

    public short getStyle() {
        return m_Style;
    }

    public void setSelectedAllShapesProps(boolean bool){
        m_IsSelectAllShape = bool;
    }

    public boolean isSelectedAllShapesProps(){
        return m_IsSelectAllShape;
    }

    public void setSelectedAreaProps(short n){
        m_sSelectedArea = n;
    }

    public short getSeletctedAreaProps(){
        return m_sSelectedArea;
    }

    public boolean isModifyColorsProps(){
        return m_IsModifyColors;
    }

    public void setModifyColorsProps(boolean bool){
        m_IsModifyColors = bool;
    }

    public void setColorProps(boolean bool){
        m_IsColor = bool;
    }

    public void setGradientProps(boolean bool){
        m_IsGradients = bool;
    }

    public boolean isGradientProps(){
        return m_IsGradients;
    }

    public void setPreDefinedGradientsProps(boolean isPreDefinedGradients){
        m_IsPreDefinedGradients = isPreDefinedGradients;
    }

    public boolean isPreDefinedGradientsProps(){
        return m_IsPreDefinedGradients;
    }

    public void setBaseColorsProps(boolean bool){
        m_IsBaseColors = bool;
    }

    public boolean isBaseColorsProps(){
        return m_IsBaseColors;
    }

    public void setBaseColorsWithGradientsProps(boolean bool){
        m_IsBaseColorsWithGradients = bool;
    }

    public boolean isBaseColorsWithGradientsProps(){
        return m_IsBaseColorsWithGradients;
    }

    public void setGradientDirectionProps(short n){
        m_sGradientDirection = n;
    }

    public short getGradientDirectionProps(){
        return m_sGradientDirection;
    }

    public void setColorProps(int color){
        m_iColor = color;
    }

    public int getColorProps(){
        return m_iColor;
    }

    public void setStartColorProps(int color){
        m_iStartColor = color;
    }

    public int getStartColorProps(){
        return m_iStartColor;
    }

    public void setEndColorProps(int color){
        m_iEndColor = color;
    }

    public int getEndColorProps(){
        return m_iEndColor;
    }

    public void setRoundedProps(short type){
        m_sRounded = type;
    }

    public short getRoundedProps(){
        return m_sRounded;
    }

    public void setTransparencyProps(short type){
        m_sTransparency = type;
    }

    public void setOutlineProps(boolean bool){
        m_IsOutline = bool;
    }

    public boolean isOutlineProps(){
        return m_IsOutline;
    }

    public void setFrameProps(boolean bool){
        m_IsFrame = bool;
    }

    public void setRoundedFrameProps(boolean bool){
        m_IsRoundedFrame = bool;
    }

    public void setShadowProps(boolean bool){
        m_IsShadow = bool;
    }

    public boolean isShadowProps(){
        return m_IsShadow;
    }

    public void setActionProps(boolean bool){
        m_IsAction = bool;
    }

    public void setTextColorProps(int color){
        m_iTextColor = color;
    }

    public int getTextColorProps(){
        return m_iTextColor;
    }

    public void setTextColorChange(boolean bool){
        m_IsTextColorChage = bool;
    }

    public boolean isTextColorChange(){
        return m_IsTextColorChage;
    }

    public void setTextFitProps(boolean bool){
        m_IsTextToFitToSize = bool;
    }

    public boolean isTextFitProps(){
        return m_IsTextToFitToSize;
    }

    public void setFontSizeProps(float fontSize){
        m_fFontSize = fontSize;
    }

    public float getFontSizeProps(){
        return m_fFontSize;
    }
}
