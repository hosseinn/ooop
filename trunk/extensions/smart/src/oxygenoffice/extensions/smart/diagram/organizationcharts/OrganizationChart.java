package oxygenoffice.extensions.smart.diagram.organizationcharts;

import com.sun.star.awt.Gradient;
import com.sun.star.awt.Point;
import com.sun.star.awt.Size;
import com.sun.star.beans.PropertyVetoException;
import com.sun.star.beans.UnknownPropertyException;
import com.sun.star.beans.XPropertySet;
import com.sun.star.container.XNamed;
import com.sun.star.drawing.ConnectorType;
import com.sun.star.drawing.FillStyle;
import com.sun.star.drawing.LineStyle;
import com.sun.star.drawing.TextFitToSizeType;
import com.sun.star.drawing.XShape;
import com.sun.star.drawing.XShapes;
import com.sun.star.frame.XFrame;
import com.sun.star.lang.IllegalArgumentException;
import com.sun.star.lang.IndexOutOfBoundsException;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.text.XText;
import com.sun.star.uno.AnyConverter;
import com.sun.star.uno.UnoRuntime;
import oxygenoffice.extensions.smart.Controller;
import oxygenoffice.extensions.smart.diagram.DataOfDiagram;
import oxygenoffice.extensions.smart.diagram.Diagram;
import oxygenoffice.extensions.smart.diagram.SchemeDefinitions;
import oxygenoffice.extensions.smart.gui.Gui;


public abstract class OrganizationChart extends Diagram {

    // rates of measure of groupShape (e.g.: 10:6)
    protected int               m_iGroupWidth;
    protected int               m_iGroupHeight;

    // rates of measure of rectangles
    //(e.g.: WIDTH:HORSPACE 2:1, HEIGHT:VERSPACE 4:3)
    protected int               m_iShapeWidth;
    protected int               m_iHorSpace;
    protected int               m_iShapeHeight;
    protected int               m_iVerSpace;
    
    //horizontal offset in the drawpage if needed
    protected int               m_iHalfDiff         = 0;

    //hierarhich tpyes
    public static final short   UNDERLING           = 0;
    public static final short   ASSOCIATE           = 1;

    //item hierarhich tpye in diagram
    protected short             m_sNewItemHType     = UNDERLING;

    //styles of property dialog
    public final static short   DEFAULT             = 0;
    public final static short   WITHOUT_OUTLINE     = 1;
    public final static short   NOT_ROUNDED         = 2;
    public final static short   WITH_SHADOW         = 3;

    public final static short   GREEN_DARK          = 4;
    public final static short   GREEN_BRIGHT        = 5;
    public final static short   BLUE_DARK           = 6;
    public final static short   BLUE_BRIGHT         = 7;
    public final static short   PURPLE_DARK         = 8;
    public final static short   PURPLE_BRIGHT       = 9;
    public final static short   ORANGE_DARK         = 10;
    public final static short   ORANGE_BRIGHT       = 11;
    public final static short   YELLOW_DARK         = 12;
    public final static short   YELLOW_BRIGHT       = 13;

    public final static short   BLUE_SCHEME         = 14;
    public final static short   AQUA_SCHEME         = 15;
    public final static short   RED_SCHEME          = 16;
    public final static short   FIRE_SCHEME         = 17;
    public final static short   SUN_SCHEME          = 18;
    public final static short   GREEN_SCHEME        = 19;
    public final static short   OLIVE_SCHEME        = 20;
    public final static short   PURPLE_SCHEME       = 21;
    public final static short   PINK_SCHEME         = 22;
    public final static short   INDIAN_SCHEME       = 23;
    public final static short   MAROON_SCHEME       = 24;
    public final static short   BROWN_SCHEME        = 25;
    public final static short   USER_DEFINE         = 26;

    public final static short   FIRST_COLORTHEMEGRADIENT_STYLE_VALUE = 4;
    public final static short   FIRST_COLORSCHEME_STYLE_VALUE = 14;

    private boolean m_IsHiddenRootElement           = false;
    
    public boolean isHiddenRootElementProp(){
        return m_IsHiddenRootElement;
    }
    
    public void setHiddenRootElementProp(boolean isHidden){
        m_IsHiddenRootElement = isHidden;
    }

    public void initRootElementHiddenProperty(){
        if(getDiagramTree().getRootItem() != null)
            m_IsHiddenRootElement = getDiagramTree().getRootItem().isHiddenElement();
    }

    public OrganizationChart(Controller controller, Gui gui, XFrame xFrame) {
        super(controller, gui, xFrame);
        m_sNewItemHType = UNDERLING;
        setDefaultProps();
    }

    @Override
    public boolean isColorSchemeStyle(short style){
        return  style == BLUE_SCHEME   || style == AQUA_SCHEME   ||
                style == RED_SCHEME    || style == FIRE_SCHEME   ||
                style == SUN_SCHEME    || style == GREEN_SCHEME  ||
                style == OLIVE_SCHEME  || style == PURPLE_SCHEME ||
                style == PINK_SCHEME   || style == INDIAN_SCHEME ||
                style == MAROON_SCHEME || style == BROWN_SCHEME;
    }

    @Override
    public short getColorModeOfSchemeStyle(short style){
        return (short)(style - FIRST_COLORSCHEME_STYLE_VALUE + Diagram.FIRST_COLORSCHEME_MODE_VALUE);
    }

    @Override
    public boolean isColorThemeGradientStyle(short style){
        return  style == GREEN_DARK  || style == GREEN_BRIGHT  ||
                style == BLUE_DARK   || style == BLUE_BRIGHT   ||
                style == PURPLE_DARK || style == PURPLE_BRIGHT ||
                style == ORANGE_DARK || style == ORANGE_BRIGHT ||
                style == YELLOW_DARK || style == YELLOW_BRIGHT;
    }

    @Override
    public short getColorModeOfThemeGradientStyle(short style){
        return (short)(style - FIRST_COLORTHEMEGRADIENT_STYLE_VALUE + Diagram.FIRST_COLORTHEMEGRADIENT_MODE_VALUE);
    }
    
    @Override
    public void initColorModeAndStyle(){
        initColorModeAndStyle(getDiagramTree().getControlShape());
    }
    
    @Override
    public void initProperties(){
        XShape xControlShape = getDiagramTree().getControlShape();
        XShape xRootShape = getDiagramTree().getRootItem().getRectangleShape();
        if(xControlShape != null && xRootShape != null){
            initProperties(xControlShape, xRootShape);
            initRootElementHiddenProperty();
        }
    }
    
    public void initProperties(XShape xControlShape, XShape xRootShape){
        setDefaultProps();
        initColorModeAndStyle();
        XPropertySet xProps = (XPropertySet)UnoRuntime.queryInterface( XPropertySet.class, xRootShape);
        try {
            if (isSimpleColorMode()){
                setColorProp(AnyConverter.toInt(xProps.getPropertyValue("FillColor")));
            }
            if(isGradientColorMode()){
                Gradient aGradient = (Gradient)xProps.getPropertyValue("FillGradient");
                setStartColorProp(aGradient.StartColor);
                setEndColorProp(aGradient.EndColor);
                if(aGradient.Angle == 900)
                    setGradientDirectionProp(Diagram.HORIZONTAL);
            }
            if(isColorThemeGradientMode()){
                setColorThemeGradientColors();
                //setShapesLineWidthProp(Diagram.LINE_WIDTH200);
                setRoundedProp(Diagram.NULL_ROUNDED);
            }
            //if(getStyleProp() == OrganizationChart.DEFAULT){ }
            if(getStyleProp() == OrganizationChart.WITHOUT_OUTLINE)
                setOutlineProp(false);
            if(getStyleProp() == OrganizationChart.NOT_ROUNDED)
                setRoundedProp(Diagram.NULL_ROUNDED);
            if(getStyleProp() == OrganizationChart.WITH_SHADOW)
                setShadowProp(true);

            if(getStyleProp() == OrganizationChart.USER_DEFINE){
                if(((LineStyle)xProps.getPropertyValue("LineStyle")).getValue() == LineStyle.NONE_value)
                    setOutlineProp(false);
                setShapesLineWidthProp(AnyConverter.toInt(xProps.getPropertyValue("LineWidth")));
                int cornerRadius = AnyConverter.toInt(xProps.getPropertyValue("CornerRadius"));
                if(cornerRadius < 200)
                    setRoundedProp(Diagram.NULL_ROUNDED);
                else if(cornerRadius < 600)
                    setRoundedProp(Diagram.MEDIUM_ROUNDED);
                else
                    setRoundedProp(Diagram.EXTRA_ROUNDED);
                if(AnyConverter.toBoolean(xProps.getPropertyValue("Shadow")))
                    setShadowProp(true);
            }

            setFontPropertyValues();

            XText xText = (XText)UnoRuntime.queryInterface(XText.class, xRootShape);
            XPropertySet xTextProps = (XPropertySet)UnoRuntime.queryInterface( XPropertySet.class, xText.createTextCursor());
            setTextColorProp(AnyConverter.toInt(xTextProps.getPropertyValue("CharColor")));

            if(getDiagramTypeName().equals("TableHierarchyDiagram"))
                setShownConnectorsProp(false);
            if(isShownConnectorsProp()){
                XShape xConnShape = ((OrganizationChart)this).getRootsConnector();
                xProps = (XPropertySet)UnoRuntime.queryInterface( XPropertySet.class, xConnShape);
                setConnectorsLineWidthProp(AnyConverter.toInt(xProps.getPropertyValue("LineWidth")));
                setConnectorColorProp(AnyConverter.toInt(xProps.getPropertyValue("LineColor")));
                if(((ConnectorType)xProps.getPropertyValue("EdgeKind")).getValue() == ConnectorType.STANDARD_value)
                    setConnectorTypeProp(Diagram.CONN_STANDARD);
                if(((ConnectorType)xProps.getPropertyValue("EdgeKind")).getValue() == ConnectorType.LINE_value)
                    setConnectorTypeProp(Diagram.CONN_LINE);
                if(((ConnectorType)xProps.getPropertyValue("EdgeKind")).getValue() == ConnectorType.LINES_value)
                    setConnectorTypeProp(Diagram.CONN_STRAIGHT);
                if(((ConnectorType)xProps.getPropertyValue("EdgeKind")).getValue() == ConnectorType.CURVE_value)
                    setConnectorTypeProp(Diagram.CONN_CURVED);
                if((AnyConverter.toString(xProps.getPropertyValue("LineStartName"))).equals("Arrow"))
                    setConnectorStartArrowProp(true);
                if((AnyConverter.toString(xProps.getPropertyValue("LineEndName"))).equals("Arrow"))
                    setConnectorEndArrowProp(true);
            }
        } catch (UnknownPropertyException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (IllegalArgumentException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
    }

    public abstract void initDiagramTree(OrganizationChartTree diagramTree);
    
    public void setNewItemHType(short n){
        m_sNewItemHType = n;
    }

    public XShapes getShapes(){
        return m_xShapes;
    }

    public int getShapeWidth(){
        return m_iShapeWidth;
    }

    public int getHorSpace(){
        return m_iHorSpace;
    }

    public int getShapeHeight(){
        return m_iShapeHeight;
    }

    public int getVerSpace(){
        return m_iVerSpace;
    }

    public void selectShapes(){
        getController().setSelectedShape((Object)m_xShapes);
    }

    @Override
    public void createDiagram(DataOfDiagram datas){
        super.createDiagram();
    }

    @Override
    public void createDiagram(){
        super.createDiagram();
        createDiagram(4);
    }

    public abstract void createDiagram(int n);

    @Override
    public void setDrawArea(){
        try {
            //allow horizontal place for shadow properties
            m_PageProps.BorderTop += (SHADOW_DIST1 + 100);
            m_DrawAreaWidth -= ( 2 * SHADOW_DIST1 + 100);
            m_DrawAreaHeight -= (SHADOW_DIST1 + 100);

            int orignGSWidth = m_DrawAreaWidth;
            if ((m_DrawAreaWidth / m_iGroupWidth) <= (m_DrawAreaHeight / m_iGroupHeight)) {
                m_DrawAreaHeight = m_DrawAreaWidth * m_iGroupHeight / m_iGroupWidth;
            } else {
                m_DrawAreaWidth = m_DrawAreaHeight * m_iGroupWidth / m_iGroupHeight;
            }
            // set new size of m_xGroupShape for Organigram
            m_xGroupShape.setSize(new Size(m_DrawAreaWidth, m_DrawAreaHeight));
            m_iHalfDiff = 0;
            if (orignGSWidth > m_DrawAreaWidth)
                m_iHalfDiff = (orignGSWidth - m_DrawAreaWidth) / 2;
            m_iHalfDiff += SHADOW_DIST1;
            m_xGroupShape.setPosition(new Point(m_PageProps.BorderLeft + m_iHalfDiff, m_PageProps.BorderTop));
        } catch (PropertyVetoException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
    }

    public int getTopShapeID(){
        int iTopShapeID = -1;
        XShape xCurrShape = null;
        String currShapeName = "";
        int shapeID;
        try {
            for( int i=0; i < m_xShapes.getCount(); i++ ){
                xCurrShape = (XShape) UnoRuntime.queryInterface(XShape.class, m_xShapes.getByIndex(i));
                currShapeName = getShapeName(xCurrShape);
                if (currShapeName.contains("RectangleShape")) {
                    shapeID = getController().getShapeID(currShapeName);
                    if (shapeID > iTopShapeID)
                        iTopShapeID = shapeID;
                }
            }
        } catch (IndexOutOfBoundsException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
        return iTopShapeID;
    }

    public abstract OrganizationChartTree getDiagramTree();

    @Override
    public void removeShape(){
        XShapes xSelectedShapes = getController().getSelectedShapes();
        XShape xShape = null;
        try{
            if(xSelectedShapes != null){
                for(int i = 0; i < xSelectedShapes.getCount(); i++){
                    xShape = (XShape) UnoRuntime.queryInterface(XShape.class, xSelectedShapes.getByIndex(i));
                    if(xShape != null)
                        removeShape(xShape);
                }
            }
        } catch (IndexOutOfBoundsException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
    }

    public void removeShape(XShape xSelectedShape){
        if(xSelectedShape != null){
            XNamed xNamed = (XNamed) UnoRuntime.queryInterface( XNamed.class, xSelectedShape );
            String selectedShapeName = xNamed.getName();
            if(selectedShapeName.contains("RectangleShape") && !selectedShapeName.contains("RectangleShape0")){

                if( selectedShapeName.endsWith("RectangleShape1") ){
                    String title = getGui().getDialogPropertyValue("Strings", "RoutShapeRemoveError.Title");
                    String message = getGui().getDialogPropertyValue("Strings", "RoutShapeRemoveError.Message");
                    getGui().showMessageBox(title, message);
                }else{
                    // clear everythin under the item in the tree
                    OrganizationChartTreeItem selectedItem = getDiagramTree().getTreeItem(xSelectedShape);

                    boolean noItem = false;
                    OrganizationChartTreeItem dadItem = selectedItem.getDad();
                    if(selectedItem.equals(dadItem.getFirstChild())){
                        if(selectedItem.getFirstSibling() != null){
                            dadItem.setFirstChild(selectedItem.getFirstSibling());
                        }else{
                            dadItem.setFirstChild(null);
                            noItem = true;
                        }
                    }else{
                        OrganizationChartTreeItem previousSibling = getDiagramTree().getPreviousSibling(selectedItem);
                        if(previousSibling != null){
                            if(selectedItem.getFirstSibling() != null)
                                previousSibling.setFirstSibling(selectedItem.getFirstSibling());
                            else
                                previousSibling.setFirstSibling(null);
                        }
                    }

                    XShape xDadShape = selectedItem.getDad().getRectangleShape();

                    if (selectedItem.getFirstChild() != null)
                        selectedItem.getFirstChild().removeItems();

                    XShape xConnShape = getDiagramTree().getDadConnectorShape(xSelectedShape);
                    if (xConnShape != null){
                        getDiagramTree().removeFromConnectors(xConnShape);
                        m_xShapes.remove(xConnShape);
                    }
                    getDiagramTree().removeFromRectangles(xSelectedShape);
                    m_xShapes.remove(xSelectedShape);
                    setNullSelectedItem(selectedItem);
                    if(isHiddenRootElementProp() && getDiagramTree().getRootItem().getRectangleShape().equals(xDadShape)){
                        if(noItem){
                            getDiagramTree().getRootItem().hideElement();
                            setHiddenRootElementProp(false);
                            getController().setSelectedShape((Object)xDadShape);
                        }else{
                            getController().setSelectedShape((Object)getDiagramTree().getRootItem().getFirstChild().getRectangleShape());
                        }
                    }else{
                        getController().setSelectedShape((Object)xDadShape);
                    }
                }
            }
        }
    }

    public XShape getRootsConnector(){
        return getDiagramTree().getRootsConnector();
    }

    public boolean isErrorInTree(){
        if(getDiagramTree() != null){
            getDiagramTree().setLists();
            if((getDiagramTree().getRectangleListSize() - 1) != getDiagramTree().getConnectorListSize())
                return true;
            if(!getDiagramTree().isValidConnectors())
                return true;
            if(getDiagramTree().setRootItem() != 1)
                return true;
        }
        return false;
    }

    public void repairDiagram(){
        if(getDiagramTree().getRectangleListSize() == 0)
            clearEmptyDiagramAndReCreate();
        getDiagramTree().repairTree();
        initDiagram();
    }

    @Override
    public void refreshDiagram(){
        getDiagramTree().refresh();
    }

    public void refreshSchemesColors(){
        getDiagramTree().refreshSchemesColors();
    }

    public void refreshConnectorProps(){
        getDiagramTree().refreshConnectorProps();
    }

    public void setSchemesColors(short level, int iSteps){
        int colorCode = getColorModeProp() - Diagram.FIRST_COLORSCHEME_MODE_VALUE;
        setStartColorProp(SchemeDefinitions.getGradientColor(colorCode, level, iSteps));
        setEndColorProp(SchemeDefinitions.getGradientColor(colorCode, level+1, iSteps));
    }

    public void setColorSettingsOfShape(XShape xShape){
        if(isAnyGradientColorMode()){
            if(isColorSchemeMode()){
                int iSteps = OrganizationChartTreeItem._maxLevel + 1;
                short level = getLevelOfShape(xShape);
                if(level != -1)
                    setSchemesColors(level, iSteps);
            }
            setGradientWithAutoAngle(xShape);
            setLineColorOfShape(xShape);
        }else{
            setColorOfShape(xShape);
            setLineColorOfShape(xShape);
        }
    }

    public short getLevelOfShape(XShape xShape){
        OrganizationChartTree diagramTree = getDiagramTree();
        if(diagramTree != null){
            OrganizationChartTreeItem item = diagramTree.getTreeItem(xShape);
            if(item != null){
                if(getController().getDiagramType() == Controller.ORGANIGRAM)
                    return item.getDeepOfItem();
                else
                    return item.getLevel();
            }
        }
        return (short)-1;
    }
    
    public void setShapeProperties(XShape xShape, String type, boolean isModifyColorOfShape) {
        if(isModifyColorOfShape){
            setModifyColorsProp(true);
            setTextColorChange(true);
            setShapeProperties(xShape, type);
            setModifyColorsProp(false);
            setTextColorChange(false);
        } else {
            setShapeProperties(xShape, type);
        }
    }

    @Override
    public void setShapeProperties(XShape xShape, String type) {
        try {
            XPropertySet xProp = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xShape);
            
            if(type.equals("BaseShape")){
                if(isTextFitProp()){
                    xProp.setPropertyValue("TextFitToSize", TextFitToSizeType.PROPORTIONAL);
                }else{
                    xProp.setPropertyValue("TextFitToSize", TextFitToSizeType.NONE);
                    xProp.setPropertyValue("CharHeight", new Float(40.0));
                }
            }
  
            if(type.equals("RectangleShape")){
                if(isModifyColorsProp())
                    setColorSettingsOfShape(xShape);
                
                if(getRoundedProp() == Diagram.NULL_ROUNDED)
                    xProp.setPropertyValue("CornerRadius", new Integer(0));
                if(getRoundedProp() == Diagram.MEDIUM_ROUNDED)
                    xProp.setPropertyValue("CornerRadius", new Integer(CORNER_RADIUS2));
                if(getRoundedProp() == Diagram.EXTRA_ROUNDED)
                    xProp.setPropertyValue("CornerRadius", new Integer(CORNER_RADIUS3));

                if(isOutlineProp()){
                    xProp.setPropertyValue("LineStyle", LineStyle.SOLID);
                    xProp.setPropertyValue("LineWidth", new Integer(getShapesLineWidhtProp()));
                }else{
                    xProp.setPropertyValue("LineStyle", LineStyle.NONE);
                }

                if(isShadowProp()){
                    xProp.setPropertyValue("Shadow", new Boolean(true));
                    xProp.setPropertyValue("ShadowXDistance", new Integer(SHADOW_DIST1));
                    xProp.setPropertyValue("ShadowYDistance", new Integer(-SHADOW_DIST1));
                    xProp.setPropertyValue("ShadowTransparence", new Integer(SHADOW_TRANSP));
                    int shadowColor = -1;
                    if(((FillStyle)xProp.getPropertyValue("FillStyle")) == FillStyle.SOLID)
                        shadowColor = AnyConverter.toInt(xProp.getPropertyValue("FillColor"));
                    else{
                        int startColor = ((Gradient)xProp.getPropertyValue("FillGradient")).StartColor;
                        int endColor = ((Gradient)xProp.getPropertyValue("FillGradient")).EndColor;
                        shadowColor = startColor < endColor ? startColor : endColor;
                    }
                    if(shadowColor == -1)
                        shadowColor = 8421504;
                    xProp.setPropertyValue("ShadowColor", new Integer(shadowColor));
                }else{
                    xProp.setPropertyValue("Shadow", new Boolean(false));
                }
                setFontPropertiesOfShape(xShape);
            }
            if(type.equals("ConnectorShape")){
                if(isTextFitProp()){
                    xProp.setPropertyValue("TextFitToSize", TextFitToSizeType.PROPORTIONAL);
                }else{
                    xProp.setPropertyValue("TextFitToSize", TextFitToSizeType.NONE);
                }
                setConnectorShapeLineProps(xShape);
            }
        } catch (IllegalArgumentException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (UnknownPropertyException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (PropertyVetoException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
    }

    @Override
    public void setFontPropertyValues(){
        try {
            XPropertySet xPropText = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, getDiagramTree().getRootItem().getRectangleShape());
            TextFitToSizeType textFit = (TextFitToSizeType)xPropText.getPropertyValue("TextFitToSize");
            if(textFit.getValue() == TextFitToSizeType.NONE_value){
                setTextFitProp(false);
            } else{
                setTextFitProp(true);
            }
            float fontSizeValue = AnyConverter.toFloat(xPropText.getPropertyValue("CharHeight"));
            setFontSizeProp(fontSizeValue);
        } catch (Exception ex) {
            System.err.println(ex.getLocalizedMessage());
        }
    }

    public void setDiagramPropertyValues(boolean isSelectAllShape, boolean isModifyColors, short sRounded, boolean isOutline, int lineWidth, boolean isShadow){
        setSelectedAllShapesProp(isSelectAllShape);
        setModifyColorsProp(isModifyColors);
        setRoundedProp(sRounded);
        setOutlineProp(isOutline);
        setShapesLineWidthProp(lineWidth);
        setShadowProp(isShadow);
    }
 
    @Override
    public void showPropertyDialog(){
        getGui().enableControlDialogWindow(false);
        short exec = getGui().executePropertiesDialog();
        if(exec == 1){  
            getGui().setPropertiesOfOrganingram();

            if(isSelectedAllShapesProp()){
                setAllShapeProperties();
                setModifyColorsProp(false);
            } else {
                if(getSeletctedAreaProp() == Diagram.SELECTED_SHAPES){
                    setSelectedShapesProperties();
                }
                if(getSeletctedAreaProp() == Diagram.SIBLING_SHAPES){
                    setSiblingShapesProperties();
                }
                if(getSeletctedAreaProp() == Diagram.BRANCH_SHAPES){
                    setBranchShapesProperties();
                }
                setModifyColorsProp(false);
                setAllShapeProperties();
            }
            
            if(getController().getDiagramType() == Controller.ORGANIGRAM){   
                getController().getDiagram().refreshDiagram();
                refreshConnectorProps();
                if(getDiagramTree().getRootItem() != null)
                    getDiagramTree().getRootItem().hideElement();
            } else {
                if(getDiagramTree().getRootItem() != null)
                    getDiagramTree().getRootItem().hideElement();
                getController().getDiagram().refreshDiagram();
            }
        } 
        getGui().enableAndSetFocusControlDialog();
    }
    
    public void setAllShapeProperties(){
        try {
            XShape xCurrShape = null;
            String currShapeName = "";
            for(int i=0; i < m_xShapes.getCount(); i++){
                xCurrShape = (XShape) UnoRuntime.queryInterface(XShape.class, m_xShapes.getByIndex(i));
                currShapeName = getShapeName(xCurrShape);
                if (currShapeName.contains("RectangleShape") && !currShapeName.endsWith("RectangleShape0"))
                    setShapeProperties(xCurrShape,"RectangleShape");
                if (currShapeName.contains("ConnectorShape"))
                    setShapeProperties(xCurrShape,"ConnectorShape");
                if (currShapeName.contains("RectangleShape0"))
                    setShapeProperties(xCurrShape,"BaseShape");
            }
        } catch (IndexOutOfBoundsException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
    }
   
    public void setSelectedShapesProperties(){
         XShape xCurrShape = null;
         String currShapeName = "";
         XShapes xShapes = getController().getSelectedShapes();
         try{
            if(xShapes != null){
                for(int i = 0; i < xShapes.getCount(); i++){
                    xCurrShape = (XShape) UnoRuntime.queryInterface(XShape.class, xShapes.getByIndex(i));
                    currShapeName = getShapeName(xCurrShape);
                    if(currShapeName.endsWith("GroupShape"))
                        setAllShapeProperties();
                    else
                        if (currShapeName.contains("RectangleShape") && !currShapeName.endsWith("RectangleShape0"))
                            setShapeProperties(xCurrShape,"RectangleShape");
                }
            }
        } catch (IndexOutOfBoundsException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
    }

    public void setSiblingShapesProperties(){
         XShape xCurrShape = null;
         String currShapeName = "";
         XShapes xShapes = getController().getSelectedShapes();

         try{
            if(xShapes != null){
                for(int i = 0; i < xShapes.getCount(); i++){
                    xCurrShape = (XShape) UnoRuntime.queryInterface(XShape.class, xShapes.getByIndex(i));
                    currShapeName = getShapeName(xCurrShape);
                    if(currShapeName.endsWith("GroupShape"))
                        setAllShapeProperties();
                    else
                        if (currShapeName.contains("RectangleShape") && !currShapeName.endsWith("RectangleShape0"))
                            setSiblingShapesProperties(xCurrShape);
                }
            }
        } catch (IndexOutOfBoundsException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
    }

    public void setSiblingShapesProperties(XShape xShape){
        OrganizationChartTree diagramTree = getDiagramTree();
        if(diagramTree != null){
            OrganizationChartTreeItem selectedItem = diagramTree.getTreeItem(xShape);
            if(selectedItem != null){
                OrganizationChartTreeItem dadItem = selectedItem.getDad();
                if(dadItem != null){
                    OrganizationChartTreeItem item = dadItem.getFirstChild();
                    while(item != null){
                        setShapeProperties(item.m_xRectangleShape, "RectangleShape");
                        item = item.getFirstSibling();
                    }
                }else{
                    setShapeProperties(xShape, "RectangleShape");
                }

            }
        }
    }

    public void setBranchShapesProperties(){
         XShape xCurrShape = null;
         String currShapeName = "";
         XShapes xShapes = getController().getSelectedShapes();

         try{
            if(xShapes != null){
                for(int i = 0; i < xShapes.getCount(); i++){
                    xCurrShape = (XShape) UnoRuntime.queryInterface(XShape.class, xShapes.getByIndex(i));
                    currShapeName = getShapeName(xCurrShape);
                    if(currShapeName.endsWith("GroupShape"))
                        setAllShapeProperties();
                    else
                        if (currShapeName.contains("RectangleShape") && !currShapeName.endsWith("RectangleShape0"))
                            setBranchShapesProperties(xCurrShape);
                }
            }
        } catch (IndexOutOfBoundsException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
    }

    public void setBranchShapesProperties(XShape xShape){
        setShapeProperties(xShape, "RectangleShape");
        OrganizationChartTree diagramTree = getDiagramTree();
        if(diagramTree != null){
            OrganizationChartTreeItem selectedItem = diagramTree.getTreeItem(xShape);
            OrganizationChartTreeItem item = selectedItem.getFirstChild();
            if(item != null)
                item.setProperties();
        }
    }

    public void setShapesGradientColor(XShape xRectangleShape, short level, int iSteps){
            if(level != -1)
                setSchemesColors(level, iSteps);
            if(getGradientDirectionProp() == Diagram.VERTICAL)
                setGradient(xRectangleShape);
            else
                setGradient(xRectangleShape, (short)900);
    }
    
    //public abstract void setConnectorShapeProps(XShape xConnectorShape, XShape xStartShape, Integer startIndex, XShape xEndShape, Integer endIndex);
    public void setConnectorShapeProps(XShape xConnectorShape, XShape xStartShape, Integer startIndex, XShape xEndShape, Integer endIndex){
        try {
            XPropertySet xProps = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xConnectorShape);
            xProps.setPropertyValue("StartShape", xStartShape);
            xProps.setPropertyValue("EndShape", xEndShape);
            xProps.setPropertyValue("StartGluePointIndex", startIndex);
            xProps.setPropertyValue("EndGluePointIndex", endIndex);
            xProps.setPropertyValue("LineColor", new Integer(getConnectorColorProp()));
            setConnectorShapeLineProps(xConnectorShape);
            if(isTextFitProp())
                xProps.setPropertyValue("TextFitToSize", TextFitToSizeType.PROPORTIONAL);
            else
                xProps.setPropertyValue("TextFitToSize", TextFitToSizeType.NONE);
        } catch (com.sun.star.beans.PropertyVetoException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (IllegalArgumentException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (UnknownPropertyException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
    }

    //public abstract void setConnectorShapeProps(XShape xConnShape, Integer start, Integer end);
    public void setConnectorShapeProps(XShape xConnectorShape, Integer startIndex, Integer endIndex){
        try {
            XPropertySet xProps = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xConnectorShape);
            xProps.setPropertyValue("StartGluePointIndex", startIndex);
            xProps.setPropertyValue("EndGluePointIndex", endIndex);
            setConnectorShapeLineProps(xConnectorShape);
        } catch (com.sun.star.beans.PropertyVetoException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (UnknownPropertyException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (IllegalArgumentException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
    }

    public void setConnectorShapeLineProps(XShape xConnectorShape){
        try {
            XPropertySet xProps = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xConnectorShape);
            xProps.setPropertyValue("LineColor",new Integer(getConnectorColorProp()));
            if(isShownConnectorsProp())
                xProps.setPropertyValue("LineStyle", LineStyle.SOLID);
            else
                xProps.setPropertyValue("LineStyle", LineStyle.NONE);
            if(getConnectorTypeProp() == Diagram.CONN_STANDARD)
                xProps.setPropertyValue("EdgeKind", ConnectorType.STANDARD);
            if(getConnectorTypeProp() == Diagram.CONN_LINE)
                xProps.setPropertyValue("EdgeKind", ConnectorType.LINE);
            if(getConnectorTypeProp() == Diagram.CONN_STRAIGHT)
                xProps.setPropertyValue("EdgeKind", ConnectorType.LINES);
            if(getConnectorTypeProp() == Diagram.CONN_CURVED)
                xProps.setPropertyValue("EdgeKind", ConnectorType.CURVE);

            if(isConnectorStartArrowProp())
                xProps.setPropertyValue("LineStartName", "Arrow");
            else
                xProps.setPropertyValue("LineStartName", "");
            if(isConnectorEndArrowProp())
                xProps.setPropertyValue("LineEndName", "Arrow");
            else
                xProps.setPropertyValue("LineEndName", "");

            xProps.setPropertyValue("LineWidth",new Integer(getConnectorsLineWidhtProp()));
            int lineWidth = getConnectorsLineWidhtProp();
            int arrowWidth = 400;
            if(lineWidth == 200)
                arrowWidth = 600;
            else if(lineWidth >= 300)
                arrowWidth = (int)(lineWidth * 2.5);
            xProps.setPropertyValue("LineStartWidth",new Integer(arrowWidth));
            xProps.setPropertyValue("LineEndWidth",new Integer(arrowWidth));
        } catch (com.sun.star.beans.PropertyVetoException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (UnknownPropertyException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (IllegalArgumentException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
    }

    public void clearEmptyDiagramAndReCreate(){
        try {
             if(m_xShapes != null){
                XShape xShape = null;
                for( int i=0; i < m_xShapes.getCount(); i++ ){
                    xShape = (XShape) UnoRuntime.queryInterface(XShape.class, m_xShapes.getByIndex(i));
                    if (xShape != null)
                        m_xShapes.remove(xShape);
                }
            }
            createDiagram(1);
        } catch (IndexOutOfBoundsException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
    }

    public void setNullSelectedItem(OrganizationChartTreeItem item){
        item.setDad(null);
        item.setFirstChild(null);
        item.setFirstSibling(null);
    }
    
}
