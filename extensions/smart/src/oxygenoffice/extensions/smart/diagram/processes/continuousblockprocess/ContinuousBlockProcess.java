package oxygenoffice.extensions.smart.diagram.processes.continuousblockprocess;

import com.sun.star.awt.Gradient;
import com.sun.star.awt.GradientStyle;
import com.sun.star.awt.Point;
import com.sun.star.awt.Size;
import com.sun.star.beans.PropertyVetoException;
import com.sun.star.beans.UnknownPropertyException;
import com.sun.star.beans.XPropertySet;
import com.sun.star.drawing.FillStyle;
import com.sun.star.drawing.LineStyle;
import com.sun.star.drawing.XShape;
import com.sun.star.frame.XFrame;
import com.sun.star.lang.IndexOutOfBoundsException;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.text.XText;
import com.sun.star.uno.AnyConverter;
import com.sun.star.uno.UnoRuntime;
import java.util.ArrayList;
import oxygenoffice.extensions.smart.Controller;
import oxygenoffice.extensions.smart.diagram.Color;
import oxygenoffice.extensions.smart.diagram.Diagram;
import oxygenoffice.extensions.smart.diagram.SchemeDefinitions;
import oxygenoffice.extensions.smart.diagram.processes.ProcessDiagram;
import oxygenoffice.extensions.smart.diagram.processes.ProcessDiagramItem;
import oxygenoffice.extensions.smart.gui.Gui;


public class ContinuousBlockProcess extends ProcessDiagram{
    
    public final static short   DEFAULT                 = 0;
    public final static short   NOT_ROUNDED             = 1;
    public final static short   WITHOUT_OUTLINE         = 2;
    public final static short   WITH_SHADOW             = 3;
    public final static short   GREEN_DARK_GRADIENT     = 4;
    public final static short   GREEN_BRIGHT_GRADIENT   = 5;
    public final static short   BLUE_DARK_GRADIENT      = 6;
    public final static short   BLUE_BRIGHT_GRADIENT    = 7;
    public final static short   PURPLE_DARK_GRADIENT    = 8;
    public final static short   PURPLE_BRIGHT_GRADIENT  = 9;
    public final static short   ORANGE_DARK_GRADIENT    = 10;
    public final static short   ORANGE_BRIGHT_GRADIENT  = 11;
    public final static short   YELLOW_DARK_GRADIENT    = 12;
    public final static short   YELLOW_BRIGHT_GRADIENT  = 13;
    public final static short   BLUE_SCHEME             = 14;
    public final static short   AQUA_SCHEME             = 15;
    public final static short   RED_SCHEME              = 16;
    public final static short   FIRE_SCHEME             = 17;
    public final static short   SUN_SCHEME              = 18;
    public final static short   GREEN_SCHEME            = 19;
    public final static short   OLIVE_SCHEME            = 20;
    public final static short   PURPLE_SCHEME           = 21;
    public final static short   PINK_SCHEME             = 22;
    public final static short   INDIAN_SCHEME           = 23;
    public final static short   MAROON_SCHEME           = 24;
    public final static short   BROWN_SCHEME            = 25;
    public final static short   USER_DEFINE             = 26;

    public final static short   FIRST_COLORTHEMEGRADIENT_STYLE_VALUE    = 4;
    public final static short   FIRST_COLORSCHEME_STYLE_VALUE           = 14;
   
    private XShape m_xArrowShape    = null;
    private int m_iArrowShapeColor = aRAINBOWCOLORS[0];

    public ContinuousBlockProcess(Controller controller, Gui gui, XFrame xFrame) {
        super(controller, gui, xFrame);
        setDefaultProps();
        setColorModeProp(Diagram.BASE_COLORS_MODE);
        m_iGroupWidth   = 10;
        m_iGroupHeight  = 7;
    }
    
    @Override
    public boolean isColorSchemeStyle(short style){
        return  style == BLUE_SCHEME || style == AQUA_SCHEME ||
                style == RED_SCHEME || style == FIRE_SCHEME ||
                style == SUN_SCHEME || style == GREEN_SCHEME ||
                style == OLIVE_SCHEME || style == PURPLE_SCHEME ||
                style == PINK_SCHEME || style == INDIAN_SCHEME ||
                style == MAROON_SCHEME || style == BROWN_SCHEME;
    }

    @Override
    public short getColorModeOfSchemeStyle(short style){
        return (short)(style - FIRST_COLORSCHEME_STYLE_VALUE + Diagram.FIRST_COLORSCHEME_MODE_VALUE);
    }
    
    @Override
    public boolean isColorThemeGradientStyle(short style){
        return  style == GREEN_DARK_GRADIENT || style == GREEN_BRIGHT_GRADIENT ||
                style == BLUE_DARK_GRADIENT || style == BLUE_BRIGHT_GRADIENT || 
                style == PURPLE_DARK_GRADIENT || style == PURPLE_BRIGHT_GRADIENT ||
                style == ORANGE_DARK_GRADIENT || style == ORANGE_BRIGHT_GRADIENT ||
                style == YELLOW_DARK_GRADIENT || style == YELLOW_BRIGHT_GRADIENT;
    }
    
    @Override
    public short getColorModeOfThemeGradientStyle(short style){
        return (short)(style - FIRST_COLORTHEMEGRADIENT_STYLE_VALUE + Diagram.FIRST_COLORTHEMEGRADIENT_MODE_VALUE);
    }
    
    @Override
    public short getUserDefineStyleValue(){
        return USER_DEFINE;
    }
    
    @Override
    public String getDiagramTypeName(){
        return "ContinuousBlockProcess";
    }
    
    public void setArrowShapeColorProp(int color){
        m_iArrowShapeColor = color;
    }
    
    public int getArrowShapeColorProp(){
        return m_iArrowShapeColor;
    }
    
    @Override
    public void createDiagram(){
        super.createDiagram();
        createDiagram(3);
    }
    
    @Override
    public void initDiagram(){
        super.initDiagram();
        if(!items.isEmpty())
            items.clear();
        try {
            XShape xCurrShape;
            String currShapeName = "";
            int currShapeID;
            for(int i = 0; i < m_xShapes.getCount(); i++ ){
                xCurrShape = (XShape) UnoRuntime.queryInterface(XShape.class, m_xShapes.getByIndex(i));
                currShapeName = getShapeName(xCurrShape);
                if(currShapeName.endsWith("PolyPolygonShape0"))
                    m_xArrowShape = xCurrShape;
                if(currShapeName.endsWith("RectangleShape0"))
                    m_xControlShape = xCurrShape;
            }
            int topShapeID = getTopShapeID();
            for(int i = 1; i <= topShapeID; i++ ){
                for(int j = 0; j < m_xShapes.getCount(); j++ ){
                    xCurrShape = (XShape) UnoRuntime.queryInterface(XShape.class, m_xShapes.getByIndex(j));
                    currShapeName = getShapeName(xCurrShape);
                    if (currShapeName.contains("RectangleShape")) {
                        currShapeID = getController().getShapeID(currShapeName);
                        if(currShapeID == i)
                            addItem(new ContinuousBlockProcessItem(this, currShapeID, xCurrShape));
                    }
                }
            }
        } catch (IndexOutOfBoundsException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
        setControlShapePropsWithoutTextProps(getControlShape());
    }
    
    @Override
    public void createControlShape(){
        m_xControlShape = createShape("RectangleShape", 0, m_PageProps.BorderLeft + m_iHalfDiff, m_PageProps.BorderTop, m_DrawAreaWidth, m_DrawAreaHeight);
        m_xShapes.add(m_xControlShape);
        setControlShapeProps(getControlShape());
    }
    
    public void setArrowShape(){
        if(m_xShapes != null){
            XShape xCurrShape = null;
            String currShapeName = "";
            try {
                for( int i=0; i < m_xShapes.getCount(); i++ ){
                    xCurrShape = (XShape) UnoRuntime.queryInterface(XShape.class, m_xShapes.getByIndex(i));
                    currShapeName = getShapeName(xCurrShape);
                    if (currShapeName.endsWith("PolyPolygonShape0"))
                        m_xArrowShape = xCurrShape;
                }
            } catch (com.sun.star.lang.IndexOutOfBoundsException ex) {
                System.err.println(ex.getLocalizedMessage());
            } catch (WrappedTargetException ex) {
                System.err.println(ex.getLocalizedMessage());
            }
        }
    }
    
    public XShape getArrowShape(){
        if(m_xArrowShape == null && m_xShapes != null){
            XShape xCurrShape = null;
            String currShapeName = "";
            try {
                for( int i=0; i < m_xShapes.getCount(); i++ ){
                    xCurrShape = (XShape) UnoRuntime.queryInterface(XShape.class, m_xShapes.getByIndex(i));
                    currShapeName = getShapeName(xCurrShape);
                    if (currShapeName.endsWith("PolyPolygonShape0"))
                        m_xArrowShape = xCurrShape;
                }
            } catch (com.sun.star.lang.IndexOutOfBoundsException ex) {
                System.err.println(ex.getLocalizedMessage());
            } catch (WrappedTargetException ex) {
                System.err.println(ex.getLocalizedMessage());
            }
        }
        return m_xArrowShape;
    }
    
    public void createArrowShape(boolean isNewColor){
        try{
            m_xArrowShape = createShape("PolyPolygonShape", 0);
            m_xShapes.add(m_xArrowShape);
            if(isNewColor){
                if(isBaseColorsMode() || isBaseColorsWithGradientMode())
                    setArrowShapeColorProp(_aBaseColors[0]);
                else
                    setArrowShapeColorProp(getColorProp());
            }
            setArrowShapeProps();
            setPosOfArrowShape();
        } catch(Exception ex){
            System.err.println(ex.getLocalizedMessage());
        }
    }
    
    public void removeArrowShape(){
        m_xShapes.remove(getArrowShape());
        m_xArrowShape = null;
    }
    
    public void setArrowShapeProps(){
        try {
            if(getArrowShape() != null){
                XPropertySet xProp = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, m_xArrowShape);
                setMoveProtectOfShape(m_xArrowShape);
                if(isColorSchemeMode() && m_iArrowShapeColor == SchemeDefinitions.aColorSchemes[getColorModeProp() - FIRST_COLORSCHEME_MODE_VALUE][1]){
                    int startColor = SchemeDefinitions.aColorSchemes[getColorModeProp() - FIRST_COLORSCHEME_MODE_VALUE][0];
                    int endColor = SchemeDefinitions.aColorSchemes[getColorModeProp() - FIRST_COLORSCHEME_MODE_VALUE][1];
                    setGradient(m_xArrowShape, GradientStyle.LINEAR, startColor, endColor, (short)900, (short)10, (short)0, (short)0, (short)100, (short)100);
                }else{
                    setColorOfShape(m_xArrowShape, m_iArrowShapeColor);
                }
                xProp.setPropertyValue("FillTransparence", new Integer(TRANSP_VALUE3));
                xProp.setPropertyValue("LineStyle", LineStyle.NONE);
                setZOrder(m_xArrowShape, "PolyPolygonShape");
            }
        } catch (Exception ex) {
            System.err.println(ex.getLocalizedMessage());
        }
    }
    
    public void setPosOfArrowShape(){
        try{
            if(getArrowShape() != null){
                Point csPoint = getControlShape().getPosition();
                Size  csSize = getControlShape().getSize();
                int x = csPoint.X;
                int y = csPoint.Y;
                int width = csSize.Width;
                int height = csSize.Height;
                Point a, b, c, d, e, f, g;
                a = new Point(x + width / 64, y + height / 4);
                b = new Point(x + width / 64 * 40, y + height / 4);
                c = new Point(x + width / 64 * 40, y);
                d = new Point(x + width / 64 * 63, y + height / 2);
                e = new Point(x + width / 64 * 40, y + height);
                f = new Point(x + width / 64 * 40, y + height / 4 * 3);
                g = new Point(x + width / 64, y + height / 4 * 3);

                Point[] points1 = new Point[7];
                points1[0] = a;
                points1[1] = b;
                points1[2] = c;
                points1[3] = d;
                points1[4] = e;
                points1[5] = f;
                points1[6] = g;

                Point[] points2 = new Point[2];
                points2[0] = a;
                points2[1] = g;

                Point[][] allPoints = new Point[2][];
                allPoints[0] = points1;
                allPoints[1] = points2;

                XPropertySet xPolygonShapeProps = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, m_xArrowShape);
                xPolygonShapeProps.setPropertyValue("PolyPolygon", allPoints);
            }
        } catch(Exception ex){
            System.err.println(ex.getLocalizedMessage());
        }
    }
    
    @Override
    public ProcessDiagramItem createItem(int shapeID) {
        return createItem(shapeID, " ");
    }
    
    @Override
    public ProcessDiagramItem createItem(int shapeID, String str){
        XShape xRectangleShape = createShape("RectangleShape", shapeID);
        m_xShapes.add(xRectangleShape);
        setMoveProtectOfShape(xRectangleShape);
        setZOrder(xRectangleShape, "RectangleShape");
        
        ContinuousBlockProcessItem item = new ContinuousBlockProcessItem(this, shapeID, xRectangleShape);
        addItem(item);
        item.setText(str);
        item.setShapesProps(true);
 
        setTextColorOfShape(item.getTextShape());
        getController().setSelectedShape((Object)xRectangleShape);

        return item;
    }
    
    @Override
    public ProcessDiagramItem createItem(int shapeID, String str, Color oColor){
        XShape xRectangleShape = createShape("RectangleShape", shapeID);
        m_xShapes.add(xRectangleShape);
        setMoveProtectOfShape(xRectangleShape);
        setZOrder(xRectangleShape, "RectangleShape");
        
        ContinuousBlockProcessItem item = new ContinuousBlockProcessItem(this, shapeID, xRectangleShape);
        addItem(item);
        item.setText(str);
        item.setShapesProps();
        item.setColor(oColor);
        item.setLineColor();
        setTextColorOfShape(item.getTextShape());
        getController().setSelectedShape((Object)xRectangleShape);
   
        return item;
    }
    
    public void setZOrder(XShape xShape, String type){
        try {
            XPropertySet xProp = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xShape);
            if(type.equals("RectangleShape"))
                xProp.setPropertyValue("ZOrder", new Integer(2));
            if(type.equals("PolyPolygonShape"))
                xProp.setPropertyValue("ZOrder", new Integer(1));
        } catch (Exception ex) {
            System.err.println(ex.getLocalizedMessage());
        }
    }
    
    @Override
    public XShape getControlShape(){
        if(m_xControlShape == null && m_xShapes != null){
            XShape xCurrShape = null;
            String currShapeName = "";
            try {
                for( int i=0; i < m_xShapes.getCount(); i++ ){
                    xCurrShape = (XShape) UnoRuntime.queryInterface(XShape.class, m_xShapes.getByIndex(i));
                    currShapeName = getShapeName(xCurrShape);
                    if (currShapeName.endsWith("RectangleShape0"))
                        m_xControlShape = xCurrShape;
                }
            } catch (com.sun.star.lang.IndexOutOfBoundsException ex) {
                System.err.println(ex.getLocalizedMessage());
            } catch (WrappedTargetException ex) {
                System.err.println(ex.getLocalizedMessage());
            }
        }
        return m_xControlShape;
    }
    
    public void setColorSettingsOfShape(XShape xShape){
        if(isColorSchemeMode() || isColorThemeGradientMode()){
            if(isColorSchemeMode()){
                int shapeID = getController().getShapeID(getShapeName(xShape));
                if(shapeID > 0)
                    setSchemesColors(shapeID);
                setGradientDirectionProp(Diagram.HORIZONTAL);
                setLineColorProp(getDefaultLineColor());
            }
            if(isColorThemeGradientMode()){
                setGradientDirectionProp(Diagram.VERTICAL);
                setColorThemeGradientColors();
            }
            setGradientWithAutoAngle(xShape);
        }else{
            if(isSimpleColorMode())
                setLineColorProp(getDefaultLineColor());    
            if(isBaseColorsMode())
                setLineColorProp(getDefaultLineColor());
            setColorOfShape(xShape);
        }
        setLineColorOfShape(xShape);
    }
    
    @Override
    public void setShapeProperties(XShape xShape, String type) {
        setFontPropertiesOfShape(xShape);
        XPropertySet xProp = null;
        try {
            xProp = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xShape);

            if(type.equals("BaseShape")){
                if(!isTextFitProp())
                    xProp.setPropertyValue("CharHeight", new Float(40.0));
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
                
                if(m_IsShadow){
                    xProp.setPropertyValue("Shadow", new Boolean(true));
                    int numOfItems = getNumOfItems();
                    int shadowDist = SHADOW_DIST1;
                    if(numOfItems >= 4 && numOfItems <= 6)
                        shadowDist = SHADOW_DIST1 * 5 / 6;
                    if(numOfItems >= 7 && numOfItems <= 9)
                        shadowDist =  SHADOW_DIST1 * 4 / 6;
                    if(numOfItems >= 10)
                        shadowDist = SHADOW_DIST1 / 2;
                    xProp.setPropertyValue("ShadowXDistance", new Integer(shadowDist));
                    xProp.setPropertyValue("ShadowYDistance", new Integer(-shadowDist));
                    xProp.setPropertyValue("ShadowTransparence", new Integer(SHADOW_TRANSP));
                    int shadowColor = -1;
                    if(((FillStyle)xProp.getPropertyValue("FillStyle")) == FillStyle.SOLID)
                        shadowColor = AnyConverter.toInt(xProp.getPropertyValue("FillColor"));
                    else
                        shadowColor = ((Gradient)xProp.getPropertyValue("FillGradient")).EndColor;
                    if(shadowColor == -1)
                        shadowColor = 8421504;
                    xProp.setPropertyValue("ShadowColor", new Integer(shadowColor));
                }else{
                    xProp.setPropertyValue("Shadow", new Boolean(false));
                }
            
                if(m_IsOutline){
                    xProp.setPropertyValue("LineStyle", LineStyle.SOLID);
                    xProp.setPropertyValue("LineWidth", new Integer(getShapesLineWidhtProp()));
                } else {
                    xProp.setPropertyValue("LineStyle", LineStyle.NONE);
                }
            }

            if(type.equals("PolyPolygonShape")){
                if(isModifyColorsProp())
                    setArrowShapeProps();
                    //setColorOfShape(xShape, getArrowShapeColorProp());
            }
           
        } catch (com.sun.star.lang.IllegalArgumentException ex) {
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
    public void refreshDiagram() {
        removeSingleItems();
        int numOfItems = getNumOfItems();
        if(numOfItems > 0){
            //init
            Point controlShapePoint = getControlShape().getPosition();
            Size controlShapeSize = getControlShape().getSize();
            for(ProcessDiagramItem item : items){
                ((ContinuousBlockProcessItem)item).setPosition(numOfItems, controlShapeSize, controlShapePoint);
                if(isColorSchemeMode())
                    setGradientColor(item.getMainShape());
            }
        }
    }
    
    public void setGradientColor(XShape xShape){
        int shapeID = getController().getShapeID(getShapeName(xShape));
        if(shapeID > 0){
            setSchemesColors(shapeID);
            setGradientDirectionProp(Diagram.HORIZONTAL);
            setGradientWithAutoAngle(xShape);
        }
    }
    
    public void setSchemesColors(int shapeID){
        int colorCode = getColorModeProp() - Diagram.FIRST_COLORSCHEME_MODE_VALUE;
        int topShapeID = getTopShapeID();
        setStartColorProp(SchemeDefinitions.getGradientColor(colorCode, shapeID - 1, topShapeID));
        setEndColorProp(SchemeDefinitions.getGradientColor(colorCode, shapeID, topShapeID));
    }
    
    @Override
    public void addShape(){
        int selectedShapeID = getSelectedShapeID();
        if(selectedShapeID > 0)
            increaseItemsIDs(selectedShapeID);
        if(selectedShapeID == -1)
            selectedShapeID = getTopShapeID();
        if(selectedShapeID == -1 ){
            super.createDiagram();
            createDiagram(1);
        }else{
            if(getTopShapeID() == 1)
                createArrowShape(false);
            createItem(selectedShapeID + 1);
            if(isBaseColorsMode())
                setColorProp(_aBaseColors[(selectedShapeID + 1) % 8]);
        }
    }

    @Override
    public int getSelectedShapeID(){
        String shapeName = getShapeName(getController().getSelectedShape());
        if(shapeName.contains("ContinuousBlockProcess") && (shapeName.contains("PolyPolygonShape") || shapeName.contains("RectangleShape"))){
            int currShapeID = getController().getShapeID(shapeName);
            if(currShapeID >= 0)
                return currShapeID;
        }
        return -1;
    }

    @Override
    public int getTopShapeID(){
        XShape xCurrShape = null;
        String currShapeName = "";
        int shapeID = -1;
        int iTopShapeID = -1;
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

    @Override
    public void removeShape() {
        int selectedShapeID = getSelectedShapeID();
        if(selectedShapeID > 0){
            XShape xSelectedShape = getController().getSelectedShape();
            ProcessDiagramItem selectedItem = getItem(xSelectedShape);
            if(selectedItem != null){
                String selectedShapeName = getShapeName(xSelectedShape);
                XShape xNextSelectedShape = getPreviousShape();
                if(!selectedShapeName.endsWith("RectangleShape0")){
                    if(getTopShapeID() == 2)
                        removeArrowShape();
                    if(selectedShapeID == 1 && getTopShapeID() == 1)
                        xNextSelectedShape = getControlShape();
                    selectedItem.removeItem();
                    decreaseItemsIDs(selectedShapeID);
                }
                if(xNextSelectedShape != null)
                    getController().setSelectedShape(xNextSelectedShape);
                if(isBaseColorsMode())
                    setColorProp(getNextColor());
            }
        }
    }
    
    @Override
    public XShape getNextShape(){
        return getNextItem().getMainShape();
    }
    
    @Override
    public XShape getPreviousShape(){
        return getPreviousItem().getMainShape();
    }

    @Override
    public XShape getPairOfMainShape(int id) {
        return null;
    }
    
    @Override
    public void showPropertyDialog(){
        getGui().enableControlDialogWindow(false);
        short exec = getGui().executePropertiesDialog();
        if(exec == 1){  
            getGui().setPropertiesOfContinuousBlockProcess();
            if(isSelectedAllShapesProp()){
                setAllShapeProperties();
                setModifyColorsProp(false);
            } else {
                setSelectedShapesProperties();
                setModifyColorsProp(false);
                setAllShapeProperties();
            }
            
            getController().getDiagram().refreshDiagram();
        } 
        getGui().enableAndSetFocusControlDialog();
    }
    
    public void setPropertiesValues(boolean isSelectAllShape, boolean isModifyColors, short sColorMode, short sRounded, boolean isOutline, int lineWidth, boolean isShadow){
        setSelectedAllShapesProp(isSelectAllShape);
        setModifyColorsProp(isModifyColors);
        setColorModeProp(sColorMode);
        setRoundedProp(sRounded);
        setOutlineProp(isOutline);
        setShapesLineWidthProp(lineWidth);
        setShadowProp(isShadow);
    }
    
    public void initArrowShapeColor(){
        XPropertySet xProps = (XPropertySet)UnoRuntime.queryInterface( XPropertySet.class, getArrowShape());
        try {
            setArrowShapeColorProp(AnyConverter.toInt(xProps.getPropertyValue("FillColor")));
        } catch (Exception ex) {
            System.err.println(ex.getLocalizedMessage());
        }
    }
    
    @Override
    public void initProperties(XShape xControlShape, ArrayList<ProcessDiagramItem> items){
        super.initProperties(xControlShape, items);
        initArrowShapeColor();
        XPropertySet xProps = (XPropertySet)UnoRuntime.queryInterface( XPropertySet.class, items.get(0).getMainShape());
        try {
            if(getStyleProp() == ContinuousBlockProcess.NOT_ROUNDED)
                setRoundedProp(Diagram.NULL_ROUNDED);
            if(getStyleProp() == ContinuousBlockProcess.WITHOUT_OUTLINE)
                setOutlineProp(false);
            if(getStyleProp() == ContinuousBlockProcess.WITH_SHADOW)
                setShadowProp(true);
            if(getStyleProp() == ContinuousBlockProcess.USER_DEFINE){
                
                int cornerRadius = AnyConverter.toInt(xProps.getPropertyValue("CornerRadius"));
                if(cornerRadius < Diagram.CORNER_RADIUS2 / 2)
                    setRoundedProp(Diagram.NULL_ROUNDED);
                if(cornerRadius >= Diagram.CORNER_RADIUS2 / 2 && cornerRadius < Diagram.CORNER_RADIUS2 + ((Diagram.CORNER_RADIUS3 - Diagram.CORNER_RADIUS2) / 2) )
                    setRoundedProp(Diagram.MEDIUM_ROUNDED);
                else
                    setRoundedProp(Diagram.EXTRA_ROUNDED);
                
                if(((LineStyle)xProps.getPropertyValue("LineStyle")).getValue() == LineStyle.NONE_value)
                    setOutlineProp(false);
                setShapesLineWidthProp(AnyConverter.toInt(xProps.getPropertyValue("LineWidth")));
                            
                if(AnyConverter.toBoolean(xProps.getPropertyValue("Shadow")))
                    setShadowProp(true);               
            }

            setFontPropertyValues();
            XText xText = (XText)UnoRuntime.queryInterface(XText.class, items.get(0).getTextShape());
            XPropertySet xTextProps = (XPropertySet)UnoRuntime.queryInterface( XPropertySet.class, xText.createTextCursor());
            setTextColorProp(AnyConverter.toInt(xTextProps.getPropertyValue("CharColor")));
        } catch (UnknownPropertyException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (com.sun.star.lang.IllegalArgumentException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.err.println(ex.getLocalizedMessage());
        } 
    }
    
}
