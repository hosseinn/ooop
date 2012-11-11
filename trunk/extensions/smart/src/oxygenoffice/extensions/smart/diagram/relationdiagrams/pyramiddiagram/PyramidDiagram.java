package oxygenoffice.extensions.smart.diagram.relationdiagrams.pyramiddiagram;

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
import com.sun.star.drawing.XShapes;
import com.sun.star.frame.XFrame;
import com.sun.star.lang.IllegalArgumentException;
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
import oxygenoffice.extensions.smart.diagram.relationdiagrams.RelationDiagram;
import oxygenoffice.extensions.smart.diagram.relationdiagrams.RelationDiagramItem;
import oxygenoffice.extensions.smart.gui.Gui;


public class PyramidDiagram extends RelationDiagram {


    protected int               m_iGroupWidth;
    protected int               m_iGroupHeight;

    public final static short   DEFAULT             = 0;
    public final static short   WITHOUT_OUTLINE     = 1;
    public final static short   WITH_SHADOW         = 2;
    public final static short   BC_WITH_GRADIENTS   = 3;   
    public final static short   GREEN_DARK      = 4;
    public final static short   GREEN_BRIGHT    = 5;
    public final static short   BLUE_DARK       = 6;
    public final static short   BLUE_BRIGHT     = 7;
    public final static short   PURPLE_DARK     = 8;
    public final static short   PURPLE_BRIGHT   = 9;
    public final static short   ORANGE_DARK     = 10;
    public final static short   ORANGE_BRIGHT   = 11;
    public final static short   YELLOW_DARK     = 12;
    public final static short   YELLOW_BRIGHT   = 13;
    public final static short   BLUE_SCHEME      = 14;
    public final static short   AQUA_SCHEME      = 15;
    public final static short   RED_SCHEME       = 16;
    public final static short   FIRE_SCHEME      = 17;
    public final static short   SUN_SCHEME       = 18;
    public final static short   GREEN_SCHEME     = 19;
    public final static short   OLIVE_SCHEME     = 20;
    public final static short   PURPLE_SCHEME    = 21;
    public final static short   PINK_SCHEME      = 22;
    public final static short   INDIAN_SCHEME    = 23;
    public final static short   MAROON_SCHEME    = 24;
    public final static short   BROWN_SCHEME     = 25;
    public final static short   USER_DEFINE         = 26;

    public final static short   FIRST_COLORTHEME_STYLE_VALUE = 4;
    public final static short   FIRST_COLORSCHEME_STYLE_VALUE = 14;
    
    

    public PyramidDiagram(Controller controller, Gui gui, XFrame xFrame) {
        super(controller, gui, xFrame);
        m_iGroupWidth   = 5;
        m_iGroupHeight  = 4;
        setDefaultProps();
        setColorModeProp(Diagram.BASE_COLORS_MODE);
    }
    
    @Override
    public short getUserDefineStyleValue(){
        return USER_DEFINE;
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
    public boolean isColorThemeStyle(short style){
        return  style == GREEN_DARK || style == GREEN_BRIGHT ||
                style == BLUE_DARK || style == BLUE_BRIGHT || 
                style == PURPLE_DARK || style == PURPLE_BRIGHT ||
                style == ORANGE_DARK || style == ORANGE_BRIGHT ||
                style == YELLOW_DARK || style == YELLOW_BRIGHT;
    }
    
    @Override
    public short getColorModeOfThemeStyle(short style){
        return (short)(style - FIRST_COLORTHEME_STYLE_VALUE + Diagram.FIRST_COLORTHEME_MODE_VALUE);
    }
    
    @Override
    public String getDiagramTypeName(){
        return "PyramidDiagram";
    }
    
    @Override
    public void initProperties(XShape xControlShape, ArrayList<RelationDiagramItem> items){
        super.initProperties(xControlShape, items);
        XPropertySet xProps = (XPropertySet)UnoRuntime.queryInterface( XPropertySet.class, items.get(0).getMainShape());
        try {
            //if(getStyleProp() == VennDiagram.DEFAULT){ }
            if(getStyleProp() == PyramidDiagram.WITHOUT_OUTLINE)
                setOutlineProp(false);
            if(getStyleProp() == PyramidDiagram.WITH_SHADOW)
                setShadowProp(true);
            //if(getStyleProp() == PyramidDiagram.BC_WITH_GRADIENTS){ }
            if(getStyleProp() == PyramidDiagram.USER_DEFINE){
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

    @Override
    public void createDiagram(){
        super.createDiagram();
        m_IsTextToFitToSize = false;
        m_fFontSize = (float)22.0;
        createDiagram(3);
    }

    @Override
    public void createControlShape(){
        XShape xControlTriangle = createShape("PolyPolygonShape", 0);
        m_xShapes.add(xControlTriangle);
        Point a = new Point(m_PageProps.BorderLeft + m_iHalfDiff, m_PageProps.BorderTop + m_DrawAreaHeight);
        Point b = new Point(m_PageProps.BorderLeft + m_DrawAreaWidth / 2 + m_iHalfDiff, m_PageProps.BorderTop);
        Point c = new Point(m_PageProps.BorderLeft + m_DrawAreaWidth + m_iHalfDiff, m_PageProps.BorderTop + m_DrawAreaHeight);
        setControlTriangleShape(xControlTriangle, a, b, c);
        setControlShapeProps(getControlShape());
    }
    
    @Override
    public void setDrawArea(){
        try {
            m_PageProps.BorderTop += (SHADOW_DIST2 / 2 + 100);
            m_DrawAreaWidth -= ( 2 * SHADOW_DIST2 + 100);
            m_DrawAreaHeight -= (SHADOW_DIST2 / 2 + 100);

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
            m_iHalfDiff += SHADOW_DIST2;
            m_xGroupShape.setPosition(new Point(m_PageProps.BorderLeft + m_iHalfDiff, m_PageProps.BorderTop));
        } catch (PropertyVetoException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
    }

    public void setControlTriangleShape(XShape xTriangleShape, Point a, Point b, Point c){
        try {
                Point[] points1 = new Point[3];
                points1[0] = new Point(a.X, a.Y);
                points1[1] = new Point(c.X, c.Y);
                points1[2] = new Point(b.X, b.Y);

                Point[] points2 = new Point[2];
                points2[0] = new Point(a.X, a.Y);
                points2[1] = new Point(b.X, b.Y);

                Point[][] allPoints = new Point[2][];
                allPoints[0] = points1;
                allPoints[1] = points2;

                XPropertySet xShapeProps = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xTriangleShape);
                xShapeProps.setPropertyValue("PolyPolygon", allPoints);
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
            createItem(selectedShapeID + 1);
            if(isBaseColorsMode() || isBaseColorsWithGradientMode())
                setColorProp(_aBaseColors[(selectedShapeID + 1) % 8]);
        }
    }

    @Override
    public RelationDiagramItem createItem(int shapeID, String str, Color oColor){
        XShape xTrapezeShape = createShape("PolyPolygonShape", shapeID);
        m_xShapes.add(xTrapezeShape);
        setMoveProtectOfShape(xTrapezeShape);

        PyramidDiagramItem item = new PyramidDiagramItem(this, shapeID, xTrapezeShape);
        addItem(item);

        item.setText(str);
        item.setShapesProps();
        item.setColor(oColor);
        item.setLineColor();
        getController().setSelectedShape(xTrapezeShape);
        return item;
    }

    @Override
    public RelationDiagramItem createItem(int shapeID, String str) {
        XShape xTrapezeShape = createShape("PolyPolygonShape", shapeID);
        m_xShapes.add(xTrapezeShape);
        setMoveProtectOfShape(xTrapezeShape);

        PyramidDiagramItem item = new PyramidDiagramItem(this, shapeID, xTrapezeShape);
        addItem(item);

        if(str.equals(""))
            str = " ";
        if(str.equals("DefaultText"))
            item.setDefaultText();
        else
            item.setText(str);
        item.setShapesProps(true);
        setTextColorOfShape(item.getTextShape());
        getController().setSelectedShape(xTrapezeShape);

        return item;
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
                    m_xControlShape = xCurrShape;
            }
            int topShapeID = getTopShapeID();
            for(int i = 1; i <= topShapeID; i++ ){
                for(int j = 0; j < m_xShapes.getCount(); j++ ){
                    xCurrShape = (XShape) UnoRuntime.queryInterface(XShape.class, m_xShapes.getByIndex(j));
                    currShapeName = getShapeName(xCurrShape);
                    if (currShapeName.contains("PolyPolygonShape")) {
                        currShapeID = getController().getShapeID(currShapeName);
                        if(currShapeID == i)
                            addItem(new PyramidDiagramItem(this, currShapeID, xCurrShape));
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
    public void refreshDiagram() {
        removeSingleItems();

        XShape xControlTriangleShape = getControlShape();
        int numOfItems = getNumOfItems();
        Size controlShapeSize = xControlTriangleShape.getSize();
        Point controlShapePos = xControlTriangleShape.getPosition();

        for(RelationDiagramItem item : items){
            ((PyramidDiagramItem)item).setPosition(numOfItems, controlShapeSize, controlShapePos);
            if(isColorSchemeMode())
                setGradientColor(item.getMainShape());
        }
    }

    public void setGradientColor(XShape xShape){
        int shapeID = getController().getShapeID(getShapeName(xShape));
        if(shapeID > 0){
            setSchemesAndBCGColors(shapeID);
            setGradient(xShape);
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
                    if(currShapeName.contains("PolyPolygonShape0"))
                        m_xControlShape = xCurrShape;
                }
            } catch (IndexOutOfBoundsException ex) {
                System.err.println(ex.getLocalizedMessage());
            } catch (WrappedTargetException ex) {
                System.err.println(ex.getLocalizedMessage());
            }
        }
        return m_xControlShape;
    }

    @Override
    public int getSelectedShapeID() {
        String shapeName = getShapeName(getController().getSelectedShape());
        if(shapeName.contains("PyramidDiagram") && shapeName.contains("PolyPolygonShape")){
            int currShapeID = getController().getShapeID(shapeName);
            if(currShapeID >= 0)
                return currShapeID;
        }
        return -1;
    }

    @Override
    public int getTopShapeID() {
        XShape xCurrShape = null;
        String currShapeName = "";
        int shapeID = -1;
        int iTopShapeID = -1;
        try {

            for(int i = 0; i < m_xShapes.getCount(); i++){
                xCurrShape = (XShape) UnoRuntime.queryInterface(XShape.class, m_xShapes.getByIndex(i));
                currShapeName = getShapeName(xCurrShape);
                if(currShapeName.contains("PolyPolygonShape")){
                    shapeID = getController().getShapeID(currShapeName);
                    if (shapeID > iTopShapeID)
                        iTopShapeID = shapeID;
                }
            }
        }catch(IndexOutOfBoundsException ex) {
            System.err.println(ex.getLocalizedMessage());
        }catch(WrappedTargetException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
        return iTopShapeID;
    }

    @Override
    public void removeShape() {
        int selectedShapeID = getSelectedShapeID();
        if(selectedShapeID > 0){
            XShape xSelectedShape = getController().getSelectedShape();
            RelationDiagramItem selectedItem = getItem(xSelectedShape);

            if(selectedItem != null){
                XShape xNextSelectedShape = null;

                if(getTopShapeID() > 1){
                    if(selectedShapeID == getTopShapeID())
                        xNextSelectedShape = getPreviousShape();
                    else
                        xNextSelectedShape = getNextShape();
                    selectedItem.removeItem();
                    decreaseItemsIDs(selectedShapeID);
                }else{
                    xNextSelectedShape = getControlShape();
                    selectedItem.removeItem();
                }

                if(xNextSelectedShape != null)
                    getController().setSelectedShape(xNextSelectedShape);
                if(isBaseColorsMode() || isBaseColorsWithGradientMode())
                    setColorProp(getNextColor());
            }
        }
    }

    @Override
    public XShape getPreviousShape() {
        String shapeName = getShapeName(getController().getSelectedShape());
        if(shapeName.contains("PolyPolygonShape"))
            return getPreviousItem().getMainShape();
        return null;
    }

    @Override
    public XShape getNextShape() {
        String shapeName = getShapeName(getController().getSelectedShape());
        if(shapeName.contains("PolyPolygonShape"))
            return getNextItem().getMainShape();
        return null;
    }

    @Override
    public void setSelectedShapesProperties(){
        XShapes xShapes = getController().getSelectedShapes();
        XShape xShape = null;
        try {
            for(int i = 0; i < xShapes.getCount(); i++){
                xShape = (XShape) UnoRuntime.queryInterface(XShape.class, xShapes.getByIndex(i));
                if(xShape != null){
                    String shapeName = getShapeName(xShape);
                    if(shapeName.contains("GroupShape")){
                        setAllShapeProperties();
                    }else{
                        if(!shapeName.endsWith("PolyPolygonShape0") && !shapeName.equals("")){
                            RelationDiagramItem item = getItem(xShape);
                            if(item != null)
                                item.setShapesProps();
                        }
                    }
                }
            }
        } catch (IndexOutOfBoundsException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
    }

    public void setPropertiesValues(boolean isSelectAllShape, boolean isModifyColors, short sColorMode, boolean isOutline, int lineWidth, boolean isShadow){
        setSelectedAllShapesProp(isSelectAllShape);
        setModifyColorsProp(isModifyColors);
        setColorModeProp(sColorMode);
        setOutlineProp(isOutline);
        setShapesLineWidthProp(lineWidth);
        setShadowProp(isShadow);
    }

    @Override
    public void showPropertyDialog(){
        getGui().enableControlDialogWindow(false);
        short exec = getGui().executePropertiesDialog();
        if(exec == 1){  
            getGui().setPropertiesOfPyramidDiagram();
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
    
    public void setColorSettingsOfShape(XShape xShape){
        if(isColorSchemeMode() || isBaseColorsWithGradientMode()){
            int shapeID = getController().getShapeID(getShapeName(xShape));
            if(isBaseColorsWithGradientMode()){
                if(shapeID > 0)
                    setSchemesAndBCGColors(shapeID);
                setGradient(xShape, GradientStyle.RECT, (short)0, (short)0, (short)50, (short)50, (short)100, (short)75);
            }
            if(isColorSchemeMode()){
                if(shapeID > 0)
                    setSchemesAndBCGColors(shapeID);
                setGradient(xShape);
            }
            setLineColorProp(getDefaultLineColor());
        }else{
            if(isSimpleColorMode())
                setLineColorProp(getDefaultLineColor());    
            if(isBaseColorsMode())
                setLineColorProp(getDefaultLineColor());
            if(isColorThemeMode())
                setColorThemeColors();

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
            if(type.equals("BaseShape"))
                if(!isTextFitProp())
                    xProp.setPropertyValue("CharHeight", new Float(40.0));
            
            if(!type.equals("BaseShape")) {
                if(isModifyColorsProp())
                    setColorSettingsOfShape(xShape);
                if(m_IsShadow){
                    xProp.setPropertyValue("Shadow", new Boolean(true));
                    xProp.setPropertyValue("ShadowXDistance", new Integer(SHADOW_DIST2));
                    xProp.setPropertyValue("ShadowYDistance", new Integer(-SHADOW_DIST2/2));
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

    public void setColorsOfGradients(int color){
        if (color != -1) {
            setStartColorProp(color);
            setEndColorProp(color);
        }
    }

    public void setSchemesAndBCGColors(int shapeID){
        --shapeID;
        if(getColorModeProp() == Diagram.BASE_COLORS_WITH_GRADIENT_MODE){
            int color = -1;
            color = _aBaseColors[shapeID % 8];
            if (color != -1) {
                setStartColorProp(color);
                setEndColorProp(color);
            }
        }
        if(isColorSchemeMode()){
            int colorCode = getColorModeProp() - Diagram.FIRST_COLORSCHEME_MODE_VALUE;
            int topShapeID = getTopShapeID();
            setStartColorProp(SchemeDefinitions.getGradientColor(colorCode, shapeID, topShapeID));
            setEndColorProp(SchemeDefinitions.getGradientColor(colorCode, shapeID + 1, topShapeID));
        }
    }
}